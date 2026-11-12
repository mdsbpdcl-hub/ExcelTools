// Runs a workflow over a batch: read → per-file steps → duplicates → combine → split/summary →
// validation → formulas → export → ZIP package with success/, review/ and error/ folders.

import JSZip from 'jszip';
import type { FileInfo, Issue, OutputFolder, OutputItem, SourceFile, Table, TableProfile } from './types';
import type { WorkflowConfig } from './workflow';
import { LIMITS } from './limits';
import { NameRegistry, applyNamingRule, sanitizeRelDir, sanitizeSegment, stripExtension } from './filenames';
import { cellToString, formatCount } from './util';
import { buildTables, readRaw } from './parse';
import { ReadError } from './parse/excel';
import { newTable, ProcessError } from './ops/helpers';
import { checkQuality } from './ops/quality';
import { cleanTable } from './ops/clean';
import { findReplaceTable } from './ops/findReplace';
import { applyColumnTools } from './ops/columns';
import { filterTable } from './ops/filter';
import { dedupeTables } from './ops/dedupe';
import { validateTable, listValues } from './ops/validate';
import { mergeTables } from './ops/merge';
import { joinTables } from './ops/join';
import { compareTables } from './ops/compare';
import { splitTable, splitTokens } from './ops/split';
import { pivotTable } from './ops/pivot';
import { profileTable } from './ops/profile';
import { recalculateWorkbook } from './formulas';
import { prepareChart, renderChartPng } from './charts';
import { buildXlsx, patchXlsx, type XlsxExtras } from './export/xlsx';
import { encodeText, toDelimited, toHtml, toJson, toXml } from './export/text';
import { tablesToPdf } from './export/pdf';

export const APP_NAME = 'Sheet Desk';
export const APP_VERSION = '1.0.0';

export interface ProgressEvent {
  phase: 'read' | 'process' | 'combine' | 'export' | 'package' | 'done';
  pct: number; // 0–100 overall
  message: string;
  fileId?: string;
  fileState?: 'reading' | 'processing' | 'exporting' | 'done' | 'review' | 'error';
}

export type Reporter = (e: ProgressEvent) => void;

export interface FileError { fileId: string; name: string; relPath: string; reason: string; suggestion: string }

export interface SheetSample { name: string; headers: string[]; rows: string[][]; totalRows: number }
export interface OutputPreview { folder: OutputFolder; name: string; sheets: SheetSample[]; note?: string }

export interface PreviewResult {
  outputs: OutputPreview[];
  profiles: TableProfile[];
  issueCount: number;
  issues: Issue[];
  warnings: string[];
  errors: FileError[];
  destructive: string[];
  stats: Record<string, number>;
  formulaNote: string;
}

export interface PackagedFile { path: string; folder: OutputFolder | 'root'; size: number; bytes: Uint8Array }

export interface RunResult {
  zip: Uint8Array;
  zipName: string;
  files: PackagedFile[];
  summary: RunSummary;
}

export interface RunSummary {
  app: string;
  version: string;
  startedAt: string;
  finishedAt: string;
  durationSeconds: number;
  processedLocallyInBrowser: true;
  settings: WorkflowConfig;
  totals: { filesSelected: number; succeeded: number; needsReview: number; failed: number; skipped: number; rowsRead: number; rowsWritten: number; outputFiles: number };
  files: {
    name: string; path: string; sizeBytes: number; status: 'success' | 'review' | 'error' | 'skipped';
    reason?: string; suggestion?: string;
    sheets: { name: string; rowsRead: number; columns: number }[];
    warnings: string[];
  }[];
  steps: Record<string, number | string>;
  formulas: { preserved: number; recalculated: number; notRecalculated: number; statement: string };
  outputs: { path: string; folder: string; sizeBytes: number }[];
  warnings: string[];
  errors: FileError[];
}

class Cancelled extends Error {}

interface Ctx {
  cfg: WorkflowConfig;
  report: Reporter;
  warnings: string[];
  errors: FileError[];
  issues: Issue[];
  stats: Record<string, number>;
  destructive: Set<string>;
  fileWarnings: Map<string, string[]>;
  fileSheets: Map<string, { name: string; rowsRead: number; columns: number }[]>;
  failed: Set<string>;
  reviewFiles: Set<string>;
  rawBytes: Map<string, ArrayBuffer>;
  reports: Table[]; // go to review/
  extraSuccess: OutputItem[];
  recalc: { recalculated: number; notRecalculated: number };
}

const bump = (ctx: Ctx, k: string, n = 1) => { ctx.stats[k] = (ctx.stats[k] ?? 0) + n; };
const fileWarn = (ctx: Ctx, id: string, w: string) => { if (!ctx.fileWarnings.has(id)) ctx.fileWarnings.set(id, []); ctx.fileWarnings.get(id)!.push(w); };

function dirOf(relPath: string): string {
  const i = relPath.lastIndexOf('/');
  return i > 0 ? relPath.slice(0, i) : '';
}

/** Steps that work on one table at a time. Throws ProcessError for problems that stop this file. */
async function perTableSteps(t: Table, ctx: Ctx, nonMatching: Table[]): Promise<void> {
  const { cfg } = ctx;
  if (cfg.errors.enabled) {
    const q = checkQuality(t, cfg.errors);
    ctx.issues.push(...q.issues);
    bump(ctx, 'Spreadsheet errors found', q.counts.spreadsheetErrors);
    bump(ctx, 'Invalid dates found', q.counts.invalidDates);
    bump(ctx, 'Unexpected values found', q.counts.unexpected);
    bump(ctx, 'Cells replaced by quality rules', q.replaced);
    if (q.issues.length) ctx.reviewFiles.add(t.fileId);
  }
  if (cfg.clean.enabled) {
    const c = cleanTable(t, cfg.clean);
    bump(ctx, 'Cells cleaned', c.changedCells);
    bump(ctx, 'Dates standardised', c.datesConverted);
    bump(ctx, 'Numbers converted', c.numbersConverted);
    if (c.datesUnreadable) fileWarn(ctx, t.fileId, `${t.sheetName}: ${c.datesUnreadable} value(s) in date columns could not be read as dates and were left unchanged.`);
    if (c.formulaCellsSkipped) fileWarn(ctx, t.fileId, `${t.sheetName}: ${c.formulaCellsSkipped} formula cell(s) were not cleaned, so their formulas are kept.`);
  }
  if (cfg.findReplace.enabled && cfg.findReplace.rules.length) {
    const r = findReplaceTable(t, cfg.findReplace.rules);
    bump(ctx, 'Text replacements', r.replacements.reduce((a, b) => a + b, 0));
    if (r.formulaCellsSkipped) fileWarn(ctx, t.fileId, `${t.sheetName}: ${r.formulaCellsSkipped} formula cell(s) matched but were left unchanged to keep their formulas.`);
  }
  if (cfg.columns.enabled) {
    const before = t.headers.length;
    const r = applyColumnTools(t, cfg.columns);
    r.warnings.forEach((w) => fileWarn(ctx, t.fileId, w));
    if (cfg.columns.remove.length || cfg.columns.keepOnlyListed) {
      if (t.headers.length < before) ctx.destructive.add(`Column tools remove columns (${t.fileName}: ${before} → ${t.headers.length}).`);
    }
  }
  if (cfg.filter.enabled && cfg.filter.conditions.length) {
    const f = filterTable(t, cfg.filter);
    f.warnings.forEach((w) => fileWarn(ctx, t.fileId, w));
    bump(ctx, 'Rows kept by filter', f.matchedCount);
    bump(ctx, 'Rows left out by filter', f.restCount);
    if (f.rest) nonMatching.push(f.rest);
    if (f.restCount && !cfg.filter.saveNonMatching) ctx.destructive.add(`The filter leaves out ${formatCount(f.restCount)} row(s) from ${t.fileName} and they are not saved (turn on "Save non-matching rows" to keep them).`);
  }
}

function itemFor(t: Table, folder: OutputFolder, cfg: WorkflowConfig, base?: string): OutputItem {
  return {
    folder,
    baseName: sanitizeSegment((base ?? stripExtension(t.fileName)) + (cfg.output.nameSuffix || ''), 'output'),
    subDir: cfg.output.keepFolders ? sanitizeRelDir(dirOf(t.relPath)) : '',
    tables: [t],
    sourceFileId: t.fileId || undefined,
    kind: 'data',
  };
}

/** Groups tables back into one output per source file (keeping sheet order). */
function groupByFile(tables: Table[], cfg: WorkflowConfig): OutputItem[] {
  const items: OutputItem[] = [];
  for (const t of tables) {
    const it = items.find((i) => i.sourceFileId && i.sourceFileId === t.fileId);
    if (it) it.tables.push(t);
    else items.push(itemFor(t, 'success', cfg));
  }
  return items;
}

/** The main processing up to (but not including) export. Used by both Preview and Run. */
async function processBatch(
  sources: SourceFile[], infos: FileInfo[], cfg: WorkflowConfig, report: Reporter, isCancelled: () => boolean,
): Promise<{ ctx: Ctx; items: OutputItem[]; profiles: TableProfile[] }> {
  const ctx: Ctx = {
    cfg, report, warnings: [], errors: [], issues: [], stats: {}, destructive: new Set(), fileWarnings: new Map(), fileSheets: new Map(),
    failed: new Set(), reviewFiles: new Set(), rawBytes: new Map(), reports: [], extraSuccess: [], recalc: { recalculated: 0, notRecalculated: 0 },
  };
  const ready = sources.filter((s) => infos.find((i) => i.id === s.id)?.status === 'ready');
  for (const info of infos) {
    if (info.status === 'error') {
      ctx.errors.push({ fileId: info.id, name: info.name, relPath: info.relPath, reason: info.reason ?? 'Could not be read.', suggestion: info.suggestion ?? '' });
      ctx.failed.add(info.id);
      report({ phase: 'read', pct: 0, message: `${info.name}: ${info.reason ?? 'could not be read'}`, fileId: info.id, fileState: 'error' });
    }
  }
  const total = ready.reduce((s, f) => s + f.size, 0);
  if (total > LIMITS.maxBatchBytes) {
    throw new ProcessError(`The selected files total ${(total / 1048576).toFixed(0)} MB; one batch can be at most ${LIMITS.maxBatchBytes / 1048576} MB.`, 'Process the files in smaller batches.');
  }
  let tables: Table[] = [];
  const nonMatching: Table[] = [];
  const withFormulas: Table[] = [];
  for (let i = 0; i < ready.length; i++) {
    if (isCancelled()) throw new Cancelled();
    const src = ready[i];
    report({ phase: 'read', pct: Math.round((i / Math.max(ready.length, 1)) * 40), message: `Reading ${src.name}`, fileId: src.id, fileState: 'reading' });
    try {
      const raw = await readRaw(src, cfg.read.files[src.id]);
      raw.warnings.forEach((w) => fileWarn(ctx, src.id, w));
      if (raw.bytes) ctx.rawBytes.set(src.id, raw.bytes);
      const built = buildTables(src, raw, cfg.read);
      built.warnings.forEach((w) => fileWarn(ctx, src.id, w));
      if (!built.tables.length) throw new ProcessError('No sheets with data were selected in this file.', 'Choose at least one sheet with data, or include hidden sheets.');
      ctx.fileSheets.set(src.id, built.tables.map((t) => ({ name: t.sheetName, rowsRead: t.rows.length, columns: t.headers.length })));
      bump(ctx, 'Rows read', built.tables.reduce((s, t) => s + t.rows.length, 0));
      report({ phase: 'process', pct: Math.round(((i + 0.5) / Math.max(ready.length, 1)) * 40), message: `Processing ${src.name}`, fileId: src.id, fileState: 'processing' });
      for (const t of built.tables) if (t.formulas) withFormulas.push(t);
      const localNon: Table[] = [];
      for (const t of built.tables) await perTableSteps(t, ctx, localNon);
      nonMatching.push(...localNon);
      tables.push(...built.tables);
      report({ phase: 'process', pct: Math.round(((i + 1) / Math.max(ready.length, 1)) * 40), message: `Processed ${src.name}`, fileId: src.id, fileState: 'processing' });
    } catch (e) {
      if (e instanceof Cancelled) throw e;
      const reason = e instanceof Error ? e.message : String(e);
      const suggestion = e instanceof ProcessError || e instanceof ReadError ? e.suggestion : 'Check the file and the workflow settings, then try again.';
      ctx.errors.push({ fileId: src.id, name: src.name, relPath: src.relPath, reason, suggestion });
      ctx.failed.add(src.id);
      report({ phase: 'process', pct: Math.round(((i + 1) / Math.max(ready.length, 1)) * 40), message: `${src.name}: ${reason}`, fileId: src.id, fileState: 'error' });
    }
  }

  if (cfg.dedupe.enabled) {
    const d = dedupeTables(tables, cfg.dedupe);
    d.warnings.forEach((w) => ctx.warnings.push(w));
    bump(ctx, 'Duplicate rows removed', d.removedCount);
    if (d.removedReport) ctx.reports.push(d.removedReport);
    if (d.removedCount && !cfg.dedupe.exportRemoved) ctx.destructive.add(`${formatCount(d.removedCount)} duplicate row(s) are removed and not saved (turn on "Save removed duplicates" to keep them).`);
    else if (d.removedCount) ctx.destructive.add(`${formatCount(d.removedCount)} duplicate row(s) are removed (a copy goes to review/).`);
  }

  // ---- combine ----
  if (isCancelled()) throw new Cancelled();
  report({ phase: 'combine', pct: 45, message: 'Combining' });
  let items: OutputItem[] = [];
  const mode = cfg.combine.mode;
  const combinedFrom = [...new Set(tables.map((t) => t.fileId))];
  const base = groupByFile(tables, cfg);
  try {
    if (mode === 'merge' && tables.length) {
      const m = mergeTables(tables, cfg.combine.merge);
      bump(ctx, 'Rows in merged output', m.table.rows.length);
      if (m.partialColumns.length) ctx.warnings.push(`Merged output: ${m.partialColumns.length} column(s) exist only in some files (${m.partialColumns.slice(0, 5).join(', ')}${m.partialColumns.length > 5 ? '…' : ''}); other files have blanks there.`);
      items.push({ folder: 'success', baseName: sanitizeSegment(cfg.combine.merge.outputName || 'Merged', 'Merged'), subDir: '', tables: [m.table], kind: 'data' });
      if (cfg.combine.keepIndividual) items.push(...base);
      tables = [m.table];
    } else if (mode === 'join' && tables.length) {
      const j = joinTables(tables, cfg.combine.join);
      bump(ctx, 'Rows matched by join', j.stats.matched);
      bump(ctx, 'Unmatched rows', j.stats.unmatchedLeft + j.stats.unmatchedRight);
      bump(ctx, 'Duplicate keys', j.stats.duplicateKeys);
      items.push({ folder: 'success', baseName: sanitizeSegment(cfg.combine.join.outputName || 'Joined', 'Joined'), subDir: '', tables: [j.table], kind: 'data' });
      if (j.unmatched.rows.length) ctx.reports.push(j.unmatched);
      if (j.duplicateKeys.rows.length) ctx.reports.push(j.duplicateKeys);
      if (cfg.combine.join.type === 'inner' && j.stats.unmatchedLeft + j.stats.unmatchedRight) ctx.destructive.add(`Inner join leaves out ${formatCount(j.stats.unmatchedLeft + j.stats.unmatchedRight)} unmatched row(s) (listed in review/).`);
      if (cfg.combine.keepIndividual) items.push(...base);
      tables = [j.table];
    } else if (mode === 'compare' && tables.length) {
      const c = compareTables(tables, cfg.combine.compare);
      c.warnings.forEach((w) => ctx.warnings.push(w));
      bump(ctx, 'Changed cells', c.stats.changedCells);
      bump(ctx, 'Rows only in first file', c.stats.rowsOnlyInFirst);
      bump(ctx, 'Rows only in other file(s)', c.stats.rowsOnlyInOther);
      const first = c.baseline[0];
      const name = cfg.combine.compare.outputName || 'Comparison';
      items.push({ ...itemFor(first, 'success', cfg, `${stripExtension(first.fileName)} - ${name} (highlighted)`), tables: c.baseline, note: 'Copy of the first file; changed cells are light yellow and carry a note with the other value.' });
      items.push({ folder: 'success', baseName: sanitizeSegment(`${name} - change log`, 'change log'), subDir: '', tables: [c.log], kind: 'report' });
      tables = c.baseline;
    } else {
      items = base;
    }
  } catch (e) {
    if (e instanceof Cancelled) throw e;
    const reason = e instanceof Error ? e.message : String(e);
    ctx.errors.push({ fileId: '', name: `${mode} step`, relPath: '', reason, suggestion: e instanceof ProcessError ? e.suggestion : 'Check the combine settings.' });
    ctx.warnings.push(`The ${mode} step failed (${reason}); each file was saved separately instead.`);
    items = base;
  }

  // ---- split / summary ----
  report({ phase: 'combine', pct: 50, message: 'Splitting and summarising' });
  const dataItems = items.filter((i) => i.kind === 'data' && !(mode === 'compare'));
  if (cfg.pivot.enabled) {
    for (const it of dataItems) {
      for (const t of it.tables) {
        try {
          const p = pivotTable(t, cfg.pivot);
          if (p.nonNumeric) fileWarn(ctx, t.fileId || 'combined', `${t.sheetName}: ${p.nonNumeric} non-numeric value(s) were ignored in sums/averages.`);
          const sheetPart = it.tables.length > 1 ? ` - ${t.sheetName}` : '';
          ctx.extraSuccess.push({ folder: 'success', baseName: sanitizeSegment(`${it.baseName}${sheetPart} - ${cfg.pivot.outputName || 'Summary'}`, 'Summary'), subDir: it.subDir, tables: [p.table], kind: 'summary' });
          bump(ctx, 'Summary tables', 1);
        } catch (e) {
          ctx.warnings.push(`Summary skipped for ${t.fileName} › ${t.sheetName}: ${e instanceof Error ? e.message : e}`);
        }
      }
    }
  }
  if (cfg.split.enabled) {
    const next: OutputItem[] = [];
    for (const it of items) {
      if (!dataItems.includes(it)) { next.push(it); continue; }
      const sheetParts: Table[] = [];
      for (const t of it.tables) {
        try {
          const parts = splitTable(t, cfg.split);
          bump(ctx, 'Split parts', parts.length);
          for (const p of parts) {
            if (cfg.split.target === 'files') {
              const name = applyNamingRule(cfg.split.naming, { ...splitTokens(t, p), file: it.baseName });
              next.push({ ...it, baseName: name, tables: [p.table], sourceFileId: undefined });
            } else {
              p.table.sheetName = it.tables.length > 1 ? `${t.sheetName} ${p.label}` : p.label;
              sheetParts.push(p.table);
            }
          }
        } catch (e) {
          ctx.warnings.push(`Split skipped for ${t.fileName} › ${t.sheetName}: ${e instanceof Error ? e.message : e} The sheet was saved without splitting.`);
          if (cfg.split.target === 'sheets') sheetParts.push(t);
          else next.push({ ...it, baseName: it.tables.length > 1 ? sanitizeSegment(`${it.baseName} - ${t.sheetName}`) : it.baseName, tables: [t] });
        }
      }
      if (cfg.split.target === 'sheets' && sheetParts.length) next.push({ ...it, tables: sheetParts, sourceFileId: undefined });
    }
    items = next;
  }
  items.push(...ctx.extraSuccess);

  // ---- validation (on the final data) ----
  report({ phase: 'combine', pct: 55, message: 'Checking rules' });
  if (cfg.validate.enabled && cfg.validate.rules.length) {
    for (const it of items) {
      if (it.kind !== 'data') continue;
      let bad = 0;
      for (const t of it.tables) {
        const v = validateTable(t, cfg.validate);
        ctx.issues.push(...v.issues);
        v.warnings.forEach((w) => ctx.warnings.push(w));
        bad += v.failedRows;
      }
      bump(ctx, 'Rows failing rules', bad);
      if (bad) {
        it.folder = 'review';
        // Remember which source files fed this output so their status says "needs review".
        for (const t of it.tables) if (t.fileId) ctx.reviewFiles.add(t.fileId);
        if (!it.sourceFileId && (mode === 'merge' || mode === 'join')) for (const id of combinedFrom) ctx.reviewFiles.add(id);
      }
    }
  }
  // Outputs built from a file with flagged problems go to review/ (combined outputs included).
  for (const it of items) {
    if (it.kind !== 'data') continue;
    if (it.sourceFileId && ctx.reviewFiles.has(it.sourceFileId)) it.folder = 'review';
    if (!it.sourceFileId && (mode === 'merge' || mode === 'join') && ctx.reviewFiles.size) it.folder = 'review';
  }

  // ---- formulas ----
  if (cfg.formulas.recalculate) {
    report({ phase: 'combine', pct: 58, message: 'Recalculating formulas' });
    for (const it of items) {
      if (!it.tables.some((t) => t.formulas && !t.structural)) continue;
      try {
        const r = await recalculateWorkbook(it.tables);
        ctx.recalc.recalculated += r.recalculated;
        ctx.recalc.notRecalculated += r.unsupported;
        ctx.issues.push(...r.issues);
        if (r.errors) ctx.warnings.push(`${it.baseName}: ${r.errors} formula(s) recalculated to an error value.`);
      } catch (e) {
        ctx.warnings.push(`${it.baseName}: formulas could not be recalculated (${e instanceof Error ? e.message : e}); stored results were kept.`);
      }
    }
  }
  for (const it of items) for (const t of it.tables) if (t.formulas && !t.structural) bump(ctx, 'Formulas kept', t.formulas.reduce((s, r) => s + r.filter(Boolean).length, 0));
  if (withFormulas.some((t) => !t.formulas) || ((mode === 'merge' || mode === 'join') && withFormulas.length)) ctx.warnings.push('Where rows or columns were moved, added or removed, formulas were replaced by their last calculated values (cell references would otherwise point to the wrong cells).');

  for (const t of nonMatching) {
    items.push({ ...itemFor(t, 'success', cfg, `${stripExtension(t.fileName)}${t.sheetName !== stripExtension(t.fileName) ? ` - ${t.sheetName}` : ''} - non-matching rows`), sourceFileId: undefined });
  }
  const profiles = items.filter((i) => i.kind === 'data').flatMap((i) => i.tables.map((t) => profileTable(t, cfg.validate.dateOrder)));
  return { ctx, items, profiles };
}

function formulaStatement(ctx: Ctx): string {
  const kept = ctx.stats['Formulas kept'] ?? 0;
  if (!kept) return 'No formulas were written to the output.';
  if (!ctx.cfg.formulas.recalculate) {
    return `${formatCount(kept)} formula(s) were preserved (copied as formulas). They were NOT recalculated by this app; the values shown are the results last saved by Excel, and Excel will recalculate them when the file is opened.`;
  }
  return `${formatCount(kept)} formula(s) were preserved. ${formatCount(ctx.recalc.recalculated)} were recalculated by this app (HyperFormula engine); ${formatCount(ctx.recalc.notRecalculated)} could not be recalculated and keep Excel's last saved results. Excel will recalculate all formulas when the file is opened.`;
}

const sampleOf = (t: Table): SheetSample => ({
  name: t.sheetName,
  headers: t.headers,
  rows: t.rows.slice(0, LIMITS.previewRows).map((r) => r.map((v) => cellToString(v))),
  totalRows: t.rows.length,
});

export async function previewBatch(sources: SourceFile[], infos: FileInfo[], cfg: WorkflowConfig, report: Reporter, isCancelled: () => boolean = () => false): Promise<PreviewResult> {
  const { ctx, items, profiles } = await processBatch(sources, infos, cfg, report, isCancelled);
  const outputs: OutputPreview[] = items.map((i) => ({ folder: i.folder, name: i.baseName, sheets: i.tables.map(sampleOf), note: i.note }));
  for (const r of ctx.reports) outputs.push({ folder: 'review', name: r.fileName, sheets: [sampleOf(r)] });
  const allWarnings = [...ctx.warnings, ...[...ctx.fileWarnings.entries()].flatMap(([id, ws]) => ws.map((w) => `${infos.find((i) => i.id === id)?.name ?? ''}: ${w}`))];
  report({ phase: 'done', pct: 100, message: 'Preview ready' });
  return {
    outputs, profiles, issueCount: ctx.issues.length, issues: ctx.issues.slice(0, 200), warnings: allWarnings, errors: ctx.errors,
    destructive: [...ctx.destructive], stats: ctx.stats, formulaNote: formulaStatement(ctx),
  };
}

function issuesTable(issues: Issue[]): Table {
  const kinds: Record<Issue['kind'], string> = { blank: 'Blank', 'invalid-date': 'Invalid date', 'unexpected-value': 'Unexpected value', 'spreadsheet-error': 'Spreadsheet error', validation: 'Rule failed', 'unsupported-formula': 'Formula not recalculated' };
  const rows = issues.slice(0, LIMITS.maxChangeLogEntries).map((i) => [kinds[i.kind], i.file, i.sheet, i.row, i.column, i.value, i.message]);
  return newTable('Issues', ['Issue', 'File', 'Sheet', 'Row', 'Column', 'Value', 'Details'], rows, 'Issues report');
}

function profileReport(profiles: TableProfile[]): Table {
  const rows = profiles.flatMap((p) => p.columnProfiles.map((c) => [p.file, p.sheet, c.name, c.type, p.rows, c.blanks, c.uniqueCapped ? `${c.unique}+` : c.unique, c.min ?? null, c.max ?? null, c.errors, c.sample.join(' | ')]));
  return newTable('Data profile', ['File', 'Sheet', 'Column', 'Type', 'Rows', 'Blanks', 'Unique values', 'Min', 'Max', 'Error values', 'Examples'], rows, 'Data profile');
}

export async function runBatch(sources: SourceFile[], infos: FileInfo[], cfg: WorkflowConfig, report: Reporter, isCancelled: () => boolean = () => false): Promise<RunResult> {
  const started = new Date();
  const { ctx, items } = await processBatch(sources, infos, cfg, report, isCancelled);
  const zip = new JSZip();
  const names = new NameRegistry();
  const files: PackagedFile[] = [];
  const add = (folder: OutputFolder | 'root', subDir: string, name: string, bytes: Uint8Array) => {
    const path = names.claim([folder === 'root' ? '' : folder, subDir, name].filter(Boolean).join('/'));
    zip.file(path, bytes);
    files.push({ path, folder, size: bytes.length, bytes });
  };
  const f = cfg.output.formats;
  const reportFormats = f.xlsx ? ['xlsx'] : ['csv'];
  const dropdowns = cfg.validate.enabled && cfg.validate.addDropdowns ? cfg.validate.rules.filter((r) => r.type === 'list').map((r) => ({ column: r.column, values: listValues(r.values) })) : [];
  const textDate = cfg.output.textDateFormat || 'DD-MM-YYYY';
  let rowsWritten = 0;

  const writeItem = async (it: OutputItem, formats: string[]) => {
    const stem = it.baseName;
    const multi = it.tables.length > 1;
    if (formats.includes('xlsx')) {
      const fmt = cfg.format.enabled ? cfg.format : { ...cfg.format, conditional: [], charts: [], summarySheet: false, profileSheet: false, protection: { ...cfg.format.protection, enabled: false } };
      const extras: XlsxExtras = { format: fmt, dropdowns: it.kind === 'data' ? dropdowns : [] };
      if (it.kind === 'data' && cfg.format.enabled) {
        if (cfg.format.summarySheet || cfg.format.charts.length) {
          const images: { png: Uint8Array; title: string }[] = [];
          const notes: string[] = [];
          for (const spec of cfg.format.charts) {
            const d = prepareChart(it.tables[0], spec);
            if (typeof d === 'string') { notes.push(d); continue; }
            const png = await renderChartPng(d);
            if (png) images.push({ png, title: d.title });
            else notes.push('Charts could not be drawn in this browser (no OffscreenCanvas support).');
            if (d.note) notes.push(`${d.title}: ${d.note}`);
          }
          if (images.length) notes.unshift('Charts are pictures: the Excel library used cannot create native, editable Excel charts.');
          const t0 = it.tables[0];
          const cards: [string, string | number][] = cfg.format.summarySheet
            ? [['Rows', it.tables.reduce((s, t) => s + t.rows.length, 0)], ['Columns', t0.headers.length], ['Sheets', it.tables.length], ['Issues flagged', ctx.issues.filter((i) => it.tables.some((t) => t.fileName === i.file && t.sheetName === i.sheet)).length], ['Created', new Date().toLocaleString('en-IN')]]
            : [];
          extras.summary = { title: `${stem} – summary`, cards, images, notes };
        }
        if (cfg.format.profileSheet) extras.reportTables = [profileReport(it.tables.map((t) => profileTable(t, cfg.validate.dateOrder)))];
      }
      const original = it.sourceFileId ? ctx.rawBytes.get(it.sourceFileId) : undefined;
      const canPatch = !!original && cfg.output.keepOriginalFormatting && it.tables.every((t) => !t.structural && t.fileId === it.sourceFileId);
      let res;
      if (canPatch) {
        try { res = await patchXlsx(original!, it.tables, extras, cfg.formulas.recalculate); }
        catch (e) { ctx.warnings.push(`${stem}: original formatting could not be kept (${e instanceof Error ? e.message : e}); a clean workbook was written instead.`); }
      }
      if (!res) res = await buildXlsx(it.tables, extras);
      res.warnings.forEach((w) => fileWarn(ctx, it.sourceFileId ?? stem, w));
      add(it.folder, it.subDir, `${stem}.xlsx`, res.bytes);
    }
    for (const kind of ['csv', 'tsv'] as const) {
      if (!formats.includes(kind)) continue;
      if (multi) fileWarn(ctx, it.sourceFileId ?? stem, `${kind.toUpperCase()} holds one sheet per file, so each sheet was saved as its own file.`);
      for (const t of it.tables) {
        const text = toDelimited(t, kind === 'tsv' ? '\t' : cfg.output.csvDelimiter || ',', textDate);
        const enc = encodeText(text, cfg.output.csvEncoding);
        if (enc.lost) fileWarn(ctx, it.sourceFileId ?? stem, `${enc.lost} character(s) cannot be written in ${cfg.output.csvEncoding} and became "?". Use UTF-8 to keep them.`);
        add(it.folder, it.subDir, `${stem}${multi ? ` - ${sanitizeSegment(t.sheetName)}` : ''}.${kind}`, enc.bytes);
      }
    }
    if (formats.includes('json')) add(it.folder, it.subDir, `${stem}.json`, new TextEncoder().encode(toJson(it.tables, cfg.output.jsonStyle)));
    if (formats.includes('xml')) add(it.folder, it.subDir, `${stem}.xml`, new TextEncoder().encode(toXml(it.tables, textDate)));
    if (formats.includes('html')) add(it.folder, it.subDir, `${stem}.html`, new TextEncoder().encode(toHtml(it.tables, stem, textDate, { orientation: cfg.output.pdf.orientation, paper: cfg.output.pdf.paper })));
    if (formats.includes('pdf')) {
      try {
        const pdf = tablesToPdf(it.tables, cfg.output.pdf, stem, textDate);
        pdf.warnings.forEach((w) => fileWarn(ctx, it.sourceFileId ?? stem, w));
        add(it.folder, it.subDir, `${stem}.pdf`, pdf.bytes);
      } catch (e) {
        ctx.warnings.push(`${stem}: PDF not created — ${e instanceof Error ? e.message : e}`);
      }
    }
    if (it.kind === 'data') rowsWritten += it.tables.reduce((s, t) => s + t.rows.length, 0);
  };

  const formats = Object.entries(f).filter(([, on]) => on).map(([k]) => k);
  if (!formats.length) throw new ProcessError('Choose at least one output format.', 'Tick Excel, CSV or another format in Output.');
  for (let i = 0; i < items.length; i++) {
    if (isCancelled()) throw new Cancelled();
    const it = items[i];
    report({ phase: 'export', pct: 60 + Math.round((i / Math.max(items.length, 1)) * 30), message: `Writing ${it.baseName}`, fileId: it.sourceFileId, fileState: 'exporting' });
    try {
      await writeItem(it, it.kind === 'report' ? (formats.includes('xlsx') ? ['xlsx'] : formats.includes('csv') ? ['csv'] : ['xlsx']) : formats);
      if (it.sourceFileId) report({ phase: 'export', pct: 60 + Math.round(((i + 1) / Math.max(items.length, 1)) * 30), message: `Saved ${it.baseName}`, fileId: it.sourceFileId, fileState: it.folder === 'review' ? 'review' : 'done' });
    } catch (e) {
      const reason = e instanceof Error ? e.message : String(e);
      ctx.errors.push({ fileId: it.sourceFileId ?? '', name: it.baseName, relPath: '', reason: `Could not write output: ${reason}`, suggestion: 'Try fewer output formats, or split the data into smaller parts.' });
      if (it.sourceFileId) { ctx.failed.add(it.sourceFileId); report({ phase: 'export', pct: 90, message: reason, fileId: it.sourceFileId, fileState: 'error' }); }
    }
  }

  report({ phase: 'package', pct: 92, message: 'Writing reports' });
  for (const r of ctx.reports) await writeItem({ folder: 'review', baseName: sanitizeSegment(r.fileName), subDir: '', tables: [r], kind: 'report' }, reportFormats);
  if (ctx.issues.length) {
    if (ctx.issues.length > LIMITS.maxChangeLogEntries) ctx.warnings.push(`There are ${formatCount(ctx.issues.length)} issues; the issues report lists the first ${formatCount(LIMITS.maxChangeLogEntries)}.`);
    await writeItem({ folder: 'review', baseName: 'Issues report', subDir: '', tables: [issuesTable(ctx.issues)], kind: 'report' }, reportFormats);
  }
  for (const err of ctx.errors) {
    const src = sources.find((s) => s.id === err.fileId);
    const subDir = cfg.output.keepFolders && err.relPath ? sanitizeRelDir(dirOf(err.relPath)) : '';
    const note = `File: ${err.relPath || err.name}\nProblem: ${err.reason}\nWhat to do: ${err.suggestion || '—'}\n`;
    add('error', subDir, `${sanitizeSegment(stripExtension(err.name))} - error.txt`, new TextEncoder().encode(note));
    if (src && cfg.output.includeFailedOriginals) add('error', subDir, sanitizeSegment(src.name), new Uint8Array(await src.file.arrayBuffer()));
  }

  const finished = new Date();
  const statusOf = (id: string): 'success' | 'review' | 'error' => (ctx.failed.has(id) ? 'error' : items.some((i) => i.sourceFileId === id && i.folder === 'review') || ctx.reviewFiles.has(id) ? 'review' : 'success');
  const settings: WorkflowConfig = JSON.parse(JSON.stringify(cfg));
  if (settings.format.protection.password) settings.format.protection.password = '(hidden)';
  const summary: RunSummary = {
    app: APP_NAME, version: APP_VERSION,
    startedAt: started.toISOString(), finishedAt: finished.toISOString(), durationSeconds: Math.round((finished.getTime() - started.getTime()) / 100) / 10,
    processedLocallyInBrowser: true,
    settings,
    totals: {
      filesSelected: infos.length,
      succeeded: infos.filter((i) => i.status === 'ready' && statusOf(i.id) === 'success').length,
      needsReview: infos.filter((i) => i.status === 'ready' && statusOf(i.id) === 'review').length,
      failed: infos.filter((i) => i.status === 'error' || ctx.failed.has(i.id)).length,
      skipped: infos.filter((i) => i.status === 'skipped').length,
      rowsRead: ctx.stats['Rows read'] ?? 0,
      rowsWritten,
      outputFiles: 0,
    },
    files: infos.map((i) => ({
      name: i.name, path: i.relPath, sizeBytes: i.size,
      status: i.status === 'skipped' ? 'skipped' : i.status === 'error' ? 'error' : statusOf(i.id),
      reason: i.reason ?? ctx.errors.find((e) => e.fileId === i.id)?.reason,
      suggestion: i.suggestion ?? ctx.errors.find((e) => e.fileId === i.id)?.suggestion,
      sheets: ctx.fileSheets.get(i.id) ?? [],
      warnings: [...i.warnings, ...(ctx.fileWarnings.get(i.id) ?? [])].filter((w, k, a) => a.indexOf(w) === k),
    })),
    steps: ctx.stats,
    formulas: { preserved: ctx.stats['Formulas kept'] ?? 0, recalculated: ctx.recalc.recalculated, notRecalculated: ctx.recalc.notRecalculated, statement: formulaStatement(ctx) },
    outputs: [],
    warnings: [...ctx.warnings, ...[...ctx.fileWarnings.entries()].filter(([id]) => !infos.some((i) => i.id === id)).flatMap(([, w]) => w)].filter((w, k, a) => a.indexOf(w) === k),
    errors: ctx.errors,
  };
  summary.outputs = files.map((x) => ({ path: x.path, folder: x.folder, sizeBytes: x.size }));
  summary.totals.outputFiles = files.length;
  add('root', '', 'run-summary.json', new TextEncoder().encode(JSON.stringify(summary, null, 2)));
  add('root', '', 'run-report.txt', new TextEncoder().encode(textReport(summary)));
  report({ phase: 'package', pct: 96, message: 'Building ZIP' });
  const zipBytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  const stamp = `${started.getFullYear()}${String(started.getMonth() + 1).padStart(2, '0')}${String(started.getDate()).padStart(2, '0')}-${String(started.getHours()).padStart(2, '0')}${String(started.getMinutes()).padStart(2, '0')}`;
  report({ phase: 'done', pct: 100, message: 'Done' });
  return { zip: zipBytes, zipName: `${sanitizeSegment(cfg.output.packageName || 'processed')}-${stamp}.zip`, files, summary };
}

export function textReport(s: RunSummary): string {
  const L: string[] = [];
  L.push(`${s.app} ${s.version} — run report`, `Started: ${s.startedAt}   Finished: ${s.finishedAt}   (${s.durationSeconds}s)`, 'All processing happened in your browser; no file was uploaded.', '');
  const t = s.totals;
  L.push(`Files selected: ${t.filesSelected}  |  succeeded: ${t.succeeded}  |  need review: ${t.needsReview}  |  failed: ${t.failed}  |  skipped: ${t.skipped}`);
  L.push(`Rows read: ${t.rowsRead}  |  rows written: ${t.rowsWritten}  |  output files: ${t.outputFiles}`, '');
  L.push('FOLDERS', '  success/  finished outputs', '  review/   outputs that broke a rule or have flagged cells, plus reports (issues, removed duplicates, unmatched rows)', '  error/    files that could not be processed, with the reason and what to do', '');
  L.push('FORMULAS', `  ${s.formulas.statement}`, '');
  if (Object.keys(s.steps).length) { L.push('WHAT WAS DONE'); for (const [k, v] of Object.entries(s.steps)) L.push(`  ${k}: ${v}`); L.push(''); }
  if (s.errors.length) { L.push('ERRORS'); for (const e of s.errors) L.push(`  ${e.relPath || e.name}: ${e.reason}${e.suggestion ? `\n    → ${e.suggestion}` : ''}`); L.push(''); }
  const fw = s.files.filter((f) => f.warnings.length);
  if (s.warnings.length || fw.length) {
    L.push('WARNINGS');
    for (const w of s.warnings) L.push(`  • ${w}`);
    for (const f of fw) for (const w of f.warnings) L.push(`  • ${f.path || f.name}: ${w}`);
    L.push('');
  }
  L.push('FILES');
  for (const f of s.files) L.push(`  [${f.status}] ${f.path || f.name}${f.reason ? ` — ${f.reason}` : ''}`);
  return L.join('\n') + '\n';
}

export { Cancelled };
