// File discovery and parsing: turns a user file into FileInfo (for the review screen) and Tables (for processing).

import type { CellValue, FileInfo, FileKind, MergeRange, SheetInfo, SourceFile, Table } from '../types';
import type { FileReadOverride, ReadOptions } from '../workflow';
import { LIMITS } from '../limits';
import { extensionOf, stripExtension } from '../filenames';
import { cellToString, isBlank, maybePlainNumber, uid, uniqueHeaders } from '../util';
import { decodeText, detectDelimiter, detectEncoding, looksBinary, parseDelimited } from './csv';
import { readExcel, ReadError, type RawSheet } from './excel';

export const SUPPORTED_EXTENSIONS: FileKind[] = ['csv', 'tsv', 'txt', 'xlsx', 'xlsm', 'xls', 'xlsb'];

export function kindFromName(name: string): FileKind | null {
  const ext = extensionOf(name) as FileKind;
  return SUPPORTED_EXTENSIONS.includes(ext) ? ext : null;
}

/** Files that should be skipped quietly during folder scans (Office lock files, system files). */
export function isIgnorableName(name: string): boolean {
  return name.startsWith('~$') || name.startsWith('.') || /^(thumbs\.db|desktop\.ini)$/i.test(name);
}

const isTextKind = (k: FileKind | null) => k === 'csv' || k === 'tsv' || k === 'txt';

function sniff(bytes: Uint8Array): 'zip' | 'ole' | 'text' | 'other' {
  if (bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04) return 'zip';
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) return 'ole';
  return looksBinary(bytes) ? 'other' : 'text';
}

/** Encrypted .xlsx files are OLE containers holding an "EncryptedPackage" stream. */
function hasEncryptedPackage(bytes: Uint8Array): boolean {
  const needle = 'EncryptedPackage';
  const n = Math.min(bytes.length, 1024 * 1024);
  outer: for (let i = 0; i < n - needle.length * 2; i++) {
    for (let j = 0; j < needle.length; j++) {
      if (bytes[i + j * 2] !== needle.charCodeAt(j) || bytes[i + j * 2 + 1] !== 0) continue outer;
    }
    return true;
  }
  return false;
}

export interface RawWorkbook {
  sheets: RawSheet[];
  encoding?: string;
  delimiter?: string;
  warnings: string[];
  /** The original bytes, kept for .xlsx files so formatting can be preserved when patching. */
  bytes?: ArrayBuffer;
}

/** Reads a file into raw grids. Throws ReadError with a plain-language reason and suggestion. */
export async function readRaw(src: SourceFile, override: FileReadOverride = {}): Promise<RawWorkbook> {
  if (src.size === 0) throw new ReadError('The file is empty.', 'Check that the right file was selected.', 'empty');
  if (src.size > LIMITS.maxFileBytes) {
    throw new ReadError(
      `The file is ${(src.size / 1048576).toFixed(0)} MB; the limit is ${LIMITS.maxFileBytes / 1048576} MB.`,
      'Split the file into smaller parts (for example by year or region) and process them as a batch.',
      'too-large',
    );
  }
  const buf = await src.file.arrayBuffer();
  const bytes = new Uint8Array(buf);
  const kind = src.kind;
  const sig = sniff(bytes);
  const warnings: string[] = [];

  if (isTextKind(kind) && sig !== 'text') {
    if (sig === 'zip' || sig === 'ole') {
      warnings.push('This file has a text extension but contains an Excel workbook; it was read as Excel.');
    } else {
      throw new ReadError('The file does not look like text data (it contains binary content).', 'Check the file type; re-export it from the source system as CSV.', 'unsupported');
    }
  }

  if (isTextKind(kind) && sig === 'text') {
    const guess = detectEncoding(bytes);
    const encoding = override.encoding || guess.encoding;
    const text = decodeText(bytes, encoding, override.encoding ? (override.encoding === guess.encoding ? guess.bomLength : 0) : guess.bomLength);
    if (!guess.confident && !override.encoding) warnings.push(`Text encoding was guessed as ${encoding}. If letters look wrong, choose another encoding.`);
    if (text.includes('�') && !override.encoding) warnings.push('Some characters could not be decoded. Try a different text encoding.');
    const delimiter = override.delimiter || detectDelimiter(text, kind ?? '');
    const rows = parseDelimited(text, delimiter);
    const width = rows.reduce((m, r) => Math.max(m, r.length), 0);
    const grid: CellValue[][] = rows.map((r) => {
      const out: CellValue[] = new Array(width).fill(null);
      for (let i = 0; i < r.length; i++) out[i] = r[i] === '' ? null : r[i];
      return out;
    });
    const ragged = rows.filter((r) => r.length !== width && !(r.length === 1 && r[0] === '')).length;
    if (ragged) warnings.push(`${ragged} line(s) have fewer columns than the widest line; missing cells were left blank.`);
    if (width <= 1 && rows.length > 1) warnings.push('Only one column was found. If the data has columns, choose the delimiter manually.');
    return {
      sheets: [{ name: stripExtension(src.name), hidden: false, grid, formulas: null, origin: { r: 0, c: 0 }, merges: [] }],
      encoding, delimiter, warnings,
    };
  }

  if (sig === 'ole' && (kind === 'xlsx' || kind === 'xlsm' || kind === 'xlsb' || hasEncryptedPackage(bytes))) {
    if (hasEncryptedPackage(bytes)) {
      throw new ReadError('The file is password-protected (encrypted).', 'Open it in Excel, remove the password (File › Info › Protect Workbook › Encrypt with Password), save, and add it again.', 'password');
    }
  }
  if ((kind === 'xlsx' || kind === 'xlsm' || kind === 'xlsb') && sig !== 'zip' && sig !== 'ole') {
    throw new ReadError('The file is not a valid Excel workbook (it may be damaged or renamed).', 'Open it in Excel and save it again as .xlsx.', 'corrupt');
  }
  const sheets = readExcel(buf);
  if (kind === 'xlsm') warnings.push('Macros in this workbook were not run and are not kept in the output.');
  if (kind === 'xls' && sig === 'text') warnings.push('This .xls file is really an HTML/XML export; it was read as a table.');
  return { sheets, warnings, bytes: kind === 'xlsx' ? buf : undefined };
}

/**
 * Picks the header row: the first row (within the first rows) that is reasonably full, almost all text,
 * and followed by data. Title and note rows above the table usually fill only one or two cells.
 */
export function detectHeaderRow(grid: CellValue[][]): number {
  const scan = grid.slice(0, LIMITS.headerScanRows);
  const filled = scan.map((r) => r.filter((v) => !isBlank(v)).length);
  const max = Math.max(0, ...filled);
  if (max === 0) return 0;
  const need = max === 1 ? 1 : Math.max(2, Math.ceil(max * 0.5));
  for (let i = 0; i < scan.length; i++) {
    if (filled[i] < need) continue;
    const texts = scan[i].filter((v) => typeof v === 'string' && v.trim() !== '').length;
    if (texts < Math.ceil(filled[i] * 0.8)) continue;
    let next = i + 1;
    while (next < scan.length && filled[next] === 0) next++;
    if (next >= grid.length || next >= scan.length || filled[next] >= Math.ceil(filled[i] * 0.5)) return i;
  }
  return Math.max(0, filled.findIndex((f) => f === max));
}

function trimmedSize(grid: CellValue[][]): { rows: number; cols: number } {
  let rows = grid.length;
  while (rows > 0 && grid[rows - 1].every(isBlank)) rows--;
  let cols = 0;
  for (let r = 0; r < rows; r++) {
    for (let c = grid[r].length - 1; c >= cols; c--) if (!isBlank(grid[r][c])) { cols = c + 1; break; }
  }
  return { rows, cols };
}

export function sheetInfo(s: RawSheet, chosenHeaderRow?: number): SheetInfo {
  const { rows, cols } = trimmedSize(s.grid);
  const hdr = detectHeaderRow(s.grid);
  const useIdx = chosenHeaderRow ? Math.max(0, chosenHeaderRow - 1 - s.origin.r) : hdr;
  const headers = rows ? uniqueHeaders((s.grid[useIdx] ?? []).slice(0, Math.max(cols, 1)).map((v) => cellToString(v))) : [];
  return {
    name: s.name,
    hidden: s.hidden,
    rowCount: rows,
    colCount: cols,
    detectedHeaderRow: s.origin.r + hdr + 1,
    headerRow: s.origin.r + useIdx + 1,
    headers,
    preview: s.grid.slice(0, 15).map((r) => r.slice(0, Math.min(cols || r.length, 30)).map((v) => (v instanceof Date ? cellToString(v) : v))),
    firstRow: s.origin.r + 1,
    hasFormulas: !!s.formulas,
    mergeCount: s.merges.length,
  };
}

export async function inspectFile(src: SourceFile, override: FileReadOverride = {}): Promise<{ info: FileInfo; raw?: RawWorkbook }> {
  const base: FileInfo = { id: src.id, name: src.name, relPath: src.relPath, size: src.size, kind: src.kind, status: 'ready', sheets: [], warnings: [] };
  if (!src.kind) {
    return { info: { ...base, status: 'skipped', reason: 'Unsupported file type.', suggestion: 'Supported: .csv, .txt, .tsv, .xlsx, .xlsm, .xls, .xlsb' } };
  }
  try {
    const raw = await readRaw(src, override);
    const sheets = raw.sheets.map((s) => sheetInfo(s, override.headerRow?.[s.name]));
    const info: FileInfo = { ...base, sheets, encoding: raw.encoding, delimiter: raw.delimiter, warnings: [...raw.warnings] };
    if (sheets.every((s) => s.rowCount === 0)) {
      info.status = 'error';
      info.reason = 'The file has no data.';
      info.suggestion = 'Check that the right file was selected.';
    }
    for (const s of sheets) {
      if (s.rowCount > LIMITS.maxRowsPerTable + 1) {
        info.status = 'error';
        info.reason = `Sheet "${s.name}" has ${s.rowCount.toLocaleString('en-IN')} rows; the limit is ${LIMITS.maxRowsPerTable.toLocaleString('en-IN')}.`;
        info.suggestion = 'Split the file into smaller parts and process them as a batch.';
      }
    }
    return { info, raw };
  } catch (e) {
    if (e instanceof ReadError) return { info: { ...base, status: 'error', reason: e.message, suggestion: e.suggestion } };
    return { info: { ...base, status: 'error', reason: `Unexpected problem: ${e instanceof Error ? e.message : String(e)}`, suggestion: 'Try saving the file again from Excel and re-adding it.' } };
  }
}

function fillMerges(grid: CellValue[][], merges: MergeRange[], origin: { r: number; c: number }): void {
  for (const m of merges) {
    const top = grid[m.s.r - origin.r]?.[m.s.c - origin.c] ?? null;
    for (let r = m.s.r; r <= m.e.r; r++) {
      const row = grid[r - origin.r];
      if (!row) continue;
      for (let c = m.s.c; c <= m.e.c; c++) {
        if (r === m.s.r && c === m.s.c) continue;
        if (c - origin.c < row.length) row[c - origin.c] = top;
      }
    }
  }
}

/** Builds Tables for the selected sheets of one file, using the header row chosen by the user (or detected). */
export function buildTables(src: SourceFile, raw: RawWorkbook, opts: ReadOptions): { tables: Table[]; warnings: string[] } {
  const ov = opts.files[src.id] ?? {};
  const warnings: string[] = [];
  const tables: Table[] = [];
  const isText = raw.encoding !== undefined;
  for (const s of raw.sheets) {
    if (ov.sheets && !ov.sheets.includes(s.name)) continue;
    if (!ov.sheets && s.hidden && !opts.includeHidden) {
      warnings.push(`Hidden sheet "${s.name}" was skipped (turn on "Include hidden sheets" to process it).`);
      continue;
    }
    const grid = s.grid.map((r) => r.slice());
    const { rows: usedRows, cols: usedCols } = trimmedSize(grid);
    const detected = detectHeaderRow(s.grid); // before merged cells are filled, so merged titles are not mistaken for headers
    if (usedRows === 0) {
      warnings.push(`Sheet "${s.name}" is empty and was skipped.`);
      continue;
    }
    if (s.merges.length) {
      if (opts.mergedCells === 'fill') fillMerges(grid, s.merges, s.origin);
      else warnings.push(`Sheet "${s.name}" has ${s.merges.length} merged cell range(s); only the first cell of each keeps its value.`);
    }
    const chosen = ov.headerRow?.[s.name];
    const hIdx = chosen ? Math.max(0, chosen - 1 - s.origin.r) : detected;
    if (hIdx >= usedRows) {
      warnings.push(`Sheet "${s.name}": the chosen header row is below the data; it was skipped.`);
      continue;
    }
    const width = Math.max(usedCols, 1);
    const headers = uniqueHeaders(grid[hIdx].slice(0, width).map((v) => cellToString(v)));
    const preamble = grid.slice(0, hIdx).map((r) => r.slice(0, width));
    const rows: CellValue[][] = [];
    const rowNums: number[] = [];
    const formulas: (string | null)[][] | undefined = s.formulas ? [] : undefined;
    let removedBlank = 0;
    for (let r = hIdx + 1; r < usedRows; r++) {
      const row = grid[r].slice(0, width);
      while (row.length < width) row.push(null);
      if (opts.skipBlankRows && row.every(isBlank)) { removedBlank++; continue; }
      if (isText && opts.convertTextNumbers) {
        for (let c = 0; c < row.length; c++) if (typeof row[c] === 'string') row[c] = maybePlainNumber(row[c] as string);
      }
      rows.push(row);
      rowNums.push(s.origin.r + r + 1);
      if (formulas && s.formulas) formulas.push(s.formulas[r].slice(0, width));
    }
    if (removedBlank) warnings.push(`Sheet "${s.name}": ${removedBlank} blank row(s) were removed.`);
    const hasFormula = formulas?.some((r) => r.some((f) => f)) ?? false;
    tables.push({
      id: uid('t'),
      fileId: src.id,
      fileName: src.name,
      relPath: src.relPath,
      sheetName: s.name,
      headers,
      rows,
      rowNums,
      formulas: hasFormula ? formulas : undefined,
      preamble,
      origin: { ...s.origin },
      headerRow: s.origin.r + hIdx + 1,
      hidden: s.hidden,
      merges: s.merges,
      structural: removedBlank > 0,
    });
  }
  return { tables, warnings };
}
