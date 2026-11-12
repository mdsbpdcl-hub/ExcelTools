import type { CellValue, Table } from '../types';
import { normHeader, uid } from '../util';

export class ProcessError extends Error {
  constructor(message: string, public suggestion = '') {
    super(message);
  }
}

/** Case- and spacing-insensitive column lookup. Returns -1 when missing. */
export function colIndex(t: { headers: string[] }, name: string): number {
  const exact = t.headers.indexOf(name);
  if (exact >= 0) return exact;
  const n = normHeader(name);
  return t.headers.findIndex((h) => normHeader(h) === n);
}

/** Resolves names to indices; an empty list means "all columns". Missing names are returned separately. */
export function resolveColumns(t: { headers: string[] }, names: string[]): { idx: number[]; missing: string[] } {
  if (!names.length) return { idx: t.headers.map((_, i) => i), missing: [] };
  const idx: number[] = [];
  const missing: string[] = [];
  for (const n of names) {
    const i = colIndex(t, n);
    if (i >= 0) idx.push(i);
    else missing.push(n);
  }
  return { idx, missing };
}

/**
 * Marks a table as re-shaped: formulas can no longer be kept because cell references would point to
 * the wrong cells. Highlights and notes are kept; callers that move rows/columns remap them.
 */
export function markStructural(t: Table): void {
  t.structural = true;
  if (t.formulas) t.formulas = undefined;
  t.merges = [];
}

function remapMarks(t: Table, map: (r: number, c: number) => [number, number] | null): void {
  for (const key of ['highlights', 'notes'] as const) {
    const src = t[key];
    if (!src?.size) continue;
    const out = new Map<string, string>();
    src.forEach((v, k) => {
      const [r, c] = k.split(':').map(Number);
      const to = map(r, c);
      if (to) out.set(`${to[0]}:${to[1]}`, v);
    });
    t[key] = out.size ? out : undefined;
  }
}

/** Highlights/notes for a new table made from some rows of `t` (used by split). */
export function marksForRows(t: Table, rows: number[]): Pick<Table, 'highlights' | 'notes'> {
  const idx = new Map(rows.map((r, i) => [r, i]));
  const pick = (m?: Map<string, string>) => {
    if (!m?.size) return undefined;
    const out = new Map<string, string>();
    m.forEach((v, k) => { const [r, c] = k.split(':').map(Number); if (idx.has(r)) out.set(`${idx.get(r)}:${c}`, v); });
    return out.size ? out : undefined;
  };
  return { highlights: pick(t.highlights), notes: pick(t.notes) };
}

/** Keeps only the given rows (by index, in order), carrying highlights and notes along. */
export function keepRows(t: Table, keep: number[]): void {
  const newIndex = new Map(keep.map((r, i) => [r, i]));
  remapMarks(t, (r, c) => (newIndex.has(r) ? [newIndex.get(r)!, c] : null));
  t.rows = keep.map((r) => t.rows[r]);
  t.rowNums = keep.map((r) => t.rowNums[r]);
  markStructural(t);
}

/** Snapshot of marked cells by column name, so they survive column changes. */
export function marksByName(t: Table): { r: number; name: string; key: 'highlights' | 'notes'; v: string }[] {
  const out: { r: number; name: string; key: 'highlights' | 'notes'; v: string }[] = [];
  for (const key of ['highlights', 'notes'] as const) {
    t[key]?.forEach((v, k) => {
      const [r, c] = k.split(':').map(Number);
      out.push({ r, name: t.headers[c], key, v });
    });
  }
  return out;
}

export function restoreMarks(t: Table, marks: ReturnType<typeof marksByName>, renamed: Map<string, string>): void {
  t.highlights = undefined;
  t.notes = undefined;
  for (const m of marks) {
    const c = t.headers.indexOf(renamed.get(m.name) ?? m.name);
    if (c < 0) continue;
    if (!t[m.key]) t[m.key] = new Map();
    t[m.key]!.set(`${m.r}:${c}`, m.v);
  }
}

export function isFormulaCell(t: Table, r: number, c: number): boolean {
  return !!t.formulas?.[r]?.[c];
}

/** Changing a formula cell's value replaces the formula with a value. */
export function setCell(t: Table, r: number, c: number, v: CellValue): void {
  t.rows[r][c] = v;
  if (t.formulas?.[r]?.[c]) t.formulas[r][c] = null;
}

export function highlight(t: Table, r: number, c: number, argb: string): void {
  if (!t.highlights) t.highlights = new Map();
  t.highlights.set(`${r}:${c}`, argb);
}

export function deriveTable(base: Table, patch: Partial<Table>): Table {
  return {
    ...base,
    id: uid('t'),
    formulas: undefined,
    highlights: undefined,
    notes: undefined,
    merges: [],
    preamble: [],
    origin: { r: 0, c: 0 },
    headerRow: 1,
    structural: true,
    ...patch,
  };
}

/** A new standalone table (reports, summaries). */
export function newTable(name: string, headers: string[], rows: CellValue[][], fileName = name): Table {
  return {
    id: uid('t'),
    fileId: '',
    fileName,
    relPath: '',
    sheetName: name,
    headers,
    rows,
    rowNums: rows.map((_, i) => i + 2),
    preamble: [],
    origin: { r: 0, c: 0 },
    headerRow: 1,
    hidden: false,
    merges: [],
    structural: true,
  };
}

export const COLORS = {
  changed: 'FFFFF2A8', // light yellow
  issue: 'FFFFD7D7', // light red
  warning: 'FFFFE7C2', // light amber
  header: 'FFE6F2EA', // light green
  headerText: 'FF173A25',
} as const;
