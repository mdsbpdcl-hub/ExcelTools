import type { CellValue, Table } from '../types';
import type { MergeOptions } from '../workflow';
import { normHeader } from '../util';
import { newTable } from './helpers';

export interface HeaderGroup {
  key: string; // normalised header
  spellings: string[];
  suggested: string;
  files: number;
}

/** Groups headers from all tables by their normalised form, for the header-mapping screen. */
export function headerGroups(tables: { headers: string[] }[]): HeaderGroup[] {
  const map = new Map<string, HeaderGroup>();
  for (const t of tables) {
    const seenHere = new Set<string>();
    for (const h of t.headers) {
      const key = normHeader(h) || h;
      let g = map.get(key);
      if (!g) { g = { key, spellings: [], suggested: h, files: 0 }; map.set(key, g); }
      if (!g.spellings.includes(h)) g.spellings.push(h);
      if (!seenHere.has(key)) { g.files++; seenHere.add(key); }
    }
  }
  return [...map.values()];
}

export interface MergeResult {
  table: Table;
  mapping: { file: string; sheet: string; from: string; to: string }[];
  partialColumns: string[];
}

/** Stacks all tables, matching columns by (mapped) header name rather than by position. */
export function mergeTables(tables: Table[], o: MergeOptions): MergeResult {
  const out: string[] = [];
  const mapping: MergeResult['mapping'] = [];
  const target = (h: string) => {
    const key = normHeader(h) || h;
    const mapped = o.headerMap[key]?.trim();
    return mapped || h;
  };
  const canonical = new Map<string, string>(); // normalised target → display name
  const perTable = tables.map((t) =>
    t.headers.map((h) => {
      const tgt = target(h);
      const nk = normHeader(tgt) || tgt;
      if (!canonical.has(nk)) { canonical.set(nk, tgt); out.push(tgt); }
      const name = canonical.get(nk)!;
      if (name !== h) mapping.push({ file: t.fileName, sheet: t.sheetName, from: h, to: name });
      return out.indexOf(name);
    }),
  );
  const extra: string[] = [];
  if (o.addSourceFile) extra.push('Source file');
  if (o.addSourceSheet) extra.push('Source sheet');
  if (o.addSourceRow) extra.push('Source row');
  const extraNames = extra.map((e) => (out.includes(e) ? `${e} (added)` : e));
  const rows: CellValue[][] = [];
  const rowNums: number[] = [];
  const presence = new Array(out.length).fill(0);
  const highlights = new Map<string, string>();
  tables.forEach((t, ti) => {
    const idx = perTable[ti];
    idx.forEach((j) => presence[j]++);
    const offset = rows.length;
    t.highlights?.forEach((argb, key) => {
      const [r, c] = key.split(':').map(Number);
      highlights.set(`${offset + r}:${idx[c]}`, argb);
    });
    t.rows.forEach((row, r) => {
      const nr: CellValue[] = new Array(out.length).fill(null);
      idx.forEach((j, c) => { if (nr[j] === null) nr[j] = row[c]; });
      if (o.addSourceFile) nr.push(t.relPath || t.fileName);
      if (o.addSourceSheet) nr.push(t.sheetName);
      if (o.addSourceRow) nr.push(t.rowNums[r]);
      rows.push(nr);
      rowNums.push(rows.length + 1);
    });
  });
  const table = newTable(o.outputName || 'Merged', [...out, ...extraNames], rows, o.outputName || 'Merged');
  table.rowNums = rowNums;
  if (highlights.size) table.highlights = highlights;
  const partialColumns = out.filter((_, j) => presence[j] < tables.length);
  return { table, mapping, partialColumns };
}
