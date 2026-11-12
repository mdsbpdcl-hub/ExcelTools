import type { CellValue, Table } from '../types';
import type { DedupeOptions } from '../workflow';
import { cellToString, isBlank } from '../util';
import { colIndex, keepRows, newTable } from './helpers';

export interface DedupeResult {
  removedCount: number;
  removedReport: Table | null;
  perTable: Map<string, number>;
  warnings: string[];
}

interface Entry { t: Table; r: number; filled: number }

function keyFn(t: Table, o: DedupeOptions): ((row: CellValue[]) => string) | null {
  const idx = o.keys.length ? o.keys.map((k) => colIndex(t, k)) : t.headers.map((_, i) => i);
  if (idx.some((i) => i < 0)) return null;
  return (row) =>
    idx
      .map((i) => {
        let s = cellToString(row[i]);
        if (o.trimSpaces) s = s.replace(/\s+/g, ' ').trim();
        if (o.ignoreCase) s = s.toLowerCase();
        return s;
      })
      .join('␟');
}

/**
 * Removes duplicate rows. Scope "file" checks each table on its own; scope "batch" checks across all tables
 * (in the order the files were added). Which copy survives is decided by `keep`.
 */
export function dedupeTables(tables: Table[], o: DedupeOptions): DedupeResult {
  const warnings: string[] = [];
  const perTable = new Map<string, number>();
  const removed: { e: Entry; kept: Entry }[] = [];
  const groups: Table[][] = o.scope === 'batch' ? [tables] : tables.map((t) => [t]);

  for (const group of groups) {
    const winners = new Map<string, Entry>();
    const losers: { e: Entry; key: string }[] = [];
    for (const t of group) {
      const key = keyFn(t, o);
      if (!key) {
        warnings.push(`${t.fileName} › ${t.sheetName}: duplicate check skipped — missing key column(s) ${o.keys.filter((k) => colIndex(t, k) < 0).join(', ')}.`);
        continue;
      }
      t.rows.forEach((row, r) => {
        if (row.every(isBlank)) return;
        const k = key(row);
        const e: Entry = { t, r, filled: row.filter((v) => !isBlank(v)).length };
        const cur = winners.get(k);
        if (!cur) { winners.set(k, e); return; }
        const replace = o.keep === 'last' || (o.keep === 'mostComplete' && e.filled > cur.filled);
        if (replace) { winners.set(k, e); losers.push({ e: cur, key: k }); }
        else losers.push({ e, key: k });
      });
    }
    for (const l of losers) removed.push({ e: l.e, kept: winners.get(l.key)! });
  }

  if (!removed.length) return { removedCount: 0, removedReport: null, perTable, warnings };

  const drop = new Map<Table, Set<number>>();
  for (const { e } of removed) {
    if (!drop.has(e.t)) drop.set(e.t, new Set());
    drop.get(e.t)!.add(e.r);
  }
  let report: Table | null = null;
  if (o.exportRemoved) {
    const dataHeaders: string[] = [];
    for (const { e } of removed) for (const h of e.t.headers) if (!dataHeaders.includes(h)) dataHeaders.push(h);
    const headers = ['Removed from file', 'Sheet', 'Source row', 'Kept copy in file', 'Kept sheet', 'Kept row', ...dataHeaders];
    const rows = removed.map(({ e, kept }) => {
      const map = new Map(e.t.headers.map((h, i) => [h, e.t.rows[e.r][i]]));
      return [e.t.fileName, e.t.sheetName, e.t.rowNums[e.r], kept.t.fileName, kept.t.sheetName, kept.t.rowNums[kept.r], ...dataHeaders.map((h) => map.get(h) ?? null)] as CellValue[];
    });
    report = newTable('Removed duplicates', headers, rows, 'Removed duplicates');
  }
  for (const [t, set] of drop) {
    perTable.set(t.id, set.size);
    keepRows(t, t.rows.map((_, i) => i).filter((i) => !set.has(i)));
  }
  return { removedCount: removed.length, removedReport: report, perTable, warnings };
}
