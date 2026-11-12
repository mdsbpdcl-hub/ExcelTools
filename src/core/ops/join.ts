import type { CellValue, Table } from '../types';
import type { JoinOptions } from '../workflow';
import { cellToString } from '../util';
import { colIndex, newTable, ProcessError } from './helpers';

export interface JoinResult {
  table: Table;
  unmatched: Table;
  duplicateKeys: Table;
  stats: { matched: number; unmatchedLeft: number; unmatchedRight: number; duplicateKeys: number };
}

const label = (t: Table) => (t.sheetName && t.sheetName !== t.fileName.replace(/\.[^.]+$/, '') ? `${t.fileName} › ${t.sheetName}` : t.fileName);

/**
 * Joins the first table with each following table on the chosen key columns (same names in every table).
 * Rows without a partner and keys that appear more than once are reported separately.
 */
export function joinTables(tables: Table[], o: JoinOptions): JoinResult {
  if (tables.length < 2) throw new ProcessError('Joining needs at least two files or sheets.', 'Add another file, or turn off Join.');
  if (!o.keys.length) throw new ProcessError('Choose at least one key column for the join.', 'Pick the column(s) that identify a row, such as Consumer No.');
  for (const t of tables) {
    const missing = o.keys.filter((k) => colIndex(t, k) < 0);
    if (missing.length) throw new ProcessError(`${label(t)} has no column ${missing.map((m) => `"${m}"`).join(', ')} to join on.`, 'Rename the column with Column tools, or choose a different key.');
  }
  const keyOf = (t: Table) => {
    const idx = o.keys.map((k) => colIndex(t, k));
    return (row: CellValue[]) =>
      idx.map((i) => {
        let s = cellToString(row[i]);
        if (o.trimSpaces) s = s.trim();
        if (o.ignoreCase) s = s.toLowerCase();
        return s;
      }).join('␟');
  };
  const unmatchedRows: CellValue[][] = [];
  const dupRows: CellValue[][] = [];
  const stats = { matched: 0, unmatchedLeft: 0, unmatchedRight: 0, duplicateKeys: 0 };

  const reportDuplicates = (t: Table, key: (r: CellValue[]) => string) => {
    const seen = new Map<string, number[]>();
    t.rows.forEach((row, r) => { const k = key(row); if (!seen.has(k)) seen.set(k, []); seen.get(k)!.push(r); });
    for (const [, list] of seen) {
      if (list.length < 2) continue;
      stats.duplicateKeys++;
      const keys = o.keys.map((k) => t.rows[list[0]][colIndex(t, k)]);
      dupRows.push([label(t), ...keys, list.length, list.map((r) => t.rowNums[r]).join(', ')]);
    }
  };

  let left = tables[0];
  let leftHeaders = [...left.headers];
  let leftRows = left.rows.map((r) => [...r]);
  let leftNums = [...left.rowNums];
  reportDuplicates(left, keyOf(left));

  for (let ti = 1; ti < tables.length; ti++) {
    const right = tables[ti];
    const rKey = keyOf(right);
    reportDuplicates(right, rKey);
    const lKeyIdx = o.keys.map((k) => colIndex({ headers: leftHeaders }, k));
    const lKey = (row: CellValue[]) => lKeyIdx.map((i) => { let s = cellToString(row[i]); if (o.trimSpaces) s = s.trim(); if (o.ignoreCase) s = s.toLowerCase(); return s; }).join('␟');
    const rIdx = new Map<string, number[]>();
    right.rows.forEach((row, r) => { const k = rKey(row); if (!rIdx.has(k)) rIdx.set(k, []); rIdx.get(k)!.push(r); });
    const rKeyCols = new Set(o.keys.map((k) => colIndex(right, k)));
    const rCols = right.headers.map((_, i) => i).filter((i) => !rKeyCols.has(i));
    const suffix = ` (${right.fileName.replace(/\.[^.]+$/, '')}${tables.filter((x) => x.fileName === right.fileName).length > 1 ? ` ${right.sheetName}` : ''})`;
    const rNames = rCols.map((i) => {
      const n = right.headers[i];
      return leftHeaders.some((h) => h.toLowerCase() === n.toLowerCase()) ? `${n}${suffix}` : n;
    });
    const used = new Set<number>();
    const rows: CellValue[][] = [];
    const nums: number[] = [];
    leftRows.forEach((row, r) => {
      const matches = rIdx.get(lKey(row));
      if (matches) {
        for (const m of matches) {
          used.add(m);
          rows.push([...row, ...rCols.map((i) => right.rows[m][i])]);
          nums.push(leftNums[r]);
          stats.matched++;
        }
      } else {
        stats.unmatchedLeft++;
        unmatchedRows.push([ti === 1 ? label(left) : 'Joined result', leftNums[r], `No match in ${label(right)}`, ...lKeyIdx.map((i) => row[i])]);
        if (o.type !== 'inner') { rows.push([...row, ...rCols.map(() => null)]); nums.push(leftNums[r]); }
      }
    });
    right.rows.forEach((row, m) => {
      if (used.has(m)) return;
      stats.unmatchedRight++;
      unmatchedRows.push([label(right), right.rowNums[m], `No match in ${ti === 1 ? label(left) : 'joined result'}`, ...o.keys.map((k) => row[colIndex(right, k)])]);
      if (o.type === 'full') {
        const nr: CellValue[] = new Array(leftHeaders.length).fill(null);
        o.keys.forEach((k, j) => { nr[lKeyIdx[j]] = row[colIndex(right, k)]; });
        rows.push([...nr, ...rCols.map((i) => row[i])]);
        nums.push(right.rowNums[m]);
      }
    });
    leftHeaders = [...leftHeaders, ...rNames];
    leftRows = rows;
    leftNums = nums;
    left = right;
  }
  const table = newTable(o.outputName || 'Joined', leftHeaders, leftRows, o.outputName || 'Joined');
  table.rowNums = leftRows.map((_, i) => i + 2);
  return {
    table,
    unmatched: newTable('Unmatched rows', ['File', 'Source row', 'Reason', ...o.keys], unmatchedRows, 'Unmatched rows'),
    duplicateKeys: newTable('Duplicate keys', ['File', ...o.keys, 'Times', 'Source rows'], dupRows, 'Duplicate keys'),
    stats,
  };
}
