// Compare: the first file is the baseline. Output is (1) a copy of the first file with changed cells
// highlighted light yellow (with the other file's value as a cell note) and (2) a change log.

import type { CellValue, Table } from '../types';
import type { CompareOptions } from '../workflow';
import { LIMITS } from '../limits';
import { cellToString, isBlank, normHeader, toNumber } from '../util';
import { colIndex, COLORS, highlight, newTable, ProcessError } from './helpers';

export interface CompareStats {
  pairs: number;
  changedCells: number;
  rowsOnlyInFirst: number;
  rowsOnlyInOther: number;
  logEntries: number;
  logTruncated: boolean;
}

export interface CompareResult {
  baseline: Table[];
  log: Table;
  stats: CompareStats;
  warnings: string[];
}

const MAX_NOTES = 5000;

function same(a: CellValue, b: CellValue, o: CompareOptions): boolean {
  if (isBlank(a) && isBlank(b)) return true;
  if (a instanceof Date || b instanceof Date) {
    if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  }
  const na = typeof a === 'number' ? a : typeof a === 'string' ? toNumber(a) : null;
  const nb = typeof b === 'number' ? b : typeof b === 'string' ? toNumber(b) : null;
  if (na !== null && nb !== null && (typeof a === 'number' || typeof b === 'number')) return Math.abs(na - nb) <= (o.numericTolerance || 0) + 1e-9 * Math.max(1, Math.abs(na));
  let sa = cellToString(a), sb = cellToString(b);
  if (o.trimSpaces) { sa = sa.replace(/\s+/g, ' ').trim(); sb = sb.replace(/\s+/g, ' ').trim(); }
  if (o.ignoreCase) { sa = sa.toLowerCase(); sb = sb.toLowerCase(); }
  return sa === sb;
}

/** Pairs sheets of the first file with sheets of another file: by name, or 1-to-1 when each has one sheet. */
function pairSheets(first: Table[], other: Table[]): { pairs: [Table, Table][]; onlyFirst: Table[]; onlyOther: Table[] } {
  if (first.length === 1 && other.length === 1) return { pairs: [[first[0], other[0]]], onlyFirst: [], onlyOther: [] };
  const pairs: [Table, Table][] = [];
  const usedOther = new Set<Table>();
  const onlyFirst: Table[] = [];
  for (const a of first) {
    const b = other.find((t) => !usedOther.has(t) && t.sheetName.toLowerCase() === a.sheetName.toLowerCase());
    if (b) { pairs.push([a, b]); usedOther.add(b); } else onlyFirst.push(a);
  }
  return { pairs, onlyFirst, onlyOther: other.filter((t) => !usedOther.has(t)) };
}

export function compareTables(tables: Table[], o: CompareOptions): CompareResult {
  const byFile: Table[][] = [];
  for (const t of tables) {
    const g = byFile.find((x) => x[0].fileId === t.fileId);
    if (g) g.push(t); else byFile.push([t]);
  }
  if (byFile.length < 2) throw new ProcessError('Compare needs at least two files.', 'Add the second file to compare against.');
  const warnings: string[] = [];
  const stats: CompareStats = { pairs: 0, changedCells: 0, rowsOnlyInFirst: 0, rowsOnlyInOther: 0, logEntries: 0, logTruncated: false };
  const keyNames = o.keys;
  const logHeaders = ['Change', 'Compared with file', 'Sheet', 'Row in first file', 'Row in other file', 'Column', ...keyNames, 'Old value (first file)', 'New value (other file)'];
  const log: CellValue[][] = [];
  const push = (row: CellValue[]) => {
    stats.logEntries++;
    if (log.length < LIMITS.maxChangeLogEntries) log.push(row); else stats.logTruncated = true;
  };
  const first = byFile[0];
  let notes = 0;
  let keyedPairs = 0;

  for (const other of byFile.slice(1)) {
    const { pairs, onlyFirst, onlyOther } = pairSheets(first, other);
    for (const t of onlyFirst) push(['Sheet only in first file', other[0].fileName, t.sheetName, null, null, null, ...keyNames.map(() => null), null, null]);
    for (const t of onlyOther) push(['Sheet only in other file', t.fileName, t.sheetName, null, null, null, ...keyNames.map(() => null), null, null]);
    for (const [a, b] of pairs) {
      stats.pairs++;
      let keyIdxA = keyNames.map((k) => colIndex(a, k));
      let keyIdxB = keyNames.map((k) => colIndex(b, k));
      const useKeys = keyNames.length > 0 && keyIdxA.every((i) => i >= 0) && keyIdxB.every((i) => i >= 0);
      if (keyNames.length && !useKeys) {
        warnings.push(`Sheet "${a.sheetName}": key column(s) ${keyNames.join(', ')} not found in both files, so rows were compared by position.`);
        keyIdxA = []; keyIdxB = [];
      }
      if (keyNames.length && useKeys) keyedPairs++;
      const colsB = new Map(b.headers.map((h, i) => [normHeader(h), i]));
      const common: [number, number][] = [];
      a.headers.forEach((h, i) => { const j = colsB.get(normHeader(h)); if (j !== undefined) common.push([i, j]); });
      const colsA = new Set(a.headers.map(normHeader));
      for (const h of a.headers) if (!colsB.has(normHeader(h))) push(['Column only in first file', b.fileName, a.sheetName, null, null, h, ...keyNames.map(() => null), null, null]);
      for (const h of b.headers) if (!colsA.has(normHeader(h))) push(['Column only in other file', b.fileName, b.sheetName, null, null, h, ...keyNames.map(() => null), null, null]);

      const keyStr = (row: CellValue[], idx: number[]) => idx.map((i) => { let s = cellToString(row[i]).trim(); if (o.ignoreCase) s = s.toLowerCase(); return s; }).join('␟');
      const matchB = new Map<number, number>(); // row in A → row in B
      const unmatchedB = new Set(b.rows.map((_, i) => i));
      if (useKeys) {
        const idxB = new Map<string, number>();
        let dups = 0;
        b.rows.forEach((row, r) => { const k = keyStr(row, keyIdxB); if (idxB.has(k)) dups++; else idxB.set(k, r); });
        if (dups) warnings.push(`${b.fileName} › ${b.sheetName}: ${dups} row(s) repeat a key; only the first of each was compared.`);
        a.rows.forEach((row, r) => { const m = idxB.get(keyStr(row, keyIdxA)); if (m !== undefined && unmatchedB.has(m)) { matchB.set(r, m); unmatchedB.delete(m); } });
      } else {
        a.rows.forEach((_, r) => { if (r < b.rows.length) { matchB.set(r, r); unmatchedB.delete(r); } });
      }
      a.rows.forEach((row, r) => {
        const keysA = keyNames.map((_, k) => (useKeys ? row[keyIdxA[k]] : null));
        const m = matchB.get(r);
        if (m === undefined) {
          stats.rowsOnlyInFirst++;
          push(['Row only in first file', b.fileName, a.sheetName, a.rowNums[r], null, null, ...keysA, null, null]);
          return;
        }
        const rowB = b.rows[m];
        for (const [i, j] of common) {
          if (same(row[i], rowB[j], o)) continue;
          stats.changedCells++;
          highlight(a, r, i, COLORS.changed);
          if (notes < MAX_NOTES) {
            if (!a.notes) a.notes = new Map();
            const prev = a.notes.get(`${r}:${i}`);
            a.notes.set(`${r}:${i}`, `${prev ? `${prev}\n` : ''}${b.fileName}: ${cellToString(rowB[j]) || '(blank)'}`);
            notes++;
          }
          push(['Changed', b.fileName, a.sheetName, a.rowNums[r], b.rowNums[m], a.headers[i], ...keysA, row[i], rowB[j]]);
        }
      });
      for (const m of unmatchedB) {
        stats.rowsOnlyInOther++;
        push(['Row only in other file', b.fileName, b.sheetName, null, b.rowNums[m], null, ...keyNames.map((_, k) => (useKeys ? b.rows[m][keyIdxB[k]] : null)), null, null]);
      }
    }
  }
  if (keyNames.length && !keyedPairs) throw new ProcessError(`Key column(s) ${keyNames.join(', ')} were not found in any pair of sheets.`, 'Check the key column names, or leave keys empty to compare by row position.');
  if (notes >= MAX_NOTES) warnings.push(`Cell notes were added for the first ${MAX_NOTES} changes only; the change log lists every change.`);
  if (stats.logTruncated) warnings.push(`The change log has ${stats.logEntries.toLocaleString('en-IN')} entries; only the first ${LIMITS.maxChangeLogEntries.toLocaleString('en-IN')} were written. Narrow the comparison (filter or select columns) to see all.`);
  return { baseline: first, log: newTable('Change log', logHeaders, log, o.outputName ? `${o.outputName} change log` : 'Change log'), stats, warnings };
}
