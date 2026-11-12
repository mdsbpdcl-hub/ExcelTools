import type { CellValue, Table } from '../types';
import type { SplitOptions } from '../workflow';
import { stripExtension } from '../filenames';
import { cellToString, isBlank } from '../util';
import { colIndex, deriveTable, marksForRows, ProcessError } from './helpers';
import { LIMITS } from '../limits';

export interface SplitPart {
  label: string; // value / part / group label used in naming
  n: number;
  table: Table;
}

/** Parses "Name, Phone; Address, City" into column groups. */
export function parseColumnGroups(s: string): string[][] {
  return s.split(/[;\n]/).map((g) => g.split(',').map((c) => c.trim()).filter(Boolean)).filter((g) => g.length);
}

export function splitTable(t: Table, o: SplitOptions): SplitPart[] {
  const parts: SplitPart[] = [];
  if (o.mode === 'rows') {
    const size = Math.max(1, Math.floor(o.rowCount || 0));
    if (!o.rowCount || o.rowCount < 1) throw new ProcessError('Rows per part must be at least 1.');
    if (t.rows.length / size > LIMITS.maxSplitParts) throw new ProcessError(`This would make ${Math.ceil(t.rows.length / size).toLocaleString('en-IN')} parts; the limit is ${LIMITS.maxSplitParts}.`, 'Use a larger number of rows per part.');
    for (let i = 0, n = 1; i < t.rows.length || (i === 0 && n === 1); i += size, n++) {
      const idx = Array.from({ length: Math.min(size, t.rows.length - i) }, (_, k) => i + k);
      parts.push({ label: `part${n}`, n, table: { ...deriveTable(t, { headers: [...t.headers], rows: idx.map((r) => t.rows[r]), rowNums: idx.map((r) => t.rowNums[r]), sheetName: `${t.sheetName} ${n}` }), ...marksForRows(t, idx) } });
      if (!t.rows.length) break;
    }
    return parts;
  }
  if (o.mode === 'values') {
    if (!o.columns.length) throw new ProcessError('Choose the column(s) whose values decide the split.');
    const idx = o.columns.map((c) => colIndex(t, c));
    const missing = o.columns.filter((_, k) => idx[k] < 0);
    if (missing.length) throw new ProcessError(`Column ${missing.join(', ')} not found in ${t.fileName} › ${t.sheetName}.`);
    const groups = new Map<string, number[]>();
    t.rows.forEach((row, r) => {
      const label = idx.map((i) => (isBlank(row[i]) ? '(blank)' : cellToString(row[i]).trim())).join(' - ');
      if (!groups.has(label)) groups.set(label, []);
      groups.get(label)!.push(r);
    });
    if (groups.size > LIMITS.maxSplitParts) throw new ProcessError(`"${o.columns.join(', ')}" has ${groups.size.toLocaleString('en-IN')} different values; splitting is limited to ${LIMITS.maxSplitParts} parts.`, 'Split by a column with fewer different values (for example Circle or Division instead of Consumer No).');
    let n = 1;
    for (const [label, rows] of groups) {
      parts.push({ label, n: n++, table: { ...deriveTable(t, { headers: [...t.headers], rows: rows.map((r) => t.rows[r]), rowNums: rows.map((r) => t.rowNums[r]), sheetName: label }), ...marksForRows(t, rows) } });
    }
    return parts;
  }
  const groups = parseColumnGroups(o.columnGroups);
  if (!groups.length) throw new ProcessError('Enter column groups such as "Name, Phone; Address, City".');
  const repeat = o.repeatColumns.map((c) => colIndex(t, c)).filter((i) => i >= 0);
  groups.forEach((g, k) => {
    const idx = g.map((c) => colIndex(t, c)).filter((i) => i >= 0);
    if (!idx.length) return;
    const cols = [...repeat.filter((i) => !idx.includes(i)), ...idx];
    const label = g.slice(0, 2).join(' ').slice(0, 40) || `group${k + 1}`;
    parts.push({
      label, n: k + 1,
      table: deriveTable(t, { headers: cols.map((i) => t.headers[i]), rows: t.rows.map((row) => cols.map((i) => row[i]) as CellValue[]), rowNums: [...t.rowNums], sheetName: label }),
    });
  });
  if (!parts.length) throw new ProcessError(`None of the column groups match columns in ${t.fileName} › ${t.sheetName}.`);
  return parts;
}

export function splitTokens(t: Table, p: SplitPart): Record<string, string | number> {
  return { file: stripExtension(t.fileName), sheet: t.sheetName, value: p.label, n: p.n, date: new Date().toISOString().slice(0, 10) };
}
