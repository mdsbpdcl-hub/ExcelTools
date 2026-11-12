import type { CellValue, ColumnProfile, Table, TableProfile } from '../types';
import { LIMITS } from '../limits';
import { cellToString, isBlank, isErrorValue, parseDateDetailed, toNumber, type DateOrder } from '../util';

export type ColumnType = ColumnProfile['type'];

function kindOf(v: CellValue, order: DateOrder): 'number' | 'date' | 'boolean' | 'text' | 'error' | 'blank' {
  if (v === null || isBlank(v)) return 'blank';
  if (typeof v === 'number') return 'number';
  if (typeof v === 'boolean') return 'boolean';
  if (v instanceof Date) return 'date';
  if (isErrorValue(v)) return 'error';
  if (toNumber(v) !== null && /\d/.test(v)) return 'number';
  if (parseDateDetailed(v, order).looksLikeDate) return 'date';
  const u = v.trim().toUpperCase();
  if (u === 'TRUE' || u === 'FALSE') return 'boolean';
  return 'text';
}

/** The type that at least 90% of non-blank, non-error values share; otherwise "mixed". */
export function inferColumnType(t: Table, c: number, order: DateOrder = 'DMY'): ColumnType {
  const counts = { number: 0, date: 0, boolean: 0, text: 0 };
  let n = 0;
  for (const row of t.rows) {
    const k = kindOf(row[c], order);
    if (k === 'blank' || k === 'error') continue;
    counts[k]++;
    n++;
  }
  if (!n) return 'empty';
  for (const k of ['number', 'date', 'boolean', 'text'] as const) if (counts[k] / n >= 0.9) return k;
  return 'mixed';
}

export function profileTable(t: Table, order: DateOrder = 'DMY'): TableProfile {
  const warnings: string[] = [];
  const columnProfiles: ColumnProfile[] = t.headers.map((name, c) => {
    const type = inferColumnType(t, c, order);
    const uniq = new Set<string>();
    let blanks = 0, errors = 0, capped = false, invalidDates = 0, unexpected = 0;
    let min: number | string | undefined, max: number | string | undefined;
    let minN = Infinity, maxN = -Infinity, minD = Infinity, maxD = -Infinity;
    const sample: string[] = [];
    for (const row of t.rows) {
      const v = row[c];
      if (isBlank(v)) { blanks++; continue; }
      if (isErrorValue(v)) { errors++; continue; }
      const s = cellToString(v);
      if (!capped) {
        uniq.add(s);
        if (uniq.size >= LIMITS.profileUniqueCap) capped = true;
      }
      if (sample.length < 3 && !sample.includes(s)) sample.push(s.length > 40 ? `${s.slice(0, 40)}…` : s);
      if (type === 'number') {
        const n = toNumber(v);
        if (n === null) unexpected++;
        else { if (n < minN) minN = n; if (n > maxN) maxN = n; }
      } else if (type === 'date') {
        const d = parseDateDetailed(v, order);
        if (d.date) { const ms = d.date.getTime(); if (ms < minD) minD = ms; if (ms > maxD) maxD = ms; }
        else if (d.looksLikeDate) invalidDates++;
        else unexpected++;
      }
    }
    if (type === 'number' && minN !== Infinity) { min = minN; max = maxN; }
    if (type === 'date' && minD !== Infinity) { min = cellToString(new Date(minD)); max = cellToString(new Date(maxD)); }
    if (t.rows.length && blanks === t.rows.length) warnings.push(`"${name}" is completely blank.`);
    else if (t.rows.length >= 10 && blanks / t.rows.length > 0.5) warnings.push(`"${name}" is more than half blank (${blanks} of ${t.rows.length}).`);
    if (errors) warnings.push(`"${name}" has ${errors} spreadsheet error value(s) such as #N/A.`);
    if (invalidDates) warnings.push(`"${name}" has ${invalidDates} value(s) that look like dates but are not valid dates.`);
    if (unexpected) warnings.push(`"${name}" is mostly ${type === 'number' ? 'numbers' : 'dates'} but has ${unexpected} other value(s).`);
    if (type === 'mixed') warnings.push(`"${name}" mixes numbers, dates and text.`);
    return { name, type, blanks, unique: uniq.size, uniqueCapped: capped, min, max, errors, sample };
  });
  const auto = t.headers.filter((h) => /^Column \d+$/.test(h));
  if (auto.length) warnings.push(`${auto.length} column(s) had no header and were named automatically (${auto.slice(0, 3).join(', ')}…).`);
  const dup = t.headers.filter((h) => / \(\d+\)$/.test(h));
  if (dup.length) warnings.push(`Repeated header names were made unique: ${dup.slice(0, 3).join(', ')}${dup.length > 3 ? '…' : ''}.`);
  return { file: t.fileName, sheet: t.sheetName, rows: t.rows.length, columns: t.headers.length, columnProfiles, warnings };
}
