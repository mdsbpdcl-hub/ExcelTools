import type { CellValue, Table } from '../types';
import type { CleanOptions } from '../workflow';
import { toTitleCase } from '../expr';
import { formatDate, parseDate, toNumber } from '../util';
import { isFormulaCell, resolveColumns, setCell } from './helpers';

const INVISIBLE = /[​-‍⁠﻿­\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g;
const NBSP = /[   　]/g;

export function cleanText(s: string, o: CleanOptions): string {
  let v = s;
  if (o.invisible) v = v.replace(INVISIBLE, '').replace(NBSP, ' ');
  if (o.lineBreaks === 'space') v = v.replace(/\s*(\r\n|\r|\n)\s*/g, ' ');
  else if (o.lineBreaks === 'remove') v = v.replace(/\r\n|\r|\n/g, '');
  if (o.collapseSpaces) v = o.lineBreaks === 'keep' ? v.replace(/[ \t]{2,}/g, ' ') : v.replace(/[ \t]+/g, ' ');
  if (o.trim) v = o.lineBreaks === 'keep' ? v.split('\n').map((l) => l.trim()).join('\n').trim() : v.trim();
  switch (o.caseMode) {
    case 'upper': v = v.toUpperCase(); break;
    case 'lower': v = v.toLowerCase(); break;
    case 'title': v = toTitleCase(v); break;
    case 'sentence': v = v.toLowerCase().replace(/(^\s*|[.!?]\s+)(\p{L})/gu, (_m, p: string, c: string) => p + c.toUpperCase()); break;
  }
  return v;
}

export interface CleanResult {
  changedCells: number;
  datesConverted: number;
  datesUnreadable: number;
  numbersConverted: number;
  formulaCellsSkipped: number;
  missingColumns: string[];
}

export function cleanTable(t: Table, o: CleanOptions): CleanResult {
  const res: CleanResult = { changedCells: 0, datesConverted: 0, datesUnreadable: 0, numbersConverted: 0, formulaCellsSkipped: 0, missingColumns: [] };
  const { idx, missing } = resolveColumns(t, o.columns);
  res.missingColumns.push(...missing);
  const textOn = o.trim || o.collapseSpaces || o.lineBreaks !== 'keep' || o.invisible || o.caseMode !== 'none';

  if (textOn) {
    for (let r = 0; r < t.rows.length; r++) {
      const row = t.rows[r];
      for (const c of idx) {
        const v = row[c];
        if (typeof v !== 'string') continue;
        if (isFormulaCell(t, r, c)) { res.formulaCellsSkipped++; continue; }
        const nv = cleanText(v, o);
        if (nv !== v) { setCell(t, r, c, nv === '' ? null : nv); res.changedCells++; }
      }
    }
  }

  if (o.numbers.enabled) {
    const n = resolveColumns(t, o.numbers.columns);
    res.missingColumns.push(...n.missing);
    for (let r = 0; r < t.rows.length; r++) {
      for (const c of n.idx) {
        const v = t.rows[r][c];
        if (typeof v !== 'string' || isFormulaCell(t, r, c)) continue;
        const num = toNumber(v);
        if (num !== null) { setCell(t, r, c, num); res.numbersConverted++; }
      }
    }
  }

  if (o.dates.enabled) {
    const d = resolveColumns(t, o.dates.columns);
    res.missingColumns.push(...d.missing);
    for (let r = 0; r < t.rows.length; r++) {
      for (const c of d.idx) {
        const v: CellValue = t.rows[r][c];
        if (v === null || typeof v === 'boolean' || isFormulaCell(t, r, c)) continue;
        if (typeof v === 'string' && v.trim() === '') continue;
        const date = parseDate(v, o.dates.inputOrder, false);
        if (!date) { res.datesUnreadable++; continue; }
        const out: CellValue = o.dates.asText ? formatDate(date, o.dates.outputFormat) : date;
        if (!(v instanceof Date && out instanceof Date && v.getTime() === out.getTime())) {
          setCell(t, r, c, out);
          res.datesConverted++;
        }
      }
    }
  }
  res.missingColumns = [...new Set(res.missingColumns)];
  return res;
}
