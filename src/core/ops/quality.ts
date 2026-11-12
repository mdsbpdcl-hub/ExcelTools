// Data-quality checks: spreadsheet errors (#N/A …), blank cells, invalid dates and unexpected values.
// For each kind the user chooses: leave unchanged, flag (highlight + list in the issues report) or replace.

import type { CellValue, Issue, Table } from '../types';
import type { ErrorOptions, IssueAction } from '../workflow';
import { cellToString, isBlank, isErrorValue, maybePlainNumber, parseDateDetailed, toNumber } from '../util';
import { COLORS, highlight, resolveColumns, setCell } from './helpers';
import { inferColumnType } from './profile';

export interface QualityResult {
  issues: Issue[];
  counts: { spreadsheetErrors: number; blanks: number; invalidDates: number; unexpected: number };
  replaced: number;
}

const replacement = (a: IssueAction): CellValue => (a.replaceWith === '' ? null : maybePlainNumber(a.replaceWith));

export function checkQuality(t: Table, o: ErrorOptions): QualityResult {
  const res: QualityResult = { issues: [], counts: { spreadsheetErrors: 0, blanks: 0, invalidDates: 0, unexpected: 0 }, replaced: 0 };
  const { idx } = resolveColumns(t, o.columns);
  const types = new Map(idx.map((c) => [c, inferColumnType(t, c, o.dateOrder)]));
  const handle = (r: number, c: number, kind: Issue['kind'], a: IssueAction, message: string) => {
    if (a.action === 'leave') return;
    const v = t.rows[r][c];
    if (a.action === 'flag') {
      highlight(t, r, c, COLORS.issue);
      res.issues.push({ kind, file: t.fileName, sheet: t.sheetName, row: t.rowNums[r], column: t.headers[c], value: cellToString(v), message });
    } else {
      setCell(t, r, c, replacement(a));
      res.replaced++;
    }
  };
  for (let r = 0; r < t.rows.length; r++) {
    const row = t.rows[r];
    for (const c of idx) {
      const v = row[c];
      if (isErrorValue(v)) {
        res.counts.spreadsheetErrors++;
        handle(r, c, 'spreadsheet-error', o.spreadsheetErrors, `Spreadsheet error ${cellToString(v)}`);
        continue;
      }
      if (isBlank(v)) {
        res.counts.blanks++;
        handle(r, c, 'blank', o.blanks, 'Blank cell');
        continue;
      }
      const type = types.get(c);
      if (type === 'date' && !(v instanceof Date)) {
        const d = parseDateDetailed(v, o.dateOrder);
        if (!d.date && d.looksLikeDate) {
          res.counts.invalidDates++;
          handle(r, c, 'invalid-date', o.invalidDates, 'Not a real calendar date');
        } else if (!d.date && typeof v !== 'number') {
          res.counts.unexpected++;
          handle(r, c, 'unexpected-value', o.unexpected, 'Expected a date');
        }
      } else if (type === 'number' && typeof v !== 'number' && toNumber(v) === null) {
        res.counts.unexpected++;
        handle(r, c, 'unexpected-value', o.unexpected, 'Expected a number');
      }
    }
  }
  return res;
}
