import type { Issue, Table } from '../types';
import type { ValidateOptions, ValidationRule } from '../workflow';
import { compileExpr, ExprError, ExprValueError, isTruthy } from '../expr';
import { cellToString, isBlank, parseDate, toNumber } from '../util';
import { colIndex, COLORS, highlight } from './helpers';

export const RULE_LABELS: Record<ValidationRule['type'], string> = {
  required: 'Required (not blank)',
  list: 'Must be one of a list (dropdown)',
  range: 'Number between min and max',
  number: 'Must be a number',
  date: 'Must be a date',
  dateRange: 'Date between two dates',
  pattern: 'Must match a pattern',
  unique: 'Must be unique',
  expression: 'Custom rule (expression)',
};

export const listValues = (s: string) => s.split(/[\n,]/).map((x) => x.trim()).filter(Boolean);

export interface ValidationResult {
  issues: Issue[];
  failedRows: number;
  warnings: string[];
}

export function validateTable(t: Table, o: ValidateOptions): ValidationResult {
  const issues: Issue[] = [];
  const warnings: string[] = [];
  const badRows = new Set<number>();
  const lookup = new Map(t.headers.map((h, i) => [h, i]));
  const add = (r: number, c: number, rule: ValidationRule, msg: string) => {
    badRows.add(r);
    if (o.markInvalidCells && c >= 0) highlight(t, r, c, COLORS.issue);
    issues.push({ kind: 'validation', file: t.fileName, sheet: t.sheetName, row: t.rowNums[r], column: c >= 0 ? t.headers[c] : rule.column, value: c >= 0 ? cellToString(t.rows[r][c]) : '', message: rule.message || msg });
  };

  for (const rule of o.rules) {
    if (rule.type === 'expression') {
      let expr;
      try { expr = compileExpr(rule.expression, t.headers); } catch (e) {
        if (e instanceof ExprError) { warnings.push(`${t.fileName} › ${t.sheetName}: rule skipped — ${e.message}`); continue; }
        throw e;
      }
      const c = rule.column ? colIndex(t, rule.column) : -1;
      t.rows.forEach((row, r) => {
        let ok = false;
        try { ok = isTruthy(expr.run((col) => row[lookup.get(col) ?? -1] ?? null)); } catch (e) { if (!(e instanceof ExprValueError)) throw e; }
        if (!ok) add(r, c, rule, `Fails rule: ${rule.expression}`);
      });
      continue;
    }
    const c = colIndex(t, rule.column);
    if (c < 0) { warnings.push(`${t.fileName} › ${t.sheetName}: column "${rule.column}" not found; its rule was skipped.`); continue; }
    switch (rule.type) {
      case 'required':
        t.rows.forEach((row, r) => { if (isBlank(row[c])) add(r, c, rule, 'Required value is blank'); });
        break;
      case 'list': {
        const allowed = new Set(listValues(rule.values).map((v) => v.toLowerCase()));
        t.rows.forEach((row, r) => { if (!isBlank(row[c]) && !allowed.has(cellToString(row[c]).trim().toLowerCase())) add(r, c, rule, `Not in the allowed list`); });
        break;
      }
      case 'number':
        t.rows.forEach((row, r) => { if (!isBlank(row[c]) && toNumber(row[c]) === null) add(r, c, rule, 'Not a number'); });
        break;
      case 'range': {
        const min = rule.min === '' ? -Infinity : toNumber(rule.min) ?? -Infinity;
        const max = rule.max === '' ? Infinity : toNumber(rule.max) ?? Infinity;
        t.rows.forEach((row, r) => {
          if (isBlank(row[c])) return;
          const n = toNumber(row[c]);
          if (n === null) add(r, c, rule, 'Not a number');
          else if (n < min || n > max) add(r, c, rule, `Outside ${rule.min || '…'} to ${rule.max || '…'}`);
        });
        break;
      }
      case 'date':
      case 'dateRange': {
        const min = rule.type === 'dateRange' && rule.min ? parseDate(rule.min, o.dateOrder) : null;
        const max = rule.type === 'dateRange' && rule.max ? parseDate(rule.max, o.dateOrder) : null;
        if (rule.type === 'dateRange' && ((rule.min && !min) || (rule.max && !max))) {
          warnings.push(`Date rule on "${rule.column}" has a date that could not be read; use DD-MM-YYYY.`);
        }
        t.rows.forEach((row, r) => {
          if (isBlank(row[c])) return;
          const d = parseDate(row[c], o.dateOrder);
          if (!d) add(r, c, rule, 'Not a valid date');
          else if ((min && d < min) || (max && d > max)) add(r, c, rule, `Outside ${rule.min || '…'} to ${rule.max || '…'}`);
        });
        break;
      }
      case 'pattern': {
        let re: RegExp;
        try { re = new RegExp(rule.pattern, 'u'); } catch { warnings.push(`Pattern for "${rule.column}" is not valid; rule skipped.`); break; }
        t.rows.forEach((row, r) => { if (!isBlank(row[c]) && !re.test(cellToString(row[c]))) add(r, c, rule, `Does not match ${rule.pattern}`); });
        break;
      }
      case 'unique': {
        const seen = new Map<string, number>();
        t.rows.forEach((row, r) => {
          if (isBlank(row[c])) return;
          const k = cellToString(row[c]).trim().toLowerCase();
          if (seen.has(k)) add(r, c, rule, `Repeats the value in row ${t.rowNums[seen.get(k)!]}`);
          else seen.set(k, r);
        });
        break;
      }
    }
  }
  return { issues, failedRows: badRows.size, warnings };
}
