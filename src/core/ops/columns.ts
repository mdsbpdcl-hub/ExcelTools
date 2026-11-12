import type { CellValue, Table } from '../types';
import type { ColumnOptions } from '../workflow';
import { compileExpr, ExprError, ExprValueError } from '../expr';
import { cellToString, isBlank, uniqueHeaders } from '../util';
import { colIndex, markStructural, marksByName, ProcessError, restoreMarks } from './helpers';

export interface ColumnResult {
  warnings: string[];
  expressionErrors: number;
}

function addColumn(t: Table, name: string, values: CellValue[], at = t.headers.length): void {
  t.headers.splice(at, 0, name);
  t.rows.forEach((row, i) => row.splice(at, 0, values[i]));
}

function removeColumn(t: Table, idx: number): void {
  t.headers.splice(idx, 1);
  t.rows.forEach((row) => row.splice(idx, 1));
}

/** Order of work: split → combine → derived → rename → remove → order/keep. */
export function applyColumnTools(t: Table, o: ColumnOptions): ColumnResult {
  const warnings: string[] = [];
  const marks = marksByName(t);
  const renamed = new Map<string, string>();
  let expressionErrors = 0;
  let reshaped = false;
  const where = `${t.fileName} › ${t.sheetName}`;

  for (const rule of o.split) {
    const i = colIndex(t, rule.column);
    if (i < 0) { warnings.push(`${where}: column "${rule.column}" not found for splitting.`); continue; }
    if (!rule.delimiter) throw new ProcessError(`Splitting "${rule.column}" needs a delimiter.`);
    const delim = rule.delimiter === '\\t' ? '\t' : rule.delimiter;
    const parts = t.rows.map((row) => (isBlank(row[i]) ? [] : cellToString(row[i]).split(delim).map((p) => p.trim())));
    const count = Math.max(rule.into.filter(Boolean).length, ...parts.map((p) => p.length), 1);
    const names = uniqueHeaders([...t.headers, ...Array.from({ length: count }, (_, k) => rule.into[k] || `${t.headers[i]} ${k + 1}`)]).slice(t.headers.length);
    const named = rule.into.filter(Boolean).length;
    if (named && count > named) warnings.push(`${where}: "${rule.column}" split into up to ${count} parts; extra parts went to extra columns.`);
    names.forEach((name, k) => addColumn(t, name, parts.map((p) => (p[k] === undefined || p[k] === '' ? null : p[k])), i + 1 + k));
    if (!rule.keepOriginal) removeColumn(t, i);
    reshaped = true;
  }

  for (const rule of o.combine) {
    const idx = rule.columns.map((c) => colIndex(t, c));
    const missing = rule.columns.filter((_, k) => idx[k] < 0);
    if (missing.length) { warnings.push(`${where}: cannot combine; missing ${missing.join(', ')}.`); continue; }
    const sep = rule.separator === '\\t' ? '\t' : rule.separator;
    const values = t.rows.map((row) => {
      const parts = idx.map((i) => row[i]).filter((v) => !(rule.skipBlanks && isBlank(v))).map(cellToString);
      const s = parts.join(sep);
      return s === '' ? null : s;
    });
    const name = uniqueHeaders([...t.headers, rule.into || rule.columns.join(' ')]).pop()!;
    addColumn(t, name, values, Math.max(...idx) + 1);
    if (!rule.keepOriginals) [...idx].sort((a, b) => b - a).forEach((i) => removeColumn(t, i));
    reshaped = true;
  }

  for (const rule of o.derived) {
    if (!rule.name || !rule.expression) continue;
    let expr;
    try {
      expr = compileExpr(rule.expression, t.headers);
    } catch (e) {
      if (e instanceof ExprError) { warnings.push(`${where}: new column "${rule.name}" skipped — ${e.message}`); continue; }
      throw e;
    }
    const lookup = new Map(t.headers.map((h, i) => [h, i]));
    const values = t.rows.map((row) => {
      try {
        const v = expr.run((col) => row[lookup.get(col) ?? -1] ?? null);
        return typeof v === 'number' && !Number.isFinite(v) ? '#NUM!' : v;
      } catch (e) {
        if (e instanceof ExprValueError) { expressionErrors++; return '#VALUE!'; }
        throw e;
      }
    });
    const existing = colIndex(t, rule.name);
    if (existing >= 0) t.rows.forEach((row, i) => { row[existing] = values[i]; });
    else addColumn(t, rule.name, values);
    reshaped = true;
  }
  if (expressionErrors) warnings.push(`${where}: ${expressionErrors} cell(s) in new columns could not be calculated and show #VALUE!.`);

  for (const { from, to } of o.rename) {
    if (!from || !to) continue;
    const i = colIndex(t, from);
    if (i < 0) continue;
    if (colIndex(t, to) >= 0 && colIndex(t, to) !== i) { warnings.push(`${where}: cannot rename "${from}" to "${to}" — that name is already used.`); continue; }
    for (const [k, v] of renamed) if (v === t.headers[i]) renamed.set(k, to);
    if (![...renamed.values()].includes(to)) renamed.set(t.headers[i], to);
    t.headers[i] = to;
  }

  const removeIdx = o.remove.map((n) => colIndex(t, n)).filter((i) => i >= 0);
  if (removeIdx.length) {
    [...new Set(removeIdx)].sort((a, b) => b - a).forEach((i) => removeColumn(t, i));
    reshaped = true;
  }

  if (o.order.length || o.keepOnlyListed) {
    const listed = o.order.map((n) => colIndex(t, n)).filter((i) => i >= 0);
    const uniqListed = [...new Set(listed)];
    const rest = t.headers.map((_, i) => i).filter((i) => !uniqListed.includes(i));
    const finalOrder = o.keepOnlyListed ? uniqListed : [...uniqListed, ...rest];
    if (o.keepOnlyListed && !finalOrder.length) {
      warnings.push(`${where}: none of the listed columns exist, so all columns were kept.`);
    } else if (finalOrder.some((v, k) => v !== k) || finalOrder.length !== t.headers.length) {
      t.headers = finalOrder.map((i) => t.headers[i]);
      t.rows = t.rows.map((row) => finalOrder.map((i) => row[i]));
      reshaped = true;
    }
  }

  if (reshaped) markStructural(t);
  if (marks.length) restoreMarks(t, marks, renamed);
  return { warnings, expressionErrors };
}
