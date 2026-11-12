import type { CellValue, Table } from '../types';
import type { FilterCondition, FilterOptions } from '../workflow';
import { compileExpr, ExprError, ExprValueError, isTruthy } from '../expr';
import { cellToString, isBlank, parseDate, toNumber, type DateOrder } from '../util';
import { colIndex, deriveTable, keepRows, ProcessError } from './helpers';

type Test = (row: CellValue[]) => boolean;

const listOf = (s: string) => s.split(/[\n,]/).map((x) => x.trim()).filter(Boolean);

function compileCondition(t: Table, c: FilterCondition, order: DateOrder): Test | null {
  if (c.op === 'expression') {
    try {
      const expr = compileExpr(c.value, t.headers);
      const lookup = new Map(t.headers.map((h, i) => [h, i]));
      return (row) => {
        try { return isTruthy(expr.run((col) => row[lookup.get(col) ?? -1] ?? null)); } catch (e) { if (e instanceof ExprValueError) return false; throw e; }
      };
    } catch (e) {
      if (e instanceof ExprError) throw new ProcessError(`Filter expression problem: ${e.message}`);
      throw e;
    }
  }
  const i = colIndex(t, c.column);
  if (i < 0) return null;
  const cs = c.caseSensitive;
  const norm = (s: string) => (cs ? s : s.toLowerCase());
  const val = norm(c.value);
  const text = (row: CellValue[]) => norm(cellToString(row[i]));
  const num = (s: string, label: string) => {
    const n = toNumber(s);
    if (n === null) throw new ProcessError(`Filter on "${c.column}": "${s}" is not a number (${label}).`);
    return n;
  };
  const date = (s: string, label: string) => {
    const d = parseDate(s, order);
    if (!d) throw new ProcessError(`Filter on "${c.column}": "${s}" is not a date (${label}). Use DD-MM-YYYY or YYYY-MM-DD.`);
    return d.getTime();
  };
  const cellNum = (row: CellValue[]) => toNumber(row[i]);
  const cellDay = (row: CellValue[]) => {
    const d = parseDate(row[i], order, false);
    return d ? Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) : null;
  };
  switch (c.op) {
    case 'equals': {
      const n = toNumber(c.value);
      return (row) => (n !== null && cellNum(row) !== null ? cellNum(row) === n : text(row) === val);
    }
    case 'notEquals': {
      const n = toNumber(c.value);
      return (row) => !(n !== null && cellNum(row) !== null ? cellNum(row) === n : text(row) === val);
    }
    case 'contains': return (row) => text(row).includes(val);
    case 'notContains': return (row) => !text(row).includes(val);
    case 'startsWith': return (row) => text(row).startsWith(val);
    case 'endsWith': return (row) => text(row).endsWith(val);
    case 'regex': {
      let re: RegExp;
      try { re = new RegExp(c.value, cs ? 'u' : 'iu'); } catch { throw new ProcessError(`Filter pattern "${c.value}" is not valid.`); }
      return (row) => re.test(cellToString(row[i]));
    }
    case 'gt': { const n = num(c.value, 'greater than'); return (row) => { const v = cellNum(row); return v !== null && v > n; }; }
    case 'gte': { const n = num(c.value, 'at least'); return (row) => { const v = cellNum(row); return v !== null && v >= n; }; }
    case 'lt': { const n = num(c.value, 'less than'); return (row) => { const v = cellNum(row); return v !== null && v < n; }; }
    case 'lte': { const n = num(c.value, 'at most'); return (row) => { const v = cellNum(row); return v !== null && v <= n; }; }
    case 'between': {
      const a = num(c.value, 'from'), b = num(c.value2, 'to');
      return (row) => { const v = cellNum(row); return v !== null && v >= Math.min(a, b) && v <= Math.max(a, b); };
    }
    case 'dateOn': { const d = date(c.value, 'date'); return (row) => cellDay(row) === d; }
    case 'dateBefore': { const d = date(c.value, 'before'); return (row) => { const v = cellDay(row); return v !== null && v < d; }; }
    case 'dateAfter': { const d = date(c.value, 'after'); return (row) => { const v = cellDay(row); return v !== null && v > d; }; }
    case 'dateBetween': {
      const a = date(c.value, 'from'), b = date(c.value2, 'to');
      return (row) => { const v = cellDay(row); return v !== null && v >= Math.min(a, b) && v <= Math.max(a, b); };
    }
    case 'blank': return (row) => isBlank(row[i]);
    case 'notBlank': return (row) => !isBlank(row[i]);
    case 'inList': { const set = new Set(listOf(c.value).map(norm)); return (row) => set.has(text(row).trim()); }
    case 'notInList': { const set = new Set(listOf(c.value).map(norm)); return (row) => !set.has(text(row).trim()); }
  }
  return null;
}

export interface FilterResult {
  matched: Table;
  rest: Table | null;
  matchedCount: number;
  restCount: number;
  warnings: string[];
}

export function filterTable(t: Table, o: FilterOptions): FilterResult {
  const warnings: string[] = [];
  const tests: Test[] = [];
  for (const c of o.conditions) {
    const test = compileCondition(t, c, o.dateOrder);
    if (!test) warnings.push(`${t.fileName} › ${t.sheetName}: filter column "${c.column}" not found; that condition was ignored.`);
    else tests.push(test);
  }
  if (!tests.length) return { matched: t, rest: null, matchedCount: t.rows.length, restCount: 0, warnings };
  const keep: number[] = [], drop: number[] = [];
  t.rows.forEach((row, r) => {
    const ok = o.match === 'all' ? tests.every((f) => f(row)) : tests.some((f) => f(row));
    (ok ? keep : drop).push(r);
  });
  let rest: Table | null = null;
  if (o.saveNonMatching && drop.length) {
    rest = deriveTable(t, { rows: drop.map((r) => t.rows[r]), rowNums: drop.map((r) => t.rowNums[r]), headers: [...t.headers] });
  }
  if (drop.length) keepRows(t, keep);
  return { matched: t, rest, matchedCount: keep.length, restCount: drop.length, warnings };
}
