// A small, safe expression language for derived columns, custom validation rules and advanced filters.
// It never uses eval/new Function: expressions are parsed into a tree and interpreted.
//
// Syntax (Excel-like):
//   [Column Name]  or  Column_Name     column reference
//   12.5  "text"  'text'  TRUE  FALSE  literals
//   + - * / ^ &                        arithmetic and text join
//   = <> != < > <= >=                  comparisons (text comparison ignores case)
//   IF(cond, a, b)  AND(..)  OR(..)  NOT(x)  and the functions listed in FUNCTIONS below.

import type { CellValue } from './types';
import { cellToString, formatDate, isBlank, parseDate, toNumber } from './util';

type Node =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'bool'; v: boolean }
  | { t: 'col'; name: string }
  | { t: 'un'; op: string; a: Node }
  | { t: 'bin'; op: string; a: Node; b: Node }
  | { t: 'call'; name: string; args: Node[] };

export class ExprError extends Error {}

interface Tok { k: 'num' | 'str' | 'col' | 'id' | 'op' | '(' | ')' | ','; v: string; pos: number }

function tokenize(src: string): Tok[] {
  const out: Tok[] = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (/\s/.test(ch)) { i++; continue; }
    if (/[0-9.]/.test(ch)) {
      const m = src.slice(i).match(/^(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?/);
      if (!m) throw new ExprError(`Unexpected "${ch}" at position ${i + 1}`);
      out.push({ k: 'num', v: m[0], pos: i }); i += m[0].length; continue;
    }
    if (ch === '"' || ch === "'") {
      let j = i + 1; let s = '';
      while (j < src.length) {
        if (src[j] === ch && src[j + 1] === ch) { s += ch; j += 2; continue; }
        if (src[j] === ch) break;
        s += src[j]; j++;
      }
      if (j >= src.length) throw new ExprError('A text value is missing its closing quote');
      out.push({ k: 'str', v: s, pos: i }); i = j + 1; continue;
    }
    if (ch === '[') {
      const j = src.indexOf(']', i);
      if (j < 0) throw new ExprError('A column name is missing its closing ]');
      out.push({ k: 'col', v: src.slice(i + 1, j).trim(), pos: i }); i = j + 1; continue;
    }
    if (/[\p{L}_]/u.test(ch)) {
      const m = src.slice(i).match(/^[\p{L}_][\p{L}\p{N}_.]*/u)!;
      out.push({ k: 'id', v: m[0], pos: i }); i += m[0].length; continue;
    }
    const two = src.slice(i, i + 2);
    if (['<=', '>=', '<>', '!=', '=='].includes(two)) { out.push({ k: 'op', v: two === '==' ? '=' : two, pos: i }); i += 2; continue; }
    if ('+-*/^&=<>'.includes(ch)) { out.push({ k: 'op', v: ch, pos: i }); i++; continue; }
    if (ch === '(' || ch === ')' || ch === ',') { out.push({ k: ch, v: ch, pos: i }); i++; continue; }
    if (ch === ';') { out.push({ k: ',', v: ',', pos: i }); i++; continue; }
    throw new ExprError(`Unexpected "${ch}" at position ${i + 1}`);
  }
  return out;
}

function parse(src: string): Node {
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  const isOp = (...ops: string[]) => peek()?.k === 'op' && ops.includes(peek().v);
  const expect = (k: Tok['k']) => {
    if (peek()?.k !== k) throw new ExprError(`Expected "${k}" ${peek() ? `near position ${peek().pos + 1}` : 'at the end'}`);
    return toks[p++];
  };
  const comparison = (): Node => {
    let a = concat();
    while (isOp('=', '<>', '!=', '<', '>', '<=', '>=')) { const op = toks[p++].v; a = { t: 'bin', op, a, b: concat() }; }
    return a;
  };
  const concat = (): Node => {
    let a = additive();
    while (isOp('&')) { p++; a = { t: 'bin', op: '&', a, b: additive() }; }
    return a;
  };
  const additive = (): Node => {
    let a = mult();
    while (isOp('+', '-')) { const op = toks[p++].v; a = { t: 'bin', op, a, b: mult() }; }
    return a;
  };
  const mult = (): Node => {
    let a = unary();
    while (isOp('*', '/')) { const op = toks[p++].v; a = { t: 'bin', op, a, b: unary() }; }
    return a;
  };
  const unary = (): Node => {
    if (isOp('-', '+')) { const op = toks[p++].v; return { t: 'un', op, a: unary() }; }
    return power();
  };
  const power = (): Node => {
    const a = primary();
    if (isOp('^')) { p++; return { t: 'bin', op: '^', a, b: unary() }; }
    return a;
  };
  const primary = (): Node => {
    const tk = peek();
    if (!tk) throw new ExprError('The expression ends too early');
    if (tk.k === 'num') { p++; return { t: 'num', v: Number(tk.v) }; }
    if (tk.k === 'str') { p++; return { t: 'str', v: tk.v }; }
    if (tk.k === 'col') { p++; return { t: 'col', name: tk.v }; }
    if (tk.k === '(') { p++; const e = comparison(); expect(')'); return e; }
    if (tk.k === 'id') {
      p++;
      const up = tk.v.toUpperCase();
      if (peek()?.k === '(') {
        p++;
        const args: Node[] = [];
        if (peek()?.k !== ')') {
          args.push(comparison());
          while (peek()?.k === ',') { p++; args.push(comparison()); }
        }
        expect(')');
        if (!FUNCTIONS[up]) throw new ExprError(`Unknown function ${tk.v}(). See the function list for what is supported.`);
        return { t: 'call', name: up, args };
      }
      if (up === 'TRUE' || up === 'FALSE') return { t: 'bool', v: up === 'TRUE' };
      return { t: 'col', name: tk.v };
    }
    throw new ExprError(`Unexpected "${tk.v}" at position ${tk.pos + 1}`);
  };
  if (!toks.length) throw new ExprError('The expression is empty');
  const node = comparison();
  if (p < toks.length) throw new ExprError(`Unexpected "${toks[p].v}" at position ${toks[p].pos + 1}`);
  return node;
}

class ValueErr extends Error {}

const num = (v: CellValue): number => {
  if (isBlank(v)) return 0;
  if (v instanceof Date) return v.getTime() / 86400000 + 25569;
  const n = toNumber(v);
  if (n === null) throw new ValueErr(`"${cellToString(v)}" is not a number`);
  return n;
};
const str = (v: CellValue) => cellToString(v);
const truthy = (v: CellValue): boolean => {
  if (typeof v === 'boolean') return v;
  if (isBlank(v)) return false;
  if (typeof v === 'number') return v !== 0;
  const s = str(v).toUpperCase();
  if (s === 'TRUE') return true;
  if (s === 'FALSE') return false;
  const n = toNumber(v);
  return n === null ? true : n !== 0;
};
const date = (v: CellValue): Date => {
  const d = parseDate(v, 'DMY', true);
  if (!d) throw new ValueErr(`"${str(v)}" is not a date`);
  return d;
};
const round = (x: number, d: number, mode: 'n' | 'u' | 'd') => {
  const f = 10 ** d;
  const y = x * f;
  const r = mode === 'n' ? Math.sign(y) * Math.round(Math.abs(y) + 1e-9) : mode === 'u' ? Math.sign(y) * Math.ceil(Math.abs(y) - 1e-9) : Math.trunc(y);
  return r / f;
};

type Fn = { min: number; max: number; help: string; f: (a: CellValue[]) => CellValue };

/** Functions available in expressions. IF/AND/OR are evaluated lazily below. */
export const FUNCTIONS: Record<string, Fn> = {
  IF: { min: 2, max: 3, help: 'IF(condition, value_if_true, value_if_false)', f: () => null },
  AND: { min: 1, max: 99, help: 'AND(a, b, ...) — true when all are true', f: (a) => a.every(truthy) },
  OR: { min: 1, max: 99, help: 'OR(a, b, ...) — true when any is true', f: (a) => a.some(truthy) },
  NOT: { min: 1, max: 1, help: 'NOT(a)', f: (a) => !truthy(a[0]) },
  IFERROR: { min: 2, max: 2, help: 'IFERROR(value, fallback)', f: () => null },
  ISBLANK: { min: 1, max: 1, help: 'ISBLANK(value)', f: (a) => isBlank(a[0]) },
  ISNUMBER: { min: 1, max: 1, help: 'ISNUMBER(value) — also true for numbers stored as text', f: (a) => toNumber(a[0]) !== null },
  ISDATE: { min: 1, max: 1, help: 'ISDATE(value)', f: (a) => parseDate(a[0], 'DMY', false) !== null },
  COALESCE: { min: 1, max: 99, help: 'COALESCE(a, b, ...) — first non-blank value', f: (a) => a.find((x) => !isBlank(x)) ?? null },
  ROUND: { min: 1, max: 2, help: 'ROUND(number, digits)', f: (a) => round(num(a[0]), a.length > 1 ? num(a[1]) : 0, 'n') },
  ROUNDUP: { min: 1, max: 2, help: 'ROUNDUP(number, digits)', f: (a) => round(num(a[0]), a.length > 1 ? num(a[1]) : 0, 'u') },
  ROUNDDOWN: { min: 1, max: 2, help: 'ROUNDDOWN(number, digits)', f: (a) => round(num(a[0]), a.length > 1 ? num(a[1]) : 0, 'd') },
  INT: { min: 1, max: 1, help: 'INT(number)', f: (a) => Math.floor(num(a[0])) },
  ABS: { min: 1, max: 1, help: 'ABS(number)', f: (a) => Math.abs(num(a[0])) },
  MOD: { min: 2, max: 2, help: 'MOD(number, divisor)', f: (a) => { const d = num(a[1]); if (d === 0) throw new ValueErr('Division by zero'); const n = num(a[0]); return n - d * Math.floor(n / d); } },
  SUM: { min: 1, max: 99, help: 'SUM(a, b, ...)', f: (a) => a.reduce<number>((s, x) => s + (isBlank(x) ? 0 : num(x)), 0) },
  AVERAGE: { min: 1, max: 99, help: 'AVERAGE(a, b, ...) — ignores blanks', f: (a) => { const v = a.filter((x) => !isBlank(x)).map(num); if (!v.length) throw new ValueErr('Nothing to average'); return v.reduce((s, x) => s + x, 0) / v.length; } },
  MIN: { min: 1, max: 99, help: 'MIN(a, b, ...)', f: (a) => Math.min(...a.filter((x) => !isBlank(x)).map(num)) },
  MAX: { min: 1, max: 99, help: 'MAX(a, b, ...)', f: (a) => Math.max(...a.filter((x) => !isBlank(x)).map(num)) },
  UPPER: { min: 1, max: 1, help: 'UPPER(text)', f: (a) => str(a[0]).toUpperCase() },
  LOWER: { min: 1, max: 1, help: 'LOWER(text)', f: (a) => str(a[0]).toLowerCase() },
  PROPER: { min: 1, max: 1, help: 'PROPER(text) — Title Case', f: (a) => toTitleCase(str(a[0])) },
  TRIM: { min: 1, max: 1, help: 'TRIM(text) — removes extra spaces', f: (a) => str(a[0]).replace(/\s+/g, ' ').trim() },
  LEN: { min: 1, max: 1, help: 'LEN(text)', f: (a) => [...str(a[0])].length },
  LEFT: { min: 1, max: 2, help: 'LEFT(text, count)', f: (a) => [...str(a[0])].slice(0, a.length > 1 ? num(a[1]) : 1).join('') },
  RIGHT: { min: 1, max: 2, help: 'RIGHT(text, count)', f: (a) => { const c = [...str(a[0])]; const n = a.length > 1 ? num(a[1]) : 1; return n <= 0 ? '' : c.slice(-n).join(''); } },
  MID: { min: 3, max: 3, help: 'MID(text, start, count) — start at 1', f: (a) => [...str(a[0])].slice(num(a[1]) - 1, num(a[1]) - 1 + num(a[2])).join('') },
  CONCAT: { min: 1, max: 99, help: 'CONCAT(a, b, ...)', f: (a) => a.map(str).join('') },
  TEXTJOIN: { min: 3, max: 99, help: 'TEXTJOIN(separator, skip_blanks, a, b, ...)', f: (a) => a.slice(2).filter((x) => !(truthy(a[1]) && isBlank(x))).map(str).join(str(a[0])) },
  SUBSTITUTE: { min: 3, max: 3, help: 'SUBSTITUTE(text, old, new)', f: (a) => str(a[0]).split(str(a[1])).join(str(a[2])) },
  CONTAINS: { min: 2, max: 2, help: 'CONTAINS(text, part) — ignores case', f: (a) => str(a[0]).toLowerCase().includes(str(a[1]).toLowerCase()) },
  STARTSWITH: { min: 2, max: 2, help: 'STARTSWITH(text, part) — ignores case', f: (a) => str(a[0]).toLowerCase().startsWith(str(a[1]).toLowerCase()) },
  ENDSWITH: { min: 2, max: 2, help: 'ENDSWITH(text, part) — ignores case', f: (a) => str(a[0]).toLowerCase().endsWith(str(a[1]).toLowerCase()) },
  REGEXMATCH: { min: 2, max: 2, help: 'REGEXMATCH(text, pattern)', f: (a) => { try { return new RegExp(str(a[1])).test(str(a[0])); } catch { throw new ValueErr('Invalid pattern'); } } },
  VALUE: { min: 1, max: 1, help: 'VALUE(text) — text to number', f: (a) => num(a[0]) },
  TEXT: { min: 2, max: 2, help: 'TEXT(value, format) — dates: "DD-MM-YYYY"; numbers: "0.00"', f: (a) => textFn(a[0], str(a[1])) },
  DATE: { min: 3, max: 3, help: 'DATE(year, month, day)', f: (a) => new Date(Date.UTC(num(a[0]), num(a[1]) - 1, num(a[2]))) },
  DATEVALUE: { min: 1, max: 1, help: 'DATEVALUE(text) — reads day/month/year text', f: (a) => date(a[0]) },
  YEAR: { min: 1, max: 1, help: 'YEAR(date)', f: (a) => date(a[0]).getUTCFullYear() },
  MONTH: { min: 1, max: 1, help: 'MONTH(date)', f: (a) => date(a[0]).getUTCMonth() + 1 },
  DAY: { min: 1, max: 1, help: 'DAY(date)', f: (a) => date(a[0]).getUTCDate() },
  TODAY: { min: 0, max: 0, help: 'TODAY()', f: () => { const n = new Date(); return new Date(Date.UTC(n.getFullYear(), n.getMonth(), n.getDate())); } },
  DAYS: { min: 2, max: 2, help: 'DAYS(end_date, start_date)', f: (a) => Math.round((date(a[0]).getTime() - date(a[1]).getTime()) / 86400000) },
};

export function toTitleCase(s: string): string {
  return s.toLowerCase().replace(/(^|[\s\-/.(])(\p{L})/gu, (_m, p: string, c: string) => p + c.toUpperCase());
}

function textFn(v: CellValue, fmt: string): string {
  if (/[YDM]/.test(fmt)) {
    const d = parseDate(v, 'DMY', true);
    return d ? formatDate(d, fmt.replace(/yyyy/i, 'YYYY').replace(/dd/i, 'DD')) : str(v);
  }
  const n = toNumber(v);
  if (n === null) return str(v);
  const dec = (fmt.split('.')[1] || '').replace(/[^0#]/g, '').length;
  const grouped = fmt.includes(',');
  return grouped ? n.toLocaleString('en-IN', { minimumFractionDigits: dec, maximumFractionDigits: dec }) : n.toFixed(dec);
}

function compare(a: CellValue, b: CellValue): number {
  if (a instanceof Date || b instanceof Date) {
    const da = parseDate(a, 'DMY', true), db = parseDate(b, 'DMY', true);
    if (da && db) return da.getTime() - db.getTime();
  }
  const na = typeof a === 'number' ? a : toNumber(a);
  const nb = typeof b === 'number' ? b : toNumber(b);
  if (na !== null && nb !== null && !isBlank(a) && !isBlank(b)) return na - nb;
  return str(a).toLowerCase().localeCompare(str(b).toLowerCase());
}

export type RowGetter = (column: string) => CellValue;

function evaluate(n: Node, get: RowGetter): CellValue {
  switch (n.t) {
    case 'num': return n.v;
    case 'str': return n.v;
    case 'bool': return n.v;
    case 'col': return get(n.name);
    case 'un': { const v = num(evaluate(n.a, get)); return n.op === '-' ? -v : v; }
    case 'bin': {
      const a = evaluate(n.a, get), b = evaluate(n.b, get);
      switch (n.op) {
        case '+': return num(a) + num(b);
        case '-': return num(a) - num(b);
        case '*': return num(a) * num(b);
        case '/': { const d = num(b); if (d === 0) throw new ValueErr('Division by zero'); return num(a) / d; }
        case '^': return num(a) ** num(b);
        case '&': return str(a) + str(b);
        case '=': return compare(a, b) === 0;
        case '<>': case '!=': return compare(a, b) !== 0;
        case '<': return compare(a, b) < 0;
        case '>': return compare(a, b) > 0;
        case '<=': return compare(a, b) <= 0;
        case '>=': return compare(a, b) >= 0;
      }
      throw new ValueErr('Unknown operator');
    }
    case 'call': {
      if (n.name === 'IF') return truthy(evaluate(n.args[0], get)) ? evaluate(n.args[1], get) : n.args[2] ? evaluate(n.args[2], get) : false;
      if (n.name === 'IFERROR') { try { return evaluate(n.args[0], get); } catch (e) { if (e instanceof ValueErr) return evaluate(n.args[1], get); throw e; } }
      if (n.name === 'AND') return n.args.every((x) => truthy(evaluate(x, get)));
      if (n.name === 'OR') return n.args.some((x) => truthy(evaluate(x, get)));
      return FUNCTIONS[n.name].f(n.args.map((x) => evaluate(x, get)));
    }
  }
}

function collect(n: Node, cols: Set<string>): void {
  if (n.t === 'col') cols.add(n.name);
  if (n.t === 'un') collect(n.a, cols);
  if (n.t === 'bin') { collect(n.a, cols); collect(n.b, cols); }
  if (n.t === 'call') {
    const fn = FUNCTIONS[n.name];
    if (n.args.length < fn.min || n.args.length > fn.max) {
      throw new ExprError(`${n.name} needs ${fn.min === fn.max ? fn.min : `${fn.min} to ${fn.max}`} value(s): ${fn.help}`);
    }
    n.args.forEach((a) => collect(a, cols));
  }
}

export interface CompiledExpr {
  source: string;
  columns: string[];
  /** Returns the result, or throws ExprValueError for a bad value in this row. */
  run: (get: RowGetter) => CellValue;
}

export class ExprValueError extends Error {}

/**
 * Parses an expression and checks its column names against `headers` (case-insensitive).
 * Throws ExprError with a plain-language message if anything is wrong.
 */
export function compileExpr(source: string, headers: string[]): CompiledExpr {
  const tree = parse(source);
  const cols = new Set<string>();
  collect(tree, cols);
  const lookup = new Map(headers.map((h) => [h.toLowerCase(), h]));
  const resolved = new Map<string, string>();
  for (const c of cols) {
    // Bare names may use _ for spaces: Bill_Date → [Bill Date]. A real column named "Bill_Date" wins.
    const real = lookup.get(c.toLowerCase()) ?? lookup.get(c.replace(/_/g, ' ').toLowerCase());
    if (!real) throw new ExprError(`Column [${c}] was not found. Available: ${headers.slice(0, 12).map((h) => `[${h}]`).join(', ')}${headers.length > 12 ? ' …' : ''}`);
    resolved.set(c, real);
  }
  return {
    source,
    columns: [...resolved.values()],
    run: (get) => {
      try {
        return evaluate(tree, (name) => get(resolved.get(name) ?? name));
      } catch (e) {
        if (e instanceof ValueErr) throw new ExprValueError(e.message);
        throw e;
      }
    },
  };
}

export function isTruthy(v: CellValue): boolean {
  return truthy(v);
}
