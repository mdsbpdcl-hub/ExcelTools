import type { CellValue, Table } from '../types';
import type { AggFn, PivotOptions } from '../workflow';
import { cellToString, compareValues, isBlank, toNumber } from '../util';
import { colIndex, newTable, ProcessError } from './helpers';

export const AGG_LABELS: Record<AggFn, string> = {
  sum: 'Sum', count: 'Count', countDistinct: 'Distinct count', avg: 'Average', min: 'Minimum', max: 'Maximum', median: 'Median',
};

const round = (n: number) => Math.round(n * 1e10) / 1e10;

/** Running aggregate for one cell of the summary. */
export class Acc {
  sum = 0;
  count = 0; // non-blank values
  numCount = 0; // numeric values
  nums: number[] = [];
  min = Infinity;
  max = -Infinity;
  distinct = new Set<string>();

  add(v: CellValue, keepNums: boolean): void {
    if (isBlank(v)) return;
    this.distinct.add(cellToString(v).trim().toLowerCase());
    this.count++;
    const n = toNumber(v);
    if (n === null) return;
    this.numCount++;
    this.sum += n;
    if (n < this.min) this.min = n;
    if (n > this.max) this.max = n;
    if (keepNums) this.nums.push(n);
  }

  absorb(o: Acc): void {
    this.sum += o.sum;
    this.count += o.count;
    this.numCount += o.numCount;
    for (const n of o.nums) this.nums.push(n);
    this.min = Math.min(this.min, o.min);
    this.max = Math.max(this.max, o.max);
    o.distinct.forEach((d) => this.distinct.add(d));
  }

  result(fn: AggFn): CellValue {
    switch (fn) {
      case 'count': return this.count;
      case 'countDistinct': return this.distinct.size;
      case 'sum': return round(this.sum);
      case 'avg': return this.numCount ? round(this.sum / this.numCount) : null;
      case 'min': return this.min === Infinity ? null : this.min;
      case 'max': return this.max === -Infinity ? null : this.max;
      case 'median': {
        if (!this.nums.length) return null;
        const s = [...this.nums].sort((a, b) => a - b);
        const m = Math.floor(s.length / 2);
        return s.length % 2 ? s[m] : round((s[m - 1] + s[m]) / 2);
      }
    }
  }
}

/** Groups by the chosen fields (optionally spreading one field across columns) and aggregates the value fields. */
export function pivotTable(t: Table, o: PivotOptions): { table: Table; nonNumeric: number } {
  if (!o.values.length) throw new ProcessError('Add at least one value to summarise (for example Sum of Amount).');
  const gIdx = o.groupBy.map((g) => colIndex(t, g));
  const pIdx = o.pivotColumn ? colIndex(t, o.pivotColumn) : -1;
  const vIdx = o.values.map((v) => colIndex(t, v.column));
  const missing = [
    ...o.groupBy.filter((_, k) => gIdx[k] < 0),
    ...o.values.filter((_, k) => vIdx[k] < 0).map((v) => v.column),
    ...(o.pivotColumn && pIdx < 0 ? [o.pivotColumn] : []),
  ];
  if (missing.length) throw new ProcessError(`Column(s) ${missing.join(', ')} not found in ${t.fileName} › ${t.sheetName}.`);

  const keepNums = o.values.some((v) => v.agg === 'median');
  const numericAgg = (a: AggFn) => a === 'sum' || a === 'avg' || a === 'min' || a === 'max' || a === 'median';
  const newAccs = () => o.values.map(() => new Acc());
  const groups = new Map<string, { keys: CellValue[]; cells: Map<string, Acc[]> }>();
  const pivotKeys = new Map<string, CellValue>();
  let nonNumeric = 0;

  for (const row of t.rows) {
    const keys = gIdx.map((i) => row[i]);
    const gk = keys.map((k) => cellToString(k).trim().toLowerCase()).join('␟');
    const pk = pIdx >= 0 ? cellToString(row[pIdx]).trim() || '(blank)' : '';
    if (pIdx >= 0 && !pivotKeys.has(pk)) pivotKeys.set(pk, row[pIdx]);
    let g = groups.get(gk);
    if (!g) { g = { keys: keys.map((k) => (isBlank(k) ? '(blank)' : k)), cells: new Map() }; groups.set(gk, g); }
    if (!g.cells.has(pk)) g.cells.set(pk, newAccs());
    const accs = g.cells.get(pk)!;
    o.values.forEach((v, k) => {
      const val = row[vIdx[k]];
      if (numericAgg(v.agg) && !isBlank(val) && toNumber(val) === null) nonNumeric++;
      accs[k].add(val, keepNums);
    });
  }

  const pkeys = pIdx >= 0 ? [...pivotKeys.keys()].sort((a, b) => compareValues(pivotKeys.get(a)!, pivotKeys.get(b)!)) : [''];
  const valueLabel = (k: number) => `${AGG_LABELS[o.values[k].agg]} of ${t.headers[vIdx[k]]}`;
  const empty = (a: AggFn): CellValue => (a === 'count' || a === 'countDistinct' || a === 'sum' ? 0 : null);
  const labelCols = o.groupBy.length ? o.groupBy.map((_, k) => t.headers[gIdx[k]]) : [''];

  const headers = [...labelCols];
  for (const pk of pkeys) o.values.forEach((_, k) => headers.push(pIdx < 0 ? valueLabel(k) : o.values.length > 1 ? `${pk} – ${valueLabel(k)}` : pk));
  const rowTotals = pIdx >= 0 && o.includeTotals;
  if (rowTotals) o.values.forEach((_, k) => headers.push(`Total – ${valueLabel(k)}`));

  const sorted = [...groups.values()].sort((a, b) => {
    for (let i = 0; i < a.keys.length; i++) { const c = compareValues(a.keys[i], b.keys[i]); if (c) return c; }
    return 0;
  });
  const colTotals = new Map(pkeys.map((pk) => [pk, newAccs()]));
  const grand = newAccs();
  const rows: CellValue[][] = sorted.map((g) => {
    const out: CellValue[] = o.groupBy.length ? [...g.keys] : ['All rows'];
    const rowAcc = newAccs();
    for (const pk of pkeys) {
      const accs = g.cells.get(pk);
      o.values.forEach((v, k) => {
        out.push(accs ? accs[k].result(v.agg) : empty(v.agg));
        if (accs) { rowAcc[k].absorb(accs[k]); colTotals.get(pk)![k].absorb(accs[k]); grand[k].absorb(accs[k]); }
      });
    }
    if (rowTotals) o.values.forEach((v, k) => out.push(rowAcc[k].result(v.agg)));
    return out;
  });
  if (o.includeTotals && o.groupBy.length) {
    const total: CellValue[] = labelCols.map((_, i) => (i === 0 ? 'Total' : null));
    for (const pk of pkeys) o.values.forEach((v, k) => total.push(colTotals.get(pk)![k].result(v.agg)));
    if (rowTotals) o.values.forEach((v, k) => total.push(grand[k].result(v.agg)));
    rows.push(total);
  }
  const name = o.outputName || 'Summary';
  return { table: newTable(name, headers, rows, `${t.fileName} ${name}`), nonNumeric };
}
