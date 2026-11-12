import { describe, expect, it } from 'vitest';
import { defaultWorkflow } from '../src/core/workflow';
import { cleanTable } from '../src/core/ops/clean';
import { findReplaceTable } from '../src/core/ops/findReplace';
import { applyColumnTools } from '../src/core/ops/columns';
import { filterTable } from '../src/core/ops/filter';
import { dedupeTables } from '../src/core/ops/dedupe';
import { validateTable } from '../src/core/ops/validate';
import { checkQuality } from '../src/core/ops/quality';
import { mergeTables, headerGroups } from '../src/core/ops/merge';
import { joinTables } from '../src/core/ops/join';
import { compareTables } from '../src/core/ops/compare';
import { splitTable, parseColumnGroups } from '../src/core/ops/split';
import { pivotTable } from '../src/core/ops/pivot';
import { profileTable } from '../src/core/ops/profile';
import { cellToString } from '../src/core/util';
import { table } from './fixtures';

const cfg = defaultWorkflow();

describe('cleaning', () => {
  it('trims, collapses spaces, removes invisible characters and changes case', () => {
    const t = table(['Name', 'City'], [['  ram​  kumar ', 'PATNA  city'], ['line\none', null]]);
    cleanTable(t, { ...cfg.clean, enabled: true, caseMode: 'title' });
    expect(t.rows[0]).toEqual(['Ram Kumar', 'Patna City']);
    expect(t.rows[1][0]).toBe('Line One');
  });
  it('standardises dates and numbers but never touches formula cells', () => {
    const t = table(['Date', 'Amount'], [['05/01/2025', '1,200'], ['31/02/2025', 'abc'], ['07-03-2025', '5']]);
    t.formulas = [[null, null], [null, null], [null, 'B1*2']];
    const r = cleanTable(t, { ...cfg.clean, enabled: true, dates: { enabled: true, columns: ['Date'], inputOrder: 'DMY', outputFormat: 'YYYY-MM-DD', asText: true }, numbers: { enabled: true, columns: ['Amount'] } });
    expect(t.rows[0]).toEqual(['2025-01-05', 1200]);
    expect(t.rows[1][0]).toBe('31/02/2025');
    expect(r.datesUnreadable).toBe(1);
    expect(t.rows[2][1]).toBe('5'); // formula cell left alone
  });
});

describe('find and replace', () => {
  it('supports partial, exact, regex and case options', () => {
    const t = table(['A'], [['Patna East'], ['patna'], ['ID-0042'], [1500]]);
    const res = findReplaceTable(t, [
      { find: 'patna', replace: 'PAT', mode: 'exact', caseSensitive: true, columns: [], sheets: [], files: [] },
      { find: 'East', replace: 'E', mode: 'partial', caseSensitive: false, columns: [], sheets: [], files: [] },
      { find: 'ID-(\\d+)', replace: '$1', mode: 'regex', caseSensitive: false, columns: [], sheets: [], files: [] },
      { find: '15', replace: '25', mode: 'partial', caseSensitive: false, columns: [], sheets: [], files: [] },
    ]);
    expect(t.rows.map((r) => r[0])).toEqual(['Patna E', 'PAT', '0042', 2500]);
    expect(res.replacements).toEqual([1, 1, 1, 1]);
  });
});

describe('column tools', () => {
  it('splits, combines, derives, renames, removes and reorders', () => {
    const t = table(['Full Name', 'Units', 'Rate', 'Junk'], [['Ram Kumar', 10, 8, 'x'], ['Sita', 5, 8, 'y']]);
    const r = applyColumnTools(t, {
      ...cfg.columns, enabled: true,
      split: [{ column: 'Full Name', delimiter: ' ', into: ['First', 'Last'], keepOriginal: false }],
      combine: [{ columns: ['Last', 'First'], separator: ', ', into: 'Sort Name', keepOriginals: true, skipBlanks: true }],
      derived: [{ name: 'Amount', expression: '[Units] * [Rate]' }],
      rename: [{ from: 'Units', to: 'kWh' }],
      remove: ['Junk'],
      order: ['Sort Name', 'Amount'],
    });
    expect(r.warnings).toEqual([]);
    expect(t.headers).toEqual(['Sort Name', 'Amount', 'First', 'Last', 'kWh', 'Rate']);
    expect(t.rows[0]).toEqual(['Kumar, Ram', 80, 'Ram', 'Kumar', 10, 8]);
    expect(t.rows[1]).toEqual(['Sita', 40, 'Sita', null, 5, 8]);
    expect(t.structural).toBe(true);
  });
});

describe('filter', () => {
  const rows = [['Patna', 100, '01/01/2025'], ['Gaya', 50, '15/02/2025'], ['Patna', 300, '10/03/2025'], ['Arrah', null, '']];
  it('combines conditions and keeps non-matching rows separately', () => {
    const t = table(['Circle', 'Amount', 'Date'], rows);
    const r = filterTable(t, { ...cfg.filter, enabled: true, saveNonMatching: true, conditions: [
      { column: 'Circle', op: 'equals', value: 'patna', value2: '', caseSensitive: false },
      { column: 'Amount', op: 'gt', value: '150', value2: '', caseSensitive: false },
    ] });
    expect(r.matchedCount).toBe(1);
    expect(r.rest!.rows.length).toBe(3);
    expect(r.rest!.rowNums).toEqual([2, 3, 5]);
  });
  it('filters by date range, list and expression', () => {
    const t1 = table(['Circle', 'Amount', 'Date'], rows);
    filterTable(t1, { ...cfg.filter, enabled: true, conditions: [{ column: 'Date', op: 'dateBetween', value: '01-02-2025', value2: '31-03-2025', caseSensitive: false }] });
    expect(t1.rows.map((r) => r[0])).toEqual(['Gaya', 'Patna']);
    const t2 = table(['Circle', 'Amount', 'Date'], rows);
    filterTable(t2, { ...cfg.filter, enabled: true, match: 'any', conditions: [{ column: 'Circle', op: 'inList', value: 'Gaya, Arrah', value2: '', caseSensitive: false }] });
    expect(t2.rows.length).toBe(2);
    const t3 = table(['Circle', 'Amount', 'Date'], rows);
    filterTable(t3, { ...cfg.filter, enabled: true, conditions: [{ column: '', op: 'expression', value: 'AND(NOT(ISBLANK(Amount)), Amount < 200)', value2: '', caseSensitive: false }] });
    expect(t3.rows.map((r) => r[1])).toEqual([100, 50]);
  });
});

describe('duplicates', () => {
  it('keeps the most complete row and reports what was removed', () => {
    const t = table(['ID', 'Name', 'Phone'], [['1', 'Ram', null], ['2', 'Sita', '98'], ['1', 'ram ', '99'], ['3', 'Mohan', null]]);
    const r = dedupeTables([t], { ...cfg.dedupe, enabled: true, keys: ['ID'], keep: 'mostComplete' });
    expect(r.removedCount).toBe(1);
    expect(t.rows.map((x) => x[2])).toEqual(['98', '99', null]);
    expect(r.removedReport!.rows[0].slice(0, 6)).toEqual(['test.xlsx', 'Sheet1', 2, 'test.xlsx', 'Sheet1', 4]);
  });
  it('works across a batch', () => {
    const a = table(['ID', 'Name'], [['1', 'A'], ['2', 'B']], { fileName: 'a.csv' });
    const b = table(['Name', 'ID'], [['B2', '2'], ['C', '3']], { fileName: 'b.csv' });
    const r = dedupeTables([a, b], { ...cfg.dedupe, enabled: true, keys: ['ID'], scope: 'batch', keep: 'first' });
    expect(r.removedCount).toBe(1);
    expect(b.rows).toEqual([['C', '3']]);
    expect(r.removedReport!.rows[0][0]).toBe('b.csv');
  });
});

describe('validation and quality', () => {
  it('applies rules and highlights invalid cells', () => {
    const t = table(['ID', 'Status', 'Units', 'Date', 'Email'], [
      ['1', 'Open', 50, '01/01/2025', 'a@b.in'],
      ['', 'Pending', 500, '31/02/2025', 'bad'],
      ['1', 'closed', -1, '10/10/2030', 'c@d.in'],
    ]);
    const r = validateTable(t, { ...cfg.validate, enabled: true, rules: [
      { column: 'ID', type: 'required', values: '', min: '', max: '', pattern: '', expression: '', message: '' },
      { column: 'ID', type: 'unique', values: '', min: '', max: '', pattern: '', expression: '', message: '' },
      { column: 'Status', type: 'list', values: 'Open, Closed', min: '', max: '', pattern: '', expression: '', message: '' },
      { column: 'Units', type: 'range', values: '', min: '0', max: '100', pattern: '', expression: '', message: '' },
      { column: 'Date', type: 'dateRange', values: '', min: '01-01-2024', max: '31-12-2026', pattern: '', expression: '', message: '' },
      { column: 'Email', type: 'pattern', values: '', min: '', max: '', pattern: '^[^@\\s]+@[^@\\s]+$', expression: '', message: '' },
      { column: '', type: 'expression', values: '', min: '', max: '', pattern: '', expression: 'Units < 400', message: 'Units too high' },
    ] });
    expect(r.failedRows).toBe(2);
    const kinds = r.issues.map((i) => `${i.row}:${i.column}`);
    expect(kinds).toEqual(expect.arrayContaining(['3:ID', '4:ID', '3:Status', '3:Units', '4:Units', '3:Date', '4:Date', '3:Email']));
    expect(r.issues.find((i) => i.message === 'Units too high')).toBeTruthy();
    expect(t.highlights?.has('1:0')).toBe(true);
  });
  it('flags or replaces spreadsheet errors, invalid dates and unexpected values', () => {
    const t = table(['Amt', 'Date'], [[10, '01/01/2025'], ['#N/A', '31/02/2025'], ['abc', '02/01/2025'], [20, '03/01/2025'], [30, '04/01/2025'], [40, '05/01/2025'], [50, '06/01/2025'], [60, '07/01/2025'], [70, '08/01/2025'], [80, '09/01/2025'], [90, '10/01/2025']]);
    const r = checkQuality(t, { ...cfg.errors, enabled: true, spreadsheetErrors: { action: 'replace', replaceWith: '0' } });
    expect(t.rows[1][0]).toBe(0);
    expect(r.counts).toMatchObject({ spreadsheetErrors: 1, invalidDates: 1, unexpected: 1 });
    expect(r.issues.map((i) => i.kind).sort()).toEqual(['invalid-date', 'unexpected-value']);
  });
  it('profiles columns', () => {
    const t = table(['n', 'd', 't'], [[1, '01/01/2025', 'a'], [2, '02/01/2025', null], ['#REF!', '03/01/2025', 'a']]);
    const p = profileTable(t);
    expect(p.columnProfiles.map((c) => c.type)).toEqual(['number', 'date', 'text']);
    expect(p.columnProfiles[0]).toMatchObject({ errors: 1, min: 1, max: 2 });
    expect(p.columnProfiles[2]).toMatchObject({ blanks: 1, unique: 1 });
  });
});

describe('merge, join, compare', () => {
  it('merges by header with user mapping and source columns', () => {
    const a = table(['Consumer No', 'Name'], [['1', 'A']], { fileName: 'a.csv', relPath: 'x/a.csv' });
    const b = table(['consumer no.', 'Cust Name', 'Phone'], [['2', 'B', '9']], { fileName: 'b.csv', relPath: 'b.csv' });
    expect(headerGroups([a, b]).find((g) => g.key === 'consumer no')!.spellings).toEqual(['Consumer No', 'consumer no.']);
    const m = mergeTables([a, b], { ...cfg.combine.merge, headerMap: { 'cust name': 'Name' }, addSourceRow: true });
    expect(m.table.headers).toEqual(['Consumer No', 'Name', 'Phone', 'Source file', 'Source row']);
    expect(m.table.rows).toEqual([['1', 'A', null, 'x/a.csv', 2], ['2', 'B', '9', 'b.csv', 2]]);
    expect(m.partialColumns).toEqual(['Phone']);
  });
  it('joins and reports unmatched rows and duplicate keys', () => {
    const a = table(['ID', 'Name'], [['1', 'A'], ['2', 'B'], ['3', 'C']], { fileName: 'a.csv' });
    const b = table(['id', 'Name', 'Amt'], [['1', 'A1', 10], ['1', 'A2', 11], ['4', 'D', 40]], { fileName: 'b.csv' });
    const j = joinTables([a, b], { ...cfg.combine.join, keys: ['ID'], type: 'full' });
    expect(j.table.headers).toEqual(['ID', 'Name', 'Name (b)', 'Amt']);
    expect(j.table.rows).toEqual([['1', 'A', 'A1', 10], ['1', 'A', 'A2', 11], ['2', 'B', null, null], ['3', 'C', null, null], ['4', null, 'D', 40]]);
    expect(j.stats).toMatchObject({ matched: 2, unmatchedLeft: 2, unmatchedRight: 1, duplicateKeys: 1 });
    expect(j.duplicateKeys.rows[0]).toEqual(['b.csv › Sheet1', '1', 2, '2, 3']);
  });
  it('compares by key, highlights changes and logs them', () => {
    const a = table(['ID', 'Name', 'Amt'], [['1', 'A', 10], ['2', 'B', 20], ['3', 'C', 30]], { fileId: 'A', fileName: 'old.xlsx' });
    const b = table(['ID', 'Amt', 'Name', 'New'], [['2', 25, 'B', 'x'], ['1', '10', 'a', 'y'], ['9', 1, 'Z', 'z']], { fileId: 'B', fileName: 'new.xlsx' });
    const c = compareTables([a, b], { ...cfg.combine.compare, keys: ['ID'] });
    expect(c.stats).toMatchObject({ changedCells: 2, rowsOnlyInFirst: 1, rowsOnlyInOther: 1 });
    expect(a.highlights && [...a.highlights.keys()].sort()).toEqual(['0:1', '1:2']);
    const changes = c.log.rows.filter((r) => r[0] === 'Changed').map((r) => [r[5], r[7], r[8]]);
    expect(changes).toEqual([['Name', 'A', 'a'], ['Amt', 20, 25]]);
    expect(c.log.rows.some((r) => r[0] === 'Column only in other file' && r[5] === 'New')).toBe(true);
    expect(a.structural).toBe(false);
    expect(a.notes?.get('1:2')).toBe('new.xlsx: 25');
  });
});

describe('split and pivot', () => {
  const t = () => table(['Region', 'Type', 'Amt'], [['North', 'A', 10], ['South', 'A', 5], ['North', 'B', 7], ['North', 'A', '3'], ['', 'B', 'x']]);
  it('splits by value, rows and column groups', () => {
    expect(splitTable(t(), { ...cfg.split, mode: 'values', columns: ['Region'] }).map((p) => [p.label, p.table.rows.length])).toEqual([['North', 3], ['South', 1], ['(blank)', 1]]);
    expect(splitTable(t(), { ...cfg.split, mode: 'rows', rowCount: 2 }).map((p) => p.table.rows.length)).toEqual([2, 2, 1]);
    expect(parseColumnGroups('Region, Type; Amt')).toEqual([['Region', 'Type'], ['Amt']]);
    const g = splitTable(t(), { ...cfg.split, mode: 'columns', columnGroups: 'Type; Amt', repeatColumns: ['Region'] });
    expect(g.map((p) => p.table.headers)).toEqual([['Region', 'Type'], ['Region', 'Amt']]);
  });
  it('summarises with groups, a pivot column and totals', () => {
    const r = pivotTable(t(), { ...cfg.pivot, enabled: true, groupBy: ['Region'], pivotColumn: 'Type', values: [{ column: 'Amt', agg: 'sum' }] });
    expect(r.table.headers).toEqual(['Region', 'A', 'B', 'Total – Sum of Amt']);
    const rows = r.table.rows.map((x) => x.map(cellToString));
    expect(rows).toEqual([['(blank)', '0', '0', '0'], ['North', '13', '7', '20'], ['South', '5', '0', '5'], ['Total', '18', '7', '25']]);
    expect(r.nonNumeric).toBe(1);
    const s = pivotTable(t(), { ...cfg.pivot, enabled: true, groupBy: ['Type'], values: [{ column: 'Amt', agg: 'count' }, { column: 'Amt', agg: 'avg' }, { column: 'Region', agg: 'countDistinct' }] });
    expect(s.table.rows).toEqual([['A', 3, 6, 2], ['B', 2, 7, 1], ['Total', 5, 6.25, 2]]);
  });
});
