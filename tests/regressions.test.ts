import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import { defaultWorkflow, upgradeWorkflow } from '../src/core/workflow';
import { buildTables, readRaw } from '../src/core/parse';
import { detectEncoding } from '../src/core/parse/csv';
import { recalculateWorkbook } from '../src/core/formulas';
import { compileExpr } from '../src/core/expr';
import { checkQuality } from '../src/core/ops/quality';
import { filterTable } from '../src/core/ops/filter';
import { dedupeTables } from '../src/core/ops/dedupe';
import { applyColumnTools } from '../src/core/ops/columns';
import { splitTable } from '../src/core/ops/split';
import { cellToString, excelSerialToDate } from '../src/core/util';
import { source, table } from './fixtures';

const cfg = defaultWorkflow();

describe('regressions found in the bug check', () => {
  it('reads dates from workbooks that use the 1904 date system', async () => {
    const wb = new ExcelJS.Workbook();
    wb.properties.date1904 = true;
    const ws = wb.addWorksheet('S');
    ws.addRow(['Date', 'Time']);
    ws.addRow([new Date(Date.UTC(2025, 0, 5)), 0.5]);
    ws.getCell('A2').numFmt = 'dd-mm-yyyy';
    ws.getCell('B2').numFmt = 'h:mm';
    const src = source('mac.xlsx', (await wb.xlsx.writeBuffer()) as ArrayBuffer);
    const { tables } = buildTables(src, await readRaw(src), cfg.read);
    expect(cellToString(tables[0].rows[0][0])).toBe('2025-01-05');
  });

  it('does not shift time-only values by a day', () => {
    expect(excelSerialToDate(0.5).toISOString()).toBe('1899-12-30T12:00:00.000Z');
    expect(cellToString(excelSerialToDate(45662))).toBe('2025-01-05');
  });

  it('detects UTF-8 when the 64 KB sample ends in the middle of a Hindi character', () => {
    const line = 'नाम,राशि\n';
    let text = '';
    while (new TextEncoder().encode(text).length < 70 * 1024) text += line;
    const bytes = new TextEncoder().encode(text);
    // Make sure the cut really falls inside a multi-byte character.
    let start = 0;
    while ((bytes[65536 + start] & 0xc0) !== 0x80) start++;
    const shifted = bytes.subarray(start + 1);
    expect(detectEncoding(shifted).encoding).toBe('utf-8');
    expect(detectEncoding(bytes).encoding).toBe('utf-8');
  });

  it('recalculation keeps text IDs and text dates as text', async () => {
    const t = table(['ID', 'Name', 'Found'], [['0012345', 'Ram', null], ['0012346', 'Sita', null], ['05/01/2025', 'x', null]]);
    t.formulas = [[null, null, 'XLOOKUP("0012346",A2:A4,B2:B4)'], [null, null, 'ISTEXT(A2)'], [null, null, 'ISTEXT(A4)']];
    t.rows[0][2] = 'old'; t.rows[1][2] = false; t.rows[2][2] = false;
    await recalculateWorkbook([t]);
    expect(t.rows.map((r) => r[2])).toEqual(['Sita', true, true]);
  });

  it('keeps flagged highlights after filtering, de-duplicating and column changes', () => {
    const rows = Array.from({ length: 12 }, (_, i) => [String(i % 10), i === 5 ? '#N/A' : i * 10, 'x']);
    const t = table(['ID', 'Amt', 'Other'], rows);
    checkQuality(t, { ...cfg.errors, enabled: true });
    expect(t.highlights?.get('5:1')).toBeTruthy();
    filterTable(t, { ...cfg.filter, enabled: true, conditions: [{ column: 'ID', op: 'notEquals', value: '0', value2: '', caseSensitive: false }] }); // drops row 0 and 10
    expect([...t.highlights!.keys()]).toEqual(['4:1']);
    dedupeTables([t], { ...cfg.dedupe, enabled: true, keys: ['ID'] }); // drops the later "1"
    expect([...t.highlights!.keys()]).toEqual(['4:1']);
    applyColumnTools(t, { ...cfg.columns, enabled: true, rename: [{ from: 'Amt', to: 'Amount' }], remove: ['Other'], order: ['Amount'] });
    expect([...t.highlights!.keys()]).toEqual(['4:0']);
    expect(t.rows[4][0]).toBe('#N/A');
    const parts = splitTable(t, { ...cfg.split, mode: 'rows', rowCount: 3 });
    expect([...parts[1].table.highlights!.keys()]).toEqual(['1:0']);
  });

  it('refuses to split into thousands of parts instead of freezing the browser', () => {
    const t = table(['ID'], Array.from({ length: 1500 }, (_, i) => [String(i)]));
    expect(() => splitTable(t, { ...cfg.split, mode: 'values', columns: ['ID'] })).toThrow(/limited to 1000/);
  });

  it('accepts underscores in column names in calculations', () => {
    expect(compileExpr('Bill_Date', ['Bill Date']).columns).toEqual(['Bill Date']);
    expect(compileExpr('Bill_Date', ['Bill_Date', 'Bill Date']).columns).toEqual(['Bill_Date']);
  });

  it('ignores null values in loaded workflow files', () => {
    const w = upgradeWorkflow({ output: { pdf: null, formats: { csv: true } }, filter: { conditions: null } });
    expect(w.output.pdf.paper).toBe('a4');
    expect(w.output.formats.csv).toBe(true);
    expect(w.filter.conditions).toEqual([]);
  });
});
