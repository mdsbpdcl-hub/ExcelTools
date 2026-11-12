import { describe, expect, it } from 'vitest';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import * as XLSX from 'xlsx';
import { defaultWorkflow } from '../src/core/workflow';
import { buildTables, inspectFile, readRaw } from '../src/core/parse';
import { checkFormula, recalculateWorkbook } from '../src/core/formulas';
import { buildXlsx, patchXlsx } from '../src/core/export/xlsx';
import { toJson, toXml, toHtml } from '../src/core/export/text';
import { tablesToPdf } from '../src/core/export/pdf';
import { previewBatch, runBatch } from '../src/core/pipeline';
import type { FileInfo, SourceFile } from '../src/core/types';
import { sampleWorkbook, source, table } from './fixtures';

const cfg = () => defaultWorkflow();

async function inspectAll(srcs: SourceFile[]): Promise<FileInfo[]> {
  return Promise.all(srcs.map(async (s) => (await inspectFile(s)).info));
}

describe('reading files', () => {
  it('finds the header row below title rows, keeps formulas, skips hidden sheets', async () => {
    const src = source('bills.xlsx', await sampleWorkbook());
    const { info } = await inspectFile(src);
    expect(info.status).toBe('ready');
    expect(info.sheets.map((s) => [s.name, s.hidden, s.detectedHeaderRow])).toEqual([['Bills', false, 2], ['Lookup', false, 1], ['Hidden', true, 1]]);
    expect(info.sheets[0].headers).toEqual(['Consumer No', 'Name', 'Units', 'Amount', 'Rate']);
    const raw = await readRaw(src);
    const { tables, warnings } = buildTables(src, raw, cfg().read);
    expect(tables.map((t) => t.sheetName)).toEqual(['Bills', 'Lookup']);
    expect(warnings.some((w) => w.includes('Hidden'))).toBe(true);
    const bills = tables[0];
    expect(bills.preamble[0][0]).toBe('Monthly bills — Patna Circle');
    expect(bills.preamble[0][1]).toBe('Monthly bills — Patna Circle'); // merged cell filled
    expect(bills.rows[0]).toEqual(['0012345', '  Ram  Kumar ', 120, 960, 8]);
    expect(bills.formulas![0][4]).toBe('D3/C3');
    expect(bills.rowNums).toEqual([3, 4, 5]);
    expect(tables[1].formulas![0][2]).toContain('XLOOKUP');
  });
  it('reads CSV with encoding/delimiter detection and keeps IDs as text', async () => {
    const bytes = new Uint8Array([...new TextEncoder().encode('ID;Name;Amt\r\n007;'), 0x52, 0xe9, 0x6d, 0x79, ...new TextEncoder().encode(';1200\r\n008;Ab;12.5\r\n')]);
    const src = source('data.csv', bytes);
    const { info } = await inspectFile(src);
    expect(info.encoding).toBe('windows-1252');
    expect(info.delimiter).toBe(';');
    const raw = await readRaw(src);
    const { tables } = buildTables(src, raw, cfg().read);
    expect(tables[0].rows).toEqual([['007', 'Rémy', 1200], ['008', 'Ab', 12.5]]);
  });
  it('reports empty, corrupt, encrypted and unsupported files clearly', async () => {
    expect((await inspectFile(source('empty.csv', ''))).info).toMatchObject({ status: 'error', reason: 'The file is empty.' });
    const bad = await inspectFile(source('broken.xlsx', 'not really a workbook'));
    expect(bad.info.status).toBe('error');
    expect(bad.info.suggestion).toMatch(/save it again/i);
    const ole = new Uint8Array(4096);
    ole.set([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]);
    'EncryptedPackage'.split('').forEach((ch, i) => { ole[600 + i * 2] = ch.charCodeAt(0); });
    const enc = await inspectFile(source('locked.xlsx', ole));
    expect(enc.info).toMatchObject({ status: 'error', reason: expect.stringMatching(/password/) });
    expect((await inspectFile(source('notes.docx', 'x'))).info.status).toBe('skipped');
  });
  it('reads .xls and .xlsb files written by SheetJS', async () => {
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['A', 'B'], [1, 'x'], [2, 'y']]), 'S');
    for (const type of ['xls', 'xlsb'] as const) {
      const out = XLSX.write(wb, { type: 'array', bookType: type === 'xls' ? 'biff8' : 'xlsb' });
      const src = source(`old.${type}`, out);
      const raw = await readRaw(src);
      const { tables } = buildTables(src, raw, cfg().read);
      expect(tables[0].rows).toEqual([[1, 'x'], [2, 'y']]);
    }
  });
});

describe('formulas', () => {
  it('recalculates supported formulas and keeps stored results for unsupported ones', async () => {
    const t = table(['A', 'B', 'C'], [[2, 3, 0], [4, 5, 0]]);
    t.formulas = [[null, null, 'A2*B2'], [null, null, 'FOOBAR(A3)']];
    t.rows[1][2] = 99;
    const r = await recalculateWorkbook([t]);
    expect(t.rows[0][2]).toBe(6);
    expect(t.rows[1][2]).toBe(99);
    expect(r.recalculated).toBe(1);
    expect(r.unsupported).toBe(1);
    expect(r.issues[0].kind).toBe('unsupported-formula');
  });
  it('handles _xlfn prefixes and cross-sheet references', async () => {
    const src = source('bills.xlsx', await sampleWorkbook());
    const { tables } = buildTables(src, await readRaw(src), cfg().read);
    tables[0].rows[0][3] = 1200; // change Amount in row 3
    const r = await recalculateWorkbook(tables);
    expect(tables[0].rows[0][4]).toBe(10); // D3/C3 = 1200/120
    expect(tables[1].rows[0][2]).toBe('Gaya'); // XLOOKUP
    expect(r.unsupported).toBe(0);
  });
  it('checks formulas for the assistant', async () => {
    expect((await checkFormula('=XLOOKUP(A2,B:B,C:C)')).message).toMatch(/can recalculate/);
    const c = await checkFormula('=XMATCH(A2,B:B)');
    expect(c.functions[0]).toEqual({ name: 'XMATCH', supported: false });
    expect((await checkFormula('=SUM(A1')).valid).toBe(false);
  });
});

describe('exports', () => {
  it('writes Excel with highlights, notes, formulas, conditional formats, dropdowns and protection', async () => {
    const t = table(['Name', 'Amt', 'Double'], [['A', 10, 20], ['B', 5, 10]], { structural: false });
    t.formulas = [[null, null, 'B2*2'], [null, null, 'B3*2']];
    t.highlights = new Map([['1:1', 'FFFFF2A8']]);
    t.notes = new Map([['1:1', 'other: 7']]);
    const fmt = { ...cfg().format, enabled: true, conditional: [{ column: 'Amt', type: 'greaterThan' as const, value: '8', value2: '', color: 'FFD7D7' }], protection: { ...cfg().format.protection, enabled: true, password: 'pw', unlockedColumns: ['Name'] } };
    const res = await buildXlsx([t], { format: fmt, dropdowns: [{ column: 'Name', values: ['A', 'B'] }] });
    expect(res.formulasWritten).toBe(2);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(res.bytes.buffer as ArrayBuffer);
    const ws = wb.worksheets[0];
    expect((ws.getCell('C2').value as ExcelJS.CellFormulaValue).formula).toBe('B2*2');
    expect((ws.getCell('B3').fill as ExcelJS.FillPattern).fgColor?.argb).toBe('FFFFF2A8');
    expect(ws.getCell('B3').note).toBeTruthy();
    expect(ws.getCell('A2').protection?.locked).toBe(false);
    const xml = await (await JSZip.loadAsync(res.bytes)).file('xl/worksheets/sheet1.xml')!.async('string');
    expect(xml).toContain('<sheetProtection');
    expect(xml).toContain('<conditionalFormatting');
    expect(xml).toContain('<dataValidation');
  });
  it('splits sheets beyond Excel\'s row limit instead of dropping rows', async () => {
    const rows = Array.from({ length: 1_048_580 }, (_, i) => [i]);
    const t = table(['N'], rows, { structural: true });
    const res = await buildXlsx([t], { format: cfg().format, dropdowns: [] });
    expect(res.warnings.join(' ')).toMatch(/continues on 2 sheets/);
    const zip = await JSZip.loadAsync(res.bytes);
    expect(Object.keys(zip.files).filter((f) => f.startsWith('xl/worksheets/sheet')).length).toBe(2);
  }, 120000);
  it('patches the original workbook and keeps its formatting', async () => {
    const original = await sampleWorkbook();
    const src = source('bills.xlsx', original);
    const raw = await readRaw(src);
    const { tables } = buildTables(src, raw, { ...cfg().read, includeHidden: true });
    tables[0].rows[0][1] = 'Ram Kumar';
    tables[0].highlights = new Map([['0:1', 'FFFFF2A8']]);
    const res = await patchXlsx(original, tables, { format: cfg().format, dropdowns: [] }, false);
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(res.bytes.buffer as ArrayBuffer);
    const ws = wb.getWorksheet('Bills')!;
    expect(ws.getCell('B3').value).toBe('Ram Kumar');
    expect(ws.getRow(2).getCell(1).font?.color?.argb).toBe('FF0000FF'); // header font kept
    expect((ws.getCell('E3').value as ExcelJS.CellFormulaValue).formula).toBe('D3/C3');
    expect(ws.getCell('A1').isMerged).toBe(true);
    expect(res.warnings[0]).toMatch(/charts, pivot tables/);
  });
  it('writes JSON, XML, HTML and PDF', () => {
    const t = table(['Name', 'Date'], [['A & B', new Date(Date.UTC(2025, 0, 5))]]);
    expect(JSON.parse(toJson([t], 'records'))).toEqual([{ Name: 'A & B', Date: '2025-01-05' }]);
    expect(toXml([t], 'DD-MM-YYYY')).toContain('<cell column="Name" type="text">A &amp; B</cell>');
    expect(toHtml([t], 'Report', 'DD-MM-YYYY')).toContain('<td>05-01-2025</td>');
    const pdf = tablesToPdf([t], cfg().output.pdf, 'Report', 'DD-MM-YYYY');
    expect(new TextDecoder().decode(pdf.bytes.slice(0, 5))).toBe('%PDF-');
    const hindi = table(['नाम'], [['राम']]);
    expect(tablesToPdf([hindi], cfg().output.pdf, 'R', 'DD-MM-YYYY').warnings[0]).toMatch(/HTML output/);
  });
});

describe('end to end', () => {
  it('cleans, merges, validates and packages success/review/error with a run summary', async () => {
    const a = source('a.csv', 'ID,Name,Units\n1, ram ,10\n2,Sita,abc\n', 'Batch/North/a.csv');
    const b = source('b.csv', 'id,Customer Name,Units\n3,Mohan,30\n1,Ram,10\n', 'Batch/South/b.csv');
    const broken = source('c.xlsx', 'garbage', 'Batch/c.xlsx');
    const other = source('readme.docx', 'x', 'Batch/readme.docx');
    const srcs = [a, b, broken, other];
    const infos = await inspectAll(srcs);
    const c = cfg();
    c.clean.enabled = true;
    c.clean.caseMode = 'title';
    c.combine.mode = 'merge';
    c.combine.merge.headerMap = { 'customer name': 'Name' };
    c.dedupe = { ...c.dedupe, enabled: true, scope: 'batch', keys: ['ID'] };
    c.validate = { ...c.validate, enabled: true, rules: [{ column: 'Units', type: 'number', values: '', min: '', max: '', pattern: '', expression: '', message: '' }] };
    c.output.formats = { ...c.output.formats, xlsx: true, csv: true };
    const preview = await previewBatch(srcs, infos, c, () => {});
    expect(preview.outputs[0].sheets[0].rows).toEqual([['1', 'Ram', '10', 'Batch/North/a.csv'], ['2', 'Sita', 'Abc', 'Batch/North/a.csv'], ['3', 'Mohan', '30', 'Batch/South/b.csv']]);
    expect(preview.destructive.join(' ')).toMatch(/duplicate/);
    const events: string[] = [];
    const run = await runBatch(srcs, infos, c, (e) => { if (e.fileState) events.push(`${e.fileId}:${e.fileState}`); });
    const paths = run.files.map((f) => f.path).sort();
    expect(paths).toEqual([
      'error/Batch/c - error.txt',
      'error/Batch/c.xlsx',
      'review/Issues report.xlsx',
      'review/Merged.csv',
      'review/Merged.xlsx',
      'review/Removed duplicates.xlsx',
      'run-report.txt',
      'run-summary.json',
    ]);
    const s = run.summary;
    expect(s.totals).toMatchObject({ filesSelected: 4, failed: 1, skipped: 1, rowsRead: 4, rowsWritten: 3 });
    expect(s.files.find((f) => f.name === 'c.xlsx')!.suggestion).toBeTruthy();
    expect(s.formulas.statement).toMatch(/No formulas/);
    const zip = await JSZip.loadAsync(run.zip);
    expect(Object.keys(zip.files)).toContain('run-summary.json');
    expect(events).toContain(`${broken.id}:error`);
  });
  it('compares two workbooks and keeps the first file\'s formatting in the highlighted copy', async () => {
    const original = await sampleWorkbook();
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(original);
    wb.getWorksheet('Bills')!.getCell('D4').value = 700;
    const changed = await wb.xlsx.writeBuffer();
    const srcs = [source('v1.xlsx', original), source('v2.xlsx', changed as ArrayBuffer)];
    const c = cfg();
    c.combine.mode = 'compare';
    c.combine.compare.keys = ['Consumer No'];
    const run = await runBatch(srcs, await inspectAll(srcs), c, () => {});
    expect(run.files.map((f) => f.path).sort()).toEqual(['run-report.txt', 'run-summary.json', 'success/Comparison - change log.xlsx', 'success/v1 - Comparison (highlighted).xlsx']);
    const copy = run.files.find((f) => f.path.includes('highlighted'))!;
    const out = new ExcelJS.Workbook();
    await out.xlsx.load(copy.bytes.buffer.slice(copy.bytes.byteOffset, copy.bytes.byteOffset + copy.bytes.byteLength) as ArrayBuffer);
    const ws = out.getWorksheet('Bills')!;
    expect((ws.getCell('D4').fill as ExcelJS.FillPattern).fgColor?.argb).toBe('FFFFF2A8');
    expect((ws.getCell('B3').fill as ExcelJS.FillPattern).fgColor?.argb).toBe('FFCCE5FF'); // original fill kept
    expect(run.summary.steps['Changed cells']).toBe(1);
    expect(run.summary.formulas.statement).toMatch(/NOT recalculated/);
  });
  it('splits by value into files and writes a summary with the chosen formats', async () => {
    const src = source('sales.csv', 'Region,Amt\nNorth,10\nSouth,5\nNorth,7\n');
    const c = cfg();
    c.split = { ...c.split, enabled: true, mode: 'values', columns: ['Region'], naming: '{file}_{value}' };
    c.pivot = { ...c.pivot, enabled: true, groupBy: ['Region'], values: [{ column: 'Amt', agg: 'sum' }] };
    c.output.formats = { ...c.output.formats, xlsx: false, csv: true, json: true };
    const run = await runBatch([src], await inspectAll([src]), c, () => {});
    expect(run.files.map((f) => f.path).sort()).toEqual([
      'run-report.txt', 'run-summary.json',
      'success/sales - Summary.csv', 'success/sales - Summary.json',
      'success/sales_North.csv', 'success/sales_North.json', 'success/sales_South.csv', 'success/sales_South.json',
    ]);
    const north = run.files.find((f) => f.path === 'success/sales_North.csv')!.bytes;
    expect([...north.slice(0, 3)]).toEqual([0xef, 0xbb, 0xbf]);
    expect(new TextDecoder().decode(north)).toBe('Region,Amt\r\nNorth,10\r\nNorth,7\r\n');
  });
});

describe('split safety', () => {
  it('keeps sheets that cannot be split instead of dropping them', async () => {
    const src = source('bills.xlsx', await sampleWorkbook());
    const c = cfg();
    c.split = { ...c.split, enabled: true, mode: 'values', columns: ['Name'], naming: '{file}_{value}' };
    const run = await runBatch([src], await inspectAll([src]), c, () => {});
    const paths = run.files.map((f) => f.path);
    expect(paths).toContain('success/bills - Lookup.xlsx');
    expect(paths.filter((p) => p.startsWith('success/bills_')).length).toBe(3);
    expect(run.summary.warnings.join(' ')).toMatch(/saved without splitting/);
  });
});
