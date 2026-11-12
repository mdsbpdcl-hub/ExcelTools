// Excel output with ExcelJS.
//  - buildXlsx: writes a new workbook from tables (used when the layout changed or the source is not .xlsx).
//  - patchXlsx: opens the original .xlsx and changes only what the workflow changed, so fonts, colours,
//    column widths and formulas stay as they were. ExcelJS does not keep some advanced features
//    (charts, pivot tables, slicers, macros); the run summary says so whenever this path is used.

import ExcelJS from 'exceljs';
import type { CellValue, Table } from '../types';
import type { FormatOptions } from '../workflow';
import { LIMITS } from '../limits';
import { safeSheetName } from '../filenames';
import { cellToString, colLetter, hasTime, isErrorValue, toNumber } from '../util';
import { colIndex, COLORS } from '../ops/helpers';

export interface XlsxExtras {
  format: FormatOptions;
  /** Dropdown lists to add to data columns (from "Must be one of a list" validation rules). */
  dropdowns: { column: string; values: string[] }[];
  /** Optional summary sheet: key numbers ("cards") and chart pictures. */
  summary?: { title: string; cards: [string, string | number][]; images: { png: Uint8Array; title: string }[]; notes: string[] };
  /** Extra report sheets (data profile). */
  reportTables?: Table[];
}

export interface XlsxResult {
  bytes: Uint8Array;
  warnings: string[];
  formulasWritten: number;
}

type XCell = ExcelJS.CellValue;

function excelValue(v: CellValue, counters: { long: number }): XCell {
  if (v === null || v === undefined) return null;
  if (typeof v === 'string') {
    if (isErrorValue(v)) return { error: v.trim() as ExcelJS.CellErrorValue['error'] };
    if (v.length > 32767) { counters.long++; return v.slice(0, 32767); }
    return v;
  }
  if (typeof v === 'number' && !Number.isFinite(v)) return { error: '#NUM!' as ExcelJS.CellErrorValue['error'] };
  return v;
}

function dateFmt(base: string, d: Date): string {
  return hasTime(d) ? `${base} hh:mm` : base;
}

const solid = (argb: string): ExcelJS.Fill => ({ type: 'pattern', pattern: 'solid', fgColor: { argb } });

function applyConditional(ws: ExcelJS.Worksheet, t: Table, firstDataRow: number, startCol: number, fmt: FormatOptions, warnings: string[]) {
  if (!t.rows.length) return;
  const last = firstDataRow + t.rows.length - 1;
  for (const rule of fmt.conditional) {
    const i = colIndex(t, rule.column);
    if (i < 0) continue;
    const col = colLetter(startCol - 1 + i);
    const ref = `${col}${firstDataRow}:${col}${last}`;
    const color = `FF${(rule.color || 'FFD7D7').replace('#', '').toUpperCase().padStart(6, '0')}`;
    const style: Partial<ExcelJS.Style> = { fill: { type: 'pattern', pattern: 'solid', bgColor: { argb: color } } };
    const num = (s: string) => toNumber(s);
    const literal = (s: string) => (num(s) !== null ? String(num(s)) : `"${s.replace(/"/g, '""')}"`);
    let r: ExcelJS.ConditionalFormattingRule | null = null;
    switch (rule.type) {
      case 'greaterThan': r = { type: 'cellIs', operator: 'greaterThan', formulae: [literal(rule.value)], style, priority: 1 }; break;
      case 'lessThan': r = { type: 'cellIs', operator: 'lessThan', formulae: [literal(rule.value)], style, priority: 1 }; break;
      case 'equal': r = { type: 'cellIs', operator: 'equal', formulae: [literal(rule.value)], style, priority: 1 }; break;
      case 'between': r = { type: 'cellIs', operator: 'between', formulae: [literal(rule.value), literal(rule.value2)], style, priority: 1 }; break;
      case 'containsText': r = { type: 'containsText', operator: 'containsText', text: rule.value, style, priority: 1 } as ExcelJS.ConditionalFormattingRule; break;
      case 'duplicates': r = { type: 'expression', formulae: [`COUNTIF($${col}$${firstDataRow}:$${col}$${last},${col}${firstDataRow})>1`], style, priority: 1 }; break;
      case 'dataBar': r = { type: 'dataBar', cfvo: [{ type: 'min' }, { type: 'max' }], color: { argb: color }, gradient: true, priority: 1 } as unknown as ExcelJS.ConditionalFormattingRule; break;
      case 'colorScale': r = { type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }], color: [{ argb: 'FFF8696B' }, { argb: 'FFFFEB84' }, { argb: 'FF63BE7B' }], priority: 1 } as unknown as ExcelJS.ConditionalFormattingRule; break;
      case 'iconSet': r = { type: 'iconSet', iconSet: '3TrafficLights1', cfvo: [{ type: 'percent', value: 0 }, { type: 'percent', value: 33 }, { type: 'percent', value: 67 }], priority: 1 } as unknown as ExcelJS.ConditionalFormattingRule; break;
    }
    if (r) ws.addConditionalFormatting({ ref, rules: [r] });
    else warnings.push(`Conditional format "${rule.type}" is not supported and was skipped.`);
  }
}

function applyDropdowns(ws: ExcelJS.Worksheet, t: Table, firstDataRow: number, startCol: number, extras: XlsxExtras, warnings: string[]) {
  const dv = (ws as unknown as { dataValidations: { add: (addr: string, v: unknown) => void } }).dataValidations;
  for (const d of extras.dropdowns) {
    const i = colIndex(t, d.column);
    if (i < 0 || !d.values.length) continue;
    const list = d.values.map((v) => v.replace(/"/g, '""')).join(',');
    if (list.length > 250) { warnings.push(`Dropdown for "${d.column}" has too many values for an Excel in-cell list (max ~255 characters); it was not added.`); continue; }
    const col = colLetter(startCol - 1 + i);
    const last = firstDataRow + Math.max(t.rows.length, 1) - 1 + 500; // room for new rows
    dv.add(`${col}${firstDataRow}:${col}${last}`, { type: 'list', allowBlank: true, formulae: [`"${list}"`], showErrorMessage: true, errorTitle: 'Not in list', error: `Choose a value from the list for ${d.column}.` });
  }
}

async function applyProtection(ws: ExcelJS.Worksheet, t: Table, firstDataRow: number, startCol: number, fmt: FormatOptions) {
  const p = fmt.protection;
  if (!p.enabled) return;
  for (const name of p.unlockedColumns) {
    const i = colIndex(t, name);
    if (i < 0) continue;
    for (let r = 0; r < t.rows.length + 500; r++) ws.getCell(firstDataRow + r, startCol + i).protection = { locked: false };
  }
  await ws.protect(p.password, {
    selectLockedCells: true, selectUnlockedCells: true,
    sort: p.allowSort, autoFilter: p.allowFilter, formatColumns: p.allowFormatColumns, formatCells: false,
    insertRows: false, deleteRows: false,
  });
}

function applyMarks(ws: ExcelJS.Worksheet, t: Table, firstDataRow: number, startCol: number) {
  t.highlights?.forEach((argb, key) => {
    const [r, c] = key.split(':').map(Number);
    ws.getCell(firstDataRow + r, startCol + c).fill = solid(argb);
  });
  t.notes?.forEach((text, key) => {
    const [r, c] = key.split(':').map(Number);
    ws.getCell(firstDataRow + r, startCol + c).note = text;
  });
}

function autoWidths(ws: ExcelJS.Worksheet, t: Table, startCol: number) {
  t.headers.forEach((h, i) => {
    let w = h.length;
    for (let r = 0; r < Math.min(t.rows.length, 300); r++) w = Math.max(w, cellToString(t.rows[r][i]).length);
    ws.getColumn(startCol + i).width = Math.min(60, Math.max(8, w + 2));
  });
}

function writeTableSheet(wb: ExcelJS.Workbook, t: Table, name: string, extras: XlsxExtras, warnings: string[], counters: { long: number; formulas: number }, rows: CellValue[][], formulas: (string | null)[][] | undefined, rowOffsetInTable: number) {
  const fmt = extras.format;
  const ws = wb.addWorksheet(name);
  const keepLayout = !t.structural && rowOffsetInTable === 0;
  const startRow = keepLayout ? t.origin.r + 1 : 1;
  const startCol = keepLayout ? t.origin.c + 1 : 1;
  const preamble = keepLayout ? t.preamble : [];
  const headerRow = startRow + preamble.length;
  const firstDataRow = headerRow + 1;
  const dateBase = fmt.excelDateFormat || 'dd-mm-yyyy';

  if (startCol === 1 && !preamble.length && startRow === 1 && !formulas) {
    ws.addRow(t.headers);
    ws.addRows(rows.map((r) => r.map((v) => excelValue(v, counters))));
    // Date formats: one pass over the date cells only.
    rows.forEach((r, ri) => r.forEach((v, ci) => { if (v instanceof Date) ws.getCell(firstDataRow + ri, ci + 1).numFmt = dateFmt(dateBase, v); }));
  } else {
    preamble.forEach((r, ri) => r.forEach((v, ci) => { if (v !== null) ws.getCell(startRow + ri, startCol + ci).value = excelValue(v, counters); }));
    t.headers.forEach((h, ci) => { ws.getCell(headerRow, startCol + ci).value = h; });
    rows.forEach((r, ri) => {
      const xr = ws.getRow(firstDataRow + ri);
      r.forEach((v, ci) => {
        const f = formulas?.[ri]?.[ci];
        const cell = xr.getCell(startCol + ci);
        if (f) {
          cell.value = { formula: f, result: (excelValue(v, counters) ?? undefined) as ExcelJS.CellFormulaValue['result'] };
          counters.formulas++;
        } else if (v !== null) cell.value = excelValue(v, counters);
        if (v instanceof Date) cell.numFmt = dateFmt(dateBase, v);
      });
    });
    if (keepLayout) for (const m of t.merges) {
      try { ws.mergeCells(m.s.r + 1, m.s.c + 1, m.e.r + 1, m.e.c + 1); } catch { /* overlapping merge: ignore */ }
    }
  }
  if (fmt.styleHeader) {
    const hr = ws.getRow(headerRow);
    t.headers.forEach((_, ci) => {
      const c = hr.getCell(startCol + ci);
      c.font = { bold: true, color: { argb: COLORS.headerText } };
      c.fill = solid(COLORS.header);
      c.border = { bottom: { style: 'thin', color: { argb: 'FFB9D9C4' } } };
    });
  }
  if (fmt.freezeHeader) ws.views = [{ state: 'frozen', ySplit: headerRow, xSplit: 0 }];
  if (fmt.autoFilter && t.headers.length) ws.autoFilter = { from: { row: headerRow, column: startCol }, to: { row: headerRow, column: startCol + t.headers.length - 1 } };
  autoWidths(ws, t, startCol);
  const view = { ...t, rows };
  if (rowOffsetInTable === 0) applyMarks(ws, t, firstDataRow, startCol);
  applyConditional(ws, view, firstDataRow, startCol, fmt, warnings);
  applyDropdowns(ws, view, firstDataRow, startCol, extras, warnings);
  return { ws, firstDataRow, startCol };
}

function addSummarySheet(wb: ExcelJS.Workbook, used: Set<string>, s: NonNullable<XlsxExtras['summary']>) {
  const ws = wb.addWorksheet(safeSheetName('Summary', used));
  ws.getColumn(1).width = 34;
  ws.getColumn(2).width = 22;
  ws.getCell('A1').value = s.title;
  ws.getCell('A1').font = { bold: true, size: 14, color: { argb: COLORS.headerText } };
  s.cards.forEach(([label, value], i) => {
    const r = 3 + i;
    ws.getCell(r, 1).value = label;
    ws.getCell(r, 2).value = value;
    ws.getCell(r, 1).fill = solid('FFF2F8F4');
    ws.getCell(r, 2).font = { bold: true };
    ws.getCell(r, 2).alignment = { horizontal: 'right' };
  });
  let row = 4 + s.cards.length;
  for (const n of s.notes) { ws.getCell(row++, 1).value = n; }
  row += 1;
  for (const img of s.images) {
    ws.getCell(row, 1).value = img.title;
    ws.getCell(row, 1).font = { bold: true };
    const id = wb.addImage({ buffer: img.png as unknown as ExcelJS.Buffer, extension: 'png' });
    ws.addImage(id, { tl: { col: 0, row }, ext: { width: 720, height: 384 } });
    row += 22;
  }
}

function addReportSheets(wb: ExcelJS.Workbook, used: Set<string>, tables: Table[], counters: { long: number }) {
  for (const t of tables) {
    const ws = wb.addWorksheet(safeSheetName(t.sheetName, used));
    ws.addRow(t.headers);
    ws.addRows(t.rows.map((r) => r.map((v) => excelValue(v, counters))));
    ws.getRow(1).font = { bold: true, color: { argb: COLORS.headerText } };
    ws.getRow(1).fill = solid(COLORS.header);
    ws.views = [{ state: 'frozen', ySplit: 1, xSplit: 0 }];
    autoWidths(ws, t, 1);
  }
}

export async function buildXlsx(tables: Table[], extras: XlsxExtras): Promise<XlsxResult> {
  const wb = new ExcelJS.Workbook();
  wb.creator = 'Sheet Desk';
  wb.created = new Date();
  const warnings: string[] = [];
  const counters = { long: 0, formulas: 0 };
  const used = new Set<string>();
  for (const t of tables) {
    if (t.headers.length > LIMITS.excelMaxColumns) throw new Error(`"${t.sheetName}" has ${t.headers.length} columns; Excel allows ${LIMITS.excelMaxColumns}.`);
    const chunk = LIMITS.excelMaxDataRows - (t.structural ? 0 : t.origin.r + t.preamble.length);
    if (t.rows.length <= chunk) {
      const { ws, firstDataRow, startCol } = writeTableSheet(wb, t, safeSheetName(t.sheetName, used), extras, warnings, counters, t.rows, t.structural ? undefined : t.formulas, 0);
      await applyProtection(ws, t, firstDataRow, startCol, extras.format);
    } else {
      const parts = Math.ceil(t.rows.length / LIMITS.excelMaxDataRows);
      warnings.push(`"${t.sheetName}" has ${t.rows.length.toLocaleString('en-IN')} rows — more than one Excel sheet can hold — so it continues on ${parts} sheets ("${t.sheetName}", "${t.sheetName} (2)", …). No rows were dropped.`);
      for (let i = 0, p = 0; i < t.rows.length; i += LIMITS.excelMaxDataRows, p++) {
        const part = { ...t, structural: true };
        const { ws, firstDataRow, startCol } = writeTableSheet(wb, part, safeSheetName(t.sheetName, used), extras, warnings, counters, t.rows.slice(i, i + LIMITS.excelMaxDataRows), undefined, i);
        await applyProtection(ws, part, firstDataRow, startCol, extras.format);
      }
    }
  }
  if (extras.summary) addSummarySheet(wb, used, extras.summary);
  if (extras.reportTables?.length) addReportSheets(wb, used, extras.reportTables, counters);
  if (counters.long) warnings.push(`${counters.long} cell(s) were longer than Excel's 32,767-character limit and were cut to that length in the Excel file (CSV/JSON outputs keep the full text).`);
  if (counters.formulas) wb.calcProperties.fullCalcOnLoad = true;
  const buf = await wb.xlsx.writeBuffer();
  return { bytes: new Uint8Array(buf as ArrayBuffer), warnings, formulasWritten: counters.formulas };
}

function excelText(v: XCell): string {
  if (v === null || v === undefined) return '';
  if (typeof v !== 'object' || v instanceof Date) return cellToString(v as CellValue);
  if ('richText' in v) return v.richText.map((x) => x.text).join('');
  if ('text' in v) return String((v as { text: unknown }).text);
  if ('result' in v) return cellToString(((v as ExcelJS.CellFormulaValue).result ?? null) as CellValue);
  if ('error' in v) return String(v.error);
  return '';
}

function sameValue(a: XCell, b: CellValue): boolean {
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();
  if ((a === null || a === undefined || a === '') && (b === null || b === '')) return true;
  if (a && typeof a === 'object' && 'error' in a) return typeof b === 'string' && b.trim() === a.error;
  if (a && typeof a === 'object' && 'richText' in a) return a.richText.map((x) => x.text).join('') === b;
  if (a && typeof a === 'object' && 'text' in a) return (a as { text: string }).text === b;
  return a === b;
}

/** Changes only what the workflow changed inside the original .xlsx, keeping its formatting. */
export async function patchXlsx(original: ArrayBuffer, tables: Table[], extras: XlsxExtras, recalculated: boolean): Promise<XlsxResult> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(original);
  const warnings: string[] = ['Original formatting was kept. Note: the Excel library used here does not keep charts, pivot tables, slicers or macros from the original file.'];
  const counters = { long: 0, formulas: 0 };
  const used = new Set(wb.worksheets.map((w) => w.name.toLowerCase()));
  const dateBase = extras.format.excelDateFormat || 'dd-mm-yyyy';
  for (const t of tables) {
    const ws = wb.getWorksheet(t.sheetName);
    if (!ws) { warnings.push(`Sheet "${t.sheetName}" was not found when re-opening the file.`); continue; }
    const headerRow = t.headerRow;
    const startCol = t.origin.c + 1;
    t.headers.forEach((h, ci) => {
      const cell = ws.getCell(headerRow, startCol + ci);
      const cur = excelText(cell.value);
      if (cur.replace(/\s+/g, ' ').trim() !== h && !/^Column \d+$/.test(h) && !/ \(\d+\)$/.test(h)) cell.value = h;
    });
    t.rows.forEach((row, ri) => {
      const xr = ws.getRow(t.rowNums[ri]);
      row.forEach((v, ci) => {
        const cell = xr.getCell(startCol + ci);
        const f = t.formulas?.[ri]?.[ci];
        const isFormula = cell.type === ExcelJS.ValueType.Formula;
        if (isFormula && f) {
          counters.formulas++;
          if (recalculated) {
            const cur = cell.value as ExcelJS.CellFormulaValue;
            cell.value = { ...cur, result: (excelValue(v, counters) ?? undefined) as ExcelJS.CellFormulaValue['result'] } as ExcelJS.CellFormulaValue;
          }
          return;
        }
        if (!isFormula && sameValue(cell.value, v)) return;
        cell.value = excelValue(v, counters);
        if (v instanceof Date && !cell.numFmt?.match(/[dmy]/i)) cell.numFmt = dateFmt(dateBase, v);
      });
    });
    const firstDataRow = t.headerRow + 1;
    applyMarks(ws, t, firstDataRow, startCol);
    applyConditional(ws, t, firstDataRow, startCol, extras.format, warnings);
    applyDropdowns(ws, t, firstDataRow, startCol, extras, warnings);
    await applyProtection(ws, t, firstDataRow, startCol, extras.format);
  }
  if (extras.summary) addSummarySheet(wb, used, extras.summary);
  if (extras.reportTables?.length) addReportSheets(wb, used, extras.reportTables, counters);
  if (counters.long) warnings.push(`${counters.long} cell(s) were longer than Excel's 32,767-character limit and were cut to that length in the Excel file.`);
  // Ask Excel to recalculate all formulas when the file is opened.
  wb.calcProperties.fullCalcOnLoad = true;
  const buf = await wb.xlsx.writeBuffer();
  return { bytes: new Uint8Array(buf as ArrayBuffer), warnings, formulasWritten: counters.formulas };
}
