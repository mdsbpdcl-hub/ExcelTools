// Formula handling.
// "Preserved" means the formula text is written back to the output workbook; Excel recalculates it on open.
// "Recalculated" means this app computed new results with HyperFormula. The two are reported separately
// and the app never says formulas were recalculated unless HyperFormula actually did it.

import type { CellValue, Issue, Table } from './types';
import { dateToExcelSerial, excelSerialToDate, isErrorValue } from './util';

type HFModule = typeof import('hyperformula');
let hfModule: Promise<HFModule> | null = null;
export function loadHyperFormula(): Promise<HFModule> {
  if (!hfModule) hfModule = import('hyperformula');
  return hfModule;
}

/** Strips the _xlfn./_xlws. prefixes Excel stores for newer functions (XLOOKUP, IFS, …). */
export function stripFutureFunctionPrefixes(f: string): string {
  return f.replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\./gi, '');
}

export function functionsUsed(formula: string): string[] {
  const clean = stripFutureFunctionPrefixes(formula).replace(/"(?:[^"]|"")*"/g, '""');
  const out = new Set<string>();
  for (const m of clean.matchAll(/([A-Za-z][A-Za-z0-9.]*)\s*\(/g)) out.add(m[1].toUpperCase());
  return [...out];
}

let registered: Set<string> | null = null;
export async function supportedFunctionNames(): Promise<Set<string>> {
  if (registered) return registered;
  const { HyperFormula } = await loadHyperFormula();
  registered = new Set(HyperFormula.getRegisteredFunctionNames('enGB'));
  return registered;
}

export interface FormulaCheck {
  valid: boolean;
  functions: { name: string; supported: boolean }[];
  message: string;
}

export async function checkFormula(formula: string): Promise<FormulaCheck> {
  const text = formula.trim().replace(/^=/, '');
  if (!text) return { valid: false, functions: [], message: 'Type a formula, for example =XLOOKUP(A2, Sheet2!A:A, Sheet2!B:B)' };
  const { HyperFormula } = await loadHyperFormula();
  const names = await supportedFunctionNames();
  const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });
  let valid = false;
  try { valid = hf.validateFormula(`=${stripFutureFunctionPrefixes(text)}`); } finally { hf.destroy(); }
  const functions = functionsUsed(text).map((name) => ({ name, supported: names.has(name) }));
  const unsupported = functions.filter((f) => !f.supported).map((f) => f.name);
  const message = !valid
    ? 'The formula has a syntax problem (check brackets, commas and quotes).'
    : unsupported.length
      ? `Excel can calculate this, but this app cannot recalculate ${unsupported.join(', ')}. The formula will be preserved and Excel will calculate it when the file is opened.`
      : 'This app can recalculate this formula.';
  return { valid, functions, message };
}

export interface RecalcResult {
  recalculated: number;
  unsupported: number;
  errors: number;
  issues: Issue[];
}

/**
 * Recalculates every formula in the tables of ONE source workbook (tables must keep their original layout).
 * Cell positions are rebuilt from origin + preamble + header so references such as Sheet2!B5 resolve correctly.
 */
export async function recalculateWorkbook(tables: Table[]): Promise<RecalcResult> {
  const res: RecalcResult = { recalculated: 0, unsupported: 0, errors: 0, issues: [] };
  const withFormulas = tables.filter((t) => t.formulas && !t.structural);
  if (!withFormulas.length) return res;
  const { HyperFormula, DetailedCellError } = await loadHyperFormula();
  const names = await supportedFunctionNames();
  const sheets: Record<string, (CellValue | string | number | boolean | null)[][]> = {};
  const dataStart = (t: Table) => t.origin.r + t.preamble.length + 1; // 0-based row of first data row
  for (const t of tables) {
    if (t.structural) continue;
    const grid: (string | number | boolean | null)[][] = [];
    const put = (r: number, c: number, v: CellValue) => {
      while (grid.length <= r) grid.push([]);
      const row = grid[r];
      while (row.length <= c) row.push(null);
      row[c] = v instanceof Date ? dateToExcelSerial(v) : v;
    };
    // Text must stay text: without the leading apostrophe the engine would turn "0012345" or
    // "05/01/2025" into numbers, which changes lookup results. Error values stay as errors.
    const text = (v: CellValue): CellValue => (typeof v === 'string' && !isErrorValue(v) ? `'${v}` : v);
    t.preamble.forEach((row, r) => row.forEach((v, c) => put(t.origin.r + r, t.origin.c + c, text(v))));
    t.headers.forEach((h, c) => put(t.origin.r + t.preamble.length, t.origin.c + c, text(h)));
    const start = dataStart(t);
    t.rows.forEach((row, r) => row.forEach((v, c) => {
      const f = t.formulas?.[r]?.[c];
      put(start + r, t.origin.c + c, f ? `=${stripFutureFunctionPrefixes(f)}` : text(v));
    }));
    sheets[t.sheetName] = grid as (string | number | boolean | null)[][];
  }
  const hf = HyperFormula.buildFromSheets(sheets as Record<string, (string | number | boolean | null)[][]>, { licenseKey: 'gpl-v3', maxRows: 2_100_000, maxColumns: 16_384 });
  try {
    for (const t of withFormulas) {
      const sheetId = hf.getSheetId(t.sheetName);
      if (sheetId === undefined) continue;
      const start = dataStart(t);
      t.formulas!.forEach((frow, r) => frow.forEach((f, c) => {
        if (!f) return;
        const v = hf.getCellValue({ sheet: sheetId, row: start + r, col: t.origin.c + c });
        const unsupported = functionsUsed(f).filter((n) => !names.has(n));
        if (v instanceof DetailedCellError) {
          if (unsupported.length) {
            res.unsupported++;
            res.issues.push({ kind: 'unsupported-formula', file: t.fileName, sheet: t.sheetName, row: t.rowNums[r], column: t.headers[c], value: `=${f}`, message: `Not recalculated: ${unsupported.join(', ')} is not supported here. The stored result was kept; Excel will recalculate on open.` });
            return; // keep Excel's cached result
          }
          const cached = t.rows[r][c];
          if (!(typeof cached === 'string' && cached.startsWith('#'))) {
            // Excel had a real result but this engine gets an error (e.g. it depends on an unsupported
            // function or on a sheet that was not included). Keep Excel's result and say so.
            res.unsupported++;
            res.issues.push({ kind: 'unsupported-formula', file: t.fileName, sheet: t.sheetName, row: t.rowNums[r], column: t.headers[c], value: `=${f}`, message: `Not recalculated (engine returned ${v.value}). The stored result was kept; Excel will recalculate on open.` });
            return;
          }
          res.errors++;
          t.rows[r][c] = v.value;
          res.recalculated++;
          return;
        }
        const old = t.rows[r][c];
        t.rows[r][c] = old instanceof Date && typeof v === 'number' ? excelSerialToDate(v) : (v as CellValue);
        res.recalculated++;
      }));
    }
  } finally {
    hf.destroy();
  }
  return res;
}

export interface FunctionDoc { name: string; category: string; syntax: string; description: string; example: string }

/** Reference shown in the Formula assistant. Support for each entry is checked against the engine at runtime. */
export const FUNCTION_REFERENCE: FunctionDoc[] = [
  { name: 'XLOOKUP', category: 'Lookup', syntax: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found])', description: 'Finds a value in one column and returns the matching value from another.', example: '=XLOOKUP(A2, Master!A:A, Master!C:C, "Not found")' },
  { name: 'VLOOKUP', category: 'Lookup', syntax: 'VLOOKUP(lookup_value, table, column_number, FALSE)', description: 'Looks down the first column of a table and returns a value from the same row.', example: '=VLOOKUP(A2, Master!A:D, 3, FALSE)' },
  { name: 'HLOOKUP', category: 'Lookup', syntax: 'HLOOKUP(lookup_value, table, row_number, FALSE)', description: 'Like VLOOKUP but searches across the first row.', example: '=HLOOKUP("Mar", A1:M5, 3, FALSE)' },
  { name: 'INDEX', category: 'Lookup', syntax: 'INDEX(range, row, [column])', description: 'Returns the value at a position in a range. Often used with MATCH.', example: '=INDEX(C:C, MATCH(A2, B:B, 0))' },
  { name: 'MATCH', category: 'Lookup', syntax: 'MATCH(value, range, 0)', description: 'Returns the position of a value in a range (0 = exact match).', example: '=MATCH("Patna", B2:B100, 0)' },
  { name: 'XMATCH', category: 'Lookup', syntax: 'XMATCH(value, range)', description: 'Newer version of MATCH.', example: '=XMATCH(A2, B:B)' },
  { name: 'FILTER', category: 'Lookup', syntax: 'FILTER(range, condition)', description: 'Returns the rows that meet a condition (dynamic array).', example: '=FILTER(A2:D100, C2:C100>1000)' },
  { name: 'IF', category: 'Logical', syntax: 'IF(condition, value_if_true, value_if_false)', description: 'Returns one value if a condition is true and another if false.', example: '=IF(C2>=1000, "High", "Normal")' },
  { name: 'IFS', category: 'Logical', syntax: 'IFS(cond1, value1, cond2, value2, …)', description: 'Checks several conditions in order.', example: '=IFS(C2>=90,"A",C2>=75,"B",TRUE,"C")' },
  { name: 'IFERROR', category: 'Logical', syntax: 'IFERROR(value, value_if_error)', description: 'Replaces errors such as #N/A with a value of your choice.', example: '=IFERROR(VLOOKUP(A2, M:N, 2, FALSE), "")' },
  { name: 'AND', category: 'Logical', syntax: 'AND(cond1, cond2, …)', description: 'TRUE when all conditions are true.', example: '=AND(B2>0, C2<>"")' },
  { name: 'OR', category: 'Logical', syntax: 'OR(cond1, cond2, …)', description: 'TRUE when any condition is true.', example: '=OR(B2="Patna", B2="Gaya")' },
  { name: 'NOT', category: 'Logical', syntax: 'NOT(condition)', description: 'Reverses TRUE/FALSE.', example: '=NOT(ISBLANK(A2))' },
  { name: 'SWITCH', category: 'Logical', syntax: 'SWITCH(value, match1, result1, …, [default])', description: 'Picks a result by matching a value.', example: '=SWITCH(B2, "N", "North", "S", "South", "Other")' },
  { name: 'SUM', category: 'Statistical', syntax: 'SUM(range)', description: 'Adds numbers.', example: '=SUM(C2:C100)' },
  { name: 'SUMIF', category: 'Statistical', syntax: 'SUMIF(range, criteria, [sum_range])', description: 'Adds numbers that meet one condition.', example: '=SUMIF(B:B, "Patna", C:C)' },
  { name: 'SUMIFS', category: 'Statistical', syntax: 'SUMIFS(sum_range, range1, criteria1, …)', description: 'Adds numbers that meet several conditions.', example: '=SUMIFS(C:C, B:B, "Patna", D:D, ">=2025-01-01")' },
  { name: 'COUNTIF', category: 'Statistical', syntax: 'COUNTIF(range, criteria)', description: 'Counts cells that meet a condition.', example: '=COUNTIF(B:B, "Pending")' },
  { name: 'COUNTIFS', category: 'Statistical', syntax: 'COUNTIFS(range1, criteria1, …)', description: 'Counts rows that meet several conditions.', example: '=COUNTIFS(B:B, "Patna", E:E, "Open")' },
  { name: 'AVERAGE', category: 'Statistical', syntax: 'AVERAGE(range)', description: 'Average of numbers.', example: '=AVERAGE(C2:C100)' },
  { name: 'AVERAGEIF', category: 'Statistical', syntax: 'AVERAGEIF(range, criteria, [average_range])', description: 'Average of numbers that meet a condition.', example: '=AVERAGEIF(B:B, "Patna", C:C)' },
  { name: 'MEDIAN', category: 'Statistical', syntax: 'MEDIAN(range)', description: 'Middle value.', example: '=MEDIAN(C2:C100)' },
  { name: 'MIN', category: 'Statistical', syntax: 'MIN(range)', description: 'Smallest number.', example: '=MIN(C2:C100)' },
  { name: 'MAX', category: 'Statistical', syntax: 'MAX(range)', description: 'Largest number.', example: '=MAX(C2:C100)' },
  { name: 'COUNTA', category: 'Statistical', syntax: 'COUNTA(range)', description: 'Counts non-blank cells.', example: '=COUNTA(A2:A100)' },
  { name: 'STDEV.S', category: 'Statistical', syntax: 'STDEV.S(range)', description: 'Sample standard deviation.', example: '=STDEV.S(C2:C100)' },
  { name: 'ROUND', category: 'Math', syntax: 'ROUND(number, digits)', description: 'Rounds to a number of digits.', example: '=ROUND(C2/1000, 2)' },
  { name: 'SUMPRODUCT', category: 'Math', syntax: 'SUMPRODUCT(range1, range2)', description: 'Multiplies ranges and adds the results.', example: '=SUMPRODUCT(C2:C10, D2:D10)' },
  { name: 'PMT', category: 'Financial', syntax: 'PMT(rate, periods, present_value)', description: 'Loan instalment per period.', example: '=PMT(9%/12, 240, -5000000)' },
  { name: 'NPV', category: 'Financial', syntax: 'NPV(rate, value1, value2, …)', description: 'Net present value of future cash flows.', example: '=NPV(10%, C2:C11)' },
  { name: 'IRR', category: 'Financial', syntax: 'IRR(values)', description: 'Internal rate of return of cash flows.', example: '=IRR(B2:B12)' },
  { name: 'FV', category: 'Financial', syntax: 'FV(rate, periods, payment, [present_value])', description: 'Future value of an investment.', example: '=FV(7%/12, 60, -10000)' },
  { name: 'PV', category: 'Financial', syntax: 'PV(rate, periods, payment)', description: 'Present value of a series of payments.', example: '=PV(8%, 10, -100000)' },
  { name: 'RATE', category: 'Financial', syntax: 'RATE(periods, payment, present_value)', description: 'Interest rate per period.', example: '=RATE(60, -12000, 500000)' },
  { name: 'XNPV', category: 'Financial', syntax: 'XNPV(rate, values, dates)', description: 'NPV for cash flows on specific dates.', example: '=XNPV(10%, B2:B10, A2:A10)' },
  { name: 'TEXT', category: 'Text', syntax: 'TEXT(value, format)', description: 'Formats a number or date as text.', example: '=TEXT(A2, "dd-mm-yyyy")' },
  { name: 'TEXTJOIN', category: 'Text', syntax: 'TEXTJOIN(separator, ignore_blanks, range)', description: 'Joins text with a separator.', example: '=TEXTJOIN(", ", TRUE, B2:D2)' },
  { name: 'CONCAT', category: 'Text', syntax: 'CONCAT(text1, text2, …)', description: 'Joins text.', example: '=CONCAT(B2, " ", C2)' },
  { name: 'LEFT', category: 'Text', syntax: 'LEFT(text, count)', description: 'First characters of text.', example: '=LEFT(A2, 4)' },
  { name: 'TRIM', category: 'Text', syntax: 'TRIM(text)', description: 'Removes extra spaces.', example: '=TRIM(B2)' },
  { name: 'DATE', category: 'Date', syntax: 'DATE(year, month, day)', description: 'Builds a date.', example: '=DATE(2025, 3, 31)' },
  { name: 'EOMONTH', category: 'Date', syntax: 'EOMONTH(date, months)', description: 'Last day of a month.', example: '=EOMONTH(A2, 0)' },
  { name: 'NETWORKDAYS', category: 'Date', syntax: 'NETWORKDAYS(start, end, [holidays])', description: 'Working days between two dates.', example: '=NETWORKDAYS(A2, B2)' },
  { name: 'DATEDIF', category: 'Date', syntax: 'DATEDIF(start, end, "d"|"m"|"y")', description: 'Difference between dates.', example: '=DATEDIF(A2, TODAY(), "m")' },
];
