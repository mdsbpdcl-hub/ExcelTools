import * as XLSX from 'xlsx';
import type { CellAddress, CellValue, MergeRange } from '../types';
import { excelSerialToDate } from '../util';

export interface RawSheet {
  name: string;
  hidden: boolean;
  /** Used range only, starting at `origin`. */
  grid: CellValue[][];
  formulas: (string | null)[][] | null;
  origin: CellAddress;
  merges: MergeRange[];
}

export class ReadError extends Error {
  constructor(message: string, public suggestion: string, public code: 'password' | 'corrupt' | 'unsupported' | 'empty' | 'too-large') {
    super(message);
  }
}

/** Reads .xlsx/.xlsm/.xls/.xlsb (and Excel-compatible HTML/XML) into plain grids. Macros are never run. */
export function readExcel(buf: ArrayBuffer): RawSheet[] {
  let wb: XLSX.WorkBook;
  try {
    wb = XLSX.read(new Uint8Array(buf), {
      type: 'array',
      cellFormula: true,
      cellDates: false,
      cellNF: true,
      cellStyles: false,
      cellHTML: false,
      dense: true,
      bookVBA: false,
      WTF: false,
    });
  } catch (e) {
    const msg = e instanceof Error ? e.message : String(e);
    if (/password|encrypt/i.test(msg)) {
      throw new ReadError('The file is password-protected.', 'Open it in Excel, remove the password (File › Info › Protect Workbook), save, and add it again.', 'password');
    }
    throw new ReadError(`The file could not be read (${msg}).`, 'Open it in Excel and save it again as .xlsx, then add the new copy.', 'corrupt');
  }
  const sheetsMeta = wb.Workbook?.Sheets ?? [];
  const date1904 = !!(wb.Workbook?.WBProps as { date1904?: boolean } | undefined)?.date1904;
  return wb.SheetNames.map((name, i) => {
    const ws = wb.Sheets[name];
    const hidden = !!sheetsMeta[i]?.Hidden;
    if (!ws || !ws['!ref']) return { name, hidden, grid: [], formulas: null, origin: { r: 0, c: 0 }, merges: [] };
    const range = XLSX.utils.decode_range(ws['!ref']);
    const dense = (ws as unknown as { '!data'?: (XLSX.CellObject | undefined)[][] })['!data'];
    const rows = range.e.r - range.s.r + 1;
    const cols = range.e.c - range.s.c + 1;
    const grid: CellValue[][] = new Array(rows);
    let formulas: (string | null)[][] | null = null;
    for (let r = 0; r < rows; r++) {
      const srcRow = dense ? dense[range.s.r + r] : undefined;
      const out: CellValue[] = new Array(cols).fill(null);
      if (srcRow) {
        for (let c = 0; c < cols; c++) {
          const cell = srcRow[range.s.c + c];
          if (!cell) continue;
          out[c] = cellValue(cell, date1904);
          if (cell.f) {
            if (!formulas) formulas = Array.from({ length: rows }, () => new Array(cols).fill(null));
            formulas[r][c] = cell.f;
          }
        }
      } else if (!dense) {
        for (let c = 0; c < cols; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r: range.s.r + r, c: range.s.c + c })] as XLSX.CellObject | undefined;
          if (!cell) continue;
          out[c] = cellValue(cell, date1904);
          if (cell.f) {
            if (!formulas) formulas = Array.from({ length: rows }, () => new Array(cols).fill(null));
            formulas[r][c] = cell.f;
          }
        }
      }
      grid[r] = out;
    }
    const merges = (ws['!merges'] ?? []).map((m) => ({ s: { r: m.s.r, c: m.s.c }, e: { r: m.e.r, c: m.e.c } }));
    return { name, hidden, grid, formulas, origin: { r: range.s.r, c: range.s.c }, merges };
  });
}

function cellValue(cell: XLSX.CellObject, date1904: boolean): CellValue {
  switch (cell.t) {
    case 'n': {
      const v = cell.v as number;
      if (cell.z && typeof cell.z === 'string' && XLSX.SSF.is_date(cell.z)) return excelSerialToDate(v, date1904);
      return v;
    }
    case 'd': return cell.v instanceof Date ? cell.v : new Date(String(cell.v));
    case 'b': return !!cell.v;
    case 'e': return cell.w ?? XLSX.utils.format_cell(cell) ?? '#VALUE!';
    case 's': return cell.v === undefined || cell.v === null ? null : String(cell.v);
    case 'z': return null;
    default: return cell.v === undefined ? null : (cell.v as CellValue);
  }
}
