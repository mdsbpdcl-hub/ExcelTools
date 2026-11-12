import ExcelJS from 'exceljs';
import type { CellValue, SourceFile, Table } from '../src/core/types';
import { kindFromName } from '../src/core/parse';

let n = 0;
export function source(name: string, content: string | Uint8Array | ArrayBuffer, relPath = name): SourceFile {
  const blob = new Blob([typeof content === 'string' ? new TextEncoder().encode(content) : content as BlobPart]);
  n++;
  return { id: `f${n}`, name, relPath, size: blob.size, kind: kindFromName(name), file: blob };
}

export function table(headers: string[], rows: CellValue[][], extra: Partial<Table> = {}): Table {
  n++;
  return {
    id: `t${n}`, fileId: extra.fileId ?? `file${n}`, fileName: 'test.xlsx', relPath: 'test.xlsx', sheetName: 'Sheet1',
    headers: [...headers], rows: rows.map((r) => [...r]), rowNums: rows.map((_, i) => i + 2),
    preamble: [], origin: { r: 0, c: 0 }, headerRow: 1, hidden: false, merges: [], structural: false, ...extra,
  };
}

/** Builds an .xlsx in memory: a title row, a header row, data, a formula column, a merged title and styles. */
export async function sampleWorkbook(): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Bills');
  ws.getCell('A1').value = 'Monthly bills — Patna Circle';
  ws.mergeCells('A1:D1');
  ws.addRow(['Consumer No', 'Name', 'Units', 'Amount']);
  ws.getRow(2).font = { bold: true, color: { argb: 'FF0000FF' } };
  ws.addRow(['0012345', '  Ram  Kumar ', 120, 960]);
  ws.addRow(['0012346', 'sita devi', 80, 640]);
  ws.addRow(['0012347', 'Mohan', 200, 1600]);
  ws.getCell('E2').value = 'Rate';
  for (let r = 3; r <= 5; r++) ws.getCell(`E${r}`).value = { formula: `D${r}/C${r}`, result: 8 };
  ws.getCell('B3').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCCE5FF' } };
  const ws2 = wb.addWorksheet('Lookup');
  ws2.addRow(['Code', 'Region']);
  ws2.addRow(['P', 'Patna']);
  ws2.addRow(['G', 'Gaya']);
  ws2.getCell('C2').value = { formula: '_xlfn.XLOOKUP("G",A2:A3,B2:B3)', result: 'Gaya' };
  const hidden = wb.addWorksheet('Hidden');
  hidden.state = 'hidden';
  hidden.addRow(['secret']);
  const buf = await wb.xlsx.writeBuffer();
  return buf as ArrayBuffer;
}
