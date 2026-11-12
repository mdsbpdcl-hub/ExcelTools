import { jsPDF } from 'jspdf';
import { autoTable } from 'jspdf-autotable';
import type { CellValue, Table } from '../types';
import type { PdfOptions } from '../workflow';
import { LIMITS } from '../limits';
import { cellToString, formatDate, hasTime } from '../util';
import { colIndex } from '../ops/helpers';

export interface PdfResult {
  bytes: Uint8Array;
  warnings: string[];
}

// jsPDF's built-in fonts only cover Latin-1. Text in other scripts (for example Hindi) cannot be drawn correctly.
const NON_LATIN = /[^\u0000-ÿ–—‘’“”•…€₹]/;

function text(v: CellValue, dateFmt: string): string {
  if (v instanceof Date) return formatDate(v, hasTime(v) ? `${dateFmt} HH:mm` : dateFmt);
  return cellToString(v);
}

/** Splits a table into page groups when the user asks for a page break on each change of a column value. */
function groups(t: Table, breakOn: string): { title: string; rows: CellValue[][] }[] {
  const i = breakOn ? colIndex(t, breakOn) : -1;
  if (i < 0) return [{ title: '', rows: t.rows }];
  const out: { title: string; rows: CellValue[][] }[] = [];
  for (const row of t.rows) {
    const key = cellToString(row[i]) || '(blank)';
    if (!out.length || out[out.length - 1].title !== key) out.push({ title: key, rows: [] });
    out[out.length - 1].rows.push(row);
  }
  return out;
}

export function tablesToPdf(tables: Table[], o: PdfOptions, docTitle: string, dateFmt: string): PdfResult {
  const warnings: string[] = [];
  const doc = new jsPDF({ orientation: o.orientation, unit: 'mm', format: o.paper, compress: true });
  const pageW = doc.internal.pageSize.getWidth();
  let first = true;
  let nonLatin = false;
  for (const t of tables) {
    if (t.rows.length > LIMITS.maxPdfRows) {
      throw new Error(`"${t.sheetName}" has ${t.rows.length.toLocaleString('en-IN')} rows; PDF export is limited to ${LIMITS.maxPdfRows.toLocaleString('en-IN')} rows per sheet. Use Excel or CSV for this sheet, or filter/split it first.`);
    }
    const usable = pageW - o.marginMm * 2;
    // "Fit to width": shrink the font until the columns fit; never below 4pt.
    let fontSize = o.fontSize;
    if (o.scale === 'fitWidth') {
      const charsPerCol = t.headers.map((h, i) => Math.min(30, Math.max(h.length, ...t.rows.slice(0, 200).map((r) => text(r[i], dateFmt).length))));
      const totalChars = charsPerCol.reduce((s, n) => s + n + 2, 0);
      const fitted = (usable / totalChars) / 0.19; // ≈ mm per char per pt for Helvetica
      fontSize = Math.max(4, Math.min(o.fontSize, Math.floor(fitted * 10) / 10));
      if (fitted < 4) warnings.push(`"${t.sheetName}" is very wide; some columns wrap. Try landscape, A3, or fewer columns.`);
    }
    for (const g of groups(t, o.breakOnColumn)) {
      const rows = g.rows.map((r) => r.map((v) => text(v, dateFmt)));
      if (!nonLatin && (t.headers.some((h) => NON_LATIN.test(h)) || rows.some((r) => r.some((s) => NON_LATIN.test(s))))) nonLatin = true;
      const chunks: string[][][] = [];
      const per = o.rowsPerPage > 0 ? o.rowsPerPage : rows.length || 1;
      for (let i = 0; i < Math.max(rows.length, 1); i += per) chunks.push(rows.slice(i, i + per));
      for (const chunk of chunks) {
        if (!first) doc.addPage();
        first = false;
        const heading = [o.title || docTitle, t.sheetName, g.title].filter(Boolean).join(' · ');
        doc.setFontSize(11);
        doc.text(heading, o.marginMm, o.marginMm + 2);
        autoTable(doc, {
          head: [t.headers],
          body: chunk,
          startY: o.marginMm + 6,
          margin: { top: o.marginMm, right: o.marginMm, bottom: o.marginMm + 4, left: o.marginMm },
          showHead: o.repeatHeader ? 'everyPage' : 'firstPage',
          styles: { fontSize, cellPadding: 1, overflow: 'linebreak', lineColor: [229, 231, 235], lineWidth: 0.1 },
          headStyles: { fillColor: [242, 248, 244], textColor: [23, 58, 37], fontStyle: 'bold' },
          tableWidth: o.scale === 'fitWidth' ? 'auto' : 'wrap',
        });
      }
    }
  }
  const pages = doc.getNumberOfPages();
  for (let p = 1; p <= pages; p++) {
    doc.setPage(p);
    doc.setFontSize(7);
    doc.setTextColor(120);
    doc.text(`Page ${p} of ${pages}`, pageW - o.marginMm, doc.internal.pageSize.getHeight() - 4, { align: 'right' });
  }
  if (nonLatin) warnings.push('Some text uses characters outside the basic Latin set (for example Hindi). PDF fonts here cannot show these correctly — use the HTML output and the browser\'s Print › Save as PDF instead.');
  return { bytes: new Uint8Array(doc.output('arraybuffer')), warnings };
}
