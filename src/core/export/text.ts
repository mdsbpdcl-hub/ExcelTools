// Text exports: CSV / TSV, JSON, XML and HTML.

import type { CellValue, Table } from '../types';
import type { OutputOptions } from '../workflow';
import { cellToString, escapeHtml, escapeXml, formatDate, hasTime } from '../util';

function textValue(v: CellValue, dateFmt: string): string {
  if (v instanceof Date) return formatDate(v, hasTime(v) ? `${dateFmt} HH:mm:ss` : dateFmt);
  return cellToString(v);
}

export function toDelimited(t: Table, delimiter: string, dateFmt: string): string {
  const needsQuote = new RegExp(`["\\r\\n${delimiter.replace(/[\\^\]-]/g, '\\$&')}]|^\\s|\\s$`);
  const esc = (s: string) => (needsQuote.test(s) ? `"${s.replace(/"/g, '""')}"` : s);
  const lines = [t.headers.map(esc).join(delimiter)];
  for (const row of t.rows) lines.push(row.map((v) => esc(textValue(v, dateFmt))).join(delimiter));
  return lines.join('\r\n') + '\r\n';
}

const CP1252_EXTRA: Record<number, number> = {
  0x20ac: 0x80, 0x201a: 0x82, 0x0192: 0x83, 0x201e: 0x84, 0x2026: 0x85, 0x2020: 0x86, 0x2021: 0x87, 0x02c6: 0x88,
  0x2030: 0x89, 0x0160: 0x8a, 0x2039: 0x8b, 0x0152: 0x8c, 0x017d: 0x8e, 0x2018: 0x91, 0x2019: 0x92, 0x201c: 0x93,
  0x201d: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97, 0x02dc: 0x98, 0x2122: 0x99, 0x0161: 0x9a, 0x203a: 0x9b,
  0x0153: 0x9c, 0x017e: 0x9e, 0x0178: 0x9f,
};

/** Encodes text for download. Returns how many characters could not be represented (written as "?"). */
export function encodeText(text: string, encoding: OutputOptions['csvEncoding']): { bytes: Uint8Array; lost: number } {
  if (encoding === 'utf-8') return { bytes: new TextEncoder().encode(text), lost: 0 };
  if (encoding === 'utf-8-bom') return { bytes: new TextEncoder().encode('﻿' + text), lost: 0 };
  if (encoding === 'utf-16le') {
    const out = new Uint8Array(2 + text.length * 2);
    out[0] = 0xff; out[1] = 0xfe;
    for (let i = 0; i < text.length; i++) { const c = text.charCodeAt(i); out[2 + i * 2] = c & 0xff; out[3 + i * 2] = c >> 8; }
    return { bytes: out, lost: 0 };
  }
  const out = new Uint8Array(text.length);
  let n = 0, lost = 0;
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    if (cp < 0x80 || (cp >= 0xa0 && cp <= 0xff)) out[n++] = cp;
    else if (CP1252_EXTRA[cp] !== undefined) out[n++] = CP1252_EXTRA[cp];
    else { out[n++] = 0x3f; lost++; }
  }
  return { bytes: out.subarray(0, n), lost };
}

function jsonValue(v: CellValue): string | number | boolean | null {
  if (v instanceof Date) return hasTime(v) ? v.toISOString() : formatDate(v, 'YYYY-MM-DD');
  if (typeof v === 'number' && !Number.isFinite(v)) return null;
  return v;
}

export function toJson(tables: Table[], style: OutputOptions['jsonStyle']): string {
  const one = (t: Table) =>
    style === 'arrays'
      ? { columns: t.headers, rows: t.rows.map((r) => r.map(jsonValue)) }
      : t.rows.map((r) => Object.fromEntries(t.headers.map((h, i) => [h, jsonValue(r[i])])));
  if (tables.length === 1) return JSON.stringify(one(tables[0]), null, 2);
  return JSON.stringify(Object.fromEntries(tables.map((t) => [t.sheetName, one(t)])), null, 2);
}

export function toXml(tables: Table[], dateFmt: string): string {
  const parts = ['<?xml version="1.0" encoding="UTF-8"?>', '<workbook>'];
  for (const t of tables) {
    parts.push(`  <sheet name="${escapeXml(t.sheetName)}">`);
    for (const row of t.rows) {
      parts.push('    <row>');
      row.forEach((v, i) => {
        const type = v instanceof Date ? 'date' : typeof v === 'number' ? 'number' : typeof v === 'boolean' ? 'boolean' : v === null ? 'blank' : 'text';
        parts.push(`      <cell column="${escapeXml(t.headers[i])}" type="${type}">${escapeXml(textValue(v, dateFmt))}</cell>`);
      });
      parts.push('    </row>');
    }
    parts.push('  </sheet>');
  }
  parts.push('</workbook>', '');
  return parts.join('\n');
}

/**
 * Self-contained HTML. Page size/orientation are set for printing, so "Print → Save as PDF" in the
 * browser gives a PDF that shows every script correctly (including Hindi/Devanagari).
 */
export function toHtml(tables: Table[], title: string, dateFmt: string, page: { orientation: string; paper: string } = { orientation: 'landscape', paper: 'A4' }): string {
  const body = tables.map((t) => {
    const head = t.headers.map((h) => `<th>${escapeHtml(h)}</th>`).join('');
    const rows = t.rows.map((r) => `<tr>${r.map((v) => `<td${typeof v === 'number' ? ' class="n"' : ''}>${escapeHtml(textValue(v, dateFmt))}</td>`).join('')}</tr>`).join('\n');
    return `<section><h2>${escapeHtml(t.sheetName)}</h2><p class="meta">${t.rows.length.toLocaleString('en-IN')} rows · ${t.headers.length} columns</p><table><thead><tr>${head}</tr></thead><tbody>\n${rows}\n</tbody></table></section>`;
  }).join('\n');
  return `<!DOCTYPE html>
<html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title>
<style>
@page { size: ${page.paper.toUpperCase()} ${page.orientation}; margin: 10mm; }
body { font-family: Inter, "Segoe UI", Roboto, "Noto Sans", "Noto Sans Devanagari", sans-serif; color: #1f2937; background: #f8f9fb; margin: 24px; font-size: 13px; }
h1 { font-size: 20px; margin: 0 0 16px; } h2 { font-size: 16px; margin: 24px 0 4px; } .meta { color: #6b7280; margin: 0 0 8px; }
section { background: #fff; border: 1px solid #e5e7eb; border-radius: 12px; padding: 16px; margin-bottom: 16px; overflow-x: auto; }
table { border-collapse: collapse; width: 100%; } th, td { border: 1px solid #e5e7eb; padding: 4px 8px; text-align: left; vertical-align: top; }
th { background: #f2f8f4; color: #173a25; position: sticky; top: 0; } td.n { text-align: right; font-variant-numeric: tabular-nums; }
@media print { body { background: #fff; margin: 0; } section { border: 0; padding: 0; page-break-after: always; } thead { display: table-header-group; } }
</style></head><body><h1>${escapeHtml(title)}</h1>
${body}
</body></html>
`;
}
