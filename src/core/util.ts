import type { CellValue } from './types';

export const ERROR_VALUE_RE =
  /^#(N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!|NULL!|SPILL!|CALC!|GETTING_DATA|FIELD!|BLOCKED!|CONNECT!|BUSY!|UNKNOWN!)$/;

let counter = 0;
export function uid(prefix = 'id'): string {
  counter += 1;
  return `${prefix}_${Date.now().toString(36)}_${counter.toString(36)}`;
}

export function isBlank(v: CellValue | undefined): boolean {
  return v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
}

export function isErrorValue(v: CellValue | undefined): boolean {
  return typeof v === 'string' && ERROR_VALUE_RE.test(v.trim());
}

const pad = (n: number, w = 2) => String(n).padStart(w, '0');
const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** All dates inside the app are treated as UTC so that output never shifts by the viewer's time zone. */
export function formatDate(d: Date, fmt = 'YYYY-MM-DD'): string {
  if (Number.isNaN(d.getTime())) return '';
  return fmt.replace(/YYYY|MMM|MM|DD|HH|mm|ss/g, (t) => {
    switch (t) {
      case 'YYYY': return String(d.getUTCFullYear());
      case 'MMM': return MONTHS[d.getUTCMonth()];
      case 'MM': return pad(d.getUTCMonth() + 1);
      case 'DD': return pad(d.getUTCDate());
      case 'HH': return pad(d.getUTCHours());
      case 'mm': return pad(d.getUTCMinutes());
      case 'ss': return pad(d.getUTCSeconds());
      default: return t;
    }
  });
}

export function hasTime(d: Date): boolean {
  return d.getUTCHours() !== 0 || d.getUTCMinutes() !== 0 || d.getUTCSeconds() !== 0;
}

export function cellToString(v: CellValue | undefined): string {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) return formatDate(v, hasTime(v) ? 'YYYY-MM-DD HH:mm:ss' : 'YYYY-MM-DD');
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return Number.isFinite(v) ? String(v) : '';
  return v;
}

/** Excel serial date → Date (UTC). Handles the 1900 leap-year quirk. */
export function excelSerialToDate(serial: number, date1904 = false): Date {
  let s = serial;
  if (date1904) s += 1462; // workbooks saved with the 1904 date system (older Mac Excel)
  else if (s >= 1 && s < 60) s += 1; // Excel wrongly treats 1900 as a leap year; time-only values (< 1) are not shifted
  return new Date(Math.round((s - 25569) * 86400000));
}

export function dateToExcelSerial(d: Date): number {
  return d.getTime() / 86400000 + 25569;
}

const PLAIN_NUMBER_RE = /^-?(0|[1-9]\d{0,14})(\.\d+)?$/;

/**
 * Conservative conversion used when reading text files: only plain numbers become numbers.
 * Values with leading zeros or more than 15 digits (IDs, account and phone numbers) stay text.
 */
export function maybePlainNumber(s: string): CellValue {
  const t = s.trim();
  if (t !== s && t === '') return s;
  if (PLAIN_NUMBER_RE.test(t) && t.replace(/[-.]/g, '').length <= 15) return Number(t);
  return s;
}

/** Lenient number parsing for calculations: handles ₹/$/€/£, Rs., thousands separators, (negatives) and %. */
export function toNumber(v: CellValue | undefined): number | null {
  if (v === null || v === undefined) return null;
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof Date) return null;
  let s = v.trim();
  if (s === '') return null;
  let neg = false;
  if (/^\(.*\)$/.test(s)) { neg = true; s = s.slice(1, -1).trim(); }
  s = s.replace(/^(rs\.?|inr|₹|\$|€|£)\s*/i, '').replace(/\s*(₹|\$|€|£)$/, '');
  let pct = false;
  if (s.endsWith('%')) { pct = true; s = s.slice(0, -1).trim(); }
  if (/^[-+]?\d{1,3}(,\d{2,3})+(\.\d+)?$/.test(s)) s = s.replace(/,/g, '');
  if (!/^[-+]?(\d+\.?\d*|\.\d+)([eE][-+]?\d+)?$/.test(s)) return null;
  let n = Number(s);
  if (!Number.isFinite(n)) return null;
  if (pct) n /= 100;
  return neg ? -n : n;
}

export type DateOrder = 'DMY' | 'MDY' | 'YMD';

function monthFromName(m: string): number {
  const i = MONTHS.findIndex((x) => x.toLowerCase() === m.slice(0, 3).toLowerCase());
  return i;
}

function makeDate(y: number, m: number, d: number, hh = 0, mi = 0, ss = 0): Date | null {
  if (y < 100) y += y < 50 ? 2000 : 1900;
  if (m < 0 || m > 11 || d < 1 || d > 31 || hh > 23 || mi > 59 || ss > 59) return null;
  const dt = new Date(Date.UTC(y, m, d, hh, mi, ss));
  if (dt.getUTCMonth() !== m || dt.getUTCDate() !== d) return null; // e.g. 31 Feb
  return dt;
}

export interface ParsedDate {
  date: Date | null;
  /** True when the text looks like a date but is not a real calendar date (e.g. 31/02/2024). */
  looksLikeDate: boolean;
}

const TIME_PART = /(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?\s*(am|pm)?)?$/i;

/** Parses common date spellings. `order` decides how 03/04/2025 is read (India default: DMY). */
export function parseDateDetailed(v: CellValue | undefined, order: DateOrder = 'DMY', allowSerial = false): ParsedDate {
  if (v instanceof Date) return { date: Number.isNaN(v.getTime()) ? null : v, looksLikeDate: true };
  if (typeof v === 'number') {
    if (allowSerial && v > 0 && v < 2958466) return { date: excelSerialToDate(v), looksLikeDate: true };
    return { date: null, looksLikeDate: false };
  }
  if (typeof v !== 'string') return { date: null, looksLikeDate: false };
  const s = v.trim();
  if (!s) return { date: null, looksLikeDate: false };
  const time = (m: RegExpMatchArray, i: number): [number, number, number] => {
    if (!m[i]) return [0, 0, 0];
    let h = Number(m[i]);
    const ap = (m[i + 3] || '').toLowerCase();
    if (ap === 'pm' && h < 12) h += 12;
    if (ap === 'am' && h === 12) h = 0;
    return [h, Number(m[i + 1]), Number(m[i + 2] || 0)];
  };
  let m = s.match(new RegExp(`^(\\d{4})[-/.](\\d{1,2})[-/.](\\d{1,2})${TIME_PART.source}`, 'i'));
  if (m) {
    const d = makeDate(+m[1], +m[2] - 1, +m[3], ...time(m, 4));
    return { date: d, looksLikeDate: true };
  }
  m = s.match(new RegExp(`^(\\d{1,2})[-/.](\\d{1,2})[-/.](\\d{2}|\\d{4})${TIME_PART.source}`, 'i'));
  if (m) {
    const a = +m[1], b = +m[2], y = +m[3];
    const d = order === 'MDY' ? makeDate(y, a - 1, b, ...time(m, 4)) : makeDate(y, b - 1, a, ...time(m, 4));
    return { date: d, looksLikeDate: true };
  }
  m = s.match(new RegExp(`^(\\d{1,2})[-\\s/.]([A-Za-z]{3,9})[-\\s/.,]+(\\d{2}|\\d{4})${TIME_PART.source}`, 'i'));
  if (m) {
    const mo = monthFromName(m[2]);
    if (mo < 0) return { date: null, looksLikeDate: false };
    return { date: makeDate(+m[3], mo, +m[1], ...time(m, 4)), looksLikeDate: true };
  }
  m = s.match(new RegExp(`^([A-Za-z]{3,9})\\.?\\s+(\\d{1,2}),?\\s+(\\d{4})${TIME_PART.source}`, 'i'));
  if (m) {
    const mo = monthFromName(m[1]);
    if (mo < 0) return { date: null, looksLikeDate: false };
    return { date: makeDate(+m[3], mo, +m[2], ...time(m, 4)), looksLikeDate: true };
  }
  return { date: null, looksLikeDate: false };
}

export function parseDate(v: CellValue | undefined, order: DateOrder = 'DMY', allowSerial = false): Date | null {
  return parseDateDetailed(v, order, allowSerial).date;
}

/** Normalised header used for matching: case, spacing and punctuation are ignored. */
export function normHeader(h: string): string {
  return h
    .normalize('NFKC')
    .toLowerCase()
    .replace(/[​-‍﻿]/g, '')
    .replace(/[^\p{L}\p{N}]+/gu, ' ')
    .trim();
}

/** Blank headers become "Column N"; repeated headers become "Name (2)", "Name (3)". */
export function uniqueHeaders(raw: string[]): string[] {
  const seen = new Map<string, number>();
  return raw.map((h, i) => {
    let name = (h ?? '').toString().replace(/\s+/g, ' ').trim();
    if (!name) name = `Column ${i + 1}`;
    const key = name.toLowerCase();
    const n = (seen.get(key) ?? 0) + 1;
    seen.set(key, n);
    return n === 1 ? name : `${name} (${n})`;
  });
}

export function colLetter(index: number): string {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function escapeXml(s: string): string {
  return s
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

export const escapeHtml = escapeXml;

export function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / 1024 / 1024).toFixed(1)} MB`;
}

/** Indian digit grouping for counts shown in the UI (1,23,456). */
export function formatCount(n: number): string {
  return n.toLocaleString('en-IN');
}

export function compareValues(a: CellValue, b: CellValue): number {
  if (isBlank(a) && isBlank(b)) return 0;
  if (isBlank(a)) return 1;
  if (isBlank(b)) return -1;
  if (a instanceof Date && b instanceof Date) return a.getTime() - b.getTime();
  const na = toNumber(a), nb = toNumber(b);
  if (na !== null && nb !== null) return na - nb;
  return cellToString(a).localeCompare(cellToString(b), undefined, { numeric: true, sensitivity: 'base' });
}

/** Lets long loops give the event loop (and the cancel button) a chance to run. */
export function yieldToEventLoop(): Promise<void> {
  return new Promise((r) => setTimeout(r, 0));
}
