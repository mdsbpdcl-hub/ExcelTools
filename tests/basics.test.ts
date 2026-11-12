import { describe, expect, it } from 'vitest';
import { cellToString, maybePlainNumber, parseDate, parseDateDetailed, toNumber, uniqueHeaders, normHeader, colLetter } from '../src/core/util';
import { NameRegistry, applyNamingRule, safeSheetName, sanitizeRelDir, sanitizeSegment } from '../src/core/filenames';
import { compileExpr, ExprError, ExprValueError } from '../src/core/expr';
import { decodeText, detectDelimiter, detectEncoding } from '../src/core/parse/csv';
import { encodeText, toDelimited } from '../src/core/export/text';
import { table } from './fixtures';

describe('values', () => {
  it('parses numbers in Indian and western formats', () => {
    expect(toNumber('1,23,456.50')).toBe(123456.5);
    expect(toNumber('₹ 2,500')).toBe(2500);
    expect(toNumber('Rs. 75')).toBe(75);
    expect(toNumber('(1,000)')).toBe(-1000);
    expect(toNumber('12.5%')).toBeCloseTo(0.125);
    expect(toNumber('12a')).toBeNull();
    expect(toNumber('')).toBeNull();
  });
  it('keeps IDs with leading zeros or many digits as text', () => {
    expect(maybePlainNumber('0012345')).toBe('0012345');
    expect(maybePlainNumber('1234567890123456')).toBe('1234567890123456');
    expect(maybePlainNumber('42')).toBe(42);
    expect(maybePlainNumber('-3.5')).toBe(-3.5);
  });
  it('reads dates day-first by default and rejects impossible dates', () => {
    expect(cellToString(parseDate('03/04/2025'))).toBe('2025-04-03');
    expect(cellToString(parseDate('03/04/2025', 'MDY'))).toBe('2025-03-04');
    expect(cellToString(parseDate('2025-12-31'))).toBe('2025-12-31');
    expect(cellToString(parseDate('5-Jan-24'))).toBe('2024-01-05');
    expect(cellToString(parseDate('Mar 7, 2023'))).toBe('2023-03-07');
    const bad = parseDateDetailed('31/02/2024');
    expect(bad.date).toBeNull();
    expect(bad.looksLikeDate).toBe(true);
    expect(parseDateDetailed('hello').looksLikeDate).toBe(false);
  });
  it('makes headers unique and names blanks', () => {
    expect(uniqueHeaders(['Name', '', 'name', 'Name'])).toEqual(['Name', 'Column 2', 'name (2)', 'Name (3)']);
    expect(normHeader(' Consumer  No. ')).toBe('consumer no');
    expect(colLetter(0)).toBe('A');
    expect(colLetter(27)).toBe('AB');
  });
});

describe('file names', () => {
  it('blocks path traversal and reserved names', () => {
    expect(sanitizeRelDir('../../etc/passwd')).toBe('etc/passwd');
    expect(sanitizeRelDir('a/../b/./c')).toBe('a/b/c');
    expect(sanitizeSegment('..')).toBe('file');
    expect(sanitizeSegment('con.txt')).toBe('_con.txt');
    expect(sanitizeSegment('a/b\\c:d*?.xlsx')).toBe('a_b_c_d__.xlsx');
  });
  it('keeps names unique and applies naming rules', () => {
    const reg = new NameRegistry();
    expect(reg.claim('success/a.xlsx')).toBe('success/a.xlsx');
    expect(reg.claim('success/A.xlsx')).toBe('success/A (2).xlsx');
    expect(applyNamingRule('{file}_{value}', { file: 'bills', value: 'Patna/East' })).toBe('bills_Patna_East');
    const used = new Set<string>();
    expect(safeSheetName('A very long sheet name that exceeds limits', used)).toHaveLength(31);
    expect(safeSheetName('Data', used)).toBe('Data');
    expect(safeSheetName('data', used)).toBe('data (2)');
  });
});

describe('expressions', () => {
  const headers = ['Units', 'Rate', 'Name', 'Bill Date'];
  const row: Record<string, unknown> = { Units: 120, Rate: '8.5', Name: 'ram kumar', 'Bill Date': '05/01/2025' };
  const get = (c: string) => (row[c] ?? null) as never;
  it('evaluates arithmetic, text and logic', () => {
    expect(compileExpr('[Units] * Rate', headers).run(get)).toBe(1020);
    expect(compileExpr('IF([Units] > 100, "High", "Low")', headers).run(get)).toBe('High');
    expect(compileExpr('PROPER([Name]) & " (" & LEN(Name) & ")"', headers).run(get)).toBe('Ram Kumar (9)');
    expect(compileExpr('YEAR([Bill Date]) * 100 + MONTH([Bill Date])', headers).run(get)).toBe(202501);
    expect(compileExpr('AND(Units >= 100, CONTAINS(Name, "KUMAR"))', headers).run(get)).toBe(true);
    expect(compileExpr('ROUND(10/3, 2)', headers).run(get)).toBe(3.33);
  });
  it('reports clear errors', () => {
    expect(() => compileExpr('[Missing] + 1', headers)).toThrow(ExprError);
    expect(() => compileExpr('FOO(1)', headers)).toThrow(/Unknown function/);
    expect(() => compileExpr('IF(1)', headers)).toThrow(/needs/);
    expect(() => compileExpr('[Name] * 2', headers).run(get)).toThrow(ExprValueError);
    expect(compileExpr('IFERROR([Name] * 2, 0)', headers).run(get)).toBe(0);
  });
});

describe('text files', () => {
  it('detects encodings', () => {
    expect(detectEncoding(new Uint8Array([0xef, 0xbb, 0xbf, 0x41])).encoding).toBe('utf-8');
    expect(detectEncoding(new Uint8Array([0xff, 0xfe, 0x41, 0])).encoding).toBe('utf-16le');
    expect(detectEncoding(new Uint8Array([0x63, 0x61, 0x66, 0xe9])).encoding).toBe('windows-1252');
    expect(decodeText(new Uint8Array([0x63, 0x61, 0x66, 0xe9]), 'windows-1252')).toBe('café');
    expect(detectEncoding(new TextEncoder().encode('नाम,राशि')).encoding).toBe('utf-8');
  });
  it('detects delimiters', () => {
    expect(detectDelimiter('a;b;c\n1;2;3\n4;5;6')).toBe(';');
    expect(detectDelimiter('a\tb\n1\t2')).toBe('\t');
    expect(detectDelimiter('"a,b";c\n"1,2";3')).toBe(';');
    expect(detectDelimiter('a|b|c\n1|2|3')).toBe('|');
  });
  it('writes CSV with quoting and encodings', () => {
    const t = table(['Name', 'Note'], [['A, B', 'say "hi"'], ['line\nbreak', null]]);
    expect(toDelimited(t, ',', 'DD-MM-YYYY')).toBe('Name,Note\r\n"A, B","say ""hi"""\r\n"line\nbreak",\r\n');
    const enc = encodeText('€ café नाम', 'windows-1252');
    expect(enc.lost).toBe(3);
    expect(encodeText('a', 'utf-8-bom').bytes[0]).toBe(0xef);
  });
});
