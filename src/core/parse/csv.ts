import Papa from 'papaparse';

export const ENCODINGS = [
  { id: 'utf-8', label: 'UTF-8 (most common)' },
  { id: 'utf-16le', label: 'UTF-16 LE (Excel "Unicode text")' },
  { id: 'utf-16be', label: 'UTF-16 BE' },
  { id: 'windows-1252', label: 'Windows-1252 / ANSI (Western)' },
  { id: 'iso-8859-1', label: 'ISO-8859-1 (Latin-1)' },
  { id: 'shift_jis', label: 'Shift-JIS (Japanese)' },
  { id: 'gb18030', label: 'GB18030 (Chinese)' },
] as const;

export const DELIMITERS = [
  { id: ',', label: 'Comma ,' },
  { id: ';', label: 'Semicolon ;' },
  { id: '\t', label: 'Tab' },
  { id: '|', label: 'Pipe |' },
  { id: '^', label: 'Caret ^' },
  { id: '~', label: 'Tilde ~' },
] as const;

export interface EncodingGuess {
  encoding: string;
  bomLength: number;
  confident: boolean;
}

/** BOM first, then strict UTF-8, then UTF-16 heuristics, then Windows-1252. */
export function detectEncoding(bytes: Uint8Array): EncodingGuess {
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) return { encoding: 'utf-8', bomLength: 3, confident: true };
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return { encoding: 'utf-16le', bomLength: 2, confident: true };
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return { encoding: 'utf-16be', bomLength: 2, confident: true };
  const sample = bytes.subarray(0, Math.min(bytes.length, 64 * 1024));
  let zeroEven = 0, zeroOdd = 0;
  for (let i = 0; i < sample.length; i++) if (sample[i] === 0) (i % 2 === 0 ? zeroEven++ : zeroOdd++);
  if (zeroOdd > sample.length * 0.3 && zeroEven < sample.length * 0.05) return { encoding: 'utf-16le', bomLength: 0, confident: false };
  if (zeroEven > sample.length * 0.3 && zeroOdd < sample.length * 0.05) return { encoding: 'utf-16be', bomLength: 0, confident: false };
  try {
    // Trim a possibly cut multi-byte character at the end of the sample.
    let end = sample.length;
    if (end === 64 * 1024) {
      while (end > 0 && (sample[end - 1] & 0xc0) === 0x80) end--;
      if (end > 0 && sample[end - 1] >= 0xc0) end--; // drop the lead byte of the cut character too
    }
    new TextDecoder('utf-8', { fatal: true }).decode(sample.subarray(0, end));
    return { encoding: 'utf-8', bomLength: 0, confident: true };
  } catch {
    return { encoding: 'windows-1252', bomLength: 0, confident: false };
  }
}

export function decodeText(bytes: Uint8Array, encoding: string, bomLength = 0): string {
  let dec: TextDecoder;
  try {
    dec = new TextDecoder(encoding);
  } catch {
    throw new Error(`This browser cannot read the "${encoding}" text encoding.`);
  }
  let text = dec.decode(bytes.subarray(bomLength));
  if (text.charCodeAt(0) === 0xfeff) text = text.slice(1);
  return text;
}

/** Finds the delimiter that splits the first lines into the most consistent number of columns. */
export function detectDelimiter(text: string, fileExt = ''): string {
  if (fileExt === 'tsv') return '\t';
  const lines = text.split(/\r\n|\n|\r/).filter((l) => l.trim() !== '').slice(0, 50);
  if (!lines.length) return ',';
  let best = ',', bestScore = -1;
  for (const d of [',', ';', '\t', '|', '^', '~']) {
    const counts = lines.map((l) => countOutsideQuotes(l, d));
    const first = counts[0];
    if (first === 0) continue;
    const consistent = counts.filter((c) => c === first).length / counts.length;
    const score = consistent * 1000 + first;
    if (score > bestScore) { bestScore = score; best = d; }
  }
  return bestScore < 0 ? (fileExt === 'csv' ? ',' : '\t') : best;
}

function countOutsideQuotes(line: string, d: string): number {
  let n = 0, q = false;
  for (const ch of line) {
    if (ch === '"') q = !q;
    else if (ch === d && !q) n++;
  }
  return n;
}

export function parseDelimited(text: string, delimiter: string): string[][] {
  const res = Papa.parse<string[]>(text, { delimiter, skipEmptyLines: false, header: false, dynamicTyping: false });
  const rows = res.data;
  // Papa returns one trailing empty row when the file ends with a newline.
  while (rows.length && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') rows.pop();
  return rows;
}

export function looksBinary(bytes: Uint8Array): boolean {
  const n = Math.min(bytes.length, 8192);
  let ctrl = 0;
  for (let i = 0; i < n; i++) {
    const b = bytes[i];
    if (b === 0 || (b < 9) || (b > 13 && b < 32 && b !== 27)) ctrl++;
  }
  return n > 0 && ctrl / n > 0.1;
}
