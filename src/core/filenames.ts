// Output names are built only from sanitised segments, so no output can escape the package
// (no "..", no absolute paths, no drive letters, no reserved Windows names).

const RESERVED = /^(con|prn|aux|nul|com[0-9]|lpt[0-9])$/i;

export function sanitizeSegment(name: string, fallback = 'file'): string {
  let s = (name ?? '')
    .normalize('NFKC')
    .replace(/[\u0000-\u001F\u007F]/g, '')
    .replace(/[<>:"/\\|?*]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^\.+/, '')
    .replace(/[. ]+$/, '');
  if (!s || s === '.' || s === '..') s = fallback;
  const base = s.split('.')[0];
  if (RESERVED.test(base)) s = `_${s}`;
  if (s.length > 120) s = s.slice(0, 120).trim();
  return s;
}

/** Sanitises a relative directory such as "Region A/2025" into safe segments. */
export function sanitizeRelDir(rel: string): string {
  return rel
    .split(/[\\/]+/)
    .filter((p) => p && p !== '.' && p !== '..')
    .map((p) => sanitizeSegment(p, 'folder'))
    .join('/');
}

export function stripExtension(name: string): string {
  const i = name.lastIndexOf('.');
  return i > 0 ? name.slice(0, i) : name;
}

export function extensionOf(name: string): string {
  const i = name.lastIndexOf('.');
  return i > 0 ? name.slice(i + 1).toLowerCase() : '';
}

/**
 * Applies a naming rule such as "{file}_{value}" or "{file}_part{n}".
 * Tokens: {file} {sheet} {value} {n} {date}
 */
export function applyNamingRule(rule: string, tokens: Record<string, string | number>): string {
  const out = (rule || '{file}').replace(/\{(\w+)\}/g, (_m, k: string) =>
    tokens[k] === undefined ? '' : String(tokens[k]),
  );
  return sanitizeSegment(out, 'output');
}

/** Keeps names unique inside one package: "a.xlsx", "a (2).xlsx", ... */
export class NameRegistry {
  private used = new Set<string>();

  claim(path: string): string {
    const lower = (p: string) => p.toLowerCase();
    if (!this.used.has(lower(path))) {
      this.used.add(lower(path));
      return path;
    }
    const dot = path.lastIndexOf('.');
    const slash = path.lastIndexOf('/');
    const hasExt = dot > slash + 1;
    const stem = hasExt ? path.slice(0, dot) : path;
    const ext = hasExt ? path.slice(dot) : '';
    for (let i = 2; ; i++) {
      const candidate = `${stem} (${i})${ext}`;
      if (!this.used.has(lower(candidate))) {
        this.used.add(lower(candidate));
        return candidate;
      }
    }
  }
}

/** Excel sheet names: max 31 chars, no []:*?/\ and unique within a workbook. */
export function safeSheetName(name: string, used: Set<string>): string {
  let base = (name || 'Sheet').replace(/[[\]:*?/\\]/g, '_').replace(/^'+|'+$/g, '').trim() || 'Sheet';
  base = base.slice(0, 31);
  let candidate = base;
  for (let i = 2; used.has(candidate.toLowerCase()); i++) {
    const suffix = ` (${i})`;
    candidate = base.slice(0, 31 - suffix.length) + suffix;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}
