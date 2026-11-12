import type { Table } from '../types';
import type { FindReplaceRule } from '../workflow';
import { cellToString, maybePlainNumber } from '../util';
import { isFormulaCell, ProcessError, resolveColumns, setCell } from './helpers';

const escapeRe = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

export function buildMatcher(rule: FindReplaceRule): RegExp {
  if (!rule.find) throw new ProcessError('A find & replace rule has nothing to find.');
  const flags = rule.caseSensitive ? 'g' : 'gi';
  try {
    if (rule.mode === 'regex') return new RegExp(rule.find, flags + 'u');
    if (rule.mode === 'exact') return new RegExp(`^${escapeRe(rule.find)}$`, flags);
    return new RegExp(escapeRe(rule.find), flags);
  } catch (e) {
    throw new ProcessError(`The pattern "${rule.find}" is not valid: ${e instanceof Error ? e.message : e}`);
  }
}

function inScope(list: string[], value: string): boolean {
  return !list.length || list.some((x) => x.toLowerCase() === value.toLowerCase());
}

/** Applies all rules in order. Formula cells are left untouched so formulas are never silently replaced. */
export function findReplaceTable(t: Table, rules: FindReplaceRule[]): { replacements: number[]; formulaCellsSkipped: number } {
  const replacements = rules.map(() => 0);
  let formulaCellsSkipped = 0;
  rules.forEach((rule, ri) => {
    if (!inScope(rule.files, t.fileName) || !inScope(rule.sheets, t.sheetName)) return;
    const re = buildMatcher(rule);
    const { idx } = resolveColumns(t, rule.columns);
    for (let r = 0; r < t.rows.length; r++) {
      for (const c of idx) {
        const v = t.rows[r][c];
        if (v === null || v instanceof Date || typeof v === 'boolean') continue;
        const s = cellToString(v);
        re.lastIndex = 0;
        if (!re.test(s)) continue;
        if (isFormulaCell(t, r, c)) { formulaCellsSkipped++; continue; }
        re.lastIndex = 0;
        const count = (s.match(re) ?? []).length;
        re.lastIndex = 0;
        // Regex mode supports $1, $2 … and $& in the replacement; other modes insert the text literally.
        const out = rule.mode === 'regex' ? s.replace(re, rule.replace) : s.replace(re, () => rule.replace);
        if (out !== s) {
          setCell(t, r, c, out === '' ? null : typeof v === 'number' ? maybePlainNumber(out) : out);
          replacements[ri] += count;
        }
      }
    }
  });
  return { replacements, formulaCellsSkipped };
}
