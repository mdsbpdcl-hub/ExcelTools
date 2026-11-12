// The complete set of user settings for one run. Everything here is plain JSON so that a
// workflow can be saved as a preset and reused for recurring tasks.

import type { DateOrder } from './util';

export interface FileReadOverride {
  /** Sheets to include (names). Undefined = all visible sheets (plus hidden ones if includeHidden). */
  sheets?: string[];
  /** Header row (1-based sheet row number) per sheet name. Undefined = auto-detect. */
  headerRow?: Record<string, number>;
  encoding?: string;
  delimiter?: string;
}

export interface ReadOptions {
  includeHidden: boolean;
  mergedCells: 'fill' | 'first';
  skipBlankRows: boolean;
  convertTextNumbers: boolean;
  files: Record<string, FileReadOverride>;
}

export type CaseMode = 'none' | 'upper' | 'lower' | 'title' | 'sentence';

export interface CleanOptions {
  enabled: boolean;
  columns: string[]; // empty = all columns
  trim: boolean;
  collapseSpaces: boolean;
  lineBreaks: 'keep' | 'space' | 'remove';
  invisible: boolean;
  caseMode: CaseMode;
  dates: { enabled: boolean; columns: string[]; inputOrder: DateOrder; outputFormat: string; asText: boolean };
  numbers: { enabled: boolean; columns: string[] };
}

export interface FindReplaceRule {
  find: string;
  replace: string;
  mode: 'partial' | 'exact' | 'regex';
  caseSensitive: boolean;
  columns: string[];
  sheets: string[];
  files: string[]; // file names; empty = all
}

export interface FindReplaceOptions {
  enabled: boolean;
  rules: FindReplaceRule[];
}

export interface SplitColumnRule { column: string; delimiter: string; into: string[]; keepOriginal: boolean }
export interface CombineColumnRule { columns: string[]; separator: string; into: string; keepOriginals: boolean; skipBlanks: boolean }
export interface DerivedColumnRule { name: string; expression: string }

export interface ColumnOptions {
  enabled: boolean;
  split: SplitColumnRule[];
  combine: CombineColumnRule[];
  derived: DerivedColumnRule[];
  rename: { from: string; to: string }[];
  remove: string[];
  /** Columns listed here come first, in this order. */
  order: string[];
  keepOnlyListed: boolean;
}

export type FilterOp =
  | 'equals' | 'notEquals' | 'contains' | 'notContains' | 'startsWith' | 'endsWith' | 'regex'
  | 'gt' | 'gte' | 'lt' | 'lte' | 'between'
  | 'dateOn' | 'dateBefore' | 'dateAfter' | 'dateBetween'
  | 'blank' | 'notBlank' | 'inList' | 'notInList' | 'expression';

export interface FilterCondition { column: string; op: FilterOp; value: string; value2: string; caseSensitive: boolean }

export interface FilterOptions {
  enabled: boolean;
  conditions: FilterCondition[];
  match: 'all' | 'any';
  saveNonMatching: boolean;
  dateOrder: DateOrder;
}

export interface IssueAction {
  action: 'leave' | 'flag' | 'replace';
  replaceWith: string;
}

/** Data-quality checks: each kind of problem can be left alone, flagged (highlighted + listed) or replaced. */
export interface ErrorOptions {
  enabled: boolean;
  columns: string[]; // empty = all
  spreadsheetErrors: IssueAction;
  blanks: IssueAction;
  invalidDates: IssueAction;
  unexpected: IssueAction;
  dateOrder: DateOrder;
}

export interface DedupeOptions {
  enabled: boolean;
  keys: string[]; // empty = whole row
  keep: 'first' | 'last' | 'mostComplete';
  scope: 'file' | 'batch';
  ignoreCase: boolean;
  trimSpaces: boolean;
  exportRemoved: boolean;
}

export type RuleType = 'required' | 'list' | 'range' | 'dateRange' | 'pattern' | 'unique' | 'expression' | 'number' | 'date';

export interface ValidationRule {
  column: string;
  type: RuleType;
  values: string;
  min: string;
  max: string;
  pattern: string;
  expression: string;
  message: string;
}

export interface ValidateOptions {
  enabled: boolean;
  rules: ValidationRule[];
  addDropdowns: boolean;
  markInvalidCells: boolean;
  dateOrder: DateOrder;
}

export interface MergeOptions {
  /** normalised header → output column name (lets the user map "Cust Name" onto "Customer Name"). */
  headerMap: Record<string, string>;
  addSourceFile: boolean;
  addSourceSheet: boolean;
  addSourceRow: boolean;
  outputName: string;
}

export interface JoinOptions {
  keys: string[];
  type: 'left' | 'inner' | 'full';
  ignoreCase: boolean;
  trimSpaces: boolean;
  outputName: string;
}

export interface CompareOptions {
  keys: string[]; // empty = compare by row position
  ignoreCase: boolean;
  trimSpaces: boolean;
  numericTolerance: number;
  outputName: string;
}

export interface CombineOptions {
  mode: 'none' | 'merge' | 'join' | 'compare';
  keepIndividual: boolean;
  merge: MergeOptions;
  join: JoinOptions;
  compare: CompareOptions;
}

export interface SplitOptions {
  enabled: boolean;
  mode: 'rows' | 'values' | 'columns';
  rowCount: number;
  columns: string[];
  columnGroups: string; // "Name, Phone; Address, City"
  repeatColumns: string[];
  naming: string;
  target: 'files' | 'sheets';
}

export type AggFn = 'sum' | 'count' | 'countDistinct' | 'avg' | 'min' | 'max' | 'median';

export interface PivotOptions {
  enabled: boolean;
  groupBy: string[];
  pivotColumn: string;
  values: { column: string; agg: AggFn }[];
  includeTotals: boolean;
  outputName: string;
}

export interface ConditionalRule {
  column: string;
  type: 'greaterThan' | 'lessThan' | 'between' | 'equal' | 'containsText' | 'dataBar' | 'colorScale' | 'iconSet' | 'duplicates';
  value: string;
  value2: string;
  color: string; // hex without #
}

export type ChartType = 'bar' | 'line' | 'scatter' | 'pie' | 'waterfall';

export interface ChartSpec {
  type: ChartType;
  category: string;
  value: string;
  agg: AggFn;
  title: string;
  maxCategories: number;
}

export interface FormatOptions {
  enabled: boolean;
  conditional: ConditionalRule[];
  charts: ChartSpec[];
  summarySheet: boolean;
  profileSheet: boolean;
  freezeHeader: boolean;
  autoFilter: boolean;
  styleHeader: boolean;
  excelDateFormat: string;
  protection: {
    enabled: boolean;
    password: string;
    unlockedColumns: string[];
    allowSort: boolean;
    allowFilter: boolean;
    allowFormatColumns: boolean;
  };
}

export interface PdfOptions {
  orientation: 'portrait' | 'landscape';
  paper: 'a4' | 'a3' | 'letter' | 'legal';
  scale: 'fitWidth' | 'actual';
  fontSize: number;
  marginMm: number;
  repeatHeader: boolean;
  breakOnColumn: string;
  rowsPerPage: number; // 0 = automatic
  title: string;
}

export interface OutputOptions {
  formats: { xlsx: boolean; csv: boolean; tsv: boolean; json: boolean; xml: boolean; html: boolean; pdf: boolean };
  keepOriginalFormatting: boolean;
  csvEncoding: 'utf-8-bom' | 'utf-8' | 'utf-16le' | 'windows-1252';
  csvDelimiter: string;
  textDateFormat: string;
  jsonStyle: 'records' | 'arrays';
  pdf: PdfOptions;
  keepFolders: boolean;
  includeFailedOriginals: boolean;
  nameSuffix: string;
  packageName: string;
}

export interface WorkflowConfig {
  version: 1;
  read: ReadOptions;
  errors: ErrorOptions;
  clean: CleanOptions;
  findReplace: FindReplaceOptions;
  columns: ColumnOptions;
  filter: FilterOptions;
  dedupe: DedupeOptions;
  validate: ValidateOptions;
  combine: CombineOptions;
  split: SplitOptions;
  pivot: PivotOptions;
  formulas: { recalculate: boolean };
  format: FormatOptions;
  output: OutputOptions;
}

export function defaultWorkflow(): WorkflowConfig {
  return {
    version: 1,
    read: { includeHidden: false, mergedCells: 'fill', skipBlankRows: false, convertTextNumbers: true, files: {} },
    errors: {
      enabled: false, columns: [], dateOrder: 'DMY',
      spreadsheetErrors: { action: 'flag', replaceWith: '' },
      blanks: { action: 'leave', replaceWith: '' },
      invalidDates: { action: 'flag', replaceWith: '' },
      unexpected: { action: 'flag', replaceWith: '' },
    },
    clean: {
      enabled: false, columns: [], trim: true, collapseSpaces: true, lineBreaks: 'space', invisible: true, caseMode: 'none',
      dates: { enabled: false, columns: [], inputOrder: 'DMY', outputFormat: 'DD-MM-YYYY', asText: false },
      numbers: { enabled: false, columns: [] },
    },
    findReplace: { enabled: false, rules: [] },
    columns: { enabled: false, split: [], combine: [], derived: [], rename: [], remove: [], order: [], keepOnlyListed: false },
    filter: { enabled: false, conditions: [], match: 'all', saveNonMatching: false, dateOrder: 'DMY' },
    dedupe: { enabled: false, keys: [], keep: 'first', scope: 'file', ignoreCase: true, trimSpaces: true, exportRemoved: true },
    validate: {
      enabled: false, rules: [], addDropdowns: false, markInvalidCells: true, dateOrder: 'DMY',
    },
    combine: {
      mode: 'none',
      keepIndividual: false,
      merge: { headerMap: {}, addSourceFile: true, addSourceSheet: false, addSourceRow: false, outputName: 'Merged' },
      join: { keys: [], type: 'left', ignoreCase: true, trimSpaces: true, outputName: 'Joined' },
      compare: { keys: [], ignoreCase: false, trimSpaces: true, numericTolerance: 0, outputName: 'Comparison' },
    },
    split: { enabled: false, mode: 'values', rowCount: 10000, columns: [], columnGroups: '', repeatColumns: [], naming: '{file}_{value}', target: 'files' },
    pivot: { enabled: false, groupBy: [], pivotColumn: '', values: [], includeTotals: true, outputName: 'Summary' },
    formulas: { recalculate: false },
    format: {
      enabled: false, conditional: [], charts: [], summarySheet: false, profileSheet: false,
      freezeHeader: true, autoFilter: true, styleHeader: true, excelDateFormat: 'dd-mm-yyyy',
      protection: { enabled: false, password: '', unlockedColumns: [], allowSort: true, allowFilter: true, allowFormatColumns: true },
    },
    output: {
      formats: { xlsx: true, csv: false, tsv: false, json: false, xml: false, html: false, pdf: false },
      keepOriginalFormatting: true,
      csvEncoding: 'utf-8-bom',
      csvDelimiter: ',',
      textDateFormat: 'DD-MM-YYYY',
      jsonStyle: 'records',
      pdf: { orientation: 'landscape', paper: 'a4', scale: 'fitWidth', fontSize: 8, marginMm: 10, repeatHeader: true, breakOnColumn: '', rowsPerPage: 0, title: '' },
      keepFolders: true,
      includeFailedOriginals: true,
      nameSuffix: '',
      packageName: 'processed',
    },
  };
}

/** Fills in any settings missing from an older or hand-edited preset. */
export function upgradeWorkflow(raw: unknown): WorkflowConfig {
  const base = defaultWorkflow();
  if (!raw || typeof raw !== 'object') return base;
  const merge = (a: unknown, b: unknown): unknown => {
    if (Array.isArray(a)) return Array.isArray(b) ? b : a;
    if (a && typeof a === 'object') {
      const out: Record<string, unknown> = { ...(a as Record<string, unknown>) };
      if (b && typeof b === 'object' && !Array.isArray(b)) {
        for (const [k, v] of Object.entries(b as Record<string, unknown>)) {
          out[k] = k in out ? merge(out[k], v) : v;
        }
      }
      return out;
    }
    return b === undefined || b === null || typeof b !== typeof a ? a : b;
  };
  const cfg = merge(base, raw) as WorkflowConfig;
  cfg.version = 1;
  return cfg;
}
