// Shared data types used by every module.

/** A single cell value after parsing. Spreadsheet errors are kept as strings such as "#N/A". */
export type CellValue = string | number | boolean | Date | null;

export interface CellAddress {
  r: number; // 0-based absolute row in the source sheet
  c: number; // 0-based absolute column in the source sheet
}

export interface MergeRange {
  s: CellAddress;
  e: CellAddress;
}

/**
 * One worksheet (or one CSV file) as a header + rows grid.
 * Rows are aligned with `rowNums` (1-based row numbers in the source sheet)
 * so every warning, change log entry and duplicate report can point back to the source.
 */
export interface Table {
  id: string;
  fileId: string;
  fileName: string;
  relPath: string;
  sheetName: string;
  headers: string[];
  rows: CellValue[][];
  rowNums: number[];
  /** Formula text (without "=") aligned with rows; present only while the layout is unchanged. */
  formulas?: (string | null)[][];
  /** Rows above the header row (titles, notes), kept so that output can mirror the source layout. */
  preamble: CellValue[][];
  /** 0-based position of the first row/column of the used range in the source sheet. */
  origin: CellAddress;
  /** 1-based row number of the header row in the source sheet. */
  headerRow: number;
  hidden: boolean;
  merges: MergeRange[];
  /** True once rows/columns were added, removed or moved, so cell positions no longer match the source. */
  structural: boolean;
  /** Cells to highlight in Excel output: key "rowIndex:colIndex" (data rows, 0-based) → ARGB colour. */
  highlights?: Map<string, string>;
  /** Cell comments for Excel output, same key format as highlights. */
  notes?: Map<string, string>;
}

export type FileKind = 'csv' | 'tsv' | 'txt' | 'xlsx' | 'xlsm' | 'xls' | 'xlsb';

export interface SourceFile {
  id: string;
  name: string;
  relPath: string;
  size: number;
  kind: FileKind | null;
  file: Blob;
}

export interface SheetInfo {
  name: string;
  hidden: boolean;
  rowCount: number;
  colCount: number;
  detectedHeaderRow: number;
  /** Header row in use (chosen by the user, or detected) and the column names it gives. */
  headerRow: number;
  headers: string[];
  preview: CellValue[][]; // first rows, raw, for the header-row picker
  firstRow: number; // sheet row number (1-based) of preview[0]
  hasFormulas: boolean;
  mergeCount: number;
}

export type FileStatus = 'ready' | 'skipped' | 'error';

export interface FileInfo {
  id: string;
  name: string;
  relPath: string;
  size: number;
  kind: FileKind | null;
  status: FileStatus;
  reason?: string;
  suggestion?: string;
  sheets: SheetInfo[];
  encoding?: string;
  delimiter?: string;
  warnings: string[];
}

export type IssueKind =
  | 'blank'
  | 'invalid-date'
  | 'unexpected-value'
  | 'spreadsheet-error'
  | 'validation'
  | 'unsupported-formula';

export interface Issue {
  kind: IssueKind;
  file: string;
  sheet: string;
  row: number; // source row number (1-based), or output row when the source row is unknown
  column: string;
  value: string;
  message: string;
}

export interface ColumnProfile {
  name: string;
  type: 'number' | 'date' | 'text' | 'boolean' | 'mixed' | 'empty';
  blanks: number;
  unique: number;
  uniqueCapped: boolean;
  min?: number | string;
  max?: number | string;
  errors: number;
  sample: string[];
}

export interface TableProfile {
  file: string;
  sheet: string;
  rows: number;
  columns: number;
  columnProfiles: ColumnProfile[];
  warnings: string[];
}

export type OutputFolder = 'success' | 'review' | 'error';

/** A logical output: one workbook (one or more sheets) that is written in each chosen format. */
export interface OutputItem {
  folder: OutputFolder;
  baseName: string;
  subDir: string; // sanitized relative sub-folder (may be '')
  tables: Table[];
  /** For xlsx sources processed without layout changes: the original bytes to patch in place. */
  sourceFileId?: string;
  note?: string;
  /** data = processed data; summary = pivot output (all formats, no validation); report = review/ reports. */
  kind?: 'data' | 'summary' | 'report';
}
