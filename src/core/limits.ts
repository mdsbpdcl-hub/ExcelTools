// Limits are shown in the app and checked before processing. Nothing is ever truncated silently:
// anything over a limit is stopped with a clear message instead.

export const LIMITS = {
  /** Largest single file the app will open. Browsers usually cannot hold much more in memory. */
  maxFileBytes: 200 * 1024 * 1024,
  /** Largest total size of one batch. */
  maxBatchBytes: 800 * 1024 * 1024,
  /** Most data rows in one worksheet or text file. */
  maxRowsPerTable: 2_000_000,
  /** Excel's own limit per worksheet (1,048,576 rows including the header). Longer outputs continue on extra sheets. */
  excelMaxDataRows: 1_048_575,
  excelMaxColumns: 16_384,
  /** PDF export is meant for printable reports; larger tables should be exported to Excel/CSV. */
  maxPdfRows: 50_000,
  /** Most parts one split may create (one file or sheet each). */
  maxSplitParts: 1000,
  /** Rows shown in previews. */
  previewRows: 50,
  /** Rows scanned to find the header row. */
  headerScanRows: 30,
  /** Unique values tracked per column in the data profile. */
  profileUniqueCap: 100_000,
  /** Change-log entries written by Compare before the log stops (the summary states the true total). */
  maxChangeLogEntries: 1_000_000,
} as const;
