# Architecture

## Stack choice

**React + TypeScript + Vite. The app is static, its processing runs in a Web Worker, and it is hosted on GitHub Pages.**

- **Why browser-only:** the requirement is that the app is usable from any computer, that uploaded files stay private, and that it is hosted from GitHub. A static site has no server that could store files. It also costs nothing on GitHub Pages, and has no processing-time or upload limits other than the computer's memory.
- **SheetJS CE** reads `.xlsx/.xlsm/.xls/.xlsb`, as well as Excel's HTML/XML exports.
- **ExcelJS** writes `.xlsx` files with styles, conditional formatting, data validation, protection, notes and images. It also re-opens an original `.xlsx` to patch it while keeping its formatting.
- **Papa Parse** handles CSV parsing, **JSZip** builds the package, and **jsPDF + autotable** creates PDFs.
- **HyperFormula** recalculates formulas when asked. It is loaded only when needed.

## Data flow

```
 UI (React, main thread)                Worker (src/worker/processor.worker.ts)
 ───────────────────────                ─────────────────────────────────────
 Files step ── inspect ───────────────▶ parse/inspectFile → FileInfo (sheets, header row, problems)
 Review step ─ preview ───────────────▶ pipeline.previewBatch → samples, profile, issues, warnings
 Process step ─ run ──────────────────▶ pipeline.runBatch → ZIP + files + run summary
            ◀── progress events (per file + overall)
 Stop = worker.terminate() + new worker (stateless, so cancelling is always safe)
```

## Pipeline (`src/core/pipeline.ts`)

1. **Read:** `parse/readRaw` turns each file into raw grids, and `parse/buildTables` turns those into a `Table`: header, rows, source row numbers, formulas, title rows and origin. A file that fails goes to `error/`, and the batch continues.
2. **Per-table steps, in a fixed order:** data problems → clean → find & replace → column tools → filter.
3. **Duplicates:** per file or across the batch.
4. **Combine:** merge, join or compare. If this step fails, each file is saved separately and the error is reported.
5. **Split and summary.**
6. **Validation:** runs on the final data. Outputs with failures go to `review/`.
7. **Formulas:** optional recalculation for tables whose layout did not change.
8. **Export:** every output is written in every chosen format. Reports go to `review/`, and `error/` gets the reasons and copies of failed files. `run-summary.json` and `run-report.txt` are added, then everything is zipped.

`Table.structural` records whether rows or columns moved. While it is `false`, formulas and positions are still valid. The original `.xlsx` can then be patched, which keeps its formatting, and recalculation can run. When it becomes `true`, formulas are converted to their stored values and a warning is added.

## Modules

| Folder | Responsibility |
|---|---|
| `src/core/parse/` | File discovery helpers, type sniffing, encoding and delimiter detection, password and corruption detection, header-row detection, sheet → Table |
| `src/core/ops/` | One pure function per transformation: `clean`, `findReplace`, `columns`, `filter`, `dedupe`, `quality`, `validate`, `profile`, `merge`, `join`, `compare`, `split`, `pivot` |
| `src/core/expr.ts` | Safe expression language (tokeniser, parser and interpreter; no `eval`) |
| `src/core/formulas.ts` | Formula recalculation, the supported-function check and the formula reference |
| `src/core/charts.ts` | Chart data preparation and drawing to PNG (OffscreenCanvas) |
| `src/core/export/` | `xlsx` (build and patch), `text` (CSV/TSV/JSON/XML/HTML and encodings), `pdf` |
| `src/core/filenames.ts` | Filename and folder sanitising, unique names, naming rules, sheet-name rules |
| `src/core/limits.ts` | All size and row limits in one place |
| `src/core/workflow.ts` | The JSON settings model (also used for saved workflows) and its upgrade function |
| `src/worker/` | Worker entry point and message types |
| `src/ui/` | React screens (`steps/`), shared components and styles |
| `tests/` | Vitest tests for core transformations, parsing, exports and full runs |

## Security notes

- No `eval` or `new Function` anywhere in app code. Calculated columns use `expr.ts`, which interprets a parsed tree.
- Macros are never loaded (`bookVBA: false`) or run.
- Output paths are built only from sanitised segments: no `..`, no absolute paths, no reserved Windows names. *Save into a folder* writes only inside a new sub-folder of the folder the user approves.
- In the production build, the Content-Security-Policy blocks network access to any other origin (`connect-src 'self'`).
- User-supplied regular expressions run inside the worker, so a pathological pattern can always be stopped with **Stop**.
