# Sheet Desk — Excel & CSV batch processing in the browser

Sheet Desk is a web app for recurring spreadsheet work. You can clean, check, reshape, merge, join, compare, split, summarise and convert many Excel and CSV files at once. Then you download the results as one organised ZIP.

**Your files stay on your computer.** Everything runs inside the browser tab. Sheet Desk has no server and no upload. The published page enforces this with a Content-Security-Policy that blocks it from sending data anywhere. GitHub Pages only serves the app's code.

---

## 1. Put it online (GitHub Pages) — step by step

You need a GitHub account. No software has to be installed on your computer.

1. **Create a repository.** On github.com, click **New repository**. Name it, for example, `SheetDesk`, then click **Create repository**. The repository can be public or private, but GitHub Pages on a private repository needs a paid GitHub plan.
2. **Upload the code.** On the new repository page, click **uploading an existing file**. Drag in *everything inside* the `sheet-desk` folder: `src`, `tests`, `samples`, `public`, `.github`, `package.json`, `index.html`, `vite.config.ts`, `tsconfig.json`, `README.md`, `ARCHITECTURE.md` and `.gitignore`. Click **Commit changes**.
   - The `.github` folder is hidden on Mac and Windows by default and is easy to miss. If the browser upload skips it, create the file on GitHub instead. Click **Add file › Create new file**, type `.github/workflows/deploy.yml` as the name, paste in the contents of that file, and commit.
3. **Turn on Pages.** Go to **Settings › Pages**. Under **Build and deployment › Source**, choose **GitHub Actions**.
4. **Wait about two minutes.** Open the **Actions** tab and watch **Test and deploy to GitHub Pages** run. It installs, runs the tests, builds and publishes. A green tick means it is live.
5. **Open the app** at `https://<your-account>.github.io/<repository-name>/`. For example: `https://mdsbpdcl-hub.github.io/SheetDesk/`.

The app is live only after step 4 shows a green tick. If the run fails, open it in the **Actions** tab to read the error.

### Updating the app later

Edit or upload changed files in the repository, either on github.com (**Add file › Upload files**, or the ✏️ pencil on a file) or with `git push`. Every change to the `main` branch is re-tested and re-published automatically. If the tests fail, nothing is published and the previous version stays online.

### Run it on your own computer (optional, for development)

Needs Node.js 20.19 or newer.

```bash
npm install
npm run dev        # http://localhost:5173
npm test           # 50 automated tests
npm run build      # production build in dist/
npm run preview    # serve the production build locally
```

> **If `npm install` cannot reach `cdn.sheetjs.com`:** SheetJS publishes its current, security-patched version only on its own CDN, not on npm. On networks that block that CDN, change the `xlsx` line in `package.json` to `"xlsx": "npm:@e965/xlsx@0.20.3"`. That is a community republish of the same version. Do **not** use the old `xlsx@0.18.5` from npm, because it has known security issues. GitHub's build servers can reach the CDN, so the deployment needs no change.

---

## 2. Using it

1. **Files.** Add files, or choose or drag in a folder; sub-folders are included. Each file is checked straight away. You'll see whether it is ready, damaged, password-protected or an unsupported type. Use ⚙ on a file to:
   - pick which sheets to process
   - click the row that holds the column names (for files with title or note rows above the header)
   - override the text encoding and delimiter of CSV/TXT files
2. **Tasks.** Tick what should happen, choose whether to combine files, and choose output formats. You can save the workflow for a recurring task. It is stored in this browser, or you can download it as a `.json` file.
3. **Settings.** Set the details for each ticked task. Column names found in your files are suggested as you type.
4. **Review.** Sheet Desk processes the whole batch in memory and shows you:
   - every output with sample rows
   - a data profile: types, blanks, unique values, min/max and warnings
   - the issues found and the formula status
   - anything that removes data. You must confirm this before you can continue.
5. **Process.** Shows progress for each file and for the batch, with an activity log. **Stop** cancels at once, and nothing is saved from a stopped run.
6. **Results.** Download the ZIP or single files. In Chrome and Edge, **Save into a folder…** writes the package into a *new* sub-folder of a folder you pick. Nothing outside that sub-folder is touched and no existing file is overwritten.

### What's in the ZIP

| Path | Contents |
|---|---|
| `success/` | Finished outputs. The original sub-folder structure is kept (this can be turned off). |
| `review/` | Outputs that broke a rule or contain flagged cells, plus reports: `Issues report`, `Removed duplicates`, `Unmatched rows`, `Duplicate keys` |
| `error/` | One `… - error.txt` per failed file, giving the reason and what to do, plus an unchanged copy of the file (this can be turned off) |
| `run-summary.json` | Machine-readable record of the run: settings (passwords hidden), files, sheets, row counts, step counts, formula status, outputs, warnings and errors |
| `run-report.txt` | The same information in plain language |

Source files are never changed, and nothing is written outside the downloaded package.

### Sample workflows (try them with the files in `samples/`)

| Goal | Files | Settings |
|---|---|---|
| Combine circle bills into one sheet | `patna-bills-2025.xlsx`, `gaya-bills-2025.xlsx`, `nalanda-bills.csv` | Tasks: *Clean up*, *Find data problems*, *Remove duplicates* (scope: across all files, key *Consumer No*), Combine: *Merge into one*. On the merge screen, type `Name` as the output name for *Customer Name*. |
| See what changed between two versions | `patna-bills-2025.xlsx`, `patna-bills-2025-revised.xlsx` | Combine: *Compare files*, key *Consumer No*. You get a highlighted copy of the first file and a change log. |
| One file per circle, with totals and a chart | `gaya-bills-2025.xlsx` | *Split into parts* by *Circle*; *Summary table* grouped by *Circle*, Sum of *Units*; *Excel extras*: bar chart of Units by Circle plus a Summary sheet. |
| See how a failed file is reported | add `not-a-spreadsheet.xlsx` to any batch | The file goes to `error/` with the reason and what to do. The rest of the batch continues. |

---

## 3. Features

**Combine and reshape**
- **Merge** files or sheets by header name, ignoring case, spaces and punctuation. A mapping table lets you join differently named columns, and you can add *Source file / sheet / row* columns.
- **Join** on key columns: left (like VLOOKUP), inner or full. Unmatched rows and duplicate keys are reported in `review/`.
- **Compare** files by key or by row position. You get a copy of the first file with changed cells highlighted **light yellow**, each carrying a note with the new value. You also get a change log listing change type, file, sheet, rows in both files, column, key values, old value and new value. When the first file is an `.xlsx`, its own formatting is kept in the highlighted copy.
- **Split** by row count, by the values of one or more columns, or by column groups. Parts can be saved as files or as sheets, with naming rules: `{file} {sheet} {value} {n} {date}`.
- **Summary (pivot):** group by fields, optionally spread one field across columns, and calculate sum, count, distinct count, average, min, max or median, with totals.
- **Filter:** text, number, date, list, blank, pattern and custom-expression conditions, joined with all/any. Non-matching rows can be saved separately.
- **Column tools:** rename, remove, reorder, keep only the listed columns, split, combine, and add calculated columns using a safe expression language (`[Units] * [Rate]`, `IF(...)`, `PROPER(...)`, `DAYS(...)`…).

**Clean and check**
- Trim spaces, collapse repeated spaces, handle line breaks, remove invisible characters, and change case.
- Standardise dates. You choose whether the data is written day-first or month-first; day-first is the default.
- Convert numbers stored as text, including `1,23,456`, `₹ 2,500`, `(1,000)` and `12%`.
- Find and replace with partial, whole-cell or regex matching, optionally case-sensitive, and limited to chosen columns, sheets or files.
- Duplicates within each file or across the batch, with keep-first, keep-last or keep-most-complete. Removed rows are saved with their source file and row.
- Data problems: spreadsheet errors (`#N/A`, `#VALUE!`, `#REF!`…), blank cells, invalid dates (such as 31/02/2025) and unexpected values. For each kind you can leave it, **flag** it (highlight and list it) or replace it.
- Validation rules: required, list (optionally added as an Excel dropdown), number range, date range, number, date, pattern, unique, and custom expression.
- A data profile is shown before export and can also be added as a sheet.

**Excel extras** (`.xlsx` output only)
- Conditional formatting: greater/less than, between, equal, contains text, duplicates, data bars, colour scales and traffic-light icons. These are real Excel rules.
- Charts (bar, line, pie, scatter, waterfall) and summary cards on a Summary sheet.
- Sheet protection with an optional password and unlocked columns.
- Styled and frozen header with filter buttons, and your choice of date format.

**Conversion**
- Inputs: `.xlsx`, `.xlsm`, `.xls`, `.xlsb`, `.csv`, `.tsv` and `.txt`. Encoding (UTF-8/16, Windows-1252 and more) and delimiter are detected automatically and can be overridden.
- Outputs: Excel, CSV, TSV, JSON, XML, HTML and PDF. The PDF has orientation, paper size, fit-to-width or actual size, font size, margins, rows per page, a page break on change of a column, and a repeated header.

---

## 4. Formulas: preserved vs recalculated

These are two different things, and the app and run report always say which one happened.

- **Preserved:** the formula text is written into the output `.xlsx`. The value shown with it is the result Excel last saved. The output is marked so that **Excel recalculates every formula when the file is opened**. This is the default. Formulas are preserved only while rows and columns stay in place.
- **Recalculated:** only when you tick *Recalculate formulas in this app*. Results are then computed by the HyperFormula engine, which supports about 420 Excel functions, including XLOOKUP, VLOOKUP, INDEX/MATCH, SUMIFS, IFS and PMT/NPV/IRR. Some formulas cannot be recalculated, for example XMATCH or a reference to a sheet that was not selected. Those keep Excel's saved result and are listed under *Formula not recalculated* in the issues report.
- **Converted to values:** formulas become values when the layout changes, for example after filtering, removing duplicates, column tools, merge, join or split. At that point cell references would point to the wrong cells. A warning says so.
- **ƒx Formula help** in the header lists common lookup, logical, statistical, financial, text and date functions, marks which ones this engine can recalculate, and checks any formula you type.

Calculated columns from *Column tools* are worked out by the app and saved as values, not as Excel formulas.

---

## 5. Known limitations (please read)

| Area | Limitation |
|---|---|
| Size | 200 MB per file, 800 MB per batch, 20 lakh (2,000,000) rows per sheet. Anything over a limit is stopped with a message, never cut off silently. Real limits depend on the computer's memory; with very large files, close other tabs or process fewer files at once. |
| Excel row limit | An output with more than 1,048,575 rows continues on extra sheets ("Data", "Data (2)") with a warning. No rows are dropped. CSV has no row limit. |
| Excel 32,767-character cell limit | Longer cells are cut to 32,767 characters *in the Excel output only*, with a warning. CSV, JSON and other outputs keep the full text. |
| Password-protected workbooks | Encrypted files cannot be opened in a browser. They are reported with instructions to remove the password first. |
| Protecting output | Sheet protection is supported, but it is **not encryption**: anyone can open and read the file. The app cannot add a password that is needed to open the file. |
| Keeping original formatting | Used for `.xlsx`/`.xlsm` sources when rows and columns were not moved. Charts, pivot tables, slicers and macros in the original are **not** kept by the Excel library (ExcelJS). `.xls`, `.xlsb` and CSV sources are rebuilt with a clean, consistent style. |
| Charts in output | Added as pictures, because ExcelJS cannot create native, editable Excel charts. |
| Macros | Never run. Macros in `.xlsm` files are not carried into the output (it is saved as `.xlsx`). No user scripts are run; calculated columns use a safe built-in expression language, not JavaScript. |
| PDF and non-Latin scripts | The built-in PDF fonts cover Latin text only. For Hindi/Devanagari, also choose **HTML** and use the browser's *Print › Save as PDF*. Page size and orientation are already set in the HTML. The app warns when this applies. |
| Folders | Browsers cannot open or write to a folder on your computer directly. *Choose a folder* reads copies of its files. *Save into a folder* (Chrome and Edge) writes into a new sub-folder you approve. Other browsers download a ZIP. |
| Merged cells | They are filled with the top-left value by default; you can switch this off. |
| Dates | Dates are processed without time zones, so they never shift by a day. Workbooks using the older Mac "1904" date system are handled. Two-digit years are read as 2000–2049 or 1950–1999. |
| Splitting | One split can make at most 1,000 parts. Above that it stops with a message, for example when splitting by Consumer No by mistake. |

---

## 6. Privacy and hosting

- **Hosting:** GitHub Pages serves static files only. There are no server-side functions, no storage and no processing-time limits. The only limit on files is the browser's memory, shown above.
- **Privacy:** files are read with the browser's File API and processed in a Web Worker in the same tab. They are kept only in that tab's memory and are gone when you close or reload it. Nothing is stored or sent anywhere, and the page's Content-Security-Policy (`connect-src 'self'`) blocks network requests to other sites.
- **Saved in the browser:** only workflows you save on purpose, in `localStorage`. They never contain file contents or passwords.
- **No backend:** none is needed, so there is no retention policy to configure. If a server is ever added for a feature that browsers cannot do, it must store uploads in temporary storage and delete uploads and results automatically after a short, stated period, for example one hour.

## 7. Licensing

Sheet Desk uses these open-source libraries: SheetJS CE (Apache-2.0), ExcelJS (MIT), Papa Parse (MIT), JSZip (MIT/GPL-3.0 dual), jsPDF and jspdf-autotable (MIT), React (MIT) and **HyperFormula (GPL-3.0)**.

Because HyperFormula is GPL-3.0, publishing this app means its source must be available under GPL-3.0-compatible terms. A public GitHub repository satisfies this if you add a GPL-3.0 `LICENSE` file. The alternatives are to buy a HyperFormula commercial licence and change `licenseKey` in `src/core/formulas.ts`, or to remove the *Recalculate formulas* feature. Please decide this before you share the app widely.

See [ARCHITECTURE.md](ARCHITECTURE.md) for how the code is organised.
