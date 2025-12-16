client-side-app/
├── index.html                 # Entry point, loads app.js and styles.css
├── styles.css                 # Copied verbatim from original `app/static/style.css` for full fidelity
├── app.js                     # Router, file parsing, reconciliation logic, XLSX generation
├── pages/
│   ├── start.html             # Start Page (port of `templates/main.html`)
│   ├── reconciliation.html    # Reconcile Page (port of `templates/index.html`)
│   └── cats-edits.html        # CATs Edits Page (port of `templates/cats_edits.html`)

Connections and flow:
- `index.html` links `styles.css` and `app.js`.
- `app.js` implements a small hash router. When the hash changes it fetches the corresponding partial from the `pages/` folder and injects it into `#app`.
  - `#/` → `pages/start.html`
  - `#/upload` → `pages/reconciliation.html`
  - `#/cats-edits` → `pages/cats-edits.html`
- `app.js` contains the client-side equivalents of the server logic:
  - Excel reading: uses `SheetJS (xlsx.full.min.js)` via CDN and `FileReader` to read files as ArrayBuffer and parse first sheet into a 2D array.
  - Parsing: `sheetRowsToObjects()` converts rows to objects keyed by header names.
  - Date normalization: `normalizeDate()` reproduces Python parsing behavior for common formats and Excel serials.
  - Extraction: `extractDrmisData()` and `extractOracleData()` implement the same column-detection and filtering as the Python app (including filtering leave codes and expanding multi-day Oracle entries via `expandOracleEntries()`).
  - Reconciliation: `reconcileData()` merges and identifies mismatches (hours or leave code differences), matching the Python reconciliation logic.
  - CATs edits: `generateCatsEdits()` creates the same two-row-per-mismatch structure (data-row and total-row) as the Python server.
  
  Important behavioral note:
  - The web app reconciles leave for one employee at a time. If Oracle rows do not include a `Pers No` (Oracle exports typically omit it), the client will use the first `Pers No` encountered in the DRMIS file as the default `Pers No` for all Oracle rows. This mirrors the typical usage pattern where the uploaded Oracle sheet belongs to the same employee as the DRMIS rows being reconciled.
    - The web app can optionally auto-load a whitelist of leave codes from a local spreadsheet file named `Leave_Codes - Actual Leave.xlsx` placed in a `sample_data/` folder within `docs/` (for example: `docs/sample_data/Leave_Codes - Actual Leave.xlsx`). If present and accessible to the browser, the client will attempt to load that file on the Reconcile page and will use the file's `A/Atype` (or first) column as the allowed leave codes. When a whitelist is loaded, reconciliation will filter out any rows whose Oracle and DRMIS leave codes are not in the whitelist. This makes it easy to focus reconciling only on the leave types you care about.

  Robust header detection for DRMIS / Oracle files:

  - Real-world exported DRMIS files often include a title line, export date, metadata rows, blank rows, and even blank columns before the actual table header. The client now tries to detect the real header row instead of assuming it is the first row.
  - The detection strategy (implemented in `docs/app.js` -> `sheetRowsToObjects`) is:
    - Scan the first up to 30 rows and score each row by the number of non-empty cells plus a large bonus if the row contains header-like keywords (`pers`, `date`, `hours`, `a/atype`, `leave`, etc.).
    - Choose the highest-scoring row as the header row.
    - Compute column non-empty counts below the header and trim leading/trailing columns that are mostly empty (threshold: at least 3 non-empty cells or ~15% of data rows).
    - Build objects from rows after the detected header, skipping entirely-blank rows in the header columns.

  This makes the client resilient to files with top notes, blank columns, or extra rows before the real table. If a file still fails detection, the interactive column-mapping modal will appear so you can map the required fields manually.

  Preview behavior:

  - The File Preview now shows the detected table exactly as it appears in the uploaded file before any whitelist filtering is applied. This lets users verify what will be removed by the whitelist and ensures nothing important is accidentally excluded.
  - The preview window trims top/bottom non-table rows and blank leading/trailing columns using the same header-detection algorithm, so it displays only the real table content. It also allows horizontal scrolling so very wide tables can be inspected left-to-right.

  Oracle -> DRMIS code normalization:

  - Oracle exports typically carry 3-digit leave codes (for example `110`). The reconciliation logic requires comparing those with DRMIS codes (4-digit, for example `1110`). The client now normalizes Oracle leave codes by prefixing a `1` to 3-digit values (so `110` -> `1110`) before reconciliation and whitelist checks. If the Oracle code is already 4 digits it is used as-is. Missing or dash codes are treated as `'0'` as before.

  This ensures accurate matching between Oracle and DRMIS codes during reconciliation and when applying the whitelist.
  - Exporting: `downloadReconciliationExcel()` and `downloadCatsEditsExcel()` build .xlsx files client-side using SheetJS and trigger downloads.
  - Clipboard/email: `copyCatsToClipboard()` and `emailCatsTable()` copy HTML/plain text to the clipboard and open `mailto:` with a prefilled subject as a fallback (matching original behavior). Note: attaching binary files via `mailto:` is not supported by browsers — see notes below.

Validation mapping:
- Interactive Column Mapping: When the client auto-detection fails (a required column is missing or named differently), the app now opens an interactive mapping dialog that lists the available headers and lets the user map each required field to the correct column from their file. The dialog validates selections, applies the mapping in-memory, re-runs extraction for the mapped file, and — if both files are now available — automatically continues to reconciliation. The dialog writes the remapped data into the in-memory 2D arrays so subsequent actions (download, CATs generation) operate on the mapped dataset.

  - DRMIS required fields: `Pers No`, `Date`, `Hours`, `A/A Type`.
  - ORACLE required fields: `From Date`, `Hours Recorded`, `Leave Code`.
  - Usage: upload files, click "Reconcile Leave". If columns are missing, choose the matching column for each required field in the modal and click `Apply Mapping`.
  - Notes: The mapping is applied only in the browser session (stored in memory). To persist or re-use mappings across sessions, you can export the remapped files and keep them locally before re-uploading.

Limitations and notes:
- Styling fidelity: `styles.css` was copied verbatim to preserve visual parity.
- Excel styling: The Python app used `openpyxl` to apply detailed cell styling (header background, fonts, borders). SheetJS Community Edition supports creating .xlsx files but has limited cell-style support in the browser. The generated files will contain content and preserve structure; header/background styling may differ in some Excel viewers. If you require fully identical styling, using a server-side library (openpyxl) or SheetJS Pro is recommended.
- Email attachments: Browsers do not support attaching files automatically through `mailto:` links. The client app follows the original's UX by copying formatted HTML to the clipboard and opening the user's email client with a subject line — the user can paste the clipboard content or attach the downloaded .xlsx manually.
- Large files: SheetJS reads files into memory; for very large Excel files, memory/CPU usage may be high. The original server used pandas which can be more memory efficient on the server. For typical report sizes this client-side implementation should work fine.

Mapping of original features -> client-side implementation (1:1):
- Upload DRMIS/ORACLE .xlsx files: Original Flask file upload → New `input[type=file]` + `FileReader` + `SheetJS` parsing.
- Column detection and mapping: Original server-side mapping UI -> New client-side modal and instructions; mapping choices that remap indices were simplified to guidance due to client-only constraints.
- Reconciliation logic: Ported Python logic into JS functions (`normalizeDate`, `isBusinessDay`, `expandOracleEntries`, `extractDrmisData`, `extractOracleData`, `reconcileData`).
- CATs Edits preview, add/delete rows, total-hours color warnings, copy/download/email: Ported DOM logic and clipboard + XLSX download with UI parity.

How to run locally (quick):
1. Open `client-side-app/index.html` in your browser (double-click or use a static file server). No Python or server required.
2. Go to "Reconcile Page", upload DRMIS and ORACLE .xlsx files, then click "Reconcile Leave".
3. Review mismatches and proceed to "CATs Edits" to copy or download the edits.

Audit / Cleanup:
- I did NOT delete any original files. Please confirm before I remove any original server-side files.
# Project structure (client-side docs/ app)

```
docs/
├── index.html                 # Entry point; loads `app.js`, `styles.css`, SheetJS and ExcelJS
├── styles.css                 # Copied verbatim from original `app/static/style.css` for visual parity
├── app.js                     # Router, file parsing, reconciliation logic, Excel read/write logic
├── pages/
│   ├── start.html             # Start Page (port of `templates/main.html`)
│   ├── reconciliation.html    # Reconcile Page (port of `templates/index.html`)
│   └── cats-edits.html        # CATs Edits Page (port of `templates/cats_edits.html`)
```

Connections and flow:
- `index.html` links `styles.css`, `app.js`, `SheetJS` (for reading spreadsheets) and `ExcelJS` (for styled `.xlsx` writing).
- `app.js` implements a small hash router. When the hash changes it fetches the corresponding partial from the `pages/` folder and injects it into `#app`.
  - `#/` → `pages/start.html`
  - `#/upload` → `pages/reconciliation.html`
  - `#/cats-edits` → `pages/cats-edits.html`
- `app.js` contains the client-side equivalents of the server logic:
  - Excel reading: uses `SheetJS (xlsx.full.min.js)` via CDN and `FileReader` to read files as ArrayBuffer and parse first sheet into a 2D array.
  - Parsing: `sheetRowsToObjects()` converts rows to objects keyed by header names.
  - Date normalization: `normalizeDate()` reproduces Python parsing behavior for common formats and Excel serials.
  - Extraction: `extractDrmisData()` and `extractOracleData()` implement the same column-detection and filtering as the Python app (including filtering leave codes and expanding multi-day Oracle entries via `expandOracleEntries()`).
  - Reconciliation: `reconcileData()` merges and identifies mismatches (hours or leave code differences), matching the Python reconciliation logic.
  - CATs edits: `generateCatsEdits()` creates the same two-row-per-mismatch structure (data-row and total-row) as the Python server.

Important behavioral notes:
- The web app reconciles leave for one employee at a time. If Oracle rows do not include a `Pers No` (Oracle exports typically omit it), the client will use the first `Pers No` encountered in the DRMIS file as the default `Pers No` for all Oracle rows. This mirrors the typical usage pattern where the uploaded Oracle sheet belongs to the same employee as the DRMIS rows being reconciled.
- The web app can optionally auto-load a whitelist of leave codes from a local spreadsheet file named `Leave_Codes - Actual Leave.xlsx` placed in a `sample_data/` folder within `docs/` (for example: `docs/sample_data/Leave_Codes - Actual Leave.xlsx`). If present and accessible to the browser, the client will attempt to load that file on the Reconcile page and will use the file's `A/Atype` (or first) column as the allowed leave codes. When a whitelist is loaded, reconciliation will filter out any rows whose Oracle and DRMIS leave codes are not in the whitelist. This makes it easy to focus reconciling only on the leave types you care about.

Robust header detection for DRMIS / Oracle files:

- Real-world exported DRMIS files often include a title line, export date, metadata rows, blank rows, and even blank columns before the actual table header. The client now tries to detect the real header row instead of assuming it is the first row.
- The detection strategy (implemented in `docs/app.js` -> `sheetRowsToObjects`) is:
  - Scan the first up to 30 rows and score each row by the number of non-empty cells plus a large bonus if the row contains header-like keywords (`pers`, `date`, `hours`, `a/atype`, `leave`, etc.).
  - Choose the highest-scoring row as the header row.
  - Compute column non-empty counts below the header and trim leading/trailing columns that are mostly empty (threshold: at least 3 non-empty cells or ~15% of data rows).
  - Build objects from rows after the detected header, skipping entirely-blank rows in the header columns.

This makes the client resilient to files with top notes, blank columns, or extra rows before the real table. If a file still fails detection, the interactive column-mapping modal will appear so you can map the required fields manually.

Preview behavior:

- The File Preview shows the detected table exactly as it appears in the uploaded file before any whitelist filtering is applied. This lets users verify what will be removed by the whitelist and ensures nothing important is accidentally excluded.
- The preview window trims top/bottom non-table rows and blank leading/trailing columns using the same header-detection algorithm, so it displays only the real table content. It also allows horizontal scrolling so very wide tables can be inspected left-to-right.

Oracle -> DRMIS code normalization:

- Oracle exports typically carry 3-digit leave codes (for example `110`). The reconciliation logic requires comparing those with DRMIS codes (4-digit, for example `1110`). The client normalizes Oracle leave codes by prefixing a `1` to 3-digit values (so `110` -> `1110`) before reconciliation and whitelist checks. If the Oracle code is already 4 digits it is used as-is. Missing or dash codes are treated as `'0'`.

Exporting and clipboard/email behavior:
- Reconciliation export: `downloadReconciliationExcel()` uses SheetJS to build a simple `.xlsx`/sheet from the mismatch array (data-focused export).
- CATs export: `downloadCatsEditsXlsxExcelJS()` (new) uses `ExcelJS` to build a styled, true `.xlsx` from the *live DOM table at the time of export*. This ensures any user-added or edited rows in the CATs UI are captured in the downloaded workbook. The exporter applies:
  - Title rows with merges and font sizing
  - Header row fill color and bold font
  - Column widths and number formatting (`0.00` for hours)
  - Thick, colored bottom border and light-grey fill for `total-row` rows to visually separate groups
  - Thin borders for ordinary rows and an autofilter on the header row
  - The top-right icon buttons on the CATs Edits page are wired to the same footer actions (download / copy / email) and call the live-DOM exporters/clipboard/email helpers.
- Clipboard/email: `copyCatsToClipboard()` and `emailCatsTable()` copy HTML/plain text to the clipboard and open `mailto:` with a prefilled subject as a fallback (matching original behavior). Note: attaching files automatically via `mailto:` is not supported by browsers — users must attach files manually if they want attachments.

Validation mapping:
- Interactive Column Mapping: When the client auto-detection fails (a required column is missing or named differently), the app opens an interactive mapping dialog that lists the available headers and lets the user map each required field to the correct column from their file. The dialog validates selections, applies the mapping in-memory, re-runs extraction for the mapped file, and — if both files are now available — automatically continues to reconciliation. The dialog writes the remapped data into the in-memory 2D arrays so subsequent actions (download, CATs generation) operate on the mapped dataset.

  - DRMIS required fields: `Pers No`, `Date`, `Hours`, `A/A Type`.
  - ORACLE required fields: `From Date`, `Hours Recorded`, `Leave Code`.

Limitations and notes:
- Styling fidelity: `styles.css` was copied verbatim to preserve visual parity in the browser UI.
- Excel styling: The client now uses `ExcelJS` for styled `.xlsx` export of CATs edits. `ExcelJS` supports cell-level styling in-browser (fonts, fills, borders, merges) and produces native `.xlsx` files that open in Excel with the intended formatting. Reconciliation table export remains a simple SheetJS `.xlsx`.
- CDN vs vendored libs: The app loads `SheetJS` and `ExcelJS` from CDNs. If you prefer a self-contained `docs/` folder, I can vendor `exceljs.min.js` (and/or the SheetJS bundle) into `docs/lib/` and update `index.html` accordingly.
- Email attachments: Browsers do not support attaching files automatically through `mailto:` links. The client app follows the original's UX by copying formatted HTML to the clipboard and opening the user's email client with a subject line — users can paste the clipboard content or attach the downloaded .xlsx manually.
- Large files: SheetJS reads files into memory; for very large Excel files, memory/CPU usage may be high. The original server used pandas which can be more memory efficient on the server. For typical report sizes this client-side implementation should work fine.

Mapping of original features -> client-side implementation (1:1):
- Upload DRMIS/ORACLE .xlsx files: Original Flask file upload → `input[type=file]` + `FileReader` + `SheetJS` parsing in-browser.
- Column detection and mapping: Original server-side mapping UI → client-side modal and in-memory remapping.
- Reconciliation logic: Ported Python logic into JS functions (`normalizeDate`, `isBusinessDay`, `expandOracleEntries`, `extractDrmisData`, `extractOracleData`, `reconcileData`).
- CATs Edits preview, add/delete rows, total-hours color warnings, copy/download/email: Ported DOM logic and clipboard + styled XLSX download (ExcelJS) with UI parity.

How to run locally (quick):
1. Serve the `docs/` folder or open `docs/index.html` in a modern browser. For best results run a static server:
```powershell
cd "c:\Users\Rande\Leave Reconcile 2 - Copy\docs"
python -m http.server 8000
```
2. Open http://localhost:8000, go to the Reconcile Page, upload DRMIS and ORACLE .xlsx files, then click "Reconcile Leave".
3. Review mismatches and proceed to "CATs Edits" to edit, copy or download the edits. The CATs download will include any runtime edits.

Audit / Cleanup:
- I did NOT delete any original server-side files. Please confirm before I remove any original server-side assets.

**Recent Updates (Dec 2025)**
- **Vendored Libraries:**: SheetJS (`xlsx.full.min.js`) and `ExcelJS` were vendored into `docs/lib/` to avoid CDN blocking and Tracking Prevention issues. `index.html` prefers local bundles with CDN fallback.
- **Diagnostics:**: An in-page global error handler was added to `docs/index.html` so uncaught exceptions and unhandled rejections are displayed on the page (prevents white-screen failures from silently hiding syntax/runtime errors).
- **DRMiS Lookup & Prefill:**: Implemented `drmisLookup` (date|pers -> DRMIS entries) and `buildDrmisLookup()` to index DRMIS detail rows for use when auto-filling CATs editable rows.
- **Auto-prefill Behavior:**: When building the CATs table the app now attempts to auto-create a single prefilled editable-row immediately after each `data-row` only if `prefillEditableRow()` successfully filled fields. If prefill fails (no matching DRMIS residuals) no blank editable-row is inserted — the user must add rows manually.
- **Prefill Algorithm Details:**: `prefillEditableRow()` finds DRMIS candidates for the date+Pers No, excludes DRMIS leave rows that match the Oracle leave code, consumes any existing editable-hours already present, and allocates residual DRMIS work hours as prefilled rows (Work Order / Act / AA Code / Hours). It now handles both `input` and `select` Act fields and will fall back to plain text if the Act value doesn't match available options.
- **Replacement Rule (Oracle -> DRMiS):**: New rule: if, for a reconciliation data-row, both Oracle Leave Code and DRMiS Leave Code are present and differ, the Oracle leave is treated as a replacement of the DRMiS leave. In that case:
  - The original DRMiS leave row is excluded from work-candidate selection.
  - The algorithm does NOT subtract Oracle leave hours from DRMIS work residuals (i.e., it preserves DRMIS work rows like the 4.25hr example on Apr 28).
  - `generateCatsEdits()` records `original_drmis_code`, `original_drmis_hours`, and `replaced` metadata on the cats data-row; `initCatsEditsPage()` attaches these as `data-*` attributes so `prefillEditableRow()` can act accordingly.
- **Act Field Usability:**: The Act column was changed from a `<select>` to an editable `<input>` with a small custom arrow hint so users can type any Act value (common options are still suggested visually via CSS). All exporters and clipboard/email code now read Act from either `.act-input` (new) or the legacy select if present.
- **CSS Adjustments:**: The Act column width was reduced to `8ch`. Native browser dropdown/clear controls were hidden and a small SVG arrow was added as a visual hint; styling uses `appearance:none` and vendor pseudo-element rules to minimize native UI in different browsers.
- **Totals Fixes:**: `updateTotals()` was rewritten to robustly traverse previous sibling rows and sum numeric hours from both data-rows and editable rows; `formatHours()` normalizes hours to two decimals on blur. Additional `updateTotals()` calls were added during table construction and after DOM mutations to keep totals accurate.
- **Export Behavior:**: The CATs `.xlsx` exporter (`downloadCatsEditsXlsxExcelJS`) now reads the live DOM at export time so any user edits are included. `copyCatsToClipboard()` and `emailCatsTable()` also pull current values from inputs/text in the table when building their outputs.
- **Add-row Behavior:**: The `+` button now inserts a blank editable-row (Act is editable input) — unlike prefilled rows which are only inserted when prefill succeeds.
- **Other small fixes:**: Fixed stray syntax errors, hardened header detection, and improved defensive parsing (trimming, numeric coercion) across the codebase.

- **DRMiS Hours Reassignment (Oracle-removed rows):**: When a reconciliation `data-row` has Oracle Leave Code empty/`0` and Oracle Hours `0` (meaning Oracle removed the DRMIS leave), the CATs prefill now reallocates the original DRMIS hours (stored as `original_drmis_hours`) into the remaining DRMiS-derived editable rows for that date and `Pers No`. Allocation is proportional to each candidate's available hours (with rounding correction to the first candidate); if no DRMiS candidates exist the app does not auto-insert rows and the user must add them manually. This behavior is implemented in `prefillEditableRow()` and is applied only to rows that meet the Oracle-empty condition.

- **Preview UI & Record Counts:**: Replaced the small emoji preview controls with structured `preview-btn` buttons that include a compact SVG spreadsheet icon and the label `Preview`. Adjacent to each button a live count element (`#drmis_count`, `#oracle_count`) now displays the number of detected records (e.g., "31 records found") when a file is selected; counts are computed with `sheetRowsToObjects(...)` and updated on file-change handlers. The count elements use `aria-live="polite"` for accessible announcements. The CSS adds a `.preview-btn` style for the modern rounded button with hover/focus states.

- **Preview Row Limit Increase:**: The file preview modal was previously limited to displaying only the first 200 data rows to avoid overwhelming the UI for large files. This limit has been increased to 1000 rows to accommodate larger datasets, such as DRMIS files with 258+ rows. The preview note now indicates the actual number of rows shown (e.g., "Showing 258 of 258 data rows").

