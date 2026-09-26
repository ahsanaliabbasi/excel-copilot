# Excel Copilot — Complete Project Context

A no-code web interface for editing real `.xlsx` files. Users click through guided steps instead of writing formulas; the app writes **native Excel formulas with cached values** into the downloaded workbook, so everything keeps recalculating in Excel.

This document captures the whole build: goals, architecture, every feature, the key design decisions, the bugs that shaped the design, how it was tested, and what is still open.

---

## 1. Goal

Let non-technical users do spreadsheet work (lookups, conditional totals, text cleanup, IF logic, dates, duplicates) without formulas, while the output file contains **real Excel formulas**, not pasted values.

Priority functions: XLOOKUP, VLOOKUP, INDEX-MATCH, CONCAT/CONCATENATE, IF, IFNA, SUM, SUMIF, COUNTIF, MAX/MAXIF, AVERAGE, LEFT/RIGHT/MID, TRIM, SUBSTITUTE, ROUND, EDATE, TODAY, SEARCH, FIND, LEN, TEXT — plus cleaning, filtering, duplicates, editing, copy/paste, and downloading the edited file.

## 2. Stack

| Layer | Tech |
|---|---|
| Backend | FastAPI + uvicorn, pandas, numpy |
| Reading xlsx | `python-calamine` (Rust reader, ~0.5 s for 75,000 rows) |
| Writing xlsx | `openpyxl` (small workbooks, keeps styles) / `xlsxwriter` (streamed, big workbooks) |
| SQL sandbox | `sqlglot` + in-memory `sqlite3` |
| Frontend | Plain HTML/CSS/JS, no build step (served by the backend with `no-store`) |

Sessions live in server memory (`MAX_SESSIONS = 4`).

## 3. Run it

```bash
cd backend
python -m venv venv
venv\Scripts\activate            # macOS/Linux: source venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

Open **http://localhost:8000** and press **Ctrl+F5** after updates. The page shows a banner if the browser holds an old copy (`FRONTEND_VERSION` vs `/api/health`).

Current versions: `FRONTEND_VERSION = 17`, `API_VERSION = 6`.

## 4. Repository layout

```
backend/
  main.py            FastAPI app: upload, paging, operations, editing, duplicates,
                     row deletion, SQL endpoints, background download, static frontend
  expr_engine.py     Step-tree engine: every node yields an Excel formula AND a Python value
  sql_engine.py      SQLite sandbox over the sheets + SQL -> Excel-formula translation
  formula_engine.py  Legacy operation helpers (col_letter, quote_sheet, ...)
  requirements.txt
frontend/
  index.html, style.css
  app.js             state, upload, download progress, modals, version checks
  grid.js            paged Excel-style grid: selection, editing, clipboard, duplicate highlight
  builder.js         step builder + the Back/Next wizard (Lookup, Totals & Counts)
  sql.js             SQL screen
  dup.js             Duplicates screen
```

## 5. Architecture

### 5.1 Big-workbook design (12 sheets x 75,000 rows)

The first version could not even upload a 35 MB / 12-sheet file (two full openpyxl loads, ~96 s each, memory blow-up). The rewrite:

- The upload is saved to disk; only the workbook's XML table of contents is read at upload (instant).
- Sheets load lazily with calamine; the first is loaded before the upload call returns, the rest on a background thread.
- Nothing is written into an openpyxl workbook while working. State is:
  - `grid[name]` — raw rows exactly as shown,
  - `writes[name]` — `{(row, col): value or "=formula"}` overlay of what should be saved (1-based worksheet coordinates),
  - `wfmt[name]` — number formats for written cells,
  - `recipes` — saved formula columns so they can be recalculated after edits.
- The grid is paged (default 1,000 rows, max 5,000) with absolute row ids; edits, undo, copy and highlights keep working across pages.
- Download runs in the background with progress. **Full mode** (small workbooks): reopen the original with openpyxl and apply the overlay (keeps styles). **Stream mode** (> 1,000,000 cells): stream every sheet with xlsxwriter from `grid` + formulas (keeps data and formulas, not cell styling).

Measured: upload 0.2 s, first sheet ~2 s, all sheets ~7 s, an operation ~1 s per 75k rows, duplicate check 0.09 s, download 35–48 s with progress.

### 5.2 The expression engine (`expr_engine.py`)

Users build a **tree of steps** (JSON). Each node type `X` has:
- `f_X` — produces the Excel formula text for a given row,
- `v_X` — computes the Python value shown in the grid.

Both must agree; every feature was cross-checked against pandas/calamine ground truth.

Node types: `col`, `lit`, `blank`, `today`, `if`, `iferr`, `join`, `textfn` (TRIM/UPPER/LOWER/PROPER/LEN/VALUE), `replace` (multi-SUBSTITUTE), `extract` (before/after text, between, first/last N, MID, before/from digit/letter), `find`, `textfmt`, `calc`, `round`, `rowagg`, `edate`, `lookup`, `agg`.

Key behaviours:
- **Lookup** — `key`, `match_col`, optional transformed sides (`key_src` / `match_left` pipes), `cases: [{where, return_type: col|value|count|agg, agg_fn, return_col}]`, `search` first / last / nth / nth_last, `not_found`, and method XLOOKUP / VLOOKUP / INDEX-MATCH with match modes (exact, next smaller/larger, wildcard).
- **Conditions** — `where` groups `{join: AND|OR, items: [leaf | group]}`. Plain AND uses classic COUNTIFS/SUMIFS/MAXIFS; anything else uses array forms (SUMPRODUCT, `_xlfn.AGGREGATE`, `INDEX/MATCH(1, INDEX((cond)*1,0),0)`, `LOOKUP(2,1/(cond),rng)`).
- **Function chains in conditions** (e.g. `LEN(TRIM(Model)) < 6`, `RIGHT(UPPER(TRIM(x)),2) = "RD"`) work on either sheet; array-mode allows only `col, lit, blank, textfn (no VALUE), replace, extract (subset), textfmt, round, calc(+ - *)`, no IF/IFERROR.
- **`only_if` / `only_else`** on any node compiles to `IF(cond, step, otherwise or "")` — "only fill rows where…".
- Modern functions need `_xlfn.` prefixes (XLOOKUP, IFNA, TEXTJOIN, MAXIFS, MINIFS, AGGREGATE) or Excel shows `#NAME?` — an earlier bug where AGGREGATE lacked the prefix was fixed.
- Count-only lookups emit a plain `=COUNTIF(...)` (no IFERROR wrapper).

### 5.3 Row-local vs global recipes

Formula columns are saved as recipes. After a cell edit, row-local recipes recalculate only the edited rows; lookups/totals recalculate the whole column. After **row deletion** every recipe is fully re-applied (`_reapply_recipe`), because a formula's own row references would otherwise point at its old position.

### 5.4 Row deletion primitive

Added for Duplicates, generic and reusable (`_delete_rows`):
- removes rows from the live `grid`, remaps `writes` / `wfmt` / loaded original formulas to the new row numbers,
- records the **original** file row numbers in `deleted_rows`, tracked through a `row_map`,
- `_load_orig_formulas` shifts freshly-read original formulas by `deleted_rows`, so stream-mode exports stay correct,
- `_export_full` replays the deletions on the reopened original with `ws.delete_rows` (descending) before applying the overlay — without this, small workbooks would silently export with the rows still present.

## 6. Features

### 6.1 Sidebar tools
Formula Builder · SQL Query · Combine Columns · Lookup / Match · IF / Decision Flow · Totals & Counts · Clean Text · Replace Text · Find & Extract · Calculate · Round Number · Dates.

### 6.2 Lookup / Match and Totals & Counts (step-by-step wizard)
One question on screen at a time, with a progress indicator (numbered, clickable dots), **Back** and **Next: <next step>**, the chosen step remembered across redraws, and an "All set — choose where the result goes" hint on the last step. A plain-English summary and live preview stay visible.

Lookup steps: which rows of this sheet get a value → where to look → how the two sheets match (link by a matching value with optional clean-up steps on either side, or "don't link"; Advanced method/match mode collapsed here) → what should come back (cases: get a value / count rows / total / average / highest / lowest / fixed value, each optionally narrowed by AND/OR conditions on the other sheet; IF / ELSE IF cases) → which row if several qualify (only when relevant) → what to show if nothing is found.

Totals & Counts steps: which rows → what to work out → which rows count.

The earlier flowchart-themed layout (coloured nodes, connectors, YES/NO) was removed at the user's request in favour of this clean wizard.

### 6.3 Guided building blocks
"Pipes": a column followed by steps — Trim, UPPER/lower/Capitalize, Length, Replace, Remove, First/Last N characters, MID, Before/After a character, Round, Add, Multiply — usable in link keys and in conditions on either sheet.

### 6.4 SQL Query
Every sheet is a table (short alias slugs, `this.<col>` = the row being filled). Run SELECT, preview UPDATE as a diff before committing, save results as a new sheet, fill a column with query values (by key or per row), and **translate** a query to Excel formulas (COUNT/SUM/AVG/MIN/MAX and plain columns with where trees). Refuses DISTINCT/HAVING/subqueries/UNION/window/ORDER BY+LIMIT/expressions/3+ tables. Sandbox: authorizer + 30 s progress handler.

### 6.5 Duplicates (`dup.js` + `/api/dup/*`)
Opened from the grid toolbar. Works on the whole sheet server-side.

1. **Key columns** — tick any combination; an id-like column (`custId`, `employee_id`, code, number) is pre-ticked. Options: ignore case/extra spaces, ignore rows where the key is empty.
2. **Overview** — total records, duplicate groups, duplicate rows (would be removed), unique records after dedup, and "View all N duplicate rows in the grid".
3. **Check one value** — one input per key column; shows the matching rows with a duplicate verdict.
4. **Duplicate groups** — paged, searchable, biggest first; clicking a group loads its rows and highlights it. Clearing the fields, or the **✕ Clear** button, returns to the neutral state.
5. **Hand-pick rows** — checkboxes on the result table; **Delete the ticked rows** or **Keep only the ticked rows**, each with a confirmation naming the row count.
6. **Copy a set of rows** — unique / duplicates / all rows in a group, with or without header, to the clipboard as TSV for pasting into Excel.
7. **Remove duplicates** — keep first or last occurrence, then delete from this sheet (confirmation first) or write a **new sheet** (unique / duplicates only / every row in a group) with a name you choose.

### 6.6 Excel-style grid
Paged view, click/drag/Shift/row/column/Ctrl+A selection, optional cell editing behind an **Allow editing** switch (off by default), undo, Excel-compatible copy/paste (whole columns come from the server), banded row colours, row hover, selected-record highlight, header-row setting, dark mode.

## 7. API summary

| Endpoint | Purpose |
|---|---|
| `POST /api/upload`, `GET /api/sheets/{id}` | upload, sheet list |
| `POST /api/sheet/header-row`, `/api/sheet/page`, `GET /api/sheet/{id}/{name}` | header row, paging |
| `POST /api/sheet/range-text` | TSV of a block (whole-column copy) |
| `POST /api/sheet/edit` | edit cells (+ recipe recompute) |
| `POST /api/sheet/delete-rows` | delete specific rows |
| `POST /api/operations/preview` / `apply` | build formulas from a step tree |
| `POST /api/dup/summary`, `/groups`, `/lookup`, `/apply`, `/export-text` | Duplicates |
| `GET /api/sql/schema/{id}`, `POST /api/sql/run`, `/translate`, `/apply-values`, `/save-sheet`, `/commit-update` | SQL |
| `POST /api/download/start/{id}`, `GET /status`, `GET /file` | background export |
| `GET /api/health` | `{status, version, frontend, active_sessions}` |

## 8. Design decisions and the problems behind them

- **Empty white modal** — `display:flex` overrode `hidden`; fixed with a global `[hidden]{display:none!important}`.
- **"Side menu not working"** — orphaned CSS fragments swallowed `.modal-overlay`; a structural CSS checker, a delegated click handler and cache-busting were added.
- **Old UI in the browser** — a plain static server let the browser cache files; the backend now serves the frontend with `no-store`, the API base is same-origin, and a version banner appears on mismatch.
- **Joined "empty" values** — a join of two empty columns produced `" "`; fixed with "is empty (or only spaces)" and skip-empty join by default.
- **Stray spaces in Replace** — find text like `" -"` silently failed; added a warning and one-click trim.
- **Big file upload** — see 5.1.
- **Text-in-number quirks in tests** — Excel `LEN(44807)` is 5; an early test counted `"44807.0"` as 7 characters. The app was right, the expectation was wrong.
- **Flowchart lookup UI → wizard** — the flowchart layout was replaced by the step-by-step wizard for clarity.
- **Natural-language "Just say it" box** — built as a deterministic sentence parser, then **removed** at the user's request.

## 9. Testing approach

- **Ground truth**: pandas / calamine computed independently; formulas and grid values compared across every row.
- **Real browser**: headless Edge driven over the DevTools protocol with a hand-written client — real mouse and keyboard events, `DOM.setFileInputFiles` for uploads, clipboard reads, screenshots.
- **Round-trip**: apply → download → reopen with openpyxl → verify rows/formulas in the actual file.
- Temporary test servers ran on port 8001 and were stopped afterwards; the user's own reloading server on 8000 was never touched.
- Test scripts were deleted after use (not part of the repo).

## 10. Known limits

- Sessions are in memory; a restart loses uploads.
- No authentication — run locally or behind your own access control.
- Generated formulas are verified against independent calculations but were not opened in every Excel version.
- Array-mode conditions cannot use IF, VALUE, digit/letter cuts, between-text or FIND; a single error value in the searched range can poison array totals.
- Mixed AND/OR inside one IF step is not supported (one join per step).
- Stream-mode downloads of very large workbooks do not keep cell styling.
- Multiple sheets with similar names (e.g. "Data Set 1", "Data Set 1 (old)") need explicit selection.

## 11. Repository status

- A local git repo exists with a single commit (`be40fd0`, "initiall commit") tracking `origin/main`.
- Remote: `https://github.com/ahsanaliabbasi/https---github.com-ahsanaliabbasi-excel-copilot.git` — the name looks malformed (a full URL pasted into the repo-name field) and the account differs from the one in the user's email. Rename to `excel-copilot` and update `origin` if that was not intended.
- Real data files (`*.xlsx`, `sample_data/`) contain personal-looking data and should not be committed to a public repo; virtual environments and `.claude/` should be git-ignored.

## 12. Possible next steps

- Apply the wizard to the remaining tools (IF / Decision Flow, Clean Text, etc.).
- Persist sessions (Redis or SQLite) instead of memory.
- Pivot tables and charts.
- Auth before any shared deployment.
- Undo for row deletions (currently only cell edits have an undo stack).
