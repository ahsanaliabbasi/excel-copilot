# Excel Copilot — Summary Rows, Sheet Tabs, New Rows/Columns, Copy to Sheet (Session Context)

Companion to [PROJECT_CONTEXT.md](PROJECT_CONTEXT.md), [UI_REDESIGN_CONTEXT.md](UI_REDESIGN_CONTEXT.md) and [WIZARD_FOCUS_AND_GRID_FILTER_CONTEXT.md](WIZARD_FOCUS_AND_GRID_FILTER_CONTEXT.md). The look is now the Mappa-style theme (near-white / charcoal, light-blue cards, pill buttons with a ↗ badge, mono micro-labels); see the chat history for that restyle.

## 1. Request
Work on data like in Excel: Sum / Avg / Min / Max on a selected column added as a separate row; create, rename (and manage) sheets; add rows and columns; copy chosen columns or rows into another sheet, then use the existing formulas (e.g. IF value > average) and filters on them.

## 2. Design
- **Summary rows are footer rows**, not data. `meta[sheet]["footer"]` = how many rows at the bottom of the grid are summary rows; `_load_sheet(..., footer)` leaves them out of the DataFrame, so sorting, filtering, duplicates, lookups and formulas built with the Formula Builder never see them. They are real rows in the exported file: `=SUM(B2:B148)` with a label in the first free column.
- `session["footers"][sheet]` = `[{fn, label, cols}]`. `_touch()` calls `_refresh_footer()`, so summary rows recalculate after any edit, paste, deleted row or formula column. Adding data below the last row goes through `_trim_tail()` + `_open_gap()`, which insert blank rows *above* the summary rows (`writes`, `wfmt`, `orig_formulas`, `row_map` shift with them; `row_map` uses `0` for "new row").
- **Sheet names** are tracked per session: `src` (current name → name inside the uploaded file), `renames` (ordered), `gone` (deleted original sheets). `_PER_SHEET` lists every per-sheet dict that must be re-keyed on rename/delete.
  - Rename rewrites `'Old'!` in every formula (`writes`, loaded `orig_formulas`, saved recipes) and, at export, in the original file's formulas (`_apply_renames`).
  - `_export_full` gives original sheets temporary titles first, so swapping/renaming can't collide; added sheets (new / copy / query result) now also write their `writes` overlay (formulas).
- **Copy to sheet** copies values (like Paste Special → Values) from the selection or every row of the current filtered/sorted view, into a new sheet or under the data of an existing one (missing columns are created with the source names).

## 3. Backend (`main.py`, `API_VERSION` 7)
| Endpoint | Purpose |
|---|---|
| `POST /api/sheet/summary` | `{cols, fn, add}` — SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, MEDIAN; `add:false` only calculates (used by the filter popover) |
| `POST /api/sheet/summary-remove` | remove one summary row |
| `POST /api/sheet/create` / `rename` / `duplicate` / `delete` | sheet tabs (≥1 sheet always remains; Excel naming rules enforced) |
| `POST /api/sheet/add-column`, `/rename-column` | new column at the right-hand end; rename a header (saved recipes follow) |
| `POST /api/sheet/copy-to` | columns/rows → new or existing sheet |
| `/api/sheet/page` | now also returns `footer` (the summary rows) |

## 4. Frontend (`FRONTEND_VERSION` 18)
- `index.html`: "+ New sheet", a tools bar (Σ Sum · Average · Min · Max · Count · More… · + Row · + Column · Copy to sheet…), `<tfoot>` in the grid.
- `app.js`: sheet rows with hover actions (rename / duplicate / delete) and double-click inline rename; `sheetPost`, `askForm`, `confirmBox` helpers.
- `grid.js`: `buildFooter()` (pinned summary rows with ✕), blank "phantom" rows for + Row (created on the first edit), paste into an empty sheet, `window.gridApi` for tools.js, and in the number filter popover the shortcuts Average / Median / Min / Max.
- `tools.js` (new): summary buttons, + Column, double-click column-name rename, Copy-to-sheet dialog.

## 5. Typical flow
Copy a column to a new sheet → click its header → **Average** (summary row) → column filter ▸ *Greater than* ▸ **Average** shortcut (or an IF step in the Formula Builder using "Total of matching rows → Average", which excludes the summary rows) → copy the filtered rows onward.

## 6. Tested
API run against the sample workbook: summaries, appending/deleting rows with summary rows present, rename (formulas rewritten in the exported file, including names with apostrophes), delete, duplicate, copy-to, export round-trip with openpyxl. Headless Edge: summary buttons, filter shortcut, copy dialog, new sheet, paste into an empty sheet, + Row, + Column, sheet rename, dark mode. No JS errors.

## 7. Limits / open items
- New rows and columns are added at the **end** only (inserting in the middle would have to rewrite formulas inside the original file).
- Whole rows/columns can now be deleted (row number / column header: right-click, or select + Delete) — see [DELETE_ROWS_AND_COLUMNS_CONTEXT.md](DELETE_ROWS_AND_COLUMNS_CONTEXT.md).
- Sheet create / rename / delete / duplicate and summary rows have no undo (cell edits still do). Stream-mode (very large) exports were not re-tested for these features.
- A row typed under the data does not get the formula columns' formulas automatically (existing behaviour of recipes for appended rows).
- Summary rows cover whole columns, not the currently filtered rows.
- `websocket-client` and `requests` were pip-installed into `backend/venv` for testing; they are not in `requirements.txt`.
