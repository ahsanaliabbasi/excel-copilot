# Excel Copilot — Delete Whole Rows / Columns (Session Context)

Companion to [SUMMARY_ROWS_AND_SHEET_TOOLS_CONTEXT.md](SUMMARY_ROWS_AND_SHEET_TOOLS_CONTEXT.md).

## 1. Request
It's easy to apply a calculation without changing the output-column dropdown away from "+ New column", which leaves a stray column behind. Fix: let the user remove any row or column — by right-clicking its row number / column header, or by selecting the whole row/column and pressing **Delete**.

## 2. Design
- Row deletion already existed (`/api/sheet/delete-rows`, used by the Duplicates screen) — this just wires it into the grid's own selection/keyboard/right-click UI.
- Column deletion is new: `_delete_cols(session, name, idxs0)` in `main.py` removes the columns from the live `grid`, shifts every `(row, col)` key in `writes` / `wfmt` / `orig_formulas` / summary-row `cols` left by however many deleted columns preceded it, and drops entries that pointed at a deleted column outright.
  - A saved Formula-Builder recipe's `out_col` is a **fixed worksheet column number**, not a name — it has to be shifted (or the recipe dropped, if its own output column was deleted) the same way, or `_reapply_recipe` silently writes into the wrong column after a delete (caught in testing: an orphaned `col_5` appeared with duplicate values until this was added).
  - A recipe that *reads* a deleted column (e.g. `AgeInYears` built from a `date_of_birth` column that then gets deleted) is left in place with its last computed values/formula — `_reapply_recipe` swallows the "column not found" error, same as it already tolerates other broken references. No crash, but it goes stale; there's no UI hint for this yet.
- **Export**: `_export_full` (small workbooks) rebuilds the file by reopening the *original* upload with openpyxl and overlaying `writes` on top — it does not read from `session["grid"]`. Row deletion already had a side-channel for this (`deleted_rows` + `row_map`, replayed as `ws.delete_rows()`); column deletion needed the same: `deleted_cols` + `col_map` (both added to `_new_session` and `_PER_SHEET`), replayed as `ws.delete_cols()` right before the rows pass. Without this the exported file kept every original column (just with `writes` silently overlaid at shifted, wrong positions) even though the live grid/API looked correct — a small workbook and a big one (stream export, which *does* build straight from `grid`) behaved differently until this was fixed.
  - A column added in-app (`+ Column`) and later deleted has no original-file counterpart, so `col_map` naturally has no entry for it (guarded by an index-bounds check) — nothing is deleted from the reopened original file for it, which is correct: there was never anything there to delete.
- Known, accepted limitation (same class as row deletion already has): an **original** formula's cell references inside its text (e.g. `=C2+30`) are never rewritten when a column shift changes what `C` now means — only the formula's *storage position* moves. Formulas the app itself generates (EXPR/recipes) don't have this problem because they're regenerated fresh from the column **name** each time.
- Right-click menu is a small new `.ctx-menu` popover in `grid.js` (same pattern as the column filter popover), shown from a `contextmenu` listener on the row-number cells and the column header cells only — not on ordinary data cells.
- **Delete key**: only removes the row/column when the *whole* row or column is selected (`G.sel === "rows" | "cols"`, i.e. you clicked the row number / column header). A normal cell selection still just clears contents, same as before. **Backspace** always just clears — it never deletes a row/column, matching Excel's own split between the two keys.

## 3. Backend (`main.py`)
| Endpoint | Purpose |
|---|---|
| `POST /api/sheet/delete-rows` | unchanged — now also reachable from the grid's own UI |
| `POST /api/sheet/delete-column` | `{cols: [0-based positions]}` — refuses to remove every column |

## 4. Frontend (`grid.js`)
- `deleteRows()` / `deleteCols()` — `confirmBox` (danger button) → `sheetPost` → `loadPage(0)` (selection resets since positions shifted).
- Right-click on a row number or column header selects it (if not already selected) and opens a one-item `.ctx-menu` ("Delete N rows/columns").
- `Delete` key: deletes when the selection kind is whole rows/cols, otherwise clears cells (unchanged). `Backspace` always clears.

## 5. Tested
Live server round-trip (upload → apply an EXPR column that depends on another column → add a SUM summary row → delete columns, including two non-contiguous ones in one call, and a column added after an earlier delete) → re-fetched the live page and downloaded the `.xlsx`, reloaded with openpyxl: headers, values, the summary formula, and the EXPR recipe's formula all landed in the right (shifted) columns; the deleted columns were physically gone from the exported file, not just hidden. Did not test: stream-mode (very large) exports with column deletion (stream export already builds from `grid` directly, so it should be unaffected, but wasn't explicitly re-run); no browser/UI test was done (no Node/browser automation available in this session) — worth a manual click-through.

## 6. Limits / open items
- Deleting a column that another (app-generated) formula depends on doesn't warn you — that formula just quietly stops updating. Surfacing this (e.g. listing dependents before the confirm dialog) would need `referenced_columns` wired into the confirm step.
- Original in-file formulas referencing a column that shifted keep their old, now-wrong cell reference text (see Design above) — same limitation row deletion already has.
- No undo for row/column deletion (matches every other structural sheet op).
- A follow-up report ("Delete only clears content") and its fix (cache-busting + a more forgiving selection check) are in [DELETE_SELECTION_FIX_CONTEXT.md](DELETE_SELECTION_FIX_CONTEXT.md).
