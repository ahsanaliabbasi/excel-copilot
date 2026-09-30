# Excel Copilot — "Delete only clears content, doesn't delete the row/column" (Session Context)

Follow-up to [DELETE_ROWS_AND_COLUMNS_CONTEXT.md](DELETE_ROWS_AND_COLUMNS_CONTEXT.md), which added whole-row/column deletion (right-click, or select + Delete).

## 1. Report
After that feature shipped, the user selected a row and pressed Delete — it cleared the row's cell contents instead of removing the row entirely.

## 2. Investigation
Re-read `grid.js`'s selection/keyboard code end to end; the row/column-delete routing (`G.sel === "rows"/"cols"` → `deleteRows()`/`deleteCols()`) was logically correct. Two real, non-exclusive causes were identified instead of a logic bug:

1. **Stale cached frontend.** The user runs a plain `python -m http.server 5500` to serve the frontend, separate from the FastAPI backend (port 8000, which explicitly sends `no-store` for the page files — see `NoCacheStatic` in `main.py`). A bare `http.server` does not send cache-busting headers, so a browser can easily keep serving the pre-fix `grid.js` across reloads even though the file on disk changed. (This process was found already running in an earlier turn of this session, started by the user, not by the assistant.)
2. **Selection method mismatch.** The original implementation only treated an explicit click on the **row number** (left gutter) or **column letter/name** (top) — i.e. `G.sel === "rows"` / `"cols"` — as "the whole row/column is selected." Dragging across every cell of a row (without touching the gutter) leaves `G.sel === "cells"`, which fell through to "just clear the selected cells." This is a very plausible way to reproduce exactly the symptom reported, and is arguably the more natural way a first-time user would try to "select a row."

Both were fixed; either one alone could explain the report, so both were addressed rather than guessing which applied.

## 3. Fixes
- **Cache-busting**: bumped `FRONTEND_VERSION` 18 → 19 in `backend/main.py`, the matching `FRONTEND_VERSION` check in `app.js`, the six `(window.FRONTEND_PARTS...).<name> = `  markers (`app` has none of its own; `grid`, `builder`, `sql`, `dup`, `tools` do), and every `?v=18` → `?v=19` query string in `index.html` (the five scripts + `style.css`). This forces a fresh fetch of every frontend file on the next load, regardless of which static server is in front of it.
- **Selection heuristic** (`grid.js`): added `isWholeRows()` / `isWholeCols()`, used by the `Delete` key handler in `onKey` in place of a bare `G.sel === "rows"/"cols"` check.
  - True whenever `G.sel` is already `"rows"`/`"cols"` (gutter click — unchanged, still the most reliable path), **or** the current `"cells"` selection rectangle spans every column (whole row) / every row (whole column) on the page.
  - Guarded so a genuine single-cell selection (`r0===r1 && c0===c1`) never counts, even in a degenerate 1-row or 1-column sheet — a plain click there still just clears the cell, it doesn't silently delete the row/column.
  - Right-click context menu was left as-is (row-number / column-header only) — it was never the reported problem, and it's already an unambiguous, deliberate action.

## 4. Guidance given to the user
- Restart the backend process and hard-refresh the browser (Ctrl+F5) to guarantee the new files load.
- Clicking the row number / column letter is still the most explicit, reliable way to select a whole row/column — it highlights the gutter and works with both Delete and right-click.

## 5. Tested
`python -c "import main"` after the `main.py` version bump (imports cleanly). No live browser re-test was possible in this session (no Node.js / browser-automation tool available in the sandbox) — the fix is a straightforward, reviewed code change, but the user should confirm directly after a hard refresh.

## 6. Open items
- If the user still sees only cell-clearing after a hard refresh + backend restart, the next thing to check is the browser console for JS errors on load (would point at an actual syntax/logic bug rather than caching), and exactly how they're serving the frontend (port 8000 vs the separate `:5500` static server).
