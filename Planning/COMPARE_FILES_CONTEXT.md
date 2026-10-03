# Excel Copilot — Compare Files (Session Context)

Companion to [PROJECT_CONTEXT.md](PROJECT_CONTEXT.md).

## 1. Request
Compare two pasted tables (e.g. Stripe export vs BigQuery export) side by side. A row must match wherever it sits (row 10 of A can match row 107 of B), matches are tinted, hovering shows/highlights the partner row and its line number, and the tool returns rows of A not in B and rows of B not in A. Both tables stay editable at all times.

## 2. Design
- **Frontend only** (`frontend/compare.js` + CSS at the end of `style.css`). No backend, no upload needed: a "Compare Files" button in the top bar opens a full-screen overlay (`window.openCompare`). Matching runs in the browser (hash-based, O(n)), debounced 220 ms after each edit.
- **Matching rule**: no Key column ticked -> a row matches when ALL mapped columns are equal. Key column(s) ticked -> rows match on the key; other mapped columns are compared and a row that differs is "changed" (differing cells highlighted). Duplicate rows pair one-to-one in order. Rows blank in all compared columns are ignored.
- **Column mapping** auto-built by header name (else by position); editable in the "Columns & keys" panel (choose counterpart, untick to ignore, tick Key).
- **Normalisation** options: ignore case, ignore extra spaces, numbers as numbers (`1.0 = 1 = 1,000.0`...). Dates are not parsed (compared as text).
- **View**: "Line up matching rows" (default; matched rows sit on the same line in both panes, gaps hatched, scroll synced, like a diff) or original order. Chips filter: All / Matching / Changed / Not found (only in A and only in B, each in its own pane).
- **Hover**: both rows highlighted + tooltip ("Match — row 107 in File B" / "Not found" / key matches but differs in ...). A muted `↔107` also sits in each row's gutter; click the row number to jump to and flash the partner.
- **Editing**: click a cell to edit (Enter/Tab/arrows move, Esc cancels, multi-cell paste inside a cell), click header text to rename, gutter `+`/`✕` insert/delete rows, header `✕` deletes a column, + Row, + Column, Paste data (replace or append), header-row toggle, Swap sides, Undo (Ctrl+Z, 40 steps).
- **Export**: copy as TSV (paste into Excel) or download CSV for only-in-A / only-in-B / matching rows.
- Virtualised rows (28 px), so 100k-row pastes stay smooth.
- `FRONTEND_VERSION` bumped 19 -> 20 (all files, `index.html` query strings, `main.py`); `compare` added to the stale-file check in `app.js`.
- GA4 events: `compare_opened`, `compare_loaded`, `compare_export`.

## 3. Tested
Headless Edge driving the real UI: pasted 200 vs 199 rows (shuffled order, one missing, one value changed, one extra) -> whole-row mode 198 matching / 2 only-in-A / 2 only-in-B; key mode 198 / 1 changed / 1 / 1; row 10 of A correctly pointed at row 190 of B. **Not tested in a browser**: cell editing, hover tooltip, jump, export, dark mode, mobile layout — needs a manual click-through.

## 4. Limits / ideas
- Compares exact text after normalisation; no fuzzy matching, no date parsing, no numeric tolerance.
- Column widths are fixed (150 px); full value shows on hover (title) and in the edit box.
- Not wired to the uploaded workbook's sheets (paste only).
