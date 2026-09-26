# Excel Copilot — Focused Wizard + Grid Filter/Sort (Session Context)

Companion to [PROJECT_CONTEXT.md](PROJECT_CONTEXT.md) and [UI_REDESIGN_CONTEXT.md](UI_REDESIGN_CONTEXT.md). Records the work done after the redesign: what was asked, what changed, decisions, and open items.

---

## 1. Request

1. **Wizard focus.** In the Lookup / Match modal (a screenshot showed the RESULT dropdown, "In plain words" box, step card, "Put result in" and Live preview all at once), show only the current step and hide everything below it, so the user can focus on the question in front of them.
2. **Column filter and sort** for every grid column:
   - Text: filter by "starts with" or "contains" a substring.
   - Numbers: filter by equal to a specific number.
   - Dates and other types: the equivalent.
   - Sort: two arrows, ascending and descending.

## 2. Wizard focus (`builder.js`, `style.css`)

Applies to the two wizard tools (Lookup / Match, Totals & Counts). Other tools are unchanged.

- `wizard()` sets `B.wizard = true` and `B.wizLast` (is this the last step) each time it renders.
- `redraw()` resets `B.wizard = false` before rendering, then reads the flags:
  - toggles `wiz-mode` / `wiz-focus` on `#builderTree`,
  - hides `#builderTail` and `#applyBtn` unless on the last step.
- `#builderTail` is a new wrapper in `openBuilder()` around "Put result in" + Live preview.
- The "In plain words" title and box carry the class `wiz-extra`.
- CSS (end of `style.css`):
  - `wiz-mode` hides the root node's "Result" head/kind-select and flattens the node box.
  - `wiz-focus` hides `.wiz-extra` and `.test-link`.
- Result: steps 1 to N-1 show only progress dots, the question, and Back / Next. On the last step the summary, output column, preview, "Test this result" link and Apply appear.
- The last-step hint ("All set — choose where the result goes below, then press Apply.") was kept, because the extras now appear below it.

## 3. Column filter and sort

### Design
A "view" of the sheet: which rows are shown and in what order. The data never changes. It runs on the server over every row, so it works across pages. It reuses the existing `ids` mechanism (absolute row ids), so row numbers stay the real sheet row numbers, and editing, undo and copy keep working.

### Backend (`main.py`)
- `POST /api/sheet/page` accepts `view: {filters: [{col, op, value}], sort: {col, dir: "asc"|"desc"}}`. The `col` values are column indexes.
- Helpers, defined just above `_grid_payload`:
  - `_col_kind(series)` returns `number`, `date` or `text`, judged from the first ~500 non-empty values.
  - `_filter_mask(...)` builds the row mask:
    - text ops: `contains`, `notcontains`, `starts`, `ends`, `equals`, `empty`, `notempty` (case-insensitive, compared on the displayed text).
    - number ops: `=`, `!=`, `>`, `>=`, `<`, `<=`, `empty`, `notempty`.
    - date ops: `on`, `before`, `after`, `empty`, `notempty` (compared as calendar dates).
  - `_sort_order(...)` sorts numerically, chronologically, or case-insensitively for text. Empty cells always go last.
  - `_view_ids(...)` applies all filters with AND, then the sort. Results are cached per session in `session["viewcache"]`, keyed by sheet, version and view; `_touch` bumps the version on data changes.
- `_grid_payload` combines the view with the Duplicates "only duplicates" ids (`np.isin`, sort order kept) and now returns `col_kinds`.
- Invalid input returns HTTP 400 with a readable message ("“a” isn't a number.", "isn't a date", wrong operator for type, column no longer exists).

### Frontend (`grid.js`, `index.html`, `style.css`)
- `G.view = {filters, sort}` and `G.kinds`. `fetchPage` sends `view` only for the current sheet. The view resets on sheet change and on `resetPage`.
- Every column header shows `▲ ▼` and a funnel icon (`headHtml`, inside `.nm-wrap` / `.hdr-tools`). Tool clicks are intercepted in the header click handler so they don't select the column.
  - Pressing the active sort arrow again clears the sort.
  - Active tools are highlighted in the accent colour; inactive ones are faint until the header is hovered.
- The funnel opens a popover (`openFilter`) that offers operators for the column's kind. The value box becomes a date picker for dates and is hidden for "is empty" / "is not empty". Enter applies; Escape or an outside click closes it. Text filters show the hint "Not case-sensitive."
- `applyView()` sets the new view, fetches page 1 and re-renders. If the server rejects it, the previous view is restored and the error is shown as a toast.
- `#viewBar` (new element above the grid) shows removable chips ("Model contains “abc”", "Sorted by custid ▼") and a "Clear all" link.
- `G.filtered` is now true for any active view, which disables paste-append beyond the last row, as it already did for the duplicates view.
- Labels: the row-count chip reads "N of M rows" when filtered; the pager adds "(filtered)".

## 4. Testing done

- Unit-style checks of `_view_ids` on a small DataFrame (text contains / starts with, number equals, date on / before, sort asc / desc with blanks, empty filter): all gave the expected ids.
- Live API test with the sample workbook on a temporary server (port 8001): filter + sort together, `col_kinds`, and the 400 error for a text value on a number column.
- Headless Edge driven over the DevTools protocol:
  - upload, switching to "Data Set 2 - Vehicles", clicking ▼ on `custid` (first rows became 1496, 1496, 1494…), and opening the filter popover (screenshot checked);
  - wizard step 1 showed no tail or Apply, and the last step showed both;
  - no JS exceptions.
- The temporary server and browser were stopped afterwards. The server on port 8000 was never touched. Test scripts stayed in the session scratchpad, not the repo.

**Not verified in the browser:** applying a text or date filter and paging through a filtered result; dark mode for the new header tools, popover and chips; combining a view with the Duplicates "only duplicates" view.

## 5. Notes and limits

- `main.py` changed, so the server must be restarted; the frontend changes need Ctrl+F5. `FRONTEND_VERSION` and `API_VERSION` were not bumped.
- A column mixing text and numbers counts as **text**; its numbers sort as text ("10" before "2"). Consider a majority-based kind detection.
- Editing a cell while sorted or filtered does not re-sort or re-filter. The row stays put until the view is reapplied.
- Multiple filters combine with AND only. There is one filter per column and a single-column sort.
- Text filters compare the displayed text, so numeric-looking text is filtered as text.
- `git status` shows `backend/__pycache__/*.pyc` as modified; it should be git-ignored.

## 6. Open items / ideas

- Multi-column sort, OR filters, and a filter value dropdown (distinct values).
- Persist the view when switching sheets.
- Keep an edited row visible under a filter, or offer "re-apply view".
- Visual check of the new header tools, popover and chips in dark mode.
- Earlier open items still stand (consolidate light-mode CSS, wizard for remaining tools, persisted sessions, undo for row deletion, auth).
