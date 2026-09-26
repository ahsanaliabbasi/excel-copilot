# Excel Copilot

A no-code web interface for editing real `.xlsx` files. Click through guided steps instead of writing formulas — the app writes **native Excel formulas** (with cached values) into the downloaded workbook, so everything keeps recalculating in Excel.

## What it does

- **Lookup / Match** — XLOOKUP, VLOOKUP, INDEX-MATCH, plus counts, totals, averages, highest/lowest per row, extra AND/OR conditions on either sheet, and IF / ELSE IF cases. A step-by-step wizard (Back / Next) walks you through it.
- **Totals & Counts** — SUMIF, COUNTIF, AVERAGEIF, MAXIFS/MINIFS with condition groups.
- **Formula Builder** — nested IF decision flows, text (TRIM, LEFT/RIGHT/MID, SUBSTITUTE, FIND/SEARCH, TEXT, LEN), rounding, dates (EDATE, TODAY), calculations.
- **SQL Query** — every sheet is a table; run SELECT/UPDATE, translate a query into Excel formulas, or save results as a new sheet.
- **Duplicates** — choose any key columns, see an overview (total / duplicate groups / unique after dedup), check one specific value, browse groups, hand-pick which rows to keep or delete, copy the unique set to the clipboard, and either delete from the sheet or write the result to a new named sheet.
- **Excel-style grid** — paging for big sheets (tested on 12 sheets × 75,000 rows), optional cell editing, copy/paste with Excel, banded rows, undo.

## Run it

Requires Python 3.10+.

```bash
cd backend
python -m venv venv
venv\Scripts\activate            # macOS/Linux: source venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --port 8000
```

Then open **http://localhost:8000** — the backend serves the frontend itself (plain HTML/CSS/JS, no build step). Press Ctrl+F5 after updating so the browser doesn't keep an old copy.

## How it is built

```
backend/
  main.py           FastAPI app: upload, paging, operations, editing, duplicates, SQL, background download
  expr_engine.py    the step-tree engine: each step produces an Excel formula AND the value shown in the grid
  sql_engine.py     SQLite sandbox over the sheets + SQL -> Excel-formula translation
  formula_engine.py legacy operation helpers
frontend/
  index.html, style.css
  app.js            state, upload, download, modals
  grid.js           paged spreadsheet view, selection, editing, clipboard
  builder.js        the wizard / step builder
  sql.js            SQL screen
  dup.js            Duplicates screen
```

Uploads are read with `python-calamine` (fast), sheets load lazily, and changes are recorded as an overlay that is written on download (`openpyxl` for small workbooks, streamed `xlsxwriter` for big ones).

## Known limits

- Sessions live in server memory; a restart loses uploaded files.
- No authentication — run it locally or behind your own access control.
- Generated formulas are checked against independent calculations, but not opened in every Excel version.
- Streamed downloads of very large workbooks keep data and formulas but not cell styling.
