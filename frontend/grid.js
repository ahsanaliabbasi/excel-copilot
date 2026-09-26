/* grid.js — spreadsheet-style sheet view, built for big sheets.
 *
 * Only one page of rows (default 1,000) is in the browser at a time; the server holds the rest.
 * Every row is addressed by its ABSOLUTE index in the sheet (G.ids), so edits, undo, duplicate
 * highlighting and copy keep working when you move between pages.
 *
 *   selection   click, drag, Shift+click, row numbers, column letters, Ctrl+A
 *   editing     switch on "Allow editing"; double-click / F2 / just type; Enter/Tab/Esc; Delete; Ctrl+Z
 *   clipboard   Ctrl+C / X / V round-trip with Excel (tab-separated text); whole columns come from the server
 *   duplicates  found on the server across every row, coloured here
 */
(function () {
  "use strict";
  (window.FRONTEND_PARTS = window.FRONTEND_PARTS || {}).grid = 17;

  const G = {
    sheet: null, columns: [], letters: [], rows: [], formulas: [], ids: [], idPos: new Map(),
    headerRow: 1, total: 0, matched: 0, offset: 0, limit: 1000, filtered: false,
    anchor: { r: 0, c: 0 }, active: { r: 0, c: 0 }, rect: { r0: 0, c0: 0, r1: 0, c1: 0 }, sel: "cells",
    painted: [], editing: null, undo: [], drag: null, canEdit: false,
    dupSpec: null, dupInfo: null, dupSummary: null, dupTouched: [],
  };

  // ------------------------------------------------------------- helpers
  const wrap = () => $("#gridWrap");
  const tbody = () => document.querySelector("#grid tbody");
  const thead = () => document.querySelector("#grid thead");
  const nRows = () => G.rows.length;
  const nCols = () => G.columns.length;
  const cellEl = (r, c) => { const tr = tbody().rows[r]; return tr ? tr.cells[c + 1] : null; };
  const rowNum = (r) => G.headerRow + 1 + G.ids[r];
  const refOf = (r, c) => `${G.letters[c]}${rowNum(r)}`;
  const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
  const fmt = (n) => Number(n).toLocaleString("en-US");
  const disp = (v) => (v === null || v === undefined ? "" : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v));
  const focusGrid = () => wrap() && wrap().focus({ preventScroll: true });
  const api = (path, body) => fetch(API_BASE + path, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
  });
  // absolute row of a row on the current page (rows past the last one continue the sheet, so pasting can append)
  const absId = (r) => (r < G.ids.length ? G.ids[r] : G.filtered ? null : G.offset + r);

  // ------------------------------------------------------------- render
  function cellClass(v, f) {
    return ((typeof v === "number" ? "num" : typeof v === "boolean" ? "bool" : "") + (f ? " has-formula" : "")).trim();
  }
  function cellHtml(v, f) {
    const cls = cellClass(v, f);
    return `<td${cls ? ` class="${cls}"` : ""}>${esc(disp(v))}</td>`;
  }
  function buildTable() {
    thead().innerHTML =
      `<tr class="letters"><th class="corner all" title="Select all"></th>` +
      G.letters.map((l, c) => `<th class="letter" data-c="${c}">${l}</th>`).join("") + `</tr>` +
      `<tr class="names"><th class="corner rn">${G.headerRow}</th>` +
      G.columns.map((n, c) => `<th class="name" data-c="${c}" title="${esc(n)}">${esc(n)}</th>`).join("") + `</tr>`;
    const parts = [];
    for (let r = 0; r < G.rows.length; r++) {
      const row = G.rows[r], fr = G.formulas[r] || [];
      let tds = "";
      for (let c = 0; c < row.length; c++) tds += cellHtml(row[c], fr[c]);
      parts.push(`<tr data-r="${r}"><td class="rn">${rowNum(r)}</td>${tds}</tr>`);
    }
    tbody().innerHTML = parts.join("");
    G.painted = [];
    G.dupTouched = [];
  }

  // data = a page from the server. opts.resetSel: start with the first cell selected.
  window.renderGrid = function (data, opts = {}) {
    const same = G.sheet === state.currentSheet && G.columns.length > 0;
    const keep = same && !opts.resetSel
      ? { active: G.active, anchor: G.anchor, rect: G.rect, sel: G.sel, top: wrap().scrollTop, left: wrap().scrollLeft } : null;
    finishEdit(false);
    if (!same) { G.undo = []; G.dupSpec = null; }
    G.sheet = state.currentSheet;
    G.columns = data.columns; G.letters = data.col_letters; G.rows = data.rows; G.formulas = data.formulas;
    G.ids = data.ids; G.idPos = new Map(data.ids.map((id, i) => [id, i]));
    G.headerRow = data.header_row; G.total = data.total_rows; G.matched = data.matched_rows;
    G.offset = data.offset; G.limit = data.limit;
    G.filtered = !!(G.dupSpec && G.dupSpec.only);
    G.dupInfo = data.dup; G.dupSummary = data.dup_summary;
    buildTable();
    if (!nRows() || !nCols()) { paintBars(); refreshLabels(); paintPager(); paintDup(); return; }
    const clampPos = (p) => ({ r: clamp(p.r, 0, nRows() - 1), c: clamp(p.c, 0, nCols() - 1) });
    if (keep) {
      G.active = clampPos(keep.active); G.anchor = clampPos(keep.anchor); G.sel = keep.sel;
      G.rect = {
        r0: clamp(keep.rect.r0, 0, nRows() - 1), r1: clamp(keep.rect.r1, 0, nRows() - 1),
        c0: clamp(keep.rect.c0, 0, nCols() - 1), c1: clamp(keep.rect.c1, 0, nCols() - 1),
      };
      wrap().scrollTop = keep.top; wrap().scrollLeft = keep.left;
    } else {
      G.active = G.anchor = { r: 0, c: 0 };
      G.rect = { r0: 0, c0: 0, r1: 0, c1: 0 };
      G.sel = "cells";
      wrap().scrollTop = 0;
    }
    paint();
    paintDup();
    refreshLabels();
    paintPager();
  };

  function refreshLabels() {
    $("#rowCountLabel").textContent = G.filtered
      ? `${fmt(G.matched)} duplicate rows of ${fmt(G.total)}` : `${fmt(G.total)} rows`;
    const meta = state.sheets.find((s) => s.name === G.sheet);
    if (meta) { meta.row_count = G.total; meta.columns = G.columns; meta.loaded = true; renderSheetList(); }
  }

  // ------------------------------------------------------------ paging
  function setBusy(on) { if (wrap()) wrap().classList.toggle("loading", on); }

  async function fetchPage(name, offset) {
    const res = await api("/api/sheet/page", {
      session_id: state.sessionId, sheet_name: name, offset, limit: G.limit,
      dup: name === G.sheet ? G.dupSpec : null,
    });
    const data = await res.json();
    if (!res.ok) throw new Error(typeof data.detail === "string" ? data.detail : "Couldn't load the rows");
    return data;
  }

  // Open a sheet (first page). Called by app.js when you pick a sheet.
  window.gridLoadSheet = async function (name, opts = {}) {
    const same = G.sheet === name;
    if (!same) G.dupSpec = null;
    const offset = same && !opts.resetPage ? G.offset : 0;
    setBusy(true);
    try {
      const data = await fetchPage(name, offset);
      window.renderGrid(data, { resetSel: !same || !!opts.resetPage });
    } catch (e) {
      toast(e.message, true);
    } finally { setBusy(false); }
  };

  async function loadPage(offset, opts = {}) {
    finishEdit(true);
    setBusy(true);
    try {
      const data = await fetchPage(G.sheet, Math.max(0, offset));
      window.renderGrid(data, { resetSel: opts.resetSel !== false });
      if (opts.selectId !== undefined) {
        const p = G.idPos.get(opts.selectId);
        if (p !== undefined) selectCell(p, G.active.c, false);
      }
    } catch (e) {
      toast(e.message, true);
    } finally { setBusy(false); }
  }

  function lastOffset() { return Math.max(0, Math.floor((G.matched - 1) / G.limit) * G.limit); }
  function paintPager() {
    const pager = $("#pager");
    const big = G.matched > Math.min(G.limit, 250) || G.offset > 0;
    pager.hidden = !big;
    if (!big) return;
    const first = G.matched ? G.offset + 1 : 0, last = G.offset + nRows();
    $("#pagerInfo").textContent = `Rows ${fmt(first)}–${fmt(last)} of ${fmt(G.matched)}${G.filtered ? " (duplicates only)" : ""}`;
    $("#pgFirst").disabled = $("#pgPrev").disabled = G.offset <= 0;
    $("#pgNext").disabled = $("#pgLast").disabled = G.offset >= lastOffset();
    $("#pgGoto").disabled = $("#pgGo").disabled = G.filtered;
    $("#pgSize").value = String(G.limit);
  }

  // ------------------------------------------------------------ selection
  function setSelection(anchor, active, rect, scroll = true, kind = "cells") {
    G.anchor = anchor; G.active = active; G.rect = rect; G.sel = kind;
    paint();
    if (scroll) { const el = cellEl(active.r, active.c); if (el) el.scrollIntoView({ block: "nearest", inline: "nearest" }); }
  }
  function selectCell(r, c, extend, scroll = true) {
    r = clamp(r, 0, nRows() - 1); c = clamp(c, 0, nCols() - 1);
    const anchor = extend ? G.anchor : { r, c };
    setSelection(anchor, { r, c }, {
      r0: Math.min(anchor.r, r), r1: Math.max(anchor.r, r), c0: Math.min(anchor.c, c), c1: Math.max(anchor.c, c),
    }, scroll);
  }
  function selectRows(r, extend) {
    const a = extend ? G.anchor.r : r;
    setSelection({ r: a, c: 0 }, { r, c: 0 }, { r0: Math.min(a, r), r1: Math.max(a, r), c0: 0, c1: nCols() - 1 }, false, "rows");
  }
  function selectCols(c, extend) {
    const a = extend ? G.anchor.c : c;
    setSelection({ r: 0, c: a }, { r: 0, c }, { r0: 0, r1: nRows() - 1, c0: Math.min(a, c), c1: Math.max(a, c) }, false, "cols");
  }
  function selectAll() {
    setSelection({ r: 0, c: 0 }, { r: 0, c: 0 }, { r0: 0, r1: nRows() - 1, c0: 0, c1: nCols() - 1 }, false, "all");
  }
  // the selection covers more rows than this page holds (whole column / whole sheet)
  const spansPages = () => (G.sel === "cols" || G.sel === "all") && !G.filtered && G.matched > nRows();

  function paint() {
    G.painted.forEach(([el, cls]) => el.classList.remove(cls));
    G.painted = [];
    if (!nRows() || !nCols()) return;
    const add = (el, cls) => { if (el) { el.classList.add(cls); G.painted.push([el, cls]); } };
    const { r0, c0, r1, c1 } = G.rect;
    const single = r0 === r1 && c0 === c1;
    const rows = tbody().rows, heads = thead().rows;
    for (let r = r0; r <= r1; r++) {
      const tr = rows[r];
      if (!tr) continue;
      add(tr.cells[0], "hl");                                   // row number
      if (!single) for (let c = c0; c <= c1; c++) add(tr.cells[c + 1], "sel");
    }
    for (let c = c0; c <= c1; c++) { add(heads[0].cells[c + 1], "hl"); add(heads[1].cells[c + 1], "hl"); }
    add(rows[G.active.r], "row-active");                          // the whole record you are on
    add(cellEl(G.active.r, G.active.c), "cell-active");
    paintBars();
  }

  function paintBars() {
    const ref = $("#cellRef"), box = $("#formulaText"), stats = $("#selStats");
    if (!nRows() || !nCols()) { ref.textContent = ""; box.textContent = ""; stats.textContent = ""; return; }
    const { r0, c0, r1, c1 } = G.rect;
    ref.textContent = r0 === r1 && c0 === c1 ? refOf(r0, c0) : `${refOf(r0, c0)}:${refOf(r1, c1)}`;
    const f = (G.formulas[G.active.r] || [])[G.active.c];
    box.textContent = f || disp(G.rows[G.active.r][G.active.c]);
    box.classList.toggle("is-formula", !!f);
    if (spansPages()) { stats.textContent = `${fmt(G.matched)} rows × ${c1 - c0 + 1} column${c1 > c0 ? "s" : ""} (all pages)`; return; }
    if (r0 === r1 && c0 === c1) { stats.textContent = ""; return; }
    let count = 0, sum = 0;
    const cells = (r1 - r0 + 1) * (c1 - c0 + 1);
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) {
      const v = G.rows[r][c];
      if (typeof v === "number") { count++; sum += v; }
    }
    const f2 = (x) => (Number.isInteger(x) ? String(x) : x.toFixed(2));
    stats.textContent = count
      ? `${fmt(cells)} cells · Sum ${f2(sum)} · Avg ${f2(sum / count)} · Count ${fmt(count)}`
      : `${fmt(cells)} cells selected`;
  }

  // ---------------------------------------------------------------- mouse
  function cellOf(target) {
    const td = target.closest && target.closest("td");
    if (!td || !td.parentElement || td.parentElement.dataset.r === undefined) return null;
    return { td, r: +td.parentElement.dataset.r, c: td.cellIndex - 1 };       // c === -1 -> row number
  }
  function initEvents() {
    const body = tbody(), head = thead();
    body.addEventListener("mousedown", (e) => {
      if (e.button !== 0) return;
      const hit = cellOf(e.target);
      if (!hit) return;
      if (G.editing && G.editing.td === hit.td) return;               // clicking inside the cell being edited
      finishEdit(true);
      e.preventDefault();
      focusGrid();
      if (hit.c < 0) { selectRows(hit.r, e.shiftKey); G.drag = "row"; }
      else { selectCell(hit.r, hit.c, e.shiftKey, false); G.drag = "cell"; }
    });
    body.addEventListener("mouseover", (e) => {
      if (!G.drag) return;
      const hit = cellOf(e.target);
      if (!hit) return;
      if (G.drag === "row") selectRows(hit.r, true);
      else if (hit.c >= 0) selectCell(hit.r, hit.c, true, false);
    });
    document.addEventListener("mouseup", () => { G.drag = null; });
    body.addEventListener("dblclick", (e) => {
      const hit = cellOf(e.target);
      if (hit && hit.c >= 0) startEdit(hit.r, hit.c);
    });
    head.addEventListener("click", (e) => {
      const th = e.target.closest("th");
      if (!th) return;
      finishEdit(true);
      focusGrid();
      if (th.classList.contains("all")) selectAll();
      else if (th.dataset.c !== undefined) selectCols(+th.dataset.c, e.shiftKey);
    });
    wrap().addEventListener("keydown", onKey);
  }

  // ------------------------------------------------------------- keyboard
  function move(dr, dc, extend) { selectCell(G.active.r + dr, G.active.c + dc, extend); }
  function onKey(e) {
    if (G.editing || !nRows() || !nCols()) return;
    const mod = e.ctrlKey || e.metaKey;
    const k = e.key;
    if (mod && (k === "z" || k === "Z")) { e.preventDefault(); undo(); return; }
    if (mod && (k === "a" || k === "A")) { e.preventDefault(); selectAll(); return; }
    if (mod) return;                                                        // Ctrl+C / X / V use the clipboard events
    const jump = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[k];
    if (jump) { e.preventDefault(); move(jump[0], jump[1], e.shiftKey); return; }
    if (k === "Tab") { e.preventDefault(); move(0, e.shiftKey ? -1 : 1, false); return; }
    if (k === "Enter") { e.preventDefault(); move(e.shiftKey ? -1 : 1, 0, false); return; }
    if (k === "Home") { e.preventDefault(); selectCell(G.active.r, 0, e.shiftKey); return; }
    if (k === "End") { e.preventDefault(); selectCell(G.active.r, nCols() - 1, e.shiftKey); return; }
    if (k === "PageDown") { e.preventDefault(); move(20, 0, e.shiftKey); return; }
    if (k === "PageUp") { e.preventDefault(); move(-20, 0, e.shiftKey); return; }
    if (k === "F2") { e.preventDefault(); startEdit(G.active.r, G.active.c); return; }
    if (k === "Delete" || k === "Backspace") { e.preventDefault(); clearSelection(); return; }
    if (k.length === 1 && !e.altKey) { e.preventDefault(); startEdit(G.active.r, G.active.c, k); }
  }

  // --------------------------------------------------------------- editing
  // The sheet is read-only until "Allow editing" is switched on.
  let lastNag = 0;
  function requireEdit() {
    if (G.canEdit) return true;
    if (Date.now() - lastNag > 2500) { toast("The sheet is read-only. Turn on “Allow editing” in the toolbar to change cells."); lastNag = Date.now(); }
    return false;
  }
  function setEditable(on) {
    if (!on) finishEdit(true);
    G.canEdit = on;
    $("#grid").classList.toggle("editable", on);
    $("#editToggle").checked = on;
    $("#editToggleLabel").textContent = on ? "Editing on" : "Allow editing";
    $("#btnPaste").disabled = !on;
    $("#btnUndo").disabled = !on;
    $("#gridHint").textContent = on
      ? "Editing is ON — click a cell and type (or double-click) · Enter/Tab move · Delete clears · Ctrl+V pastes · Ctrl+Z undoes"
      : "Read-only — turn on “Allow editing” to change cells · drag or Shift+click to select · Ctrl+C copies to Excel";
  }

  function startEdit(r, c, initial) {
    if (r >= nRows() || c >= nCols()) return;
    if (!requireEdit()) return;
    finishEdit(true);
    const td = cellEl(r, c);
    const input = document.createElement("input");
    input.className = "cell-input";
    input.value = initial !== undefined ? initial : disp(G.rows[r][c]);
    G.editing = { r, c, td, input, old: td.textContent };
    td.classList.add("editing");
    td.textContent = "";
    td.append(input);
    input.focus();
    input.setSelectionRange(input.value.length, input.value.length);
    input.addEventListener("input", () => { $("#formulaText").textContent = input.value; $("#formulaText").classList.remove("is-formula"); });
    input.addEventListener("keydown", (e) => {
      e.stopPropagation();
      if (e.key === "Enter") { e.preventDefault(); finishEdit(true); move(e.shiftKey ? -1 : 1, 0, false); }
      else if (e.key === "Tab") { e.preventDefault(); finishEdit(true); move(0, e.shiftKey ? -1 : 1, false); }
      else if (e.key === "Escape") { e.preventDefault(); finishEdit(false); paintBars(); }
    });
    input.addEventListener("blur", () => { if (G.editing && G.editing.input === input) finishEdit(true); });
  }

  function finishEdit(commit) {
    const ed = G.editing;
    if (!ed) return;
    G.editing = null;
    const text = ed.input.value;
    ed.td.classList.remove("editing");
    ed.td.textContent = ed.old;                    // the real value comes back from the server
    if (document.activeElement === document.body || ed.input === document.activeElement) focusGrid();
    if (!commit || text === ed.old) { paintBars(); return; }
    if (text.startsWith("=")) {
      toast("Type values in the grid. To create formulas use the Formula Builder on the left.", true);
      paintBars();
      return;
    }
    applyEdits([{ r: ed.r, c: ed.c, value: text }]);
  }

  // ---------------------------------------------------- talking to the server
  // edit = { r (row on this page) | id (absolute row), c, value, formula? }
  let chain = Promise.resolve();
  function applyEdits(edits, opts = {}) {
    chain = chain.then(() => doApply(edits, opts)).catch(() => false);
    return chain;
  }
  async function doApply(edits, opts) {
    if (!G.sheet) return false;
    const list = edits.map((e) => ({ ...e, id: e.id !== undefined ? e.id : absId(e.r) })).filter((e) => e.id !== null && e.id >= 0);
    if (!list.length) return false;
    const prev = list.map((e) => {
      const p = G.idPos.get(e.id);
      return { id: e.id, c: e.c, value: p !== undefined ? G.rows[p][e.c] : null,
               formula: p !== undefined ? (G.formulas[p] || [])[e.c] || null : null };
    });
    let res, data;
    try {
      res = await api("/api/sheet/edit", {
        session_id: state.sessionId, sheet_name: G.sheet,
        edits: list.map((e) => ({ row: e.id, col: e.c, value: e.value ?? null, formula: e.formula || null })),
      });
      data = await res.json();
    } catch (err) {
      toast("Can't reach the server: " + err.message, true);
      return false;
    }
    if (!res.ok) { toast(typeof data.detail === "string" ? data.detail : "That edit was rejected.", true); return false; }
    if (!opts.isUndo) { G.undo.push(prev); if (G.undo.length > 100) G.undo.shift(); }
    if (data.reload || G.dupSpec) {                 // rows were added, or duplicate groups may have changed
      await loadPage(G.offset, { resetSel: false });
    } else {
      data.changes.forEach((ch) => {
        const r = G.idPos.get(ch.row);
        if (r === undefined || ch.c >= nCols()) return;
        G.rows[r][ch.c] = ch.value;
        (G.formulas[r] = G.formulas[r] || [])[ch.c] = ch.formula;
        const td = cellEl(r, ch.c);
        if (td && !(G.editing && G.editing.td === td)) {
          td.textContent = disp(ch.value);
          td.className = cellClass(ch.value, ch.formula) + (td.classList.contains("cell-active") ? " cell-active" : "") + (td.classList.contains("sel") ? " sel" : "");
        }
      });
      G.total = data.total_rows;
      paintBars();
      refreshLabels();
    }
    return true;
  }

  function clearSelection() {
    if (!requireEdit()) return;
    const { r0, c0, r1, c1 } = G.rect, edits = [];
    for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) if (disp(G.rows[r][c]) !== "") edits.push({ r, c, value: null });
    if (spansPages()) toast(`Cleared the ${fmt(edits.length)} cells on this page — other pages weren't touched.`);
    if (edits.length) applyEdits(edits);
  }
  function undo() {
    if (!requireEdit()) return;
    const prev = G.undo.pop();
    if (!prev) { toast("Nothing to undo"); return; }
    applyEdits(prev.map((p) => ({ id: p.id, c: p.c, value: p.value, formula: p.formula })), { isUndo: true })
      .then((ok) => ok && toast("Undone"));
  }

  // ------------------------------------------------------------- clipboard
  const quoteField = (s) => (/[\t\n\r"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s);
  const selectionSize = () => (G.rect.r1 - G.rect.r0 + 1) * (G.rect.c1 - G.rect.c0 + 1);
  function selectionText(withHeaders) {
    const { r0, c0, r1, c1 } = G.rect, lines = [];
    if (withHeaders) lines.push(G.columns.slice(c0, c1 + 1).map((n) => quoteField(String(n))).join("\t"));
    for (let r = r0; r <= r1; r++) lines.push(G.rows[r].slice(c0, c1 + 1).map((v) => quoteField(disp(v))).join("\t"));
    return lines.join("\n");
  }
  // whole columns / sheet come from the server so the copy includes every page
  async function copyPayload(withHeaders) {
    if (!spansPages()) return { text: selectionText(withHeaders), cells: selectionSize() };
    const res = await api("/api/sheet/range-text", {
      session_id: state.sessionId, sheet_name: G.sheet, c0: G.rect.c0, c1: G.rect.c1, headers: withHeaders,
    });
    const data = await res.json();
    if (!res.ok) throw new Error(typeof data.detail === "string" ? data.detail : "Couldn't copy that");
    return data;
  }
  async function copyText(text) {
    try { await navigator.clipboard.writeText(text); return true; }
    catch (e) {
      const ta = document.createElement("textarea");
      ta.value = text; ta.style.cssText = "position:fixed;opacity:0";
      document.body.append(ta); ta.select();
      let ok = false;
      try { ok = document.execCommand("copy"); } catch (e2) { /* ignore */ }
      ta.remove(); focusGrid();
      return ok;
    }
  }
  async function copySelection(withHeaders) {
    if (!nRows()) return;
    try {
      if (spansPages()) toast("Collecting every row…");
      const { text, cells } = await copyPayload(withHeaders);
      const ok = await copyText(text);
      toast(ok ? `Copied ${fmt(cells)} cell${cells === 1 ? "" : "s"}${withHeaders ? " with column headers" : ""}` : "Copy failed — try Ctrl+C", !ok);
    } catch (e) { toast(e.message, true); }
  }

  function parseTSV(text) {
    const rows = [];
    let row = [], cur = "", i = 0, inQ = false;
    while (i < text.length) {
      const ch = text[i];
      if (inQ) {
        if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i += 2; continue; } inQ = false; i++; continue; }
        cur += ch; i++; continue;
      }
      if (ch === '"' && cur === "") { inQ = true; i++; continue; }
      if (ch === "\t") { row.push(cur); cur = ""; i++; continue; }
      if (ch === "\r") { i++; continue; }
      if (ch === "\n") { row.push(cur); rows.push(row); row = []; cur = ""; i++; continue; }
      cur += ch; i++;
    }
    row.push(cur); rows.push(row);
    return rows;
  }
  function pasteText(text) {
    if (!text || !nRows() || !nCols()) return;
    if (!requireEdit()) return;
    const rows = parseTSV(text);
    if (rows.length > 1 && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === "") rows.pop();
    const { r0, c0, r1, c1 } = G.rect, edits = [];
    let clipped = false, endR = r0, endC = c0;
    if (rows.length === 1 && rows[0].length === 1 && (r1 > r0 || c1 > c0)) {          // one value -> fill the selection
      for (let r = r0; r <= r1; r++) for (let c = c0; c <= c1; c++) edits.push({ r, c, value: rows[0][0] });
      endR = r1; endC = c1;
      if (spansPages()) toast("Filled the rows on this page — other pages weren't touched.");
    } else {
      rows.forEach((row, i) => row.forEach((v, j) => {
        const c = c0 + j;
        if (c >= nCols() || absId(r0 + i) === null) { clipped = true; return; }
        edits.push({ r: r0 + i, c, value: v });
        endR = Math.max(endR, r0 + i); endC = Math.max(endC, c);
      }));
    }
    if (!edits.length) return;
    applyEdits(edits).then((ok) => {
      if (!ok) return;
      const rr = Math.min(endR, nRows() - 1);
      setSelection({ r: r0, c: c0 }, { r: r0, c: c0 }, { r0, c0, r1: rr, c1: endC }, false);
      toast(`Pasted ${fmt(edits.length)} cell${edits.length === 1 ? "" : "s"}` + (clipped ? " — cells beyond the last column were skipped" : ""));
    });
  }

  const gridReady = () =>
    G.columns.length && !G.editing && $("#modalOverlay").hidden &&
    !/^(INPUT|TEXTAREA|SELECT)$/.test((document.activeElement || {}).tagName || "");
  document.addEventListener("copy", (e) => {
    if (!gridReady()) return;
    e.preventDefault();
    if (spansPages()) { copySelection(false); return; }
    e.clipboardData.setData("text/plain", selectionText(false));
    toast(`Copied ${fmt(selectionSize())} cell${selectionSize() === 1 ? "" : "s"}`);
  });
  document.addEventListener("cut", (e) => {
    if (!gridReady()) return;
    e.preventDefault();
    if (spansPages()) { copySelection(false); return; }
    e.clipboardData.setData("text/plain", selectionText(false));
    if (!G.canEdit) { toast(`Copied ${fmt(selectionSize())} cell${selectionSize() === 1 ? "" : "s"} (read-only, so nothing was removed)`); return; }
    clearSelection();
    toast(`Cut ${fmt(selectionSize())} cells`);
  });
  document.addEventListener("paste", (e) => {
    if (!gridReady()) return;
    e.preventDefault();
    pasteText(e.clipboardData.getData("text/plain"));
  });

  // ------------------------------------------------------------ duplicates
  // The server finds them across every row; this only colours the rows on the current page.
  function paintDup() {
    G.dupTouched.forEach(([el, cls]) => { if (cls === "--h") el.style.removeProperty("--h"); else el.classList.remove(cls); });
    G.dupTouched = [];
    tbody().querySelectorAll("td.rn[data-dup]").forEach((td) => { delete td.dataset.dup; });
    if (!G.dupInfo || !G.dupSpec) { paintDupResult(null); return; }
    const touch = (el, cls) => { el.classList.add(cls); G.dupTouched.push([el, cls]); };
    G.dupInfo.forEach((info, r) => {
      if (!info) return;
      const tr = tbody().rows[r];
      if (!tr) return;
      touch(tr, "dup");
      tr.style.setProperty("--h", Math.round((info[0] * 137.5) % 360));
      G.dupTouched.push([tr, "--h"]);
      tr.cells[0].dataset.dup = "×" + info[1];
      G.dupSpec.cols.forEach((c) => { if (tr.cells[c + 1]) touch(tr.cells[c + 1], "dupkey"); });
    });
    paintDupResult(G.dupSummary);
  }
  function paintDupResult(res) {
    const out = $("#dupResult");
    if (!out) return;
    if (!res) { out.textContent = ""; out.className = "dup-result"; return; }
    out.textContent = res.groups
      ? `${fmt(res.rows)} rows share their values with another row — ${fmt(res.groups)} group${res.groups === 1 ? "" : "s"} (checked all ${fmt(res.checked)} rows).`
      : `No duplicates found in ${fmt(res.checked)} rows.`;
    out.className = "dup-result " + (res.groups ? "found" : "none");
  }

  // The Duplicates screen (dup.js) drives the in-grid highlight/filter through these two calls,
  // so "view these rows in the grid" stays in sync with everything else the grid already does
  // (paging, the row-number badge, colouring the key columns).
  window.gridSetDupView = async function (spec) {
    if (state.currentSheet !== G.sheet) await window.gridLoadSheet(state.currentSheet);
    G.dupSpec = spec;
    await loadPage(0);
  };
  window.gridClearDupView = async function () {
    if (!G.dupSpec) return;
    G.dupSpec = null;
    await loadPage(0);
  };

  // ----------------------------------------------------------------- toolbar
  initEvents();
  $("#btnCopy").addEventListener("click", () => copySelection(false));
  $("#btnCopyH").addEventListener("click", () => copySelection(true));
  $("#btnPaste").addEventListener("click", async () => {
    try { pasteText(await navigator.clipboard.readText()); }
    catch (e) { toast("Your browser blocked clipboard access — click a cell and press Ctrl+V instead.", true); }
  });
  const bandSelect = $("#bandSelect");
  let savedBand = "blue";
  try { savedBand = localStorage.getItem("rowBand") || "blue"; } catch (e) { /* storage blocked */ }
  bandSelect.value = savedBand;
  $("#grid").dataset.band = bandSelect.value;
  bandSelect.addEventListener("change", () => {
    $("#grid").dataset.band = bandSelect.value;
    try { localStorage.setItem("rowBand", bandSelect.value); } catch (e) { /* ignore */ }
  });
  $("#editToggle").addEventListener("change", (e) => setEditable(e.target.checked));
  setEditable(false);
  $("#btnUndo").addEventListener("click", undo);
  $("#btnDup").addEventListener("click", () => { if (typeof window.openDup === "function") window.openDup(); });

  // pager
  $("#pgFirst").addEventListener("click", () => loadPage(0));
  $("#pgPrev").addEventListener("click", () => loadPage(G.offset - G.limit));
  $("#pgNext").addEventListener("click", () => loadPage(G.offset + G.limit));
  $("#pgLast").addEventListener("click", () => loadPage(lastOffset()));
  $("#pgSize").addEventListener("change", (e) => { G.limit = parseInt(e.target.value, 10) || 1000; loadPage(0); });
  const goto = () => {
    const row = parseInt($("#pgGoto").value, 10);
    if (!row) return;
    const id = clamp(row - G.headerRow - 1, 0, Math.max(0, G.matched - 1));
    loadPage(Math.floor(id / G.limit) * G.limit, { selectId: id });
  };
  $("#pgGo").addEventListener("click", goto);
  $("#pgGoto").addEventListener("keydown", (e) => { if (e.key === "Enter") goto(); });
})();
