/* tools.js — working on the data like in Excel: summary rows (Sum / Average / Min / Max ...), new rows and columns,
 * renaming a column, and copying columns / rows into another sheet.
 * Sheet tabs (create / rename / duplicate / delete) live in app.js. Everything here goes through the server so it works
 * on every row of a big sheet, and the results are real Excel formulas in the downloaded file.
 */
(function () {
  "use strict";
  (window.FRONTEND_PARTS = window.FRONTEND_PARTS || {}).tools = 18;

  const grid = () => window.gridApi;
  const fmtNum = (v) => (typeof v === "number" ? v.toLocaleString("en-US", { maximumFractionDigits: 4 }) : String(v));
  const ready = () => {
    if (!state.sessionId || !grid() || !grid().ready()) { toast("Open a sheet first.", true); return false; }
    return true;
  };

  // ------------------------------------------------------------- summary rows
  // Works on the selected column(s): the result is added as a row pinned under the data, as a real formula
  // (=SUM(B2:B148)) that keeps updating when the data changes.
  async function summarize(fn) {
    if (!ready()) return;
    const sel = grid().selection(), cols = [];
    for (let c = sel.c0; c <= sel.c1; c++) cols.push(c);
    try {
      const d = await sheetPost("/api/sheet/summary", { sheet_name: grid().sheet(), cols, fn });
      await grid().reload();
      const shown = d.results.map((r) => `${r.name} = ${fmtNum(r.value)}`).join(" · ");
      toast(`${d.label}: ${shown} — added as a summary row` +
            (d.skipped.length ? ` (skipped, no numbers: ${d.skipped.join(", ")})` : ""));
    } catch (e) { toast(e.message, true); }
  }
  document.querySelectorAll(".sum-btn").forEach((b) => b.addEventListener("click", () => summarize(b.dataset.fn)));
  $("#sumMore").addEventListener("change", (e) => {
    const fn = e.target.value;
    e.target.value = "";
    if (fn) summarize(fn);
  });
  $("#grid").addEventListener("click", async (e) => {
    const x = e.target.closest(".sum-x");
    if (!x) return;
    e.stopPropagation();
    try {
      await sheetPost("/api/sheet/summary-remove", { sheet_name: grid().sheet(), index: +x.closest("tr").dataset.k });
      await grid().reload();
      toast("Summary row removed");
    } catch (err) { toast(err.message, true); }
  });

  // ------------------------------------------------------------- new row / new column
  $("#btnAddRow").addEventListener("click", () => { if (ready()) grid().addRows(1); });
  $("#btnAddCol").addEventListener("click", async () => {
    if (!ready()) return;
    const meta = currentSheetMeta();
    const vals = await askForm("New column", [
      { key: "name", label: "Column name", value: `Column ${(meta && meta.columns.length || 0) + 1}` },
    ], "Add column", "The column is added at the right-hand end. Fill it by typing, pasting, or with a formula from the Operations list.");
    if (!vals) return;
    try {
      const d = await sheetPost("/api/sheet/add-column", { sheet_name: grid().sheet(), name: vals.name });
      await grid().reload();
      grid().selectCol(d.index);
      toast(`Added column “${d.name}”`);
    } catch (e) { toast(e.message, true); }
  });

  // ------------------------------------------------------------- rename a column (double-click its name)
  $("#grid thead").addEventListener("dblclick", (e) => {
    const th = e.target.closest("th.name");
    if (!th || e.target.closest(".hdr-tools") || e.target.closest("input")) return;
    const c = +th.dataset.c, nm = th.querySelector(".nm"), old = grid().columns()[c];
    if (!nm || old === undefined) return;
    const input = el("input", { className: "col-rename", value: old });
    nm.replaceWith(input);
    input.focus(); input.select();
    ["click", "mousedown", "dblclick"].forEach((t) => input.addEventListener(t, (ev) => ev.stopPropagation()));
    let done = false;
    const finish = async (commit) => {
      if (done) return;
      done = true;
      const v = input.value.trim();
      if (!commit || !v || v === old) { input.replaceWith(nm); return; }
      try {
        await sheetPost("/api/sheet/rename-column", { sheet_name: grid().sheet(), col: c, name: v });
        await grid().reload();
        toast(`Renamed “${old}” to “${v}”`);
      } catch (err) { input.replaceWith(nm); toast(err.message, true); }
    };
    input.addEventListener("keydown", (ev) => {
      ev.stopPropagation();
      if (ev.key === "Enter") { ev.preventDefault(); finish(true); }
      else if (ev.key === "Escape") { ev.preventDefault(); finish(false); }
    });
    input.addEventListener("blur", () => finish(true));
  });

  // ------------------------------------------------------------- copy columns / rows to another sheet
  function copyToSheet() {
    if (!ready()) return;
    const g = grid(), sel = g.selection(), names = g.columns();
    const nCols = sel.c1 - sel.c0 + 1;
    const wholeColumns = sel.kind === "cols" || sel.kind === "all" || (sel.r0 === sel.r1 && sel.c0 === sel.c1);   // a single cell counts as "this column"
    const selRows = sel.pageIds.length;
    const canAll = !g.dupOnly();
    const others = state.sheets.map((s) => s.name);
    const colList = names.slice(sel.c0, sel.c1 + 1);
    const html = `<h2>Copy to another sheet</h2>
      <p class="modal-sub">Copies <b>${nCols} column${nCols === 1 ? "" : "s"}</b> (${esc(colList.slice(0, 3).join(", "))}${colList.length > 3 ? "…" : ""})
        as values — like Paste Special → Values. The original stays where it is.</p>
      <div class="field"><label>Which rows</label><div class="radio-row">
        <label class="radio"><input type="radio" name="ctRows" value="all" ${canAll && wholeColumns ? "checked" : ""} ${canAll ? "" : "disabled"}> All rows (${Number(g.matched()).toLocaleString("en-US")}${g.view() ? ", as filtered" : ""})</label>
        <label class="radio"><input type="radio" name="ctRows" value="sel" ${!(canAll && wholeColumns) ? "checked" : ""}> Only the selected rows (${selRows.toLocaleString("en-US")})</label>
      </div></div>
      <div class="field"><label>Copy into</label><div class="radio-row">
        <label class="radio"><input type="radio" name="ctDest" value="new" checked> A new sheet</label>
        <label class="radio"><input type="radio" name="ctDest" value="existing"> An existing sheet</label>
      </div></div>
      <div id="ctNew" class="field"><label>New sheet name</label><input id="ctName" type="text" maxlength="31" placeholder="Leave empty for an automatic name"></div>
      <div id="ctExisting" hidden>
        <div class="field-row">
          <div class="field"><label>Sheet</label><select id="ctSheet">${others.map((n) => `<option value="${esc(n)}">${esc(n)}</option>`).join("")}</select></div>
          <div class="field"><label>Start at column</label><select id="ctCol"></select></div>
        </div>
        <label class="check"><input type="checkbox" id="ctHead"> Also paste the column names, as a row</label>
        <div class="hint">The values go under the last row of that sheet.</div>
      </div>
      <label class="check" style="margin-top:12px"><input type="checkbox" id="ctGo" checked> Go to that sheet afterwards</label>
      <div class="modal-actions"><button class="btn secondary" type="button" data-x="cancel">Cancel</button><button class="btn primary" type="button" data-x="ok">Copy</button></div>`;
    openModal(html, (root) => {
      const q = (s) => root.querySelector(s);
      const dest = () => root.querySelector("input[name=ctDest]:checked").value;
      const fillCols = () => {
        const meta = state.sheets.find((s) => s.name === q("#ctSheet").value);
        const cs = (meta && meta.columns) || [];
        q("#ctCol").innerHTML = cs.map((n, i) => `<option value="${i}">${esc(n)}</option>`).join("") +
          `<option value="${cs.length}">${cs.length ? "(a new column at the end)" : "First column"}</option>`;
        q("#ctCol").value = "0";
      };
      const sync = () => { q("#ctNew").hidden = dest() !== "new"; q("#ctExisting").hidden = dest() !== "existing"; };
      root.querySelectorAll("input[name=ctDest]").forEach((r) => r.addEventListener("change", sync));
      q("#ctSheet").addEventListener("change", fillCols);
      fillCols(); sync();
      q("[data-x=cancel]").addEventListener("click", closeModal);
      q("[data-x=ok]").addEventListener("click", async () => {
        const rowsMode = q("input[name=ctRows]:checked").value;
        const body = {
          sheet_name: g.sheet(), cols: Array.from({ length: nCols }, (_, i) => sel.c0 + i),
          rows: rowsMode === "sel" ? sel.pageIds : null, view: rowsMode === "all" ? g.view() : null,
          dest_mode: dest(), dest_name: dest() === "new" ? q("#ctName").value : q("#ctSheet").value,
          dest_col: parseInt(q("#ctCol").value, 10) || 0, include_headers: q("#ctHead").checked,
        };
        try {
          const d = await sheetPost("/api/sheet/copy-to", body);
          closeModal();
          applySheetReply(d);
          if (q("#ctGo").checked) await selectSheet(d.name, { resetPage: true });
          toast(`Copied ${Number(d.rows).toLocaleString("en-US")} rows × ${d.cols} column${d.cols === 1 ? "" : "s"} to “${d.name}”`);
        } catch (err) { toast(err.message, true); }
      });
    });
  }
  $("#btnCopyTo").addEventListener("click", copyToSheet);
})();
