/* sql.js — the SQL screen.
 *
 * Every sheet is a table. Run SELECTs to look at data, UPDATEs are previewed as a diff before they change anything,
 * and a query that maps onto an Excel formula can be turned into one (the same steps the Formula Builder shows).
 * Anything else can still be written into the sheet as plain values, or saved as a new sheet.
 */
(function () {
  "use strict";
  (window.FRONTEND_PARTS = window.FRONTEND_PARTS || {}).sql = 17;

  let schema = [];
  let draft = "";
  let last = null;                       // the last SELECT result (for copy / fill / save)
  const fmt = (n) => Number(n).toLocaleString("en-US");
  const api = (path, body) => fetch(API_BASE + path, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
  });
  const detail = async (res, fallback) => {
    try { const d = await res.json(); return typeof d.detail === "string" ? d.detail : fallback; } catch (e) { return fallback; }
  };
  const q = (name) => (/^[A-Za-z_]\w*$/.test(name) ? name : '"' + String(name).replace(/"/g, '""') + '"');
  const dq = (name) => '"' + String(name).replace(/"/g, '""') + '"';

  // ------------------------------------------------------------ examples
  function examples() {
    const cur = state.currentSheet;
    const base = schema.find((t) => t.name === cur);
    if (!base) return [];
    const other = schema.find((t) => t.name !== cur && t.columns.some((c) => base.columns.some((b) => b.toLowerCase() === c.toLowerCase())))
      || schema.find((t) => t.name !== cur) || base;
    const bkey = base.columns.find((c) => other.columns.some((o) => o.toLowerCase() === c.toLowerCase())) || base.columns[0];
    const okey = other.columns.find((c) => c.toLowerCase() === String(bkey).toLowerCase()) || other.columns[0];
    const pick = (t, type, skip) => { const i = t.types.findIndex((x, k) => x === type && t.columns[k] !== skip); return i >= 0 ? t.columns[i] : null; };
    const onum = pick(other, "num", okey), otext = pick(other, "text", okey) || other.columns.find((c) => c !== okey) || okey;
    const btext = pick(base, "text", bkey) || base.columns[0];
    const N = dq(other.name), B = dq(base.name);
    const out = [
      ["Count matching rows for each row (COUNTIF)",
       `SELECT COUNT(*) AS matches\nFROM ${N} o\nWHERE o.${q(okey)} = this.${q(bkey)}`],
      ["Join two sheets and count (LEFT JOIN … GROUP BY)",
       `SELECT b.${q(bkey)}, COUNT(o.${q(okey)}) AS matches\nFROM ${B} b\nLEFT JOIN ${N} o ON o.${q(okey)} = b.${q(bkey)}\nGROUP BY b.${q(bkey)}`],
      ["First matching value (lookup)",
       `SELECT o.${q(otext)}\nFROM ${N} o\nWHERE o.${q(okey)} = this.${q(bkey)}\nLIMIT 1`],
      ["The 2nd matching value",
       `SELECT o.${q(otext)}\nFROM ${N} o\nWHERE o.${q(okey)} = this.${q(bkey)}\nLIMIT 1 OFFSET 1`],
    ];
    if (onum) {
      out.push(["Count with extra conditions (AND / OR)",
        `SELECT COUNT(*) AS matches\nFROM ${N} o\nWHERE o.${q(okey)} = this.${q(bkey)}\n  AND (o.${q(onum)} > 0 OR o.${q(otext)} IS NULL)`]);
      out.push(["Total of a column with a condition (SUMIF)",
        `SELECT SUM(o.${q(onum)}) AS total\nFROM ${N} o\nWHERE o.${q(okey)} = this.${q(bkey)}\n  AND o.${q(onum)} > 0`]);
    }
    out.push(["Rows that appear in one sheet but not the other",
      `SELECT b.*\nFROM ${B} b\nWHERE b.${q(bkey)} NOT IN (SELECT ${q(okey)} FROM ${N} WHERE ${q(okey)} IS NOT NULL)`]);
    out.push(["Duplicates in this sheet",
      `SELECT ${q(bkey)}, COUNT(*) AS times\nFROM ${B}\nGROUP BY ${q(bkey)}\nHAVING COUNT(*) > 1\nORDER BY times DESC`]);
    out.push(["Change data (UPDATE — previewed first)",
      `UPDATE ${B}\nSET ${q(btext)} = UPPER(TRIM(${q(btext)}))\nWHERE ${q(btext)} IS NOT NULL`]);
    return out;
  }

  // -------------------------------------------------------------- screen
  window.openSql = function () {
    if (!state.currentSheet) { toast("Upload a file first", true); return; }
    last = null;
    openModal(`
      <h2>SQL Query</h2>
      <p class="modal-sub">Every sheet is a table and every header is a column. <b>this.&lt;column&gt;</b> means “the row being filled in”.
        Run a query to see the result, turn it into a live Excel formula when it maps to one, or write the result into the sheet as values.</p>
      <div class="sql-layout">
        <aside class="sql-schema" id="sqlSchema"><div class="hint">Loading sheets…</div></aside>
        <div class="sql-main">
          <div class="sql-toolbar">
            <select id="sqlExamples"><option value="">Examples…</option></select>
            <button class="btn primary" id="sqlRun" type="button">Run ▶</button>
            <span class="sql-kbd">Ctrl+Enter</span>
            <span id="sqlStatus" class="sql-status"></span>
          </div>
          <textarea id="sqlText" spellcheck="false" placeholder='SELECT COUNT(*) FROM "Other sheet" o WHERE o.id = this.id'></textarea>
          <div id="sqlOut"></div>
        </div>
      </div>
      <div class="modal-actions"><button class="btn secondary" onclick="closeModal()">Close</button></div>
    `, mount, true);
    $("#modalBody").classList.add("xwide");
  };

  async function mount(root) {
    const text = root.querySelector("#sqlText");
    root.querySelector("#sqlRun").addEventListener("click", run);
    text.addEventListener("keydown", (e) => {
      if ((e.ctrlKey || e.metaKey) && e.key === "Enter") { e.preventDefault(); run(); }
      if (e.key === "Tab") { e.preventDefault(); insert("  "); }
    });
    text.addEventListener("input", () => { draft = text.value; });
    try {
      const res = await fetch(`${API_BASE}/api/sql/schema/${state.sessionId}`);
      schema = (await res.json()).tables;
    } catch (e) {
      root.querySelector("#sqlSchema").textContent = "Couldn't load the sheets: " + e.message;
      return;
    }
    renderSchema(root);
    const ex = examples();
    const sel = root.querySelector("#sqlExamples");
    ex.forEach(([label], i) => { const o = document.createElement("option"); o.value = i; o.textContent = label; sel.append(o); });
    sel.addEventListener("change", () => {
      if (sel.value === "") return;
      text.value = draft = ex[+sel.value][1];
      sel.value = "";
      text.focus();
    });
    text.value = draft || (ex[0] ? ex[0][1] : "");
    text.focus();
  }

  function insert(s) {
    const t = $("#sqlText");
    const a = t.selectionStart, b = t.selectionEnd;
    t.value = t.value.slice(0, a) + s + t.value.slice(b);
    t.selectionStart = t.selectionEnd = a + s.length;
    draft = t.value;
    t.focus();
  }

  function renderSchema(root) {
    const box = root.querySelector("#sqlSchema");
    box.replaceChildren();
    schema.forEach((t) => {
      const d = document.createElement("details");
      d.className = "sql-table";
      d.open = t.name === state.currentSheet;
      const sum = document.createElement("summary");
      sum.innerHTML = `<button type="button" class="sql-name" title="Insert into the query">${esc(t.name)}</button>` +
        (t.name === state.currentSheet ? `<span class="sql-badge">this sheet</span>` : "") +
        `<span class="sql-rows">${fmt(t.rows)} rows</span>`;
      sum.querySelector("button").addEventListener("click", (e) => { e.preventDefault(); insert(dq(t.name)); });
      d.append(sum);
      const cols = document.createElement("div");
      cols.className = "sql-cols";
      t.columns.forEach((c, i) => {
        const b = document.createElement("button");
        b.type = "button"; b.className = "sql-col t-" + t.types[i]; b.textContent = c; b.title = t.types[i];
        b.addEventListener("click", () => insert(q(c)));
        cols.append(b);
      });
      const alias = document.createElement("div");
      alias.className = "hint";
      alias.textContent = "short name: " + t.alias;
      d.append(cols, alias);
      box.append(d);
    });
  }

  // ----------------------------------------------------------------- run
  const setStatus = (s) => { const el = $("#sqlStatus"); if (el) el.textContent = s || ""; };
  const out = () => $("#sqlOut");

  async function run() {
    const sql = $("#sqlText").value;
    draft = sql;
    const btn = $("#sqlRun");
    btn.disabled = true;
    setStatus("Running…");
    out().replaceChildren();
    try {
      const res = await api("/api/sql/run", { session_id: state.sessionId, sheet_name: state.currentSheet, sql, limit: 1000 });
      if (!res.ok) { showError(await detail(res, "The query failed.")); return; }
      const data = await res.json();
      setStatus("");
      if (data.kind === "update") renderUpdate(data, sql); else renderSelect(data, sql);
    } catch (e) {
      showError("Can't reach the server: " + e.message);
    } finally {
      btn.disabled = false;
      if ($("#sqlStatus") && $("#sqlStatus").textContent === "Running…") setStatus("");
    }
  }

  function showError(msg) {
    setStatus("");
    const d = document.createElement("div");
    d.className = "sql-error";
    d.textContent = msg;
    out().replaceChildren(d);
  }

  function cellHtml(v) {
    if (v === null || v === undefined) return `<td class="null">NULL</td>`;
    return `<td${typeof v === "number" ? ' class="num"' : ""}>${esc(String(v))}</td>`;
  }

  // ---------------------------------------------------------------- select
  function renderSelect(data, sql) {
    last = { data, sql };
    const shown = data.rows.slice(0, 300);
    const wrap = document.createElement("div");
    wrap.innerHTML = `
      <div class="sql-meta"><b>${fmt(data.total ?? data.rows.length)}</b> row${data.total === 1 ? "" : "s"} · ${data.ms} ms
        ${data.rows.length > shown.length ? ` · showing the first ${shown.length}` : ""}${data.truncated ? " · the full result is used when you fill or save it" : ""}</div>
      <div class="sql-table-wrap"><table class="sql-result"><thead><tr>${data.columns.map((c) => `<th>${esc(c)}</th>`).join("")}</tr></thead>
        <tbody>${shown.map((r) => `<tr>${r.map(cellHtml).join("")}</tr>`).join("")}</tbody></table></div>
      <div class="sql-actions">
        <button class="btn tiny" id="sqlCopy" type="button">Copy results</button>
        <span class="sql-save"><input id="sqlSheetName" value="Query result" title="Name for the new sheet"><button class="btn tiny" id="sqlSave" type="button">Save as a new sheet</button></span>
      </div>
      <div id="sqlFormulaBox" class="sql-box"><span class="hint">Checking whether this can be an Excel formula…</span></div>
      <details class="sql-box" id="sqlFillBox"><summary>Fill “${esc(state.currentSheet)}” with these results as values</summary><div id="sqlFillBody"></div></details>`;
    out().replaceChildren(wrap);
    if (!data.rows.length) wrap.querySelector(".sql-table-wrap").insertAdjacentHTML("beforeend", `<div class="hint" style="padding:10px">The query returned no rows.</div>`);
    wrap.querySelector("#sqlCopy").addEventListener("click", copyResults);
    wrap.querySelector("#sqlSave").addEventListener("click", saveSheet);
    renderFill(data);
    translate(sql);
  }

  async function copyResults() {
    const { data } = last;
    const quote = (s) => (/[\t\n\r"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s);
    const text = [data.columns.map(quote).join("\t")].concat(data.rows.map((r) => r.map((v) => quote(v === null ? "" : String(v))).join("\t"))).join("\n");
    try { await navigator.clipboard.writeText(text); toast(`Copied ${fmt(data.rows.length)} rows — paste them into Excel`); }
    catch (e) { toast("Your browser blocked clipboard access.", true); }
  }

  async function saveSheet() {
    const name = $("#sqlSheetName").value.trim() || "Query result";
    const btn = $("#sqlSave");
    btn.disabled = true;
    try {
      const res = await api("/api/sql/save-sheet", { session_id: state.sessionId, sheet_name: state.currentSheet, sql: last.sql, name });
      if (!res.ok) { toast(await detail(res, "Couldn't save the sheet"), true); return; }
      const data = await res.json();
      data.sheets.forEach((s) => { if (!state.sheets.find((x) => x.name === s.name)) state.sheets.push(s); });
      renderSheetList();
      toast(`Saved ${fmt(data.rows)} rows as the new sheet “${data.name}”`);
      const r = await fetch(`${API_BASE}/api/sql/schema/${state.sessionId}`);
      schema = (await r.json()).tables;
      renderSchema($("#modalBody"));
    } catch (e) { toast(e.message, true); }
    finally { btn.disabled = false; }
  }

  // ---------------------------------------------------- Excel formula panel
  async function translate(sql) {
    const box = $("#sqlFormulaBox");
    let t;
    try {
      const res = await api("/api/sql/translate", { session_id: state.sessionId, sheet_name: state.currentSheet, sql });
      t = await res.json();
    } catch (e) { if (box) box.textContent = "Couldn't check: " + e.message; return; }
    if (!$("#sqlFormulaBox") || !last || last.sql !== sql) return;              // the query changed while we were waiting
    if (!t.ok) {
      box.className = "sql-box no";
      box.innerHTML = `<div class="sql-box-title">Not an Excel formula</div><div>${esc(t.error)}</div>
        <div class="hint">You can still fill the results in as values, or save them as a sheet.</div>`;
      return;
    }
    box.className = "sql-box ok";
    const outs = t.outputs;
    box.innerHTML = `
      <div class="sql-box-title">✓ This can be a live Excel formula in “${esc(state.currentSheet)}”</div>
      ${outs.map((o, i) => `
        <div class="sql-out">
          <label>Put it in column <input data-i="${i}" class="sql-outname" value="${esc(o.name)}"></label>
          <code class="formula-preview">${esc(o.formula)}</code>
          <div class="hint">First rows: ${o.values.map((v) => esc(v === null || v === "" ? "(empty)" : String(v))).join(", ")} …</div>
          <button class="link-btn" data-open="${i}" type="button">Open in the Formula Builder</button>
        </div>`).join("")}
      ${t.notes.map((n) => `<div class="hint">• ${esc(n)}</div>`).join("")}
      <div class="sql-actions"><button class="btn primary tiny" id="sqlApplyFormula" type="button">Apply as live formula${outs.length > 1 ? "s" : ""}</button></div>`;
    const names = () => outs.map((o, i) => ({ ...o, name: box.querySelector(`.sql-outname[data-i="${i}"]`).value.trim() || o.name }));
    box.querySelector("#sqlApplyFormula").addEventListener("click", () => applyFormulas(names()));
    box.querySelectorAll("[data-open]").forEach((b) => b.addEventListener("click", () => {
      const o = names()[+b.dataset.open];
      closeModal();
      openBuilderWith(o.node, o.name, "Formula from SQL");
    }));
  }

  async function applyFormulas(outs) {
    const btn = $("#sqlApplyFormula");
    btn.disabled = true;
    let lastData = null;
    try {
      for (const o of outs) {
        const res = await api("/api/operations/apply", {
          session_id: state.sessionId, sheet_name: state.currentSheet, operation: "EXPR",
          params: { expr: o.node }, output_column: o.name, label: "SQL",
        });
        if (!res.ok) { toast(`“${o.name}”: ` + await detail(res, "couldn't be applied"), true); btn.disabled = false; return; }
        lastData = await res.json();
        state.history.unshift({ op: "SQL", output: lastData.output_column, formula: lastData.sample_formula || "(values)", sheet: state.currentSheet });
      }
      const meta = currentSheetMeta();
      meta.columns = lastData.columns; meta.row_count = lastData.total_rows;
      renderHistory();
      closeModal();
      await selectSheet(state.currentSheet);
      toast(`Added ${outs.length} live formula column${outs.length > 1 ? "s" : ""}: ${outs.map((o) => o.name).join(", ")}`);
    } catch (e) { toast(e.message, true); btn.disabled = false; }
  }

  // ------------------------------------------------------- fill as values
  function renderFill(data) {
    const body = $("#sqlFillBody");
    const sheetCols = (currentSheetMeta().columns || []);
    const opt = (list, sel) => list.map((c) => `<option value="${esc(c)}" ${c === sel ? "selected" : ""}>${esc(c)}</option>`).join("");
    if (data.per_row) {
      body.innerHTML = `<div class="hint">This query gives one value for each row of the sheet.</div>
        <div class="field-row"><div class="field"><label>Write the value into column</label><input id="fillOut" value="sql_result" list="fillCols"></div></div>
        <datalist id="fillCols">${sheetCols.map((c) => `<option value="${esc(c)}">`).join("")}</datalist>
        <button class="btn tiny primary" id="sqlFill" type="button">Fill as values</button>`;
      body.querySelector("#sqlFill").addEventListener("click", () => fill({
        mode: "by_row", columns: [{ result: "result", output: $("#fillOut").value.trim() }],
      }));
      return;
    }
    const keyGuess = data.columns[0];
    const baseGuess = sheetCols.find((c) => c.toLowerCase() === String(keyGuess).toLowerCase()) || sheetCols[0];
    body.innerHTML = `
      <div class="field-row">
        <div class="field"><label>Result column that identifies the record</label><select id="fillKeyRes">${opt(data.columns, keyGuess)}</select></div>
        <div class="field"><label>…is the same value as this column of “${esc(state.currentSheet)}”</label><select id="fillKeyBase">${opt(sheetCols, baseGuess)}</select></div>
      </div>
      <div class="sql-fill-cols">${data.columns.map((c, i) => `
        <label class="check"><input type="checkbox" class="fill-pick" data-c="${esc(c)}" ${i > 0 ? "checked" : ""}> write <b>${esc(c)}</b> into column
          <input class="fill-name" data-c="${esc(c)}" value="${esc(c)}" list="fillCols"></label>`).join("")}</div>
      <datalist id="fillCols">${sheetCols.map((c) => `<option value="${esc(c)}">`).join("")}</datalist>
      <div class="hint">Rows without a matching record are left empty. If a record appears several times in the result, its first row is used.</div>
      <button class="btn tiny primary" id="sqlFill" type="button">Fill as values</button>`;
    body.querySelector("#sqlFill").addEventListener("click", () => {
      const cols = [...body.querySelectorAll(".fill-pick:checked")].map((p) => ({
        result: p.dataset.c, output: body.querySelector(`.fill-name[data-c="${p.dataset.c}"]`).value.trim(),
      }));
      fill({ mode: "by_key", key_result: $("#fillKeyRes").value, key_base: $("#fillKeyBase").value, columns: cols });
    });
  }

  async function fill(extra) {
    const btn = $("#sqlFill");
    btn.disabled = true;
    try {
      const res = await api("/api/sql/apply-values", { session_id: state.sessionId, sheet_name: state.currentSheet, sql: last.sql, ...extra });
      if (!res.ok) { toast(await detail(res, "Couldn't fill the sheet"), true); btn.disabled = false; return; }
      const data = await res.json();
      const meta = currentSheetMeta();
      meta.columns = data.columns; meta.row_count = data.total_rows;
      state.history.unshift({ op: "SQL (values)", output: data.written.join(", "), formula: "(values)", sheet: state.currentSheet });
      renderHistory();
      closeModal();
      await selectSheet(state.currentSheet);
      toast(`Filled ${data.written.join(", ")} with values from the query`);
    } catch (e) { toast(e.message, true); btn.disabled = false; }
  }

  // ---------------------------------------------------------------- update
  function renderUpdate(data, sql) {
    last = null;
    const headerRow = (state.sheets.find((s) => s.name === data.sheet) || {}).header_row || 1;
    const wrap = document.createElement("div");
    const none = data.cells_changed === 0;
    wrap.innerHTML = `
      <div class="sql-box ${none ? "no" : "warn"}">
        <div class="sql-box-title">${none ? "This UPDATE wouldn't change anything" : "Preview — nothing has changed yet"}</div>
        <div>${fmt(data.rows_matched)} row${data.rows_matched === 1 ? "" : "s"} matched in “${esc(data.sheet)}”;
          <b>${fmt(data.cells_changed)}</b> cell${data.cells_changed === 1 ? "" : "s"} in <b>${fmt(data.rows_changed)}</b> row${data.rows_changed === 1 ? "" : "s"} would change.</div>
      </div>
      ${none ? "" : `
      <div class="sql-table-wrap"><table class="sql-result"><thead><tr><th>Sheet row</th><th>Column</th><th>Before</th><th>After</th></tr></thead>
        <tbody>${data.changes.map((c) => `<tr><td class="num">${c.row + headerRow + 1}</td><td>${esc(c.column)}</td>
          ${cellHtml(c.before)}${cellHtml(c.after)}</tr>`).join("")}</tbody></table></div>
      ${data.cells_changed > data.changes.length ? `<div class="hint">Showing the first ${data.changes.length} of ${fmt(data.cells_changed)} changes.</div>` : ""}
      <div class="sql-actions"><button class="btn primary tiny" id="sqlCommit" type="button">Apply these ${fmt(data.cells_changed)} changes</button>
        <button class="btn tiny" id="sqlDiscard" type="button">Discard</button></div>`}`;
    out().replaceChildren(wrap);
    const commit = wrap.querySelector("#sqlCommit");
    if (!commit) return;
    wrap.querySelector("#sqlDiscard").addEventListener("click", () => out().replaceChildren());
    commit.addEventListener("click", async () => {
      commit.disabled = true;
      try {
        const res = await api("/api/sql/commit-update", { session_id: state.sessionId, sql });
        if (!res.ok) { toast(await detail(res, "The update failed"), true); commit.disabled = false; return; }
        const r = await res.json();
        state.history.unshift({ op: "SQL UPDATE", output: r.sheet, formula: `${r.cells} cells changed`, sheet: r.sheet });
        renderHistory();
        closeModal();
        await selectSheet(state.currentSheet);
        toast(`Updated ${fmt(r.cells)} cells in ${fmt(r.rows)} rows of “${r.sheet}”`);
      } catch (e) { toast(e.message, true); commit.disabled = false; }
    });
  }
})();
