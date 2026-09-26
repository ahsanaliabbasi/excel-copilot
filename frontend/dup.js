/* dup.js — the Duplicates screen.
 *
 * Everything here runs against the whole sheet on the server, not just the page in the grid,
 * so it works the same on 100 rows or 75,000. Nothing changes on disk (or in the sheet) until
 * you press Apply, and deleting rows from the current sheet always asks you to confirm first.
 */
(function () {
  "use strict";
  (window.FRONTEND_PARTS = window.FRONTEND_PARTS || {}).dup = 17;

  const fmt = (n) => Number(n || 0).toLocaleString("en-US");
  const api = (path, body) => fetch(API_BASE + path, {
    method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body),
  });
  const detail = async (res, fallback) => {
    try { const d = await res.json(); return typeof d.detail === "string" ? d.detail : fallback; } catch (e) { return fallback; }
  };
  const disp = (v) => (v === null || v === undefined ? "" : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : String(v));
  const quote = (s) => (/[\t\n\r"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s);
  const rowsToTsv = (columns, rows, headers) => {
    const lines = headers ? [columns.map(quote).join("\t")] : [];
    rows.forEach((r) => lines.push(r.map((v) => quote(disp(v))).join("\t")));
    return lines.join("\n");
  };
  async function copyText(text, rows, label) {
    try { await navigator.clipboard.writeText(text); toast(`Copied ${fmt(rows)} ${label || "row"}${rows === 1 ? "" : "s"} — paste straight into Excel`); }
    catch (e) { toast("Your browser blocked clipboard access.", true); }
  }

  let cols = [];                 // the sheet's column names for the open sheet
  let letters = [];
  let D = null;                  // {sheet, groupOffset}
  const CONTENT_LABEL = { unique: "Unique records (one per group)", duplicates: "Duplicate rows (the extra copies)", groups: "All rows in a duplicate group" };

  window.openDup = function () {
    if (!state.currentSheet) { toast("Upload a file and pick a sheet first", true); return; }
    const meta = currentSheetMeta();
    cols = meta.columns; letters = colLetters(cols.length);
    D = { sheet: state.currentSheet, groupOffset: 0 };
    const sel = defaultCols();
    openModal(`
      <h2>Duplicates</h2>
      <p class="modal-sub">Works across the whole sheet “${esc(D.sheet)}” — ${fmt(meta.row_count)} rows — not just the page you can see. Nothing changes until you press Apply.</p>

      <div class="dup-sub-title">1. Which columns make a row a duplicate?</div>
      <div class="dup-sub">Two rows are duplicates when every ticked column matches.</div>
      <div class="dup-cols" id="dCols">${cols.map((n, c) =>
        `<label class="check"><input type="checkbox" data-c="${c}" ${sel.includes(c) ? "checked" : ""}> ${esc(n)} <span class="dup-letter">${letters[c]}</span></label>`).join("")}</div>
      <div class="dup-quick">
        <button type="button" class="link-btn" id="dAll">All columns</button>
        <button type="button" class="link-btn" id="dNone">None</button>
        <label class="check"><input type="checkbox" id="dNorm" checked> Ignore case &amp; extra spaces</label>
        <label class="check"><input type="checkbox" id="dSkip" checked> Ignore rows where the key is empty</label>
      </div>

      <div id="dOverview" class="dup-stats"></div>

      <div class="dup-grid">
        <div>
          <div class="dup-sub-title">Check one value</div>
          <div class="dup-sub">See whether a specific record — like one employee id — has duplicates.</div>
          <div id="dFindInputs" class="dup-find-inputs"></div>
          <div class="dup-actions"><button type="button" class="btn secondary tiny" id="dFindBtn">Find</button><button type="button" class="link-btn" id="dFindClear" hidden>✕ Clear</button></div>
          <div id="dFindResult"><div class="hint">Type a value above, or pick a group on the right, to see its rows here.</div></div>
        </div>
        <div>
          <div class="dup-sub-title">Duplicate groups <span class="dup-sub-title-n" id="dGroupsN"></span></div>
          <input id="dGroupSearch" class="dup-search" placeholder="Search a value…">
          <div id="dGroupList" class="dup-group-list"></div>
          <div id="dGroupPager" class="dup-group-pager"></div>
        </div>
      </div>

      <div class="dup-sub-title">2. Copy a set of rows</div>
      <div class="dup-actions dup-copy-row">
        <select id="dCopyWhich">
          <option value="unique">Unique records (one per group)</option>
          <option value="duplicates">Duplicate rows (the extra copies)</option>
          <option value="groups">All rows in a duplicate group</option>
        </select>
        <label class="check"><input type="checkbox" id="dCopyHeaders" checked> Include header row</label>
        <button type="button" class="btn secondary tiny" id="dCopyBtn">Copy to clipboard</button>
      </div>

      <div class="dup-sub-title">3. Remove duplicates</div>
      <div class="dup-sub">Keep the <span id="dKeepWord">first</span> row of each group.</div>
      <div class="dup-actions">
        <label class="check"><input type="radio" name="dKeep" value="first" checked> Keep the first occurrence</label>
        <label class="check"><input type="radio" name="dKeep" value="last"> Keep the last occurrence</label>
      </div>
      <div class="dup-actions">
        <label class="check"><input type="radio" name="dTarget" value="existing" checked> Delete the extra rows from this sheet</label>
        <label class="check"><input type="radio" name="dTarget" value="new"> Put the result in a new sheet instead</label>
      </div>
      <div id="dNewBox" class="dup-new-box" hidden>
        <label class="check"><input type="radio" name="dContent" value="unique" checked> Unique records only</label>
        <label class="check"><input type="radio" name="dContent" value="duplicates"> Duplicate rows only (what would be removed)</label>
        <label class="check"><input type="radio" name="dContent" value="groups"> Every row that is part of a duplicate group</label>
        <input id="dNewName" type="text" placeholder="New sheet name">
      </div>
      <div id="dConfirm" class="dup-confirm" hidden></div>

      <div class="modal-actions">
        <button class="btn secondary" onclick="closeModal()">Close</button>
        <button class="btn primary" id="dApplyBtn">Apply</button>
      </div>
    `, mount, true);
    $("#modalBody").classList.add("xwide");
  };

  function colLetters(n) {
    const out = [];
    for (let i = 0; i < n; i++) { let s = "", x = i; do { s = String.fromCharCode(65 + (x % 26)) + s; x = Math.floor(x / 26) - 1; } while (x >= 0); out.push(s); }
    return out;
  }
  // A sensible starting key: an "id"-looking column if there is one (custId, employee_id, OrderID, code, number …),
  // else every column.
  function defaultCols() {
    const snake = (n) => n.replace(/([a-z0-9])([A-Z])/g, "$1_$2").toLowerCase();
    const idLike = cols.map((n, c) => [n, c]).filter(([n]) => /(^|_)(id|code|number)$/.test(snake(n)));
    return idLike.length ? idLike.map(([, c]) => c) : cols.map((_, c) => c);
  }

  function mount(root) {
    const timers = {};
    const debounce = (key, fn, ms = 250) => { clearTimeout(timers[key]); timers[key] = setTimeout(fn, ms); };

    const ticked = () => [...root.querySelectorAll("#dCols input:checked")].map((i) => +i.dataset.c);
    const spec = () => ({
      session_id: state.sessionId, sheet_name: D.sheet, cols: ticked(),
      norm: $("#dNorm").checked, skip_empty: $("#dSkip").checked,
    });
    const keep = () => root.querySelector("input[name=dKeep]:checked").value;
    const target = () => root.querySelector("input[name=dTarget]:checked").value;
    const content = () => root.querySelector("input[name=dContent]:checked").value;

    let lastSummary = null;

    async function refreshAll() {
      buildFindInputs();
      clearFind();
      await Promise.all([refreshOverview(), refreshGroups(0)]);
    }

    async function refreshOverview() {
      const box = $("#dOverview");
      if (!ticked().length) {
        box.innerHTML = `<div class="hint">Tick at least one column above to see the overview.</div>`;
        lastSummary = null;
        setApplyState();
        return;
      }
      box.innerHTML = `<div class="hint">Checking every row…</div>`;
      let s;
      try {
        const res = await api("/api/dup/summary", spec());
        if (!res.ok) { box.innerHTML = `<div class="hint">${esc(await detail(res, "Couldn't check for duplicates"))}</div>`; return; }
        s = await res.json();
      } catch (e) { box.innerHTML = `<div class="hint">${esc(e.message)}</div>`; return; }
      if (JSON.stringify(spec()) !== JSON.stringify(spec())) return;   // (kept for symmetry with other screens)
      lastSummary = s;
      const tile = (n, label, cls) => `<div class="dup-stat${cls ? " " + cls : ""}"><div class="dup-stat-n">${fmt(n)}</div><div class="dup-stat-l">${label}</div></div>`;
      box.innerHTML = `
        <div class="dup-stat-row">
          ${tile(s.total_rows, "Total records")}
          ${tile(s.groups, "Duplicate groups", s.groups ? "warn" : "")}
          ${tile(s.extra_rows, "Duplicate rows (would be removed)", s.extra_rows ? "warn" : "")}
          ${tile(s.unique_rows, "Unique records after dedup", "good")}
        </div>
        ${s.groups ? `<button type="button" class="link-btn" id="dViewAll">View all ${fmt(s.duplicate_rows)} duplicate rows in the grid →</button>` : `<div class="dup-result none">No duplicate rows found using ${fmt(s.key_columns)} key column${s.key_columns === 1 ? "" : "s"}.</div>`}`;
      const viewAll = $("#dViewAll");
      if (viewAll) viewAll.onclick = async () => { await window.gridSetDupView({ ...spec(), only: true }); closeModal(); };
      setApplyState();
    }

    function setApplyState() {
      const btn = $("#dApplyBtn");
      const has = !!(lastSummary && lastSummary.groups);
      btn.disabled = !has;
      btn.title = has ? "" : "No duplicates found with the current key columns";
      $("#dCopyBtn").disabled = !ticked().length;
    }

    async function refreshGroups(offset) {
      D.groupOffset = offset;
      const list = $("#dGroupList");
      if (!ticked().length) { list.innerHTML = ""; $("#dGroupsN").textContent = ""; $("#dGroupPager").innerHTML = ""; return; }
      list.innerHTML = `<div class="hint">Loading…</div>`;
      const body = { ...spec(), offset, limit: 25, search: $("#dGroupSearch").value.trim() };
      let g;
      try {
        const res = await api("/api/dup/groups", body);
        if (!res.ok) { list.innerHTML = `<div class="hint">${esc(await detail(res, "Couldn't load groups"))}</div>`; return; }
        g = await res.json();
      } catch (e) { list.innerHTML = `<div class="hint">${esc(e.message)}</div>`; return; }
      $("#dGroupsN").textContent = g.groups_with_duplicates ? `(${fmt(g.total_groups)}${body.search ? " matching" : ""})` : "";
      if (!g.items.length) {
        list.innerHTML = `<div class="hint">${body.search ? "No groups match your search." : "No duplicate groups."}</div>`;
      } else {
        list.innerHTML = g.items.map((it) => `
          <button type="button" class="dup-group" data-key='${esc(JSON.stringify(it.key))}'>
            <span class="dup-group-key">${it.key.map((k) => esc(k || "(empty)")).join(" / ")}</span>
            <span class="dup-group-n">×${it.size}</span>
          </button>`).join("");
        list.querySelectorAll(".dup-group").forEach((b) => b.addEventListener("click", () => showRows(JSON.parse(b.dataset.key), b)));
      }
      const pager = $("#dGroupPager");
      if (g.total_groups > 25) {
        pager.innerHTML = `<button type="button" class="link-btn" id="dGPrev" ${offset <= 0 ? "disabled" : ""}>← Prev</button>
          <span>${fmt(offset + 1)}–${fmt(Math.min(offset + 25, g.total_groups))} of ${fmt(g.total_groups)}</span>
          <button type="button" class="link-btn" id="dGNext" ${offset + 25 >= g.total_groups ? "disabled" : ""}>Next →</button>`;
        $("#dGPrev").onclick = () => refreshGroups(Math.max(0, offset - 25));
        $("#dGNext").onclick = () => refreshGroups(offset + 25);
      } else pager.innerHTML = "";
    }

    function buildFindInputs() {
      const box = $("#dFindInputs");
      const c = ticked();
      if (!c.length) { box.innerHTML = `<div class="hint">Tick at least one column above.</div>`; return; }
      box.innerHTML = c.map((ci) => `<label class="dup-find-field">${esc(cols[ci])}<input type="text" data-c="${ci}" placeholder="value…"></label>`).join("");
      box.querySelectorAll("input").forEach((i) => {
        i.addEventListener("keydown", (e) => { if (e.key === "Enter") doFind(); });
        i.addEventListener("input", () => {
          setActiveGroup(null);                          // typing by hand no longer matches "the" clicked group
          const inputs = [...box.querySelectorAll("input")];
          if (!inputs.some((x) => x.value.trim())) clearFind();   // every field emptied -> back to the neutral state
        });
      });
    }

    let lastValues = null, lastResult = null, activeGroupBtn = null;

    function setActiveGroup(btn) {
      if (activeGroupBtn) activeGroupBtn.classList.remove("on");
      activeGroupBtn = btn || null;
      if (activeGroupBtn) activeGroupBtn.classList.add("on");
    }
    // Back to the "nothing picked yet" state: clears the inputs, the result table and any highlighted group.
    function clearFind() {
      root.querySelectorAll("#dFindInputs input").forEach((i) => { i.value = ""; });
      lastValues = null; lastResult = null;
      setActiveGroup(null);
      $("#dFindResult").innerHTML = `<div class="hint">Type a value above, or pick a group on the right, to see its rows here.</div>`;
      $("#dFindClear").hidden = true;
    }

    async function doFind() {
      const c = ticked();
      if (!c.length) return;
      const inputs = [...root.querySelectorAll("#dFindInputs input")];
      const values = c.map((ci) => (inputs.find((i) => +i.dataset.c === ci) || {}).value || "");
      if (!values.some((v) => v.trim())) { clearFind(); return; }
      await runFind(values);
    }
    async function showRows(values, btn) {
      const inputs = [...root.querySelectorAll("#dFindInputs input")];
      ticked().forEach((ci, i) => { const inp = inputs.find((x) => +x.dataset.c === ci); if (inp) inp.value = values[i] || ""; });
      setActiveGroup(btn || null);
      await runFind(values);
    }
    async function runFind(values) {
      lastValues = values;
      $("#dFindClear").hidden = false;
      const out = $("#dFindResult");
      out.innerHTML = `<div class="hint">Searching…</div>`;
      let r;
      try {
        const res = await api("/api/dup/lookup", { ...spec(), values });
        if (!res.ok) { out.innerHTML = `<div class="hint">${esc(await detail(res, "Couldn't search"))}</div>`; return; }
        r = await res.json();
      } catch (e) { out.innerHTML = `<div class="hint">${esc(e.message)}</div>`; return; }
      lastResult = r;
      if (!r.matches) { out.innerHTML = `<div class="dup-result none">No rows match that value.</div>`; return; }
      const msg = r.is_duplicate
        ? `<div class="dup-result found">${fmt(r.matches)} rows share this value — that's a duplicate. Tick the ones to keep or remove below.</div>`
        : `<div class="dup-result none">1 row matches — not a duplicate.</div>`;
      const capped = r.matches > 200;
      const shown = r.rows.slice(0, 200), shownIds = r.ids.slice(0, 200);
      out.innerHTML = `${msg}
        <div class="sql-table-wrap"><table class="sql-result dup-pick-table"><thead><tr><th class="dup-pick-th"><input type="checkbox" id="dFindAll" title="Select all"></th>${r.columns.map((c) => `<th>${esc(c)}</th>`).join("")}</tr></thead>
        <tbody>${shown.map((row, i) => `<tr><td class="dup-pick-td"><input type="checkbox" class="dup-pick" data-id="${shownIds[i]}"></td>${row.map((v, j) => `<td${typeof v === "number" ? ' class="num"' : ""}>${esc(disp(v))}</td>`).join("")}</tr>`).join("")}</tbody></table></div>
        ${capped ? `<div class="hint">Showing the first ${shown.length} of ${fmt(r.matches)} rows — row picking only works within what's shown.</div>` : ""}
        <div class="dup-actions">
          <button type="button" class="btn secondary tiny" id="dFindCopy">Copy these ${fmt(r.matches)} rows</button>
          <span class="dup-pick-n" id="dPickN"></span>
          <button type="button" class="btn danger tiny" id="dDeleteChecked" disabled>Delete the ticked rows</button>
          <button type="button" class="btn secondary tiny" id="dKeepChecked" disabled>Keep only the ticked rows</button>
        </div>
        <div id="dFindConfirm" class="dup-confirm" hidden></div>`;
      $("#dFindCopy").onclick = () => copyText(rowsToTsv(r.columns, r.rows, true), r.rows.length, "row");

      const boxes = () => [...root.querySelectorAll(".dup-pick")];
      const checkedIds = () => boxes().filter((b) => b.checked).map((b) => +b.dataset.id);
      const refreshPickState = () => {
        const n = checkedIds().length;
        $("#dPickN").textContent = n ? `${fmt(n)} selected` : "";
        $("#dDeleteChecked").disabled = n === 0;
        $("#dKeepChecked").disabled = n === 0 || n === boxes().length;   // "keep only" needs something left to remove
        $("#dFindConfirm").hidden = true;
      };
      $("#dFindAll").addEventListener("change", (e) => { boxes().forEach((b) => (b.checked = e.target.checked)); refreshPickState(); });
      boxes().forEach((b) => b.addEventListener("change", refreshPickState));

      const askThenDelete = (idsToDelete, verb) => {
        const n = idsToDelete.length;
        $("#dFindConfirm").hidden = false;
        $("#dFindConfirm").innerHTML = `
          <div class="dup-confirm-text">This permanently deletes ${fmt(n)} row${n === 1 ? "" : "s"} from “${esc(D.sheet)}” — ${esc(verb)}.</div>
          <div class="dup-actions"><button type="button" class="btn secondary tiny" id="dFindConfCancel">Cancel</button>
          <button type="button" class="btn danger tiny" id="dFindConfGo">Yes, delete ${fmt(n)} row${n === 1 ? "" : "s"}</button></div>`;
        $("#dFindConfCancel").onclick = () => { $("#dFindConfirm").hidden = true; };
        $("#dFindConfGo").onclick = () => deleteRows(idsToDelete);
      };
      $("#dDeleteChecked").onclick = () => askThenDelete(checkedIds(), "the rows you ticked");
      $("#dKeepChecked").onclick = () => askThenDelete(boxes().filter((b) => !b.checked).map((b) => +b.dataset.id), "everything except the rows you ticked");
    }

    async function deleteRows(ids) {
      try {
        const res = await api("/api/sheet/delete-rows", { session_id: state.sessionId, sheet_name: D.sheet, rows: ids });
        if (!res.ok) { toast(await detail(res, "Couldn't delete those rows"), true); return; }
        const r = await res.json();
        toast(`Removed ${fmt(r.removed)} row${r.removed === 1 ? "" : "s"} from “${r.sheet}”`);
        (r.sheets || []).forEach((s) => { const ex = state.sheets.find((x) => x.name === s.name); if (ex) Object.assign(ex, s); });
        renderSheetList();
        if (state.currentSheet === D.sheet) await window.gridLoadSheet(D.sheet, { resetPage: true });
        await refreshOverview();
        await refreshGroups(D.groupOffset);
        if (lastValues) await runFind(lastValues); else $("#dFindResult").replaceChildren();
      } catch (e) { toast(e.message, true); }
    }

    // ---- wiring ----
    root.querySelectorAll("#dCols input, #dNorm, #dSkip").forEach((i) => i.addEventListener("change", () => debounce("setup", refreshAll, 150)));
    $("#dAll").onclick = () => { root.querySelectorAll("#dCols input").forEach((i) => (i.checked = true)); refreshAll(); };
    $("#dNone").onclick = () => { root.querySelectorAll("#dCols input").forEach((i) => (i.checked = false)); refreshAll(); };
    $("#dFindBtn").onclick = doFind;
    $("#dFindClear").onclick = clearFind;
    $("#dGroupSearch").addEventListener("input", () => debounce("search", () => refreshGroups(0), 250));

    root.querySelectorAll("input[name=dKeep]").forEach((i) => i.addEventListener("change", () => { $("#dKeepWord").textContent = keep() === "last" ? "last" : "first"; }));
    root.querySelectorAll("input[name=dTarget]").forEach((i) => i.addEventListener("change", () => {
      $("#dNewBox").hidden = target() !== "new";
      $("#dApplyBtn").textContent = target() === "new" ? "Create sheet" : "Apply";
      $("#dConfirm").hidden = true;
    }));
    if (!$("#dNewName").value) $("#dNewName").placeholder = `${D.sheet} - unique`;

    $("#dCopyBtn").onclick = async () => {
      const which = $("#dCopyWhich").value;
      const btn = $("#dCopyBtn");
      btn.disabled = true;
      try {
        const res = await api("/api/dup/export-text", { ...spec(), keep: keep(), which, headers: $("#dCopyHeaders").checked });
        if (!res.ok) { toast(await detail(res, "Couldn't copy"), true); return; }
        const r = await res.json();
        if (!r.rows) { toast("Nothing to copy — that set is empty.", true); return; }
        await copyText(r.text, r.rows, "record");
      } catch (e) { toast(e.message, true); }
      finally { btn.disabled = false; setApplyState(); }
    };

    $("#dApplyBtn").onclick = async () => {
      if (target() === "existing") {
        if (!$("#dConfirm").hidden) return;                     // already showing the confirmation
        const extra = lastSummary ? lastSummary.extra_rows : 0;
        $("#dConfirm").hidden = false;
        $("#dConfirm").innerHTML = `
          <div class="dup-confirm-text">This permanently deletes ${fmt(extra)} row${extra === 1 ? "" : "s"} from “${esc(D.sheet)}”, keeping the ${keep()} row of each duplicate group.</div>
          <div class="dup-actions"><button type="button" class="btn secondary tiny" id="dConfCancel">Cancel</button>
          <button type="button" class="btn danger tiny" id="dConfGo">Yes, delete ${fmt(extra)} row${extra === 1 ? "" : "s"}</button></div>`;
        $("#dConfCancel").onclick = () => { $("#dConfirm").hidden = true; };
        $("#dConfGo").onclick = doApply;
        return;
      }
      await doApply();
    };

    async function doApply() {
      const btn = $("#dApplyBtn");
      btn.disabled = true;
      try {
        const body = { ...spec(), keep: keep(), target: target(), new_sheet_name: $("#dNewName").value.trim(), content: content() };
        const res = await api("/api/dup/apply", body);
        if (!res.ok) { toast(await detail(res, "Couldn't apply"), true); return; }
        const r = await res.json();
        $("#dConfirm").hidden = true;
        if (r.mode === "existing") {
          toast(r.removed ? `Removed ${fmt(r.removed)} duplicate rows from “${r.sheet}”` : "No rows needed to be removed");
          if (state.currentSheet === r.sheet) await window.gridLoadSheet(r.sheet, { resetPage: true });
          (r.sheets || []).forEach((s) => { const ex = state.sheets.find((x) => x.name === s.name); if (ex) Object.assign(ex, s); });
        } else {
          toast(`Created “${r.sheet}” with ${fmt(r.rows)} rows (${CONTENT_LABEL[content()].toLowerCase()})`);
          (r.sheets || []).forEach((s) => { if (!state.sheets.find((x) => x.name === s.name)) state.sheets.push(s); });
        }
        renderSheetList();
        await refreshAll();
      } catch (e) { toast(e.message, true); }
      finally { setApplyState(); }
    }

    refreshAll();
  }
})();
