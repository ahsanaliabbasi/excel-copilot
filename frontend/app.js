// When this page is served by the backend itself (http://localhost:8000) the API is on the same origin;
// otherwise (opened from another static server or from a file) use the default backend address.
const API_BASE = window.EXCEL_COPILOT_API || (function () {
  if (location.protocol === "http:" || location.protocol === "https:") {
    try {
      const x = new XMLHttpRequest();
      x.open("GET", location.origin + "/api/health", false);
      x.send();
      if (x.status === 200 && JSON.parse(x.responseText).status === "ok") return location.origin;
    } catch (e) { /* not the backend */ }
  }
  return "http://localhost:8000";
})();

let state = {
  sessionId: null,
  sheets: [],          // [{name, columns, row_count, header_row}]
  currentSheet: null,
  history: [],
};

const $ = (sel) => document.querySelector(sel);
const el = (tag, opts = {}) => Object.assign(document.createElement(tag), opts);

// Fires a Google Analytics custom event if GA4 is loaded (index.html); a no-op otherwise
// (e.g. running locally without a Measurement ID), so this is always safe to call.
function trackEvent(name, params = {}) {
  try { if (typeof gtag === "function") gtag("event", name, params); } catch (e) { /* analytics must never break the app */ }
}

function toast(msg, isError = false) {
  const t = $("#toast");
  t.textContent = msg;
  t.className = "toast" + (isError ? " error" : "");
  t.hidden = false;
  clearTimeout(toast._t);
  toast._t = setTimeout(() => (t.hidden = true), 3500);
}

// ---------------------------------------------------------------------------
// Upload — three ways in: the top-bar button, the "Browse" button in the
// empty state, and drag-and-drop onto the drop zone (this last one never
// opens the OS file picker, so it works even if that dialog is glitchy on
// your machine/remote desktop).
// ---------------------------------------------------------------------------
function openFilePicker() {
  const input = $("#fileInput");
  input.value = "";                    // so choosing the same file again still fires "change"
  input.click();
}
$("#uploadBtn").addEventListener("click", openFilePicker);
const uploadBtn2 = $("#uploadBtn2");
if (uploadBtn2) uploadBtn2.addEventListener("click", (e) => { e.stopPropagation(); openFilePicker(); });
document.querySelectorAll("[data-upload]").forEach((b) => b.addEventListener("click", openFilePicker));
// clicking anywhere in the drop zone (icon, text, empty space) uploads too
const dz = $("#dropZone");
if (dz) {
  dz.addEventListener("click", openFilePicker);
  dz.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); openFilePicker(); } });
}

$("#fileInput").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (file) handleFile(file);
});

// Drag and drop: works anywhere on the page. The drop zone lights up while a file is dragged over the
// window; the enter/leave counter stops it flickering when the pointer crosses child elements.
const dropZone = $("#dropZone");
let dragDepth = 0;
const hasFiles = (e) => e.dataTransfer && Array.from(e.dataTransfer.types || []).includes("Files");
function takeDroppedFile(e) {
  const dt = e.dataTransfer;
  let file = dt && dt.files && dt.files[0];
  if (!file && dt && dt.items) {                       // some browsers only expose the file through items
    for (const it of dt.items) { if (it.kind === "file") { file = it.getAsFile(); if (file) break; } }
  }
  if (!file) {
    toast("No file received. Drag the .xlsx from File Explorer (not from inside the editor), or use Browse.", true);
    return;
  }
  if (!/\.xls[xm]$/i.test(file.name)) { toast("Please drop a .xlsx file", true); return; }
  handleFile(file);
}
window.addEventListener("dragenter", (e) => {
  if (!hasFiles(e)) return;
  e.preventDefault();
  dragDepth++;
  if (dropZone) dropZone.classList.add("dragover");
});
window.addEventListener("dragover", (e) => {
  e.preventDefault();                                   // required, or the browser refuses the drop
  if (e.dataTransfer) e.dataTransfer.dropEffect = "copy";
});
window.addEventListener("dragleave", (e) => {
  if (!hasFiles(e)) return;
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0 && dropZone) dropZone.classList.remove("dragover");
});
window.addEventListener("drop", (e) => {
  e.preventDefault();                                   // otherwise the browser opens/downloads the file itself
  dragDepth = 0;
  if (dropZone) dropZone.classList.remove("dragover");
  takeDroppedFile(e);
});

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// A slim status line under the top bar (loading / building progress). showInfo(null) hides it.
function showInfo(msg) {
  let bar = $("#infoBar");
  if (!msg) { if (bar) bar.remove(); return; }
  if (!bar) {
    bar = el("div", { id: "infoBar", className: "info-bar" });
    document.querySelector(".topbar").after(bar);
  }
  bar.textContent = msg;
}

async function errorText(res, fallback) {
  try { const d = await res.json(); return typeof d.detail === "string" ? d.detail : fallback; } catch (e) { return fallback; }
}

async function handleFile(file) {
  const fd = new FormData();
  fd.append("file", file);
  const mb = (file.size / 1e6).toFixed(1);
  await loadWorkbook(`Uploading ${file.name} (${mb} MB)…`, `${API_BASE}/api/upload`, { method: "POST", body: fd }, file.name);
}

async function loadWorkbook(info, url, init, name) {
  try {
    showInfo(info);
    const res = await fetch(url, init);
    if (res.status === 404 || res.status === 405) throw new Error("the server is out of date — stop it and start it again");
    if (!res.ok) throw new Error(await errorText(res, `The server rejected the file (${res.status})`));
    const data = await res.json();
    state.sessionId = data.session_id;
    state.sheets = data.sheets;
    state.history = [];
    renderHistory();
    $("#emptyState").hidden = true;
    $("#app").hidden = false;
    $("#downloadBtn").disabled = false;
    renderSheetList();
    await selectSheet(state.sheets[0].name, { resetPage: true });
    toast(`Loaded ${name} — ${data.sheets.length} sheet(s)` +
          (data.mode === "stream" ? " · large workbook mode" : ""));
    trackEvent("upload_workbook", { sheet_count: data.sheets.length, mode: data.mode || "full" });
    pollSheets();                       // the other sheets keep loading in the background
  } catch (err) {
    showInfo(null);
    toast("Upload failed: " + (err.message === "Failed to fetch" ? "can't reach the backend — is it running?" : err.message), true);
  }
}

// Sheets after the first one load on the server in the background; keep the sidebar up to date.
async function pollSheets() {
  const sid = state.sessionId;
  for (;;) {
    let d;
    try { d = await (await fetch(`${API_BASE}/api/sheets/${sid}`)).json(); } catch (e) { showInfo(null); return; }
    if (state.sessionId !== sid) return;                       // a newer file was uploaded
    d.sheets.forEach((s) => {
      const m = state.sheets.find((x) => x.name === s.name);
      if (m && s.loaded && !(s.name === state.currentSheet && m.loaded)) {
        m.columns = s.columns; m.row_count = s.row_count; m.loaded = true;
      }
    });
    renderSheetList();
    if (d.ready) {
      showInfo(null);
      const bad = Object.keys(d.errors || {});
      if (bad.length) toast(`Couldn't read: ${bad.join(", ")}`, true);
      return;
    }
    showInfo(`Reading sheets… ${d.loaded} of ${d.total} ready (you can already work on the open sheet)`);
    await sleep(700);
  }
}

// Download: the server builds the file in the background; show its progress, then fetch it.
$("#downloadBtn").addEventListener("click", async () => {
  if (!state.sessionId) return;
  const btn = $("#downloadBtn"), label = btn.textContent, sid = state.sessionId;
  btn.disabled = true;
  try {
    let st = await (await fetch(`${API_BASE}/api/download/start/${sid}`, { method: "POST" })).json();
    while (st.state === "running") {
      btn.textContent = `Preparing… ${st.done}/${st.total}`;
      showInfo(`Building your file: ${st.phase} (${st.done} of ${st.total} sheets)`);
      await sleep(700);
      st = await (await fetch(`${API_BASE}/api/download/status/${sid}`)).json();
    }
    if (st.state !== "done") throw new Error(st.error || "The file couldn't be built");
    window.location.href = `${API_BASE}/api/download/file/${sid}`;
    trackEvent("download_workbook", { mode: st.mode || "full" });
    if (st.mode === "stream") toast("Large workbook: every value and formula is saved; cell colours and fonts are not.");
  } catch (err) {
    toast("Download failed: " + err.message, true);
  } finally {
    btn.disabled = false; btn.textContent = label; showInfo(null);
  }
});

// ---------------------------------------------------------------------------
// Sheet list + grid rendering
// ---------------------------------------------------------------------------
// Every call to the server that changes the workbook goes through here (throws the server's message on failure)
async function sheetPost(path, body) {
  const res = await fetch(API_BASE + path, {
    method: "POST", headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ session_id: state.sessionId, ...body }),
  });
  let data = {};
  try { data = await res.json(); } catch (e) { /* not JSON */ }
  if (!res.ok) throw new Error(typeof data.detail === "string" ? data.detail : `The server said no (${res.status})`);
  return data;
}

const SA_ICONS = {
  rename: '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4z"/></svg>',
  dup: '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="12" height="12" rx="2"/><path d="M5 15V5a2 2 0 0 1 2-2h10"/></svg>',
  del: '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M6 6l1 14h10l1-14"/></svg>',
};
let renamingSheet = false;

function renderSheetList() {
  if (renamingSheet) return;                                     // don't wipe the box being typed in
  const container = $("#sheetList");
  container.innerHTML = "";
  state.sheets.forEach((s) => {
    const row = el("div", { className: "sheet-row" });
    const btn = el("button", { className: "sheet-btn", type: "button", title: "Double-click to rename" });
    btn.append(el("span", { className: "sheet-name", textContent: s.name }),
               el("span", { className: "sheet-count", textContent: `${s.loaded === false ? "~" : ""}${Number(s.row_count).toLocaleString("en-US")}` }));
    btn.dataset.name = s.name;
    if (s.name === state.currentSheet) row.classList.add("active");
    btn.addEventListener("click", () => selectSheet(s.name));
    btn.addEventListener("dblclick", () => startRenameSheet(row, s.name));
    const acts = el("span", { className: "sheet-actions" });
    [["rename", "Rename sheet"], ["dup", "Duplicate sheet"], ["del", "Delete sheet"]].forEach(([act, title]) => {
      const b = el("button", { className: "sa", type: "button", title, innerHTML: SA_ICONS[act] });
      b.addEventListener("click", (e) => {
        e.stopPropagation();
        if (act === "rename") startRenameSheet(row, s.name);
        else if (act === "dup") duplicateSheet(s.name);
        else deleteSheet(s.name);
      });
      acts.append(b);
    });
    row.append(btn, acts);
    container.appendChild(row);
  });
}

function applySheetReply(data) {
  state.sheets = data.sheets;
  renderSheetList();
}

function startRenameSheet(row, name) {
  renamingSheet = true;
  const input = el("input", { className: "sheet-rename", value: name, maxLength: 31 });
  row.replaceChildren(input);
  input.focus(); input.select();
  let done = false;
  const finish = async (commit) => {
    if (done) return;
    done = true;
    renamingSheet = false;
    const next = input.value.trim();
    if (!commit || !next || next === name) { renderSheetList(); return; }
    try {
      const data = await sheetPost("/api/sheet/rename", { sheet_name: name, new_name: next });
      const wasCurrent = state.currentSheet === name;
      applySheetReply(data);
      state.history.forEach((h) => { if (h.sheet === name) h.sheet = data.name; });
      renderHistory();
      if (wasCurrent) await selectSheet(data.name, { resetPage: true });
      toast(`Renamed to “${data.name}” — formulas that used it were updated`);
    } catch (err) { renderSheetList(); toast(err.message, true); }
  };
  input.addEventListener("keydown", (e) => {
    if (e.key === "Enter") { e.preventDefault(); finish(true); }
    else if (e.key === "Escape") { e.preventDefault(); finish(false); }
  });
  input.addEventListener("blur", () => finish(true));
}

async function duplicateSheet(name) {
  try {
    const data = await sheetPost("/api/sheet/duplicate", { sheet_name: name });
    applySheetReply(data);
    await selectSheet(data.name, { resetPage: true });
    toast(`Copied “${name}” to “${data.name}”`);
  } catch (err) { toast(err.message, true); }
}

async function deleteSheet(name) {
  const ok = await confirmBox(`Delete “${name}”?`,
    "The sheet and everything on it is removed. Formulas in other sheets that use it will show #REF! in Excel. This can't be undone.",
    "Delete sheet");
  if (!ok) return;
  try {
    const data = await sheetPost("/api/sheet/delete", { sheet_name: name });
    applySheetReply(data);
    if (state.currentSheet === name) await selectSheet(data.current, { resetPage: true });
    toast(`Deleted “${name}”`);
  } catch (err) { toast(err.message, true); }
}

async function createSheet() {
  if (!state.sessionId) { toast("Upload a workbook first", true); return; }
  const vals = await askForm("New sheet", [
    { key: "name", label: "Sheet name", value: `Sheet ${state.sheets.length + 1}` },
    { key: "columns", label: "Number of columns", value: "6", type: "number", min: 1, max: 200,
      hint: "Rename a column by double-clicking its name; add more with “+ Column”." },
  ], "Create sheet");
  if (!vals) return;
  try {
    const data = await sheetPost("/api/sheet/create", { name: vals.name, columns: parseInt(vals.columns, 10) || 6 });
    applySheetReply(data);
    await selectSheet(data.name, { resetPage: true });
    if (window.gridApi) window.gridApi.enableEditing();
    toast(`Created “${data.name}” — press + Row or paste data from Excel`);
  } catch (err) { toast(err.message, true); }
}
document.addEventListener("click", (e) => { if (e.target.closest && e.target.closest("#addSheetBtn")) createSheet(); });

// small dialogs (built on the same modal as the tools). fields: {key, label, value, type?, options?, hint?}
function askForm(title, fields, okText = "OK", intro = "") {
  return new Promise((resolve) => {
    const html = `<h2>${esc(title)}</h2>${intro ? `<p class="modal-sub">${intro}</p>` : ""}` +
      fields.map((f) => `<div class="field"><label>${esc(f.label)}</label>` +
        (f.options
          ? `<select data-k="${f.key}">${f.options.map((o) => `<option value="${esc(o[0])}"${o[0] === f.value ? " selected" : ""}>${esc(o[1])}</option>`).join("")}</select>`
          : `<input data-k="${f.key}" type="${f.type || "text"}" value="${esc(f.value ?? "")}"${f.min !== undefined ? ` min="${f.min}"` : ""}${f.max !== undefined ? ` max="${f.max}"` : ""}>`) +
        (f.hint ? `<div class="hint">${f.hint}</div>` : "") + `</div>`).join("") +
      `<div class="modal-actions"><button class="btn secondary" type="button" data-x="cancel">Cancel</button><button class="btn primary" type="button" data-x="ok">${esc(okText)}</button></div>`;
    openModal(html, (root) => {
      const read = () => Object.fromEntries([...root.querySelectorAll("[data-k]")].map((i) => [i.dataset.k, i.value]));
      const overlay = $("#modalOverlay");
      const done = (v) => {
        document.removeEventListener("keydown", onKey, true);
        overlay.removeEventListener("click", onAway);
        closeModal(); resolve(v);
      };
      const onKey = (e) => { if (e.key === "Escape") { e.stopPropagation(); done(null); } };
      const onAway = (e) => { if (e.target === overlay) done(null); };
      document.addEventListener("keydown", onKey, true);
      overlay.addEventListener("click", onAway);
      root.querySelector("[data-x=cancel]").addEventListener("click", () => done(null));
      root.querySelector("[data-x=ok]").addEventListener("click", () => done(read()));
      root.querySelectorAll("input").forEach((i) => i.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); done(read()); } }));
      const first = root.querySelector("input, select");
      if (first) { first.focus(); if (first.select) first.select(); }
    });
  });
}

function confirmBox(title, text, okText = "OK") {
  return new Promise((resolve) => {
    openModal(`<h2>${esc(title)}</h2><p class="modal-sub">${esc(text)}</p>` +
      `<div class="modal-actions"><button class="btn secondary" type="button" data-x="cancel">Cancel</button><button class="btn danger" type="button" data-x="ok">${esc(okText)}</button></div>`, (root) => {
      const done = (v) => { closeModal(); resolve(v); };
      root.querySelector("[data-x=cancel]").addEventListener("click", () => done(false));
      root.querySelector("[data-x=ok]").addEventListener("click", () => done(true));
      root.querySelector("[data-x=cancel]").focus();
    });
  });
}

async function selectSheet(name, opts = {}) {
  state.currentSheet = name;
  renderSheetList();
  const meta = state.sheets.find((s) => s.name === name);
  $("#headerRowInput").value = meta.header_row;
  $("#currentSheetLabel").textContent = name;
  await gridLoadSheet(name, opts);            // grid.js — fetches the page and draws it
}

$("#applyHeaderRowBtn").addEventListener("click", async () => {
  const headerRow = parseInt($("#headerRowInput").value, 10) || 1;
  const res = await fetch(`${API_BASE}/api/sheet/header-row`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ session_id: state.sessionId, sheet_name: state.currentSheet, header_row: headerRow }),
  });
  if (!res.ok) { toast("Failed to set header row", true); return; }
  const data = await res.json();
  const meta = state.sheets.find((s) => s.name === state.currentSheet);
  meta.header_row = headerRow;
  meta.columns = data.columns;
  meta.row_count = data.row_count;
  selectSheet(state.currentSheet, { resetPage: true });
  toast("Header row updated");
});

function currentSheetMeta() {
  return state.sheets.find((s) => s.name === state.currentSheet);
}

// ---------------------------------------------------------------------------
// Modal helpers
// ---------------------------------------------------------------------------
function esc(v) {
  return String(v).replace(/&/g, "&amp;").replace(/"/g, "&quot;").replace(/</g, "&lt;");
}

// "Put result in" — any existing column (overwrites it) or a brand-new one.
function outputField(meta, defaultName) {
  return `<div class="field">
    <label>Put result in</label>
    <div class="output-row">
    <select id="outputSel">
      <option value="__new__">+ New column</option>
      ${meta.columns.map((c) => `<option value="${esc(c)}">${esc(c)}</option>`).join("")}
    </select>
    <input id="outputCol" type="text" value="${esc(defaultName)}" placeholder="New column name">
    </div>
    <div class="hint" id="outputHint"></div>
  </div>`;
}
function wireOutput(root) {
  const sel = root.querySelector("#outputSel");
  if (!sel) return;
  const input = root.querySelector("#outputCol");
  const hint = root.querySelector("#outputHint");
  const refresh = () => {
    const isNew = sel.value === "__new__";
    input.hidden = !isNew;
    hint.textContent = isNew
      ? "Added as a new column at the end."
      : `Replaces the values in "${sel.value}". If it is also an input, values are written instead of a formula.`;
  };
  sel.addEventListener("change", refresh);
  refresh();
}
function readOutput(root, fallback) {
  const sel = root.querySelector("#outputSel");
  if (sel && sel.value !== "__new__") return sel.value;
  return root.querySelector("#outputCol").value.trim() || fallback;
}

function openModal(html, onMount, wide = false) {
  $("#modalBody").className = "modal" + (wide ? " wide" : "");
  $("#modalBody").innerHTML = html;
  $("#modalOverlay").hidden = false;
  wireOutput($("#modalBody"));
  if (onMount) onMount($("#modalBody"));
}
function closeModal() {
  $("#modalOverlay").hidden = true;
}
$("#modalOverlay").addEventListener("click", (e) => {
  if (e.target.id === "modalOverlay") closeModal();
});

function columnOptions(columns, selected) {
  return columns.map((c) => `<option value="${c}" ${c === selected ? "selected" : ""}>${c}</option>`).join("");
}
function sheetOptions(selected) {
  return state.sheets.map((s) => `<option value="${s.name}" ${s.name === selected ? "selected" : ""}>${s.name}</option>`).join("");
}

// ---------------------------------------------------------------------------
// Operation button wiring
// ---------------------------------------------------------------------------
// One delegated listener (works for any click on a menu item, including its icon and text)
document.addEventListener("click", (e) => {
  const btn = e.target.closest && e.target.closest(".op-btn");
  if (!btn) return;
  if (!state.currentSheet) { toast("Upload a file and pick a sheet first", true); return; }
  try {
    if (typeof openBuilder !== "function") throw new Error("builder.js did not load — press Ctrl+F5");
    trackEvent("tool_opened", { tool: btn.dataset.op });
    if (btn.dataset.op === "SQL") {
      if (typeof openSql !== "function") throw new Error("sql.js did not load — press Ctrl+F5");
      openSql();
    } else {
      openBuilder(btn.dataset.op);
    }
  } catch (err) {
    console.error(err);
    toast("Couldn't open this tool: " + err.message, true);
  }
});

async function submitOperation(payload) {
  try {
    const res = await fetch(`${API_BASE}/api/operations/apply`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    if (!res.ok) {
      const err = await res.json();
      throw new Error(err.detail || "Operation failed");
    }
    const data = await res.json();
    closeModal();
    const meta = currentSheetMeta();
    meta.columns = data.columns;
    meta.row_count = data.total_rows;
    selectSheet(state.currentSheet);
    state.history.unshift({
      op: payload.label || payload.operation,
      output: data.output_column,
      formula: data.sample_formula || "(values)",
      sheet: payload.sheet_name,
    });
    renderHistory();
    trackEvent("operation_applied", { operation: payload.operation });
    toast(data.static
      ? `Wrote values to "${data.output_column}" (it is also an input, so no formula)`
      : `Applied to column "${data.output_column}"`);
  } catch (err) {
    toast(err.message, true);
  }
}

function renderHistory() {
  const container = $("#historyList");
  container.innerHTML = "";
  if (state.history.length === 0) {
    container.innerHTML = '<div class="hint">No changes yet.</div>';
    return;
  }
  state.history.forEach((h) => {
    const div = el("div", { className: "history-item" });
    div.innerHTML = `<strong>${h.op}</strong> → ${h.sheet}.${h.output}<br><code>${h.formula || ""}</code>`;
    container.appendChild(div);
  });
}

window.closeModal = closeModal;

// ---------------------------------------------------------------------------
// Light / dark theme (remembered in localStorage)
// ---------------------------------------------------------------------------
(function () {
  const btn = $("#themeToggle");
  const root = document.documentElement;
  const sync = () => (btn.textContent = root.getAttribute("data-theme") === "dark" ? "Light mode" : "Dark mode");
  btn.addEventListener("click", () => {
    const next = root.getAttribute("data-theme") === "dark" ? "light" : "dark";
    root.setAttribute("data-theme", next);
    try { localStorage.setItem("theme", next); } catch (e) {}
    sync();
  });
  sync();
})();

// ---------------------------------------------------------------------------
// Backend check: an old server that was never restarted fails quietly, so say so.
// ---------------------------------------------------------------------------
const EXPECTED_API_VERSION = 7;
const FRONTEND_VERSION = 20;               // grid.js and builder.js must report the same number

function showBanner(msg) {
  const old = document.querySelector(".banner[data-kind='" + msg.slice(0, 12) + "']");
  if (old) return;
  const bar = el("div", { className: "banner", textContent: msg });
  bar.dataset.kind = msg.slice(0, 12);
  document.body.insertBefore(bar, document.body.firstChild);
}

// A browser that keeps OLD copies of some script files makes the page fail silently, so check.
window.addEventListener("load", () => {
  const parts = window.FRONTEND_PARTS || {};
  const stale = ["grid", "builder", "sql", "dup", "tools", "compare"].filter((k) => parts[k] !== FRONTEND_VERSION).map((k) => k + ".js");
  if (stale.length) {
    showBanner(`Your browser is using out-of-date copies of ${stale.join(" and ")}. Press Ctrl+F5 to reload the page files.`);
  }
});

(async function checkBackend() {
  try {
    const h = await (await fetch(`${API_BASE}/api/health`)).json();
    if ((h.frontend || 0) > FRONTEND_VERSION) {
      showBanner("Your browser is showing an OLD copy of this page. Press Ctrl+F5 (or open http://localhost:8000 directly) to load the new one.");
    }
    if ((h.version || 0) < EXPECTED_API_VERSION) {
      showBanner("The backend is running an OLD version of the code, so some tools and cell editing won't work. " +
                 "Stop it (Ctrl+C) and start it again: uvicorn main:app --reload --port 8000");
    }
  } catch (e) {
    showBanner(`Can't reach the backend at ${API_BASE}. Start it from the backend folder: uvicorn main:app --reload --port 8000`);
  }
})();
