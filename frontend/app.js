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
// A fresh, un-hidden input is created for every click: some embedded browsers (IDE previews, remote
// desktops) ignore .click() on a display:none input. The original input is kept as a fallback.
function openFilePicker() {
  try {
    const inp = document.createElement("input");
    inp.type = "file"; inp.accept = ".xlsx";
    inp.style.cssText = "position:fixed;left:-9999px;top:0;opacity:0";
    inp.addEventListener("change", () => { if (inp.files[0]) handleFile(inp.files[0]); inp.remove(); });
    document.body.appendChild(inp);
    inp.click();
  } catch (e) { $("#fileInput").click(); }
}
$("#uploadBtn").addEventListener("click", openFilePicker);
const uploadBtn2 = $("#uploadBtn2");
if (uploadBtn2) uploadBtn2.addEventListener("click", openFilePicker);

$("#fileInput").addEventListener("change", (e) => {
  const file = e.target.files[0];
  if (file) handleFile(file);
});

const dropZone = $("#dropZone");
if (dropZone) {
  ["dragenter", "dragover"].forEach((evt) =>
    dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add("dragover");
    })
  );
  ["dragleave", "drop"].forEach((evt) =>
    dropZone.addEventListener(evt, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove("dragover");
    })
  );
  dropZone.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files && e.dataTransfer.files[0];
    if (!file) return;
    if (!file.name.toLowerCase().endsWith(".xlsx")) {
      toast("Please drop a .xlsx file", true);
      return;
    }
    handleFile(file);
  });
}

// Also allow dropping anywhere on the page once the app is loaded, as a
// safety net (some browsers only fire drop reliably on the whole window).
window.addEventListener("dragover", (e) => e.preventDefault());
window.addEventListener("drop", (e) => {
  if (e.target.closest && e.target.closest("#dropZone")) return; // already handled above
  e.preventDefault();
  const file = e.dataTransfer.files && e.dataTransfer.files[0];
  if (file && file.name.toLowerCase().endsWith(".xlsx")) handleFile(file);
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
  try {
    showInfo(`Uploading ${file.name} (${mb} MB)…`);
    const res = await fetch(`${API_BASE}/api/upload`, { method: "POST", body: fd });
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
    toast(`Loaded ${file.name} — ${data.sheets.length} sheet(s)` +
          (data.mode === "stream" ? " · large workbook mode" : ""));
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
function renderSheetList() {
  const container = $("#sheetList");
  container.innerHTML = "";
  state.sheets.forEach((s) => {
    const btn = el("button", { className: "sheet-btn", textContent: `${s.name} (${s.loaded === false ? "~" : ""}${Number(s.row_count).toLocaleString("en-US")})` });
    btn.dataset.name = s.name;
    if (s.name === state.currentSheet) btn.classList.add("active");
    btn.addEventListener("click", () => selectSheet(s.name));
    container.appendChild(btn);
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
const EXPECTED_API_VERSION = 6;
const FRONTEND_VERSION = 17;               // grid.js and builder.js must report the same number

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
  const stale = ["grid", "builder", "sql", "dup"].filter((k) => parts[k] !== FRONTEND_VERSION).map((k) => k + ".js");
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
