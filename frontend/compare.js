/* compare.js — side-by-side file comparison.
 *
 * Two editable tables (paste data from Excel / CSV). Rows are matched wherever they sit (row 10 of A can match row 107 of B):
 *  - no "Key" column ticked  -> a row matches when ALL mapped columns are equal;
 *  - some "Key" columns ticked -> rows match on the key, and the other mapped columns are compared ("changed" rows).
 * Duplicate rows are paired one-to-one in order. Everything runs in the browser and re-compares after each edit.
 */
(function () {
  "use strict";
  (window.FRONTEND_PARTS = window.FRONTEND_PARTS || {}).compare = 20;

  const RH = 28, HH = 32, COLW = 150, GUT = 104;
  const ST = { BLANK: 0, SAME: 1, DIFF: 2, ONLY: 3 };
  const esc = (s) => String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const fmt = (n) => Number(n).toLocaleString("en-US");

  const newSide = (name) => ({ name, head: [], rows: [], hasHeader: true });
  const S = { A: newSide("File A"), B: newSide("File B") };
  const opt = { nocase: true, trim: true, num: true };
  const view = { filter: "all", aligned: true };
  let mapping = [], userMapped = false;
  let R = null;                                   // compare result
  let hov = null, sel = null, flash = null, editing = null, timer = null, undoStack = [], root = null, tip = null;
  const P = {};                                   // panes

  // ----------------------------------------------------------------- parsing
  function parseText(text) {
    text = String(text).replace(/\r\n?/g, "\n");
    const first = text.split("\n", 1)[0];
    let d = "\t";
    if (!first.includes("\t")) d = first.split(";").length > first.split(",").length ? ";" : (first.includes(",") ? "," : "\t");
    const out = []; let row = [], cur = "", q = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (q) {
        if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; } else cur += ch;
      } else if (ch === '"' && cur === "") q = true;
      else if (ch === d) { row.push(cur); cur = ""; }
      else if (ch === "\n") { row.push(cur); out.push(row); row = []; cur = ""; }
      else cur += ch;
    }
    if (cur !== "" || row.length) { row.push(cur); out.push(row); }
    while (out.length && out[out.length - 1].every((c) => c === "")) out.pop();
    return out;
  }
  const toTSV = (rows) => rows.map((r) => r.map((c) => (/[\t\n"]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c)).join("\t")).join("\n");
  const toCSV = (rows) => rows.map((r) => r.map((c) => (/[,\n"]/.test(c) ? '"' + c.replace(/"/g, '""') + '"' : c)).join(",")).join("\n");

  // -------------------------------------------------------------- data ops
  const colName = (i) => "Column " + (i + 1);
  const snap = () => ({
    A: { head: S.A.head.slice(), rows: S.A.rows.slice(), hasHeader: S.A.hasHeader },
    B: { head: S.B.head.slice(), rows: S.B.rows.slice(), hasHeader: S.B.hasHeader },
  });
  function pushUndo() { undoStack.push(snap()); if (undoStack.length > 40) undoStack.shift(); }
  function undo() {
    const s = undoStack.pop();
    if (!s) { toast("Nothing to undo"); return; }
    Object.assign(S.A, s.A); Object.assign(S.B, s.B);
    syncMapping(); changed(true);
  }

  function load(k, matrix, append) {
    if (!matrix.length) return;
    pushUndo();
    const s = S[k], w = Math.max(...matrix.map((r) => r.length));
    const pad = (r, n) => { const x = r.slice(); while (x.length < n) x.push(""); return x; };
    let head, rows;
    if (append && s.head.length) {
      const n = Math.max(w, s.head.length);
      head = pad(s.head, n).map((h, i) => h || colName(i));
      rows = s.rows.map((r) => pad(r, n)).concat(matrix.map((r) => pad(r, n)));
    } else if (s.hasHeader) {
      head = pad(matrix[0], w).map((h, i) => String(h).trim() || colName(i));
      rows = matrix.slice(1).map((r) => pad(r, w));
    } else {
      head = Array.from({ length: w }, (_, i) => colName(i));
      rows = matrix.map((r) => pad(r, w));
    }
    s.head = head; s.rows = rows;
    syncMapping(true); changed(true);
    trackEvent("compare_loaded", { side: k, rows: rows.length, cols: head.length });
  }
  function setHeaderMode(k, on) {
    const s = S[k]; if (on === s.hasHeader) return;
    pushUndo(); s.hasHeader = on;
    if (s.head.length) {
      if (on && s.rows.length) { s.head = s.rows[0].map((h, i) => String(h).trim() || colName(i)); s.rows = s.rows.slice(1); }
      else if (!on) { s.rows = [s.head.slice()].concat(s.rows); s.head = s.head.map((_, i) => colName(i)); }
    }
    syncMapping(true); changed(true);
  }
  function setCell(k, r, c, v) {
    const s = S[k]; if (!s.rows[r] || s.rows[r][c] === v) return false;
    pushUndo(); const row = s.rows[r].slice(); row[c] = v; s.rows[r] = row; return true;
  }
  function pasteBlock(k, r, c, matrix) {
    const s = S[k]; pushUndo();
    const needCols = c + Math.max(...matrix.map((x) => x.length));
    if (needCols > s.head.length) {
      for (let i = s.head.length; i < needCols; i++) s.head.push(colName(i));
      s.rows = s.rows.map((row) => { const x = row.slice(); while (x.length < needCols) x.push(""); return x; });
    }
    matrix.forEach((vals, dr) => {
      const ri = r + dr;
      while (ri >= s.rows.length) s.rows.push(new Array(s.head.length).fill(""));
      const row = s.rows[ri].slice();
      vals.forEach((v, dc) => { row[c + dc] = v; });
      s.rows[ri] = row;
    });
    syncMapping(); changed(true);
  }
  const addRow = (k, after) => { const s = S[k]; pushUndo(); s.rows.splice(after == null ? s.rows.length : after + 1, 0, new Array(s.head.length).fill("")); changed(true); };
  const delRow = (k, r) => { pushUndo(); S[k].rows.splice(r, 1); changed(true); };
  function addCol(k) {
    const s = S[k]; pushUndo();
    s.head = s.head.concat(colName(s.head.length)); s.rows = s.rows.map((r) => r.concat(""));
    syncMapping(); changed(true);
  }
  function delCol(k, c) {
    const s = S[k]; if (s.head.length <= 1) { toast("A table needs at least one column", true); return; }
    pushUndo();
    s.head = s.head.filter((_, i) => i !== c); s.rows = s.rows.map((r) => r.filter((_, i) => i !== c));
    mapping = mapping.filter((m) => !(k === "A" && m.a === c));
    mapping.forEach((m) => {
      if (k === "A" && m.a > c) m.a--;
      if (k === "B") { if (m.b === c) { m.b = -1; m.use = false; } else if (m.b > c) m.b--; }
    });
    syncMapping(); changed(true);
  }

  // ---------------------------------------------------------- column mapping
  function syncMapping(force) {
    const a = S.A.head, b = S.B.head;
    if (!a.length || !b.length) { mapping = []; return; }
    if (userMapped && !force) {
      mapping = mapping.filter((m) => m.a < a.length).map((m) => (m.b >= b.length ? { ...m, b: -1, use: false } : m));
      for (let i = 0; i < a.length; i++) if (!mapping.some((m) => m.a === i)) mapping.push({ a: i, b: -1, use: false, key: false });
      mapping.sort((x, y) => x.a - y.a);
      return;
    }
    userMapped = false;
    const key = (h) => String(h).trim().toLowerCase().replace(/[\s_\-]+/g, "");
    const byName = new Map(); b.forEach((h, i) => { const k = key(h); if (k && !byName.has(k)) byName.set(k, i); });
    const named = a.some((h) => byName.has(key(h))) && S.A.hasHeader && S.B.hasHeader;
    const used = new Set();
    mapping = a.map((h, i) => {
      let j = -1;
      if (named) { const t = byName.get(key(h)); if (t != null && !used.has(t)) j = t; }
      else if (i < b.length) j = i;
      if (j >= 0) used.add(j);
      return { a: i, b: j, use: j >= 0, key: false };
    });
  }

  // ----------------------------------------------------------------- compare
  function norm(v) {
    let s = v == null ? "" : String(v);
    if (opt.trim) s = s.trim().replace(/\s+/g, " ");
    if (opt.num) {
      const t = s.replace(/[,\s$€£]/g, "");
      if (t !== "" && /^[-+]?(\d+\.?\d*|\.\d+)(e[-+]?\d+)?$/i.test(t)) s = String(Number(t));
    }
    return opt.nocase ? s.toLowerCase() : s;
  }
  function compare() {
    const act = mapping.filter((m) => m.use && m.a >= 0 && m.b >= 0);
    if (!act.length || !S.A.head.length || !S.B.head.length) { R = null; return; }
    const keys = act.filter((m) => m.key), sig = keys.length ? keys : act, rest = keys.length ? act.filter((m) => !m.key) : [];
    const nA = S.A.rows.length, nB = S.B.rows.length;
    const out = {
      keyMode: keys.length > 0, act, keys, rest,
      A: { st: new Uint8Array(nA), m: new Int32Array(nA).fill(-1), dc: new Map() },
      B: { st: new Uint8Array(nB), m: new Int32Array(nB).fill(-1), dc: new Map() },
    };
    const sigOf = (row, side) => sig.map((p) => norm(row[side === "A" ? p.a : p.b])).join("\u001f");
    const isBlank = (row, side) => act.every((p) => norm(row[side === "A" ? p.a : p.b]) === "");
    const noKey = (row, side) => keys.length > 0 && keys.every((p) => norm(row[side === "A" ? p.a : p.b]) === "");
    const pool = new Map();
    for (let j = 0; j < nB; j++) {
      const row = S.B.rows[j];
      if (isBlank(row, "B")) { out.B.st[j] = ST.BLANK; continue; }
      if (noKey(row, "B")) { out.B.st[j] = ST.ONLY; continue; }
      out.B.st[j] = ST.ONLY;
      const s = sigOf(row, "B"); let q = pool.get(s);
      if (!q) pool.set(s, (q = { l: [], p: 0 }));
      q.l.push(j);
    }
    for (let i = 0; i < nA; i++) {
      const row = S.A.rows[i];
      if (isBlank(row, "A")) { out.A.st[i] = ST.BLANK; continue; }
      out.A.st[i] = ST.ONLY;
      if (noKey(row, "A")) continue;
      const q = pool.get(sigOf(row, "A"));
      if (!q || q.p >= q.l.length) continue;
      const j = q.l[q.p++], other = S.B.rows[j];
      out.A.m[i] = j; out.B.m[j] = i;
      const da = [], db = [];
      rest.forEach((p) => { if (norm(row[p.a]) !== norm(other[p.b])) { da.push(p.a); db.push(p.b); } });
      if (da.length) { out.A.st[i] = out.B.st[j] = ST.DIFF; out.A.dc.set(i, da); out.B.dc.set(j, db); }
      else out.A.st[i] = out.B.st[j] = ST.SAME;
    }
    const cnt = { same: 0, diff: 0, onlyA: 0, onlyB: 0 };
    out.A.st.forEach((s) => { if (s === ST.SAME) cnt.same++; else if (s === ST.DIFF) cnt.diff++; else if (s === ST.ONLY) cnt.onlyA++; });
    out.B.st.forEach((s) => { if (s === ST.ONLY) cnt.onlyB++; });
    out.cnt = cnt; R = out;
  }

  // ------------------------------------------------------------------- views
  function buildViews() {
    const f = view.filter, nA = S.A.rows.length, nB = S.B.rows.length;
    const stA = (i) => (R ? R.A.st[i] : ST.ONLY), stB = (j) => (R ? R.B.st[j] : ST.ONLY);
    const pass = (st) => f === "all" || (f === "same" && st === ST.SAME) || (f === "diff" && st === ST.DIFF) || (f === "only" && st === ST.ONLY);
    P.aligned = !!R && view.aligned && f !== "only";
    if (P.aligned) {
      const sl = [];
      for (let i = 0; i < nA; i++) sl.push([i, R.A.m[i]]);
      for (let j = 0; j < nB; j++) if (R.B.m[j] < 0) sl.push([-1, j]);
      const keep = f === "all" ? sl : sl.filter((s) => pass(s[0] >= 0 ? stA(s[0]) : stB(s[1])));
      P.A.view = keep.map((s) => s[0]); P.B.view = keep.map((s) => s[1]);
    } else {
      P.A.view = []; P.B.view = [];
      for (let i = 0; i < nA; i++) if (pass(stA(i))) P.A.view.push(i);
      for (let j = 0; j < nB; j++) if (pass(stB(j))) P.B.view.push(j);
    }
    ["A", "B"].forEach((k) => {
      const pos = new Int32Array(S[k].rows.length).fill(-1);
      P[k].view.forEach((d, vp) => { if (d >= 0) pos[d] = vp; });
      P[k].pos = pos;
    });
  }

  // --------------------------------------------------------------- rendering
  const other = (k) => (k === "A" ? "B" : "A");
  function rowHtml(k, vp) {
    const s = S[k], d = P[k].view[vp], top = HH + vp * RH;
    if (d < 0) return `<div class="cmp-row gap" data-vp="${vp}" style="top:${top}px"><div class="cmp-gut"></div></div>`;
    const st = R ? R[k].st[d] : ST.ONLY, m = R ? R[k].m[d] : -1;
    let cls = "cmp-row";
    if (R) cls += st === ST.SAME ? " same" : st === ST.DIFF ? " diff" : st === ST.ONLY ? " only" : " blank";
    if (hov && hov.k === k && hov.i === d) cls += " hv";
    else if (hov && hov.k === other(k) && R && R[hov.k].m[hov.i] === d) cls += " pr";
    if (flash && flash.k === k && flash.i === d) cls += " flash";
    if (isSel(k, d)) cls += " sel";
    const dc = R && R[k].dc.get(d), row = s.rows[d];
    let cells = "";
    for (let c = 0; c < s.head.length; c++) {
      const mp = mapping.find((x) => (k === "A" ? x.a : x.b) === c);
      const used = mp && mp.use && mp.a >= 0 && mp.b >= 0;
      cells += `<div class="cmp-cell${dc && dc.includes(c) ? " dc" : ""}${used ? "" : " nu"}" data-c="${c}" title="${esc(row[c])}">${esc(row[c])}</div>`;
    }
    const ref = m >= 0 ? `<span class="ref" title="Matches row ${m + 1} in ${esc(S[other(k)].name)}">↔${m + 1}</span>` : "";
    return `<div class="${cls}" data-vp="${vp}" data-i="${d}" style="top:${top}px"><div class="cmp-gut"><span class="n">${d + 1}</span>${ref}` +
      `<span class="ra"><b data-act="add" title="Insert a row below">+</b><b data-act="del" title="Delete this row">✕</b></span></div>${cells}</div>`;
  }
  function headHtml(k) {
    const s = S[k];
    return `<div class="cmp-hrow" style="height:${HH}px"><div class="cmp-gut"><span class="n">#</span></div>` + s.head.map((h, c) => {
      const mp = mapping.find((x) => (k === "A" ? x.a : x.b) === c);
      const used = mp && mp.use && mp.a >= 0 && mp.b >= 0;
      return `<div class="cmp-hcell${used ? "" : " nu"}" data-c="${c}" title="${used ? (mp.key ? "Key column" : "Compared column") : "Not compared (not mapped)"}">` +
        `${mp && mp.key && used ? '<i class="key">KEY</i>' : ""}<span class="hn">${esc(h)}</span><b data-act="delcol" title="Delete this column">✕</b></div>`;
    }).join("") + "</div>";
  }
  function renderPane(k, force) {
    const p = P[k], s = S[k], n = p.view.length;
    const empty = !s.head.length;
    p.el.classList.toggle("is-empty", empty);
    if (empty) { p.content.style.height = "0"; p.rows.innerHTML = ""; p.hd.innerHTML = ""; return; }
    p.content.style.width = GUT + s.head.length * COLW + "px";
    p.content.style.height = HH + n * RH + "px";
    const top = p.body.scrollTop, h = p.body.clientHeight || 400;
    const a = Math.max(0, Math.floor((top - HH) / RH) - 6), b = Math.min(n, Math.ceil((top + h) / RH) + 6);
    const key = a + ":" + b;
    if (!force && p.key === key) return;
    p.key = key;
    if (force || p.headDirty !== false) { p.hd.innerHTML = headHtml(k); p.headDirty = false; }
    let html = "";
    for (let vp = a; vp < b; vp++) html += rowHtml(k, vp);
    p.rows.innerHTML = html;
  }
  function renderAll(force = true) {
    if (sel && (!S[sel.k].rows[sel.i])) sel = null;
    buildViews();
    ["A", "B"].forEach((k) => { P[k].headDirty = true; renderPane(k, force); });
    renderBar(); renderMapping();
  }
  function renderBar() {
    const c = R && R.cnt;
    const set = (id, n) => { root.querySelector(id).textContent = n == null ? "–" : fmt(n); };
    set("#cmpNSame", c && c.same); set("#cmpNDiff", c && c.diff); set("#cmpNA", c && c.onlyA); set("#cmpNB", c && c.onlyB);
    root.querySelectorAll(".cmp-chip").forEach((b) => b.classList.toggle("on", b.dataset.f === view.filter));
    root.querySelector("#cmpChipDiff").hidden = !(R && R.keyMode);
    ["A", "B"].forEach((k) => { P[k].count.textContent = fmt(S[k].rows.length) + " rows · " + fmt(S[k].head.length) + " cols"; });
    let msg;
    if (!S.A.head.length || !S.B.head.length) msg = "Paste data into both sides to compare.";
    else if (!R) msg = "Tick at least one column to compare (Columns panel).";
    else if (R.keyMode) msg = `Matching on key <b>${R.keys.map((p) => esc(S.A.head[p.a])).join(" + ")}</b>; ${R.rest.length} other column(s) are compared and shown as “changed”.`;
    else msg = `A row matches when all <b>${R.act.length}</b> mapped column(s) are equal — wherever it sits in the other file.`;
    root.querySelector("#cmpMode").innerHTML = msg;
    root.querySelector("#cmpUndo").disabled = !undoStack.length;
  }
  function renderMapping() {
    const box = root.querySelector("#cmpMapBody");
    if (!mapping.length) { box.innerHTML = '<div class="hint">Load both sides first.</div>'; return; }
    const opts = (sel) => '<option value="-1">— none —</option>' + S.B.head.map((h, i) => `<option value="${i}"${i === sel ? " selected" : ""}>${esc(h)}</option>`).join("");
    box.innerHTML = '<table class="cmp-map"><thead><tr><th>Compare</th><th>' + esc(S.A.name) + ' column</th><th></th><th>' + esc(S.B.name) +
      ' column</th><th title="Match rows on this column instead of on the whole row">Key</th></tr></thead><tbody>' +
      mapping.map((m, i) => `<tr data-m="${i}"><td><input type="checkbox" data-f="use"${m.use && m.b >= 0 ? " checked" : ""}${m.b < 0 ? " disabled" : ""}></td>` +
        `<td>${esc(S.A.head[m.a])}</td><td>→</td><td><select data-f="b">${opts(m.b)}</select></td>` +
        `<td><input type="checkbox" data-f="key"${m.key ? " checked" : ""}${m.b < 0 ? " disabled" : ""}></td></tr>`).join("") + "</tbody></table>";
  }

  // --------------------------------------------------------------- editing
  function commitEdit(save = true) {
    if (!editing) return;
    const e = editing; editing = null;
    const v = e.input.value;
    if (e.kind === "cell") {
      if (save && setCell(e.k, e.r, e.c, v)) schedule();
      e.cellEl.textContent = S[e.k].rows[e.r] ? S[e.k].rows[e.r][e.c] : "";
      e.cellEl.classList.remove("editing");
    } else {
      if (save && v.trim() && v !== S[e.k].head[e.c]) { pushUndo(); S[e.k].head[e.c] = v.trim(); syncMapping(); schedule(); }
      renderAll(true);
    }
  }
  function startCell(k, vp, c) {
    commitEdit();
    const p = P[k], d = p.view[vp]; if (d == null || d < 0) return;
    const cellEl = p.rows.querySelector(`.cmp-row[data-vp="${vp}"] .cmp-cell[data-c="${c}"]`); if (!cellEl) return;
    const input = document.createElement("input");
    input.className = "cmp-input"; input.value = S[k].rows[d][c];
    cellEl.textContent = ""; cellEl.classList.add("editing"); cellEl.appendChild(input);
    editing = { kind: "cell", k, vp, r: d, c, input, cellEl };
    input.focus(); input.select();
    input.addEventListener("blur", () => { if (editing && editing.input === input) commitEdit(); });
    input.addEventListener("keydown", (ev) => {
      if (ev.key === "Escape") { ev.preventDefault(); commitEdit(false); }
      else if (ev.key === "Enter") { ev.preventDefault(); moveEdit(k, vp, c, ev.shiftKey ? -1 : 1, 0); }
      else if (ev.key === "Tab") { ev.preventDefault(); moveEdit(k, vp, c, 0, ev.shiftKey ? -1 : 1); }
      else if (ev.key === "ArrowDown") { ev.preventDefault(); moveEdit(k, vp, c, 1, 0); }
      else if (ev.key === "ArrowUp") { ev.preventDefault(); moveEdit(k, vp, c, -1, 0); }
    });
    input.addEventListener("paste", (ev) => {
      const t = (ev.clipboardData || window.clipboardData).getData("text");
      if (!/[\t\n]/.test(t.replace(/\n$/, ""))) return;
      ev.preventDefault();
      const m = parseText(t); editing = null;
      pasteBlock(k, d, c, m);
    });
  }
  function moveEdit(k, vp, c, dr, dc) {
    const p = P[k], s = S[k]; commitEdit();
    let nv = vp + dr, nc = c + dc;
    if (nc >= s.head.length) { nc = 0; nv++; } else if (nc < 0) { nc = s.head.length - 1; nv--; }
    while (nv >= 0 && nv < p.view.length && p.view[nv] < 0) nv += dr || 1;
    if (nv < 0 || nv >= p.view.length) return;
    const t = HH + nv * RH;
    if (t < p.body.scrollTop + HH) p.body.scrollTop = t - HH;
    else if (t + RH > p.body.scrollTop + p.body.clientHeight) p.body.scrollTop = t + RH - p.body.clientHeight;
    renderPane(k, true);
    startCell(k, nv, nc);
  }
  function startHead(k, c, nameEl) {
    commitEdit();
    const input = document.createElement("input");
    input.className = "cmp-input"; input.value = S[k].head[c];
    nameEl.textContent = ""; nameEl.appendChild(input);
    editing = { kind: "head", k, c, input };
    input.focus(); input.select();
    input.addEventListener("blur", () => { if (editing && editing.input === input) commitEdit(); });
    input.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") { ev.preventDefault(); commitEdit(); }
      else if (ev.key === "Escape") { ev.preventDefault(); commitEdit(false); }
    });
  }

  function schedule() {
    clearTimeout(timer);
    timer = setTimeout(function run() {
      if (editing) { timer = setTimeout(run, 400); return; }
      compare(); renderAll(true);
    }, 220);
  }
  function changed(now) {
    if (now && !editing) { clearTimeout(timer); compare(); renderAll(true); } else schedule();
  }

  // ------------------------------------------------------------- hover / jump
  function isSel(k, d) {
    if (!sel) return false;
    return (sel.k === k && sel.i === d) || (R && sel.k === other(k) && R[sel.k].m[sel.i] === d);
  }
  function applySel() {
    ["A", "B"].forEach((x) => P[x].rows.querySelectorAll(".cmp-row").forEach((r) => {
      if (r.dataset.i != null) r.classList.toggle("sel", isSel(x, +r.dataset.i));
    }));
  }
  function revealRow(k, d) {
    const p = P[k], vp = p.pos[d]; if (vp == null || vp < 0) return;
    const t = HH + vp * RH;
    if (t < p.body.scrollTop + HH || t + RH > p.body.scrollTop + p.body.clientHeight) {
      p.body.scrollTop = Math.max(0, t - p.body.clientHeight / 2); renderPane(k, true);
    }
  }
  // click a row: it and its partner in the other file are selected (orange) and the partner is scrolled into view
  function select(k, d, toggle) {
    if (toggle && sel && sel.k === k && sel.i === d) sel = null; else sel = { k, i: d };
    applySel();
    if (sel && R) {
      const m = R[k].m[d];
      if (m >= 0) revealRow(other(k), m); else toast("Not found in " + S[other(k)].name);
    }
  }
  function setHover(k, i) {
    if (hov && hov.k === k && hov.i === i) return;
    hov = i == null ? null : { k, i };
    ["A", "B"].forEach((x) => {
      P[x].rows.querySelectorAll(".hv,.pr").forEach((r) => r.classList.remove("hv", "pr"));
      if (!hov) return;
      const sel = x === hov.k ? hov.i : (R ? R[hov.k].m[hov.i] : -1);
      if (sel < 0) return;
      const r = P[x].rows.querySelector(`.cmp-row[data-i="${sel}"]`);
      if (r) r.classList.add(x === hov.k ? "hv" : "pr");
    });
  }
  function tipText(k, d) {
    if (!R) return "";
    const st = R[k].st[d], o = other(k), on = esc(S[o].name), m = R[k].m[d];
    if (st === ST.BLANK) return "";
    if (st === ST.SAME) return `<b>Match</b> — row ${m + 1} in ${on}`;
    if (st === ST.ONLY) return `<b>Not found</b> in ${on}`;
    const dc = R[k].dc.get(d) || [], dm = R[o].dc.get(m) || [];
    const lines = dc.slice(0, 5).map((c, n) => `<div>${esc(S[k].head[c])}: “${esc(S[k].rows[d][c])}” vs “${esc(S[o].rows[m][dm[n]])}”</div>`).join("");
    return `<b>Key matches</b> row ${m + 1} in ${on}, but differs in ${dc.length} column(s):${lines}${dc.length > 5 ? "<div>…</div>" : ""}`;
  }
  function jump(k, d) {
    const m = R ? R[k].m[d] : -1, o = other(k);
    if (m < 0) { toast("This row has no match in " + S[o].name); return; }
    const vp = P[o].pos[m];
    if (vp < 0) { toast("The matching row is hidden by the current filter"); return; }
    const p = P[o];
    p.body.scrollTop = Math.max(0, HH + vp * RH - p.body.clientHeight / 2);
    flash = { k: o, i: m }; renderPane(o, true);
    setTimeout(() => { flash = null; const r = p.rows.querySelector(".flash"); if (r) r.classList.remove("flash"); }, 1400);
  }

  // ------------------------------------------------------------------- export
  function rowsFor(k, status) {
    const out = [];
    S[k].rows.forEach((r, i) => { if (R && R[k].st[i] === status) out.push(r); });
    return out;
  }
  function copyText(text, msg) {
    (navigator.clipboard ? navigator.clipboard.writeText(text) : Promise.reject()).catch(() => {
      const t = document.createElement("textarea"); t.value = text; document.body.appendChild(t); t.select(); document.execCommand("copy"); t.remove();
    }).then(() => toast(msg));
  }
  function download(text, name) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob(["﻿" + text], { type: "text/csv" }));
    a.download = name; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }
  // one sheet: matching rows, then changed, then only in A, then only in B (A columns | B columns side by side)
  function compiled() {
    const A = S.A, B = S.B, ea = A.head.map(() => ""), eb = B.head.map(() => "");
    const head = ["Status", "Row in " + A.name, "Row in " + B.name].concat(A.head.map((h) => A.name + ": " + h), B.head.map((h) => B.name + ": " + h));
    const groups = [[], [], [], []], names = ["Match", "Changed", "Only in " + A.name, "Only in " + B.name];
    A.rows.forEach((r, i) => {
      const st = R.A.st[i]; if (st === ST.BLANK) return;
      const m = R.A.m[i];
      const g = st === ST.SAME ? 0 : st === ST.DIFF ? 1 : 2;
      groups[g].push([names[g], i + 1, m >= 0 ? m + 1 : ""].concat(r, m >= 0 ? B.rows[m] : eb));
    });
    B.rows.forEach((r, j) => { if (R.B.st[j] === ST.ONLY) groups[3].push([names[3], "", j + 1].concat(ea, r)); });
    const rows = [].concat(...groups);
    return { all: [head].concat(rows), n: rows.length };
  }
  function doExport(v) {
    if (!v) return;
    const [mode, what, k] = v.split(":");
    if (what === "compiled") {
      if (!R) { toast("Nothing to export yet", true); return; }
      const c = compiled(); if (!c.n) { toast("No rows to export"); return; }
      if (mode === "copy") copyText(toTSV(c.all), `Copied the compiled sheet (${fmt(c.n)} rows) — paste into Excel`);
      else download(toCSV(c.all), "compared-compiled.csv");
      trackEvent("compare_export", { mode, what, rows: c.n });
      return;
    }
    const st = what === "only" ? ST.ONLY : what === "same" ? ST.SAME : ST.DIFF;
    if (!R) { toast("Nothing to export yet", true); return; }
    const rows = rowsFor(k, st);
    if (!rows.length) { toast("No rows to export"); return; }
    const all = [S[k].head].concat(rows);
    if (mode === "copy") copyText(toTSV(all), `Copied ${fmt(rows.length)} rows — paste into Excel`);
    else download(toCSV(all), `${S[k].name.replace(/[^\w\- ]+/g, "")}-${what}.csv`);
    trackEvent("compare_export", { mode, what, rows: rows.length });
  }

  // -------------------------------------------------------------------- UI
  function pasteDialog(k) {
    const dlg = document.createElement("div");
    dlg.className = "cmp-dlg";
    dlg.innerHTML = `<div class="cmp-dlg-box"><h3>Paste data into ${esc(S[k].name)}</h3>
      <p class="hint">Copy cells in Excel / Google Sheets / a CSV, then press Ctrl+V in the box.</p>
      <textarea spellcheck="false" placeholder="Paste here…"></textarea>
      <div class="cmp-dlg-row"><label><input type="radio" name="cmpPm" value="replace" checked> Replace the current data</label>
      <label><input type="radio" name="cmpPm" value="append"> Add under the current data</label></div>
      <div class="cmp-dlg-row end"><button class="btn" type="button" data-x>Cancel</button><button class="btn primary" type="button" data-ok>Load</button></div></div>`;
    root.appendChild(dlg);
    const ta = dlg.querySelector("textarea"); ta.focus();
    const close = () => dlg.remove();
    dlg.querySelector("[data-x]").onclick = close;
    dlg.querySelector("[data-ok]").onclick = () => {
      const m = parseText(ta.value);
      if (!m.length) { toast("Nothing to load", true); return; }
      load(k, m, dlg.querySelector("input[name=cmpPm]:checked").value === "append"); close();
    };
  }
  function buildPane(k) {
    const el = document.createElement("section");
    el.className = "cmp-pane"; el.dataset.k = k;
    el.innerHTML = `<div class="cmp-ph"><input class="cmp-name" value="${esc(S[k].name)}" aria-label="Name of this file" spellcheck="false">
        <span class="cmp-count"></span>
        <label class="cmp-chk" title="Untick if the first line is data, not column names"><input type="checkbox" data-act="hdr" checked> First row is header</label>
        <span class="cmp-ph-sp"></span>
        <button class="tb-btn" type="button" data-act="paste">Paste data</button>
        <button class="tb-btn" type="button" data-act="addrow">+ Row</button>
        <button class="tb-btn" type="button" data-act="addcol">+ Column</button>
        <button class="tb-btn" type="button" data-act="clear">Clear</button></div>
      <div class="cmp-body"><div class="cmp-content"><div class="cmp-hd"></div><div class="cmp-rows"></div></div>
        <div class="cmp-emptybox"><div><b>Paste your ${esc(S[k].name)} data here</b><br>Click, then press Ctrl+V — straight from Excel, Sheets or a CSV</div>
        <textarea spellcheck="false" aria-label="Paste data"></textarea></div></div>`;
    const p = P[k] = {
      el, view: [], pos: new Int32Array(0), body: el.querySelector(".cmp-body"), content: el.querySelector(".cmp-content"),
      hd: el.querySelector(".cmp-hd"), rows: el.querySelector(".cmp-rows"), count: el.querySelector(".cmp-count"),
    };
    p.body.addEventListener("scroll", () => {
      const o = P[other(k)];
      if (P.aligned && o.body.scrollTop !== p.body.scrollTop) { o.body.scrollTop = p.body.scrollTop; renderPane(other(k)); }
      renderPane(k);
    });
    el.querySelector(".cmp-name").addEventListener("input", (e) => { S[k].name = e.target.value || "File " + k; renderMapping(); });
    el.querySelector(".cmp-emptybox textarea").addEventListener("paste", (e) => {
      e.preventDefault();
      const m = parseText((e.clipboardData || window.clipboardData).getData("text"));
      if (m.length) load(k, m, false); else toast("Nothing to paste", true);
    });
    el.querySelector(".cmp-ph").addEventListener("click", (e) => {
      const b = e.target.closest("[data-act]"); if (!b) return;
      const a = b.dataset.act;
      if (a === "paste") pasteDialog(k);
      else if (a === "addrow") { if (S[k].head.length) { addRow(k); const pp = P[k]; pp.body.scrollTop = pp.body.scrollHeight; renderPane(k, true); } }
      else if (a === "addcol") { if (S[k].head.length) addCol(k); }
      else if (a === "clear") { pushUndo(); S[k].head = []; S[k].rows = []; syncMapping(true); changed(true); }
    });
    el.querySelector("[data-act=hdr]").addEventListener("change", (e) => setHeaderMode(k, e.target.checked));
    // rows: hover, click, edit
    p.rows.addEventListener("mouseover", (e) => {
      const r = e.target.closest(".cmp-row"); if (!r || r.dataset.i == null) { setHover(null); tip.hidden = true; return; }
      const d = +r.dataset.i; setHover(k, d);
      const t = tipText(k, d); if (t) { tip.innerHTML = t; tip.hidden = false; } else tip.hidden = true;
    });
    p.rows.addEventListener("mousemove", (e) => {
      if (tip.hidden) return;
      const x = Math.min(e.clientX + 14, window.innerWidth - tip.offsetWidth - 8), y = Math.min(e.clientY + 18, window.innerHeight - tip.offsetHeight - 8);
      tip.style.left = x + "px"; tip.style.top = y + "px";
    });
    p.rows.addEventListener("mouseleave", () => { setHover(null); tip.hidden = true; });
    p.rows.addEventListener("mousedown", (e) => {
      const row = e.target.closest(".cmp-row"); if (!row || row.dataset.i == null) return;
      const d = +row.dataset.i, vp = +row.dataset.vp;
      const act = e.target.closest("[data-act]");
      if (act) { e.preventDefault(); commitEdit(); if (act.dataset.act === "del") delRow(k, d); else addRow(k, d); return; }
      if (e.target.closest(".cmp-gut")) { e.preventDefault(); commitEdit(); select(k, d, true); if (sel && R && R[k].m[d] >= 0) jump(k, d); return; }
      const cell = e.target.closest(".cmp-cell");
      if (cell && !cell.classList.contains("editing")) { e.preventDefault(); select(k, d, false); startCell(k, vp, +cell.dataset.c); }
    });
    p.hd.addEventListener("mousedown", (e) => {
      const hc = e.target.closest(".cmp-hcell"); if (!hc) return;
      const c = +hc.dataset.c;
      if (e.target.closest("[data-act=delcol]")) { e.preventDefault(); commitEdit(); delCol(k, c); return; }
      const hn = e.target.closest(".hn");
      if (hn) { e.preventDefault(); startHead(k, c, hn); }
    });
    return el;
  }

  function build() {
    root = document.createElement("div");
    root.id = "cmpScreen"; root.className = "cmp-screen"; root.hidden = true;
    root.innerHTML = `
      <div class="cmp-top">
        <div class="cmp-title"><b>Compare files</b><span>Paste two tables, edit them freely — matching rows are found wherever they sit.</span></div>
        <div class="cmp-top-actions">
          <button class="tb-btn" id="cmpUndo" type="button" title="Undo (Ctrl+Z)">Undo</button>
          <button class="tb-btn" id="cmpSwap" type="button" title="Swap the two sides">⇄ Swap</button>
          <button class="btn" id="cmpClose" type="button">Close</button>
        </div>
      </div>
      <div class="cmp-opts">
        <label class="cmp-chk"><input type="checkbox" id="cmpOptCase" checked> Ignore UPPER / lower case</label>
        <label class="cmp-chk"><input type="checkbox" id="cmpOptTrim" checked> Ignore extra spaces</label>
        <label class="cmp-chk" title="So 10, 10.0 and 1,000 / 1000 are equal"><input type="checkbox" id="cmpOptNum" checked> Numbers as numbers (1.0 = 1)</label>
        <span class="cmp-sep"></span>
        <label class="cmp-chk" title="On: matching rows sit on the same line in both files (like a diff). Off: each file keeps its own order."><input type="checkbox" id="cmpAlign" checked> Line up matching rows</label>
        <button class="tb-btn" id="cmpMapBtn" type="button">Columns &amp; keys ▾</button>
        <select id="cmpExport" class="gt-select" title="Copy or download results">
          <option value="">Export…</option>
          <option value="csv:compiled:">Download CSV — compiled sheet (matching → only in A → only in B)</option>
          <option value="copy:compiled:">Copy compiled sheet</option>
          <option value="copy:only:A">Copy rows only in A</option><option value="copy:only:B">Copy rows only in B</option>
          <option value="csv:only:A">Download CSV — only in A</option><option value="csv:only:B">Download CSV — only in B</option>
          <option value="copy:same:A">Copy matching rows (from A)</option><option value="csv:same:A">Download CSV — matching rows</option>
        </select>
      </div>
      <div class="cmp-map-panel" id="cmpMapPanel" hidden>
        <p class="hint">Choose which column of one file corresponds to which column of the other. Untick a pair to ignore it.
          Tick <b>Key</b> on one or more pairs (e.g. an order id) to match rows by that key and flag rows whose <i>other</i> columns differ. With no key, whole rows are matched.</p>
        <div id="cmpMapBody"></div>
      </div>
      <div class="cmp-sum">
        <button class="cmp-chip" data-f="all" type="button">All rows</button>
        <button class="cmp-chip same" data-f="same" type="button">Matching <b id="cmpNSame">–</b></button>
        <button class="cmp-chip diff" data-f="diff" id="cmpChipDiff" type="button" hidden>Changed <b id="cmpNDiff">–</b></button>
        <button class="cmp-chip only" data-f="only" type="button" title="Show only the rows that have no match, in each file">Not found — only in A <b id="cmpNA">–</b> · only in B <b id="cmpNB">–</b></button>
        <span class="cmp-mode" id="cmpMode"></span>
      </div>
      <div class="cmp-panes" id="cmpPanes"></div>`;
    document.body.appendChild(root);
    tip = document.createElement("div"); tip.className = "cmp-tip"; tip.hidden = true; document.body.appendChild(tip);
    const panes = root.querySelector("#cmpPanes");
    panes.appendChild(buildPane("A")); panes.appendChild(buildPane("B"));

    root.querySelector("#cmpClose").onclick = close;
    root.querySelector("#cmpUndo").onclick = undo;
    root.querySelector("#cmpSwap").onclick = () => {
      pushUndo();
      const t = { ...S.A }; Object.assign(S.A, S.B); Object.assign(S.B, t);
      ["A", "B"].forEach((k) => { P[k].el.querySelector(".cmp-name").value = S[k].name; P[k].el.querySelector("[data-act=hdr]").checked = S[k].hasHeader; });
      mapping = mapping.map((m) => ({ ...m, a: m.b, b: m.a })).filter((m) => m.a >= 0).sort((x, y) => x.a - y.a);
      syncMapping(); changed(true);
    };
    const optBind = (id, key) => root.querySelector(id).addEventListener("change", (e) => { opt[key] = e.target.checked; changed(true); });
    optBind("#cmpOptCase", "nocase"); optBind("#cmpOptTrim", "trim"); optBind("#cmpOptNum", "num");
    root.querySelector("#cmpAlign").addEventListener("change", (e) => { view.aligned = e.target.checked; renderAll(true); });
    root.querySelector("#cmpMapBtn").onclick = () => { const m = root.querySelector("#cmpMapPanel"); m.hidden = !m.hidden; renderMapping(); };
    root.querySelector("#cmpExport").addEventListener("change", (e) => { doExport(e.target.value); e.target.value = ""; });
    root.querySelector(".cmp-sum").addEventListener("click", (e) => {
      const b = e.target.closest(".cmp-chip"); if (!b) return;
      view.filter = view.filter === b.dataset.f ? "all" : b.dataset.f;
      P.A.body.scrollTop = 0; P.B.body.scrollTop = 0; renderAll(true);
    });
    root.querySelector("#cmpMapBody").addEventListener("change", (e) => {
      const tr = e.target.closest("tr[data-m]"); if (!tr) return;
      const m = mapping[+tr.dataset.m], f = e.target.dataset.f;
      userMapped = true;
      if (f === "use") m.use = e.target.checked;
      else if (f === "key") m.key = e.target.checked;
      else if (f === "b") { m.b = +e.target.value; m.use = m.b >= 0; if (m.b < 0) m.key = false; mapping.forEach((o) => { if (o !== m && m.b >= 0 && o.b === m.b) { o.b = -1; o.use = false; o.key = false; } }); }
      changed(true);
    });
    root.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && !editing && !root.querySelector(".cmp-dlg")) { if (sel) { sel = null; applySel(); } else close(); }
      else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "z" && !(e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement)) { e.preventDefault(); undo(); }
    });
    window.addEventListener("resize", () => { if (!root.hidden) renderAll(true); });
  }

  function open() {
    if (!root) build();
    root.hidden = false; document.body.classList.add("cmp-open");
    renderAll(true);
    trackEvent("compare_opened");
  }
  function close() { commitEdit(); root.hidden = true; document.body.classList.remove("cmp-open"); tip.hidden = true; }

  window.openCompare = open;
  const btn = document.getElementById("compareBtn");
  if (btn) btn.addEventListener("click", open);
})();
