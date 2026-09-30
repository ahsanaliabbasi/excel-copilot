/* builder.js — decision-tree ("flow chart") formula builder.
 *
 * The whole formula is a tree of small steps. Every step is a box with a
 * "what is this?" dropdown; a Decision step has a YES branch and a NO branch and
 * each branch is itself a box, so decisions nest as deep as you like.
 * The tree is sent as JSON to /api/operations/{preview,apply}; the backend
 * (expr_engine.py) turns it into a real Excel formula per row.
 */
(function () {
  "use strict";
  (window.FRONTEND_PARTS = window.FRONTEND_PARTS || {}).builder = 19;

  // ------------------------------------------------------------------ data
  const KIND_GROUPS = [
    ["Basics", [["col", "A column's value"], ["lit", "A typed value"], ["blank", "Empty (blank)"]]],
    ["Decisions", [
      ["if", "Decision: IF … THEN … ELSE"],
      ["iferr", "If it fails / not found, use another value (IFERROR / IFNA)"],
    ]],
    ["Text", [
      ["join", "Join values together (CONCAT)"],
      ["textfn", "Clean / change text (TRIM, UPPER, LOWER, LEN …)"],
      ["replace", "Replace text — several at once (SUBSTITUTE)"],
      ["extract", "Find & extract part of a text (LEFT, RIGHT, MID, FIND …)"],
      ["find", "Position of a text (FIND / SEARCH)"],
      ["textfmt", "Format as text (TEXT)"],
    ]],
    ["Numbers", [
      ["calc", "Calculate (+ − × ÷)"],
      ["round", "Round (ROUND)"],
      ["rowagg", "SUM / MAX / MIN / AVERAGE of values in this row"],
    ]],
    ["Dates", [
      ["dateadd", "Add / subtract days, weeks, months, quarters or years"],
      ["datediff", "Difference between two dates (DATEDIF)"],
      ["today", "Today's date (TODAY)"],
      ["textfmt", "Change date format (TEXT)"],
    ]],
    ["Look up & totals", [
      ["lookup", "Look up from a sheet (XLOOKUP / VLOOKUP / INDEX-MATCH)"],
      ["agg", "Total of matching rows (SUMIF / COUNTIF / AVERAGEIF / MAXIF)"],
    ]],
  ];

  const OPS = [
    ["blank", "is empty (or only spaces)"], ["notblank", "is not empty"],
    ["=", "equals"], ["<>", "does not equal"],
    [">", "is greater than"], ["<", "is less than"], [">=", "is greater or equal"], ["<=", "is less or equal"],
    ["contains", "contains"], ["notcontains", "does not contain"],
    ["starts", "starts with"], ["ends", "ends with"],
  ];
  const NO_RIGHT = new Set(["blank", "notblank"]);

  const EXTRACT_MODES = [
    ["before_text", "Everything BEFORE the first … (a character or text)"],
    ["after_text", "Everything AFTER the first … (a character or text)"],
    ["between_text", "Everything BETWEEN two characters / texts"],
    ["before_digit", "Everything before the first DIGIT (the letters part)"],
    ["from_digit", "Everything from the first DIGIT onwards"],
    ["before_letter", "Everything before the first LETTER (the digits part)"],
    ["from_letter", "Everything from the first LETTER onwards"],
    ["first_n", "The first N characters (LEFT)"],
    ["last_n", "The last N characters (RIGHT)"],
    ["mid", "N characters starting at a position (MID)"],
  ];
  const EXTRACT_HINTS = {
    before_text: "'John Smith' with a space → 'John'. If the text isn't found, the whole value is kept.",
    after_text: "'John Smith' with a space → 'Smith'. If the text isn't found, the result is empty.",
    between_text: "'a[bc]d' between [ and ] → 'bc'. Empty if either end isn't found.",
    before_digit: "'ABC123' → 'ABC'.",
    from_digit: "'ABC123' → '123'.",
    before_letter: "'123ABC' → '123'.",
    from_letter: "'123ABC' → 'ABC'.",
    first_n: "'Hello' with 2 → 'He'.",
    last_n: "'Hello' with 2 → 'lo'.",
    mid: "'Hello' from position 2, 3 characters → 'ell'.",
  };
  const DELIMS = [
    ["", "Nothing"], [" ", "Space"], [", ", "Comma + space"], ["-", "Dash"],
    ["/", "Slash"], ["_", "Underscore"], ["|", "Pipe"], ["custom", "Custom…"],
  ];
  const FORMATS = [
    ["yyyy-mm-dd", "Date: 2026-03-20"], ["dd/mm/yyyy", "Date: 20/03/2026"], ["mm/dd/yyyy", "Date: 03/20/2026"],
    ["dd mmm yyyy", "Date: 20 Mar 2026"], ["mmm yyyy", "Date: Mar 2026"], ["mmmm d, yyyy", "Date: March 20, 2026"],
    ["0", "Number: whole"], ["0.00", "Number: 2 decimals"], ["#,##0", "Number: 1,234"],
    ["#,##0.00", "Number: 1,234.50"], ["0%", "Percent"], ["000000", "Zero-padded (6 digits)"],
    ["custom", "Custom…"],
  ];
  const CHIPS = [[" ", "space"], ["-", "-"], ["@", "@"], [",", ","], ["/", "/"], [".", "."], ["_", "_"], [":", ":"]];

  // Which step each sidebar button starts with
  const PRESETS = {
    BUILDER: "if", CONCAT: "join", LOOKUP: "lookup", IF: "if", CONDITIONAL_AGG: "agg",
    TEXT_CLEAN: "textfn", REPLACE: "replace", EXTRACT: "extract", CALC: "calc", ROUND: "round", DATE: "dateadd",
  };
  const TITLES = {
    BUILDER: "Formula Builder", CONCAT: "Combine Columns", LOOKUP: "Lookup / Match Values",
    IF: "IF / Decision Flow", CONDITIONAL_AGG: "Totals & Counts (SUMIF / COUNTIF …)", TEXT_CLEAN: "Clean Text",
    REPLACE: "Replace Text", EXTRACT: "Find & Extract Text", CALC: "Calculate", ROUND: "Round Number", DATE: "Dates",
  };
  const DEFAULT_OUT = {
    CONCAT: "combined", LOOKUP: "looked_up_value", IF: "if_result", BUILDER: "result", CONDITIONAL_AGG: "agg_result",
    TEXT_CLEAN: "cleaned", REPLACE: "replaced", EXTRACT: "extracted", CALC: "calculated", ROUND: "rounded", DATE: "date_result",
  };

  let B = null; // { root, op }

  // --------------------------------------------------------------- helpers
  function h(tag, cls, kids) {
    const n = document.createElement(tag);
    if (cls) n.className = cls;
    (kids || []).forEach((k) => k != null && n.append(k));
    return n;
  }
  const txt = (s) => document.createTextNode(s);
  function withText(tag, cls, s) { const n = h(tag, cls); n.textContent = s; return n; }
  const hint = (s) => withText("div", "hint", s);

  function fld(label, control, hintText) {
    const f = h("div", "field");
    f.append(withText("label", null, label), control);
    if (hintText) f.append(hint(hintText));
    return f;
  }
  function selectEl(options, value, onChange) {
    const s = h("select");
    options.forEach(([v, l]) => {
      const o = h("option");
      o.value = v; o.textContent = l; o.selected = v === value;
      s.append(o);
    });
    s.addEventListener("change", () => onChange(s.value));
    return s;
  }
  function inputEl(value, onInput, o = {}) {
    const i = h("input");
    i.type = o.type || "text";
    i.value = value ?? "";
    if (o.placeholder) i.placeholder = o.placeholder;
    if (o.min != null) i.min = o.min;
    i.addEventListener("input", () => onInput(i.value));
    return i;
  }
  function numInput(obj, key, min) {
    return inputEl(obj[key], (v) => { obj[key] = parseInt(v, 10) || 0; changed(); }, { type: "number", min });
  }
  function iconBtn(label, title, onClick) {
    const b = withText("button", "icon-btn", label);
    b.type = "button"; b.title = title;
    b.addEventListener("click", onClick);
    return b;
  }
  function linkBtn(label, onClick) {
    const b = withText("button", "link-btn", label);
    b.type = "button";
    b.addEventListener("click", onClick);
    return b;
  }
  function check(label, checked, onChange) {
    const l = h("label", "check");
    const i = h("input"); i.type = "checkbox"; i.checked = !!checked;
    i.addEventListener("change", () => onChange(i.checked));
    l.append(i, txt(" " + label));
    return l;
  }
  const curCols = () => currentSheetMeta().columns;
  const sheetCols = (name) => (state.sheets.find((s) => s.name === name) || currentSheetMeta()).columns;
  const sheetOpts = () => state.sheets.map((s) => [s.name, s.name]);
  function colSelect(cols, value, onChange) {
    return selectEl(cols.map((c) => [c, c]), value, onChange);
  }
  function textWithChips(obj, key, placeholder) {
    const wrap = h("div", "chips-wrap");
    const input = inputEl(obj[key], (v) => { obj[key] = v; changed(); }, { placeholder });
    const chips = h("div", "chips");
    CHIPS.forEach(([val, lab]) => {
      const c = withText("button", "chip-btn", lab);
      c.type = "button";
      c.addEventListener("click", () => { input.value = val; obj[key] = val; changed(); });
      chips.append(c);
    });
    wrap.append(input, chips);
    return wrap;
  }

  // ----------------------------------------------------------- node makers
  const colNode = (i) => { const c = curCols(); return { type: "col", name: c[Math.min(i, c.length - 1)] }; };
  const litNode = (v) => ({ type: "lit", value: v });
  const makeCmp = () => ({ left: colNode(0), op: "blank", right: litNode("") });
  const makeCond = () => ({ join: "AND", items: [makeCmp()] });
  // A "where" group: { join: "AND" | "OR", items: [ leaf | group ] },  leaf = { col, op, value }.
  // The leaf's column belongs to the sheet being searched; its value is a step evaluated on the current row.
  const makeLeaf = (cols) => ({ col: cols[0], pipe: [], op: "=", value: colNode(0) }); // default: equals this row's first column
  const makeCrit = () => makeLeaf(curCols());
  const makeGroup = (cols, join = "AND") => ({ join, items: [makeLeaf(cols)] });
  function resetLeafCols(w, col) {
    if (!w) return;
    if (w.items) w.items.forEach((i) => resetLeafCols(i, col)); else w.col = col;
  }

  function makeNode(kind) {
    const other = state.sheets.find((s) => s.name !== state.currentSheet) || currentSheetMeta();
    const oc = other.columns;
    switch (kind) {
      case "col": return colNode(0);
      case "lit": return litNode("");
      case "blank": return { type: "blank" };
      case "if": return { type: "if", cond: makeCond(), then: litNode("Yes"), else: litNode("No") };
      case "iferr": return { type: "iferr", fn: "IFERROR", value: colNode(0), fallback: litNode("") };
      case "join": return { type: "join", parts: [colNode(0), colNode(1)], delimiter: " ", delim_mode: " ", skip_blank: true };
      case "textfn": return { type: "textfn", fn: "TRIM", source: colNode(0) };
      case "replace": return { type: "replace", source: colNode(0), pairs: [{ find: " ", replace: "-" }] };
      case "extract": return { type: "extract", source: colNode(0), mode: "before_text", text: " ", text2: "", n: 3, start: 1 };
      case "find": return { type: "find", what: litNode(""), within: colNode(0), case_sensitive: false };
      case "textfmt": return { type: "textfmt", source: colNode(0), fmt: "yyyy-mm-dd", fmt_mode: "yyyy-mm-dd" };
      case "calc": return { type: "calc", left: colNode(0), op: "*", right: litNode("1") };
      case "round": return { type: "round", fn: "ROUND", source: colNode(0), digits: 2 };
      case "rowagg": return { type: "rowagg", fn: "SUM", items: [colNode(0), colNode(1)] };
      case "edate": return { type: "edate", source: colNode(0), months: litNode("1") };
      case "today": return { type: "today" };
      case "dateadd": return { type: "dateadd", source: colNode(0), unit: "MONTHS", amount: litNode("1") };
      case "datediff": return { type: "datediff", start: colNode(0), end: { type: "today" }, unit: "YEARS" };
      case "lookup":
        return { type: "lookup", method: "XLOOKUP", match: "exact", search: "first", nth: litNode("2"), key: colNode(0), key_src: { col: curCols()[0], pipe: [] }, match_pipe: [], sheet: other.name,
                 match_col: oc[0], cases: [{ where: null, return_col: oc[1] || oc[0], return_type: "col" }], not_found: litNode("Not Found") };
      case "agg":
        return { type: "agg", fn: "SUM", sheet: state.currentSheet, value_col: curCols()[0], where: makeGroup(curCols()) };
    }
    throw new Error("Unknown step " + kind);
  }
  function replaceNode(node, fresh) {
    Object.keys(node).forEach((k) => delete node[k]);
    Object.assign(node, fresh);
  }

  // -------------------------------------------------------------- rendering
  const INLINE = new Set(["col", "lit", "blank", "today"]);
  const NUMERIC = new Set(["calc", "round", "rowagg", "find", "agg", "datediff"]);
  const TESTABLE = new Set(["textfn", "replace", "extract", "find", "textfmt", "calc", "round", "rowagg", "lookup", "agg", "join", "edate", "dateadd", "datediff"]);
  const UNIT_OPTS = [["DAYS", "Days"], ["WEEKS", "Weeks"], ["MONTHS", "Months"], ["QUARTERS", "Quarters"], ["YEARS", "Years"]];
  const AMOUNT_LABEL = {
    DAYS: "Days to add (negative = subtract)", WEEKS: "Weeks to add (negative = subtract)",
    MONTHS: "Months to add (negative = subtract)", QUARTERS: "Quarters to add (negative = subtract)",
    YEARS: "Years to add (negative = subtract)",
  };

  // Wraps a finished value in an IF so its result can be tested (greater / less / equals / empty ...)
  const isNumeric = (node) => NUMERIC.has(node.type) || (node.type === "textfn" && ["LEN", "VALUE"].includes(node.fn));
  function testLink(node) {
    const numeric = isNumeric(node);
    const b = withText("button", "link-btn test-link",
      numeric ? "\u21b3 Compare this number (greater than, less than, equals …) and give Yes / No answers"
              : "\u21b3 Test this result (equals, contains, is empty …) and give Yes / No answers");
    b.type = "button";
    b.addEventListener("click", () => {
      const inner = JSON.parse(JSON.stringify(node));
      const numeric = isNumeric(inner);
      replaceNode(node, {
        type: "if",
        cond: { join: "AND", items: [{ left: inner, op: numeric ? ">" : "=", right: litNode(numeric ? "0" : "") }] },
        then: litNode("Yes"), else: litNode("No"),
      });
      redraw();
    });
    return b;
  }

  function renderNode(node, title, opts = {}) {
    const box = h("div", "node kind-" + node.type + (INLINE.has(node.type) ? " inline" : ""));
    const head = h("div", "node-head");
    if (title) head.append(withText("span", "node-title", title));
    head.append(kindSelect(node));
    box.append(head);
    const body = h("div", "node-body");
    RENDER[node.type](node, body);
    if (opts.test && TESTABLE.has(node.type)) body.append(testLink(node));
    if (body.childNodes.length) box.append(body);
    return box;
  }
  function kindSelect(node) {
    const s = h("select", "kind-select");
    KIND_GROUPS.forEach(([g, items]) => {
      const og = h("optgroup");
      og.label = g;
      items.forEach(([v, l]) => {
        const o = h("option");
        o.value = v; o.textContent = l; o.selected = v === node.type;
        og.append(o);
      });
      s.append(og);
    });
    s.addEventListener("change", () => { replaceNode(node, makeNode(s.value)); redraw(); });
    return s;
  }
  function branch(cls, tag, title, child) {
    const w = h("div", "branch " + cls);
    w.append(withText("div", "branch-tag", tag), renderNode(child, title, { test: true }));
    return w;
  }
  function renderCond(cond) {
    const c = h("div", "cond");
    const head = h("div", "cond-head", [withText("span", "cond-word", "IF")]);
    if (cond.items.length > 1) {
      head.append(selectEl([["AND", "ALL of these are true"], ["OR", "ANY of these is true"]], cond.join,
        (v) => { cond.join = v; changed(); }));
    }
    c.append(head);
    cond.items.forEach((it, idx) => {
      const row = h("div", "cmp");
      row.append(renderNode(it.left, null));
      row.append(selectEl(OPS, it.op, (v) => {
        it.op = v;
        if (!it.right) it.right = litNode("");
        redraw();
      }));
      if (!NO_RIGHT.has(it.op)) row.append(renderNode(it.right, null));
      if (cond.items.length > 1) row.append(iconBtn("✕", "Remove this condition", () => { cond.items.splice(idx, 1); redraw(); }));
      c.append(row);
    });
    c.append(linkBtn("+ Add another condition", () => { cond.items.push(makeCmp()); redraw(); }));
    return c;
  }
  function listOf(items, label, makeItem, b) {
    items.forEach((p, i) => {
      const w = h("div", "list-item");
      w.append(renderNode(p, `${label} ${i + 1}`));
      if (items.length > 1) w.append(iconBtn("✕", "Remove", () => { items.splice(i, 1); redraw(); }));
      b.append(w);
    });
    b.append(linkBtn("+ Add another value", () => { items.push(makeItem()); redraw(); }));
  }

  // ======================================================================================
  //  Guided building blocks
  // ======================================================================================
  // A "pipe" is a column followed by steps:  first_name -> Trim -> Replace "-" with "" -> Last 5 characters -> Length
  const STEP_MENU = [
    ["", "+ then\u2026"],
    ["TRIM", "Remove extra spaces  (TRIM)"], ["UPPER", "Make UPPERCASE"], ["LOWER", "Make lowercase"], ["PROPER", "Capitalize Each Word"],
    ["LEN", "Count the characters  (LEN)"],
    ["REPLACE", "Replace some text\u2026  (SUBSTITUTE)"], ["REMOVE", "Remove some text\u2026  (SUBSTITUTE with nothing)"],
    ["LEFT", "Keep the first N characters  (LEFT)"], ["RIGHT", "Keep the last N characters  (RIGHT)"],
    ["MID", "Keep N characters from a position  (MID)"],
    ["BEFORE", "Keep the text before a character"], ["AFTER", "Keep the text after a character"],
    ["ROUND", "Round to N decimals  (ROUND)"], ["ADD", "Add a number  (+)"], ["MUL", "Multiply by a number  (\u00d7)"],
  ];
  const STEP_DEFAULTS = {
    TRIM: {}, UPPER: {}, LOWER: {}, PROPER: {}, LEN: {}, REPLACE: { find: "", to: "" }, REMOVE: { find: "" },
    LEFT: { n: 5 }, RIGHT: { n: 5 }, MID: { start: 1, n: 5 }, BEFORE: { text: " " }, AFTER: { text: " " },
    ROUND: { digits: 0 }, ADD: { value: 0 }, MUL: { value: 1 },
  };
  const STEP_WORD = {
    TRIM: "Trim", UPPER: "UPPERCASE", LOWER: "lowercase", PROPER: "Capitalize", LEN: "Length", REPLACE: "Replace", REMOVE: "Remove",
    LEFT: "First", RIGHT: "Last", MID: "Middle", BEFORE: "Before", AFTER: "After", ROUND: "Round", ADD: "Add", MUL: "Multiply by",
  };

  function pipeToNode(base, pipe) {
    return (pipe || []).reduce((src, st) => {
      switch (st.fn) {
        case "TRIM": case "UPPER": case "LOWER": case "PROPER": case "LEN": return { type: "textfn", fn: st.fn, source: src };
        case "REPLACE": return { type: "replace", source: src, pairs: [{ find: st.find || "", replace: st.to || "" }] };
        case "REMOVE": return { type: "replace", source: src, pairs: [{ find: st.find || "", replace: "" }] };
        case "LEFT": return { type: "extract", mode: "first_n", n: +st.n || 0, source: src };
        case "RIGHT": return { type: "extract", mode: "last_n", n: +st.n || 0, source: src };
        case "MID": return { type: "extract", mode: "mid", start: +st.start || 1, n: +st.n || 0, source: src };
        case "BEFORE": return { type: "extract", mode: "before_text", text: st.text || "", source: src };
        case "AFTER": return { type: "extract", mode: "after_text", text: st.text || "", source: src };
        case "ROUND": return { type: "round", fn: "ROUND", digits: +st.digits || 0, source: src };
        case "ADD": return { type: "calc", op: "+", left: src, right: { type: "lit", value: String(st.value ?? 0) } };
        case "MUL": return { type: "calc", op: "*", left: src, right: { type: "lit", value: String(st.value ?? 1) } };
      }
      return src;
    }, base);
  }
  function describePipe(col, pipe) {
    return (pipe || []).reduce((x, st) => {
      switch (st.fn) {
        case "TRIM": case "UPPER": case "LOWER": case "PROPER": case "LEN": return `${st.fn}(${x})`;
        case "REPLACE": return `SUBSTITUTE(${x}, "${st.find || ""}", "${st.to || ""}")`;
        case "REMOVE": return `SUBSTITUTE(${x}, "${st.find || ""}", "")`;
        case "LEFT": return `LEFT(${x}, ${st.n})`;
        case "RIGHT": return `RIGHT(${x}, ${st.n})`;
        case "MID": return `MID(${x}, ${st.start}, ${st.n})`;
        case "BEFORE": return `the part of ${x} before "${st.text}"`;
        case "AFTER": return `the part of ${x} after "${st.text}"`;
        case "ROUND": return `ROUND(${x}, ${st.digits})`;
        case "ADD": return `${x} + ${st.value}`;
        case "MUL": return `${x} \u00d7 ${st.value}`;
      }
      return x;
    }, col);
  }

  function smallInput(st, key, width, onChange, placeholder, type) {
    const i = inputEl(st[key], (v) => { st[key] = type === "number" ? (v === "" ? "" : Number(v)) : v; onChange(false); },
      { type: type || "text", placeholder });
    i.style.width = width;
    return i;
  }

  // holder = { col, pipe:[...] }   onChange(structural)  is called after every edit
  function renderPipe(cols, holder, onChange) {
    if (!holder.pipe) holder.pipe = [];
    if (!cols.includes(holder.col)) holder.col = cols[0];
    const box = h("div", "pipe");
    box.append(colSelect(cols, holder.col, (v) => { holder.col = v; onChange(false); }));
    holder.pipe.forEach((st, i) => {
      box.append(withText("span", "pipe-arrow", "\u2192"));
      const chip = h("span", "pstep");
      chip.append(txt(STEP_WORD[st.fn] + " "));
      if (st.fn === "REPLACE") chip.append(smallInput(st, "find", "70px", onChange, "text"), txt(" with "), smallInput(st, "to", "70px", onChange, "nothing"));
      else if (st.fn === "REMOVE") chip.append(smallInput(st, "find", "80px", onChange, "text to remove"));
      else if (st.fn === "LEFT" || st.fn === "RIGHT") chip.append(smallInput(st, "n", "52px", onChange, "", "number"), txt(" characters"));
      else if (st.fn === "MID") chip.append(smallInput(st, "n", "52px", onChange, "", "number"), txt(" characters from position "), smallInput(st, "start", "52px", onChange, "", "number"));
      else if (st.fn === "BEFORE" || st.fn === "AFTER") chip.append(smallInput(st, "text", "70px", onChange, "e.g. a space"));
      else if (st.fn === "ROUND") chip.append(smallInput(st, "digits", "52px", onChange, "", "number"), txt(" decimals"));
      else if (st.fn === "ADD" || st.fn === "MUL") chip.append(smallInput(st, "value", "64px", onChange, "", "number"));
      const x = withText("button", "pstep-x", "\u2715");
      x.type = "button"; x.title = "Remove this step";
      x.addEventListener("click", () => { holder.pipe.splice(i, 1); onChange(true); });
      chip.append(x);
      box.append(chip);
    });
    box.append(withText("span", "pipe-arrow", "\u2192"));
    const add = selectEl(STEP_MENU, "", (v) => {
      if (!v) return;
      holder.pipe.push({ fn: v, ...STEP_DEFAULTS[v] });
      onChange(true);
    });
    add.className = "pipe-add";
    box.append(add);
    return box;
  }

  const ALLANY = [["AND", "ALL of these are true"], ["OR", "ANY of these is true"]];
  let radioSeq = 0;
  function radioRow(options, value, onChange) {
    const name = "rg" + (++radioSeq), row = h("div", "radio-row");
    options.forEach(([v, label]) => {
      const l = h("label", "radio"), i = h("input");
      i.type = "radio"; i.name = name; i.checked = v === value;
      i.addEventListener("change", () => { if (i.checked) onChange(v); });
      l.append(i, txt(" " + label));
      row.append(l);
    });
    return row;
  }
  function guideStep(num, title, help, build) {
    const step = h("div", "gstep"), body = h("div", "gbody");
    body.append(withText("div", "gtitle", title));
    if (help) body.append(withText("div", "ghelp", help));
    build(body);
    step.append(withText("div", "gnum", String(num)), body);
    return step;
  }

  // ---- a step-by-step wizard: one question on screen at a time, with Back / Next ----
  // steps = [{ key, title, help, build(container) }]. The current step is remembered on B so it survives redraws.
  function wizard(steps) {
    let idx = steps.findIndex((s2) => s2.key === B.wizKey);
    if (idx < 0) idx = 0;
    const go = (k) => { B.wizKey = steps[k].key; redraw(); };
    const cur = steps[idx];
    B.wizard = true; B.wizLast = idx === steps.length - 1;
    const box = h("div", "wiz");
    const dots = h("div", "wiz-dots");
    steps.forEach((st, k) => {
      const d = h("button", "wiz-dot" + (k === idx ? " on" : k < idx ? " done" : ""));
      d.type = "button"; d.title = st.title; d.textContent = String(k + 1);
      d.addEventListener("click", () => go(k));
      dots.append(d);
    });
    box.append(h("div", "wiz-progress", [withText("span", "wiz-count", `Step ${idx + 1} of ${steps.length}`), dots]));
    const card = h("div", "wiz-card");
    card.append(withText("div", "wiz-title", cur.title));
    if (cur.help) card.append(withText("div", "wiz-help", cur.help));
    const body = h("div", "wiz-body");
    cur.build(body);
    card.append(body);
    box.append(card);
    const back = withText("button", "btn secondary wiz-back", "\u2190 Back");
    back.type = "button"; back.disabled = idx === 0;
    back.addEventListener("click", () => go(idx - 1));
    const nav = h("div", "wiz-nav", [back]);
    if (idx < steps.length - 1) {
      const next = withText("button", "btn primary wiz-next", `Next: ${steps[idx + 1].short || steps[idx + 1].title} \u2192`);
      next.type = "button";
      next.addEventListener("click", () => go(idx + 1));
      nav.append(next);
    } else {
      nav.append(withText("span", "wiz-done", "All set \u2014 choose where the result goes below, then press Apply."));
    }
    box.append(nav);
    return box;
  }

  // ---- conditions on the rows of ANOTHER sheet (leaf = column -> steps -> compare) ----
  const makeFilterLeaf = (cols) => ({ col: cols[0], pipe: [], op: "=", value: litNode("") });
  function syncLeaf(l) {
    const p = l.pipe || [];
    if (p.length) l.left = pipeToNode({ type: "col", name: l.col }, p); else delete l.left;
  }
  function renderFilter(cols, group, title, opts = {}) {
    if (!group.items) {                                     // a lone condition -> make it a group
      const leaf = { ...group };
      Object.keys(group).forEach((k) => delete group[k]);
      Object.assign(group, { join: "AND", items: [leaf] });
    }
    for (let k = group.items.length - 1; k >= 0; k--) {     // a group whose last condition was removed disappears
      if (group.items[k].items && !group.items[k].items.length) group.items.splice(k, 1);
    }
    const box = h("div", "filter" + (opts.nested ? " nested" : ""));
    const head = h("div", "filter-head", [withText("span", "filter-title", title)]);
    if (group.items.length > 1) head.append(selectEl(ALLANY, group.join || "AND", (v) => { group.join = v; redraw(); }));
    if (opts.onRemove) head.append(iconBtn("\u2715", "Remove this group", opts.onRemove));
    box.append(head);
    group.items.forEach((it, i) => {
      if (it.items) {
        box.append(renderFilter(cols, it, (group.join || "AND") === "OR" ? "\u2026 any of" : "\u2026 all of",
          { nested: true, onRemove: () => { group.items.splice(i, 1); redraw(); } }));
        return;
      }
      if (!it.pipe) it.pipe = [];
      const row = h("div", "cond");
      row.append(renderPipe(cols, it, (structural) => { syncLeaf(it); if (structural) redraw(); else changed(); }));
      const line = h("div", "cmp");
      line.append(selectEl(OPS, it.op, (v) => { it.op = v; if (!it.value) it.value = litNode(""); redraw(); }));
      if (!NO_RIGHT.has(it.op)) line.append(renderNode(it.value, null));
      line.append(iconBtn("\u2715", "Remove this condition", () => { group.items.splice(i, 1); redraw(); }));
      row.append(line);
      box.append(row);
    });
    if (!group.items.length && opts.emptyHint) box.append(hint(opts.emptyHint));
    box.append(h("div", "filter-add", [
      linkBtn("+ Add a condition", () => { group.items.push(makeFilterLeaf(cols)); redraw(); }),
      linkBtn("+ Add an ALL / ANY group", () => { group.items.push({ join: (group.join || "AND") === "OR" ? "AND" : "OR", items: [makeFilterLeaf(cols)] }); redraw(); }),
    ]));
    return box;
  }

  // ---- conditions on the rows of THIS sheet: which rows should get a value at all ----
  const syncItem = (it) => { it.left = pipeToNode({ type: "col", name: it.col }, it.pipe || []); };
  function newThisItem() {
    const it = { col: curCols()[0], pipe: [], op: "=", right: litNode("") };
    syncItem(it);
    return it;
  }
  function renderThisRows(n, cur) {
    const on = !!(n.only_if && n.only_if.items);
    const wrap = h("div", "gwrap");
    wrap.append(radioRow([[false, `Every row of \u201c${cur}\u201d`], [true, "Only rows that meet conditions\u2026"]], on, (v) => {
      if (v) n.only_if = { join: "AND", items: [newThisItem()] };
      else { delete n.only_if; delete n.only_else; }
      redraw();
    }));
    if (!on) return wrap;
    const oi = n.only_if;
    if (oi.items.length > 1) wrap.append(selectEl(ALLANY, oi.join || "AND", (v) => { oi.join = v; redraw(); }));
    oi.items.forEach((it, i) => {
      if (!it.pipe) it.pipe = [];
      const row = h("div", "cond");
      row.append(renderPipe(curCols(), it, (structural) => { syncItem(it); if (structural) redraw(); else changed(); }));
      const line = h("div", "cmp");
      line.append(selectEl(OPS, it.op, (v) => { it.op = v; if (!it.right) it.right = litNode(""); redraw(); }));
      if (!NO_RIGHT.has(it.op)) line.append(renderNode(it.right || (it.right = litNode("")), null));
      line.append(iconBtn("\u2715", "Remove this condition", () => {
        oi.items.splice(i, 1);
        if (!oi.items.length) { delete n.only_if; delete n.only_else; }
        redraw();
      }));
      row.append(line);
      wrap.append(row);
    });
    wrap.append(linkBtn("+ Add another condition", () => { oi.items.push(newThisItem()); redraw(); }));
    wrap.append(check("Show a value in the other rows instead of leaving them empty", !!n.only_else, (yes) => {
      if (yes) n.only_else = litNode(""); else delete n.only_else;
      redraw();
    }));
    if (n.only_else) wrap.append(renderNode(n.only_else, "Value for the other rows"));
    return wrap;
  }

  // ---- a lookup is a link + a list of cases: IF the row meets <conditions> THEN give <column | value | count | total> ----
  function normalizeLookup(n) {
    if (!n.cases) {                                           // older shape: one where + one return column
      n.cases = [{ where: n.where || null, return_col: n.return_col, return_type: "col" }];
      delete n.where; delete n.return_col;
    }
    n.cases.forEach((c) => { if (!c.return_type) c.return_type = "col"; });
    if (!n.match_pipe) n.match_pipe = [];
    if (n.key_src === undefined && n.key && n.key.type === "col" && !n.match_left) n.key_src = { col: n.key.name, pipe: [] };
  }
  function syncLink(n) {
    if (!n.key_src) return;
    n.key = pipeToNode({ type: "col", name: n.key_src.col }, n.key_src.pipe);
    if (n.match_pipe.length) n.match_left = pipeToNode({ type: "col", name: n.match_col }, n.match_pipe); else delete n.match_left;
  }
  const leafCount = (w) => (!w ? 0 : w.items ? w.items.reduce((a, i) => a + leafCount(i), 0) : 1);
  const newCase = (cols) => ({ where: { join: "AND", items: [] }, return_col: cols[1] || cols[0], return_type: "col" });

  // ---- plain-English description of what the step will do ----
  const opWord = (op) => (OPS.find((o) => o[0] === op) || [0, op])[1].replace(" (or only spaces)", "");
  function nodeText(x) {
    if (!x) return "\u2026";
    if (x.type === "col") return `this row's ${x.name}`;
    if (x.type === "lit") return x.value === "" ? "(empty text)" : `\u201c${x.value}\u201d`;
    return "a calculated value";
  }
  function whereText(w) {
    if (!w) return "";
    if (w.items) {
      const parts = w.items.map(whereText).filter(Boolean);
      if (!parts.length) return "";
      const j = (w.join || "AND") === "OR" ? " or " : " and ";
      return parts.length > 1 ? "(" + parts.join(j) + ")" : parts[0];
    }
    return `${describePipe(w.col, w.pipe)} ${opWord(w.op)}${NO_RIGHT.has(w.op) ? "" : " " + nodeText(w.value)}`;
  }
  function onlyIfText(n) {
    if (!(n.only_if && n.only_if.items && n.only_if.items.length)) return "";
    const j = (n.only_if.join || "AND") === "OR" ? " or " : " and ";
    return n.only_if.items.map((it) => `${describePipe(it.col, it.pipe)} ${opWord(it.op)}${NO_RIGHT.has(it.op) ? "" : " " + nodeText(it.right)}`).join(j);
  }
  function returnText(c, n) {
    if (c.return_type === "count") return "the number of matching rows";
    if (c.return_type === "agg") {
      const w = { SUM: "total", AVERAGE: "average", MAX: "highest value", MIN: "lowest value" }[c.agg_fn || "SUM"];
      return `the ${w} of \u201c${c.return_col}\u201d across all matching rows`;
    }
    if (c.return_type === "value") return nodeText(c.return_value);
    return `the value in \u201c${c.return_col}\u201d`;
  }
  function describeLookup(n) {
    const cur = state.currentSheet;
    let t = `For each row of \u201c${cur}\u201d`;
    const only = onlyIfText(n);
    if (only) t += ` (only rows where ${only})`;
    t += `, look in \u201c${n.sheet}\u201d`;
    if (n.key) {
      const mine = n.key_src ? describePipe(n.key_src.col, n.key_src.pipe) : "the chosen value";
      t += ` at the rows where ${describePipe(n.match_col, n.match_pipe)} equals ${mine}`;
    } else t += " at the rows";
    const parts = n.cases.map((c, i) => {
      const w = whereText(c.where);
      return `${i ? "otherwise " : ""}${w ? `if a row also has ${w}, ` : ""}give ${returnText(c, n)}`;
    });
    t += " \u2014 " + parts.join("; ") + ".";
    if (n.not_found) t += ` If nothing is found: ${nodeText(n.not_found)}.`;
    return t;
  }
  function describeAgg(n) {
    const cur = state.currentSheet;
    let t = `For each row of \u201c${cur}\u201d`;
    const only = onlyIfText(n);
    if (only) t += ` (only rows where ${only})`;
    const what = { SUM: "add up", COUNT: "count the rows", AVERAGE: "average", MAX: "find the highest value", MIN: "find the lowest value" }[n.fn];
    t += `, ${what}${n.fn === "COUNT" ? "" : ` of \u201c${n.value_col}\u201d`} in \u201c${n.sheet}\u201d`;
    const w = whereText(n.where);
    t += w ? ` for the rows where ${w}.` : " (every row).";
    return t;
  }

  function renderCase(n, cs, i, cols, hasKey) {
    if (!cs.where) cs.where = { join: "AND", items: [] };
    const box = h("div", "case");
    const head = h("div", "case-head", [withText("span", "case-word", i === 0 ? "IF" : "ELSE IF"),
      withText("span", "case-sub", hasKey ? "a row of the other sheet matches the link" : "a row of the other sheet")]);
    if (n.cases.length > 1) head.append(iconBtn("\u2715", "Remove this case", () => { n.cases.splice(i, 1); redraw(); }));
    box.append(head);

    const has = leafCount(cs.where) > 0;
    box.append(radioRow([[false, hasKey ? "\u2026any row that matches the link counts" : "\u2026any row counts"],
                         [true, `\u2026only rows of \u201c${n.sheet}\u201d that ALSO meet conditions`]], has, (v) => {
      cs.where = { join: "AND", items: v ? [makeFilterLeaf(cols)] : [] };
      redraw();
    }));
    if (has) box.append(renderFilter(cols, cs.where, `rows of \u201c${n.sheet}\u201d where\u2026`));

    const tileKey = cs.return_type === "agg" ? "agg:" + (cs.agg_fn || "SUM") : cs.return_type;
    const TILES = [
      ["col", "\u2192", "Get a value", "a column of the matching row"],
      ["count", "#", "Count rows", "how many rows match"],
      ["agg:SUM", "\u03a3", "Total", "add up a column"],
      ["agg:AVERAGE", "\u00f8", "Average", "average of a column"],
      ["agg:MAX", "\u2191", "Highest", "biggest value"],
      ["agg:MIN", "\u2193", "Lowest", "smallest value"],
      ["value", "\u270e", "Fixed value", "type your own"],
    ];
    box.append(withText("div", "gsub", "THEN take\u2026"));
    const tiles = h("div", "tiles");
    TILES.forEach(([k, ic, t, sub]) => {
      const b2 = h("button", "tile" + (k === tileKey ? " on" : ""), [withText("span", "tile-ic", ic), withText("span", "tile-t", t), withText("span", "tile-s", sub)]);
      b2.type = "button";
      b2.addEventListener("click", () => {
        if (k.startsWith("agg:")) { cs.return_type = "agg"; cs.agg_fn = k.slice(4); } else cs.return_type = k;
        if (k === "value" && !cs.return_value) cs.return_value = litNode("");
        redraw();
      });
      tiles.append(b2);
    });
    box.append(tiles);
    if (cs.return_type === "value") {
      box.append(renderNode(cs.return_value || (cs.return_value = litNode("")), "Value to use"));
    } else if (cs.return_type === "count") {
      box.append(hint(hasKey ? "For every row of this sheet, counts the rows of the other sheet with the same link value that also meet the conditions above (a group-by count). Rows with none get 0."
                             : "Counts the rows of that sheet that meet the conditions above."));
    } else if (cs.return_type === "agg") {
      box.append(h("div", "field-row", [fld("Of column", colSelect(cols, cs.return_col, (v) => { cs.return_col = v; changed(); }))]),
        hint("Uses every matching row, not just the first."));
    } else {
      box.append(h("div", "field-row", [fld("Column to take", colSelect(cols, cs.return_col, (v) => { cs.return_col = v; changed(); }))]));
    }
    return box;
  }

  const RENDER = {
    col(n, b) {
      b.append(fld("Column", colSelect(curCols(), n.name, (v) => { n.name = v; changed(); })));
    },
    lit(n, b) {
      b.append(fld("Value", inputEl(n.value, (v) => { n.value = v; changed(); }, { placeholder: "type text or a number" })));
    },
    blank(n, b) { b.append(hint("Leaves the cell empty.")); },

    if(n, b) {
      b.append(renderCond(n.cond));
      const br = h("div", "branches");
      br.append(branch("yes", "YES", "Then use", n.then), branch("no", "NO", "Otherwise use", n.else));
      b.append(br);
    },
    iferr(n, b) {
      b.append(fld("Catch", selectEl([["IFERROR", "any error (IFERROR)"], ["IFNA", "only #N/A — 'not found' (IFNA)"]],
        n.fn, (v) => { n.fn = v; changed(); })));
      b.append(renderNode(n.value, "Try this value"));
      b.append(renderNode(n.fallback, "If it fails, use"));
    },

    join(n, b) {
      listOf(n.parts, "Value", () => colNode(0), b);
      const mode = n.delim_mode ?? " ";
      b.append(fld("Put between values", selectEl(DELIMS, mode, (v) => {
        n.delim_mode = v;
        if (v !== "custom") n.delimiter = v;
        redraw();
      })));
      if (mode === "custom") {
        b.append(fld("Custom separator", inputEl(n.delimiter, (v) => { n.delimiter = v; changed(); }, { placeholder: "e.g.  &  or  ->" })));
      }
      b.append(check("Skip empty values (no double separators)", n.skip_blank, (v) => { n.skip_blank = v; changed(); }));
    },
    textfn(n, b) {
      b.append(renderNode(n.source, "Text"));
      b.append(fld("Do this", selectEl([
        ["TRIM", "Remove extra spaces (TRIM)"], ["UPPER", "UPPERCASE"], ["LOWER", "lowercase"],
        ["PROPER", "Capitalize Each Word (PROPER)"], ["LEN", "Count characters (LEN)"], ["VALUE", "Convert text to a number (VALUE)"],
      ], n.fn, (v) => { n.fn = v; redraw(); })));
    },
    replace(n, b) {
      b.append(renderNode(n.source, "Text to change"));
      const list = h("div", "pairs");
      const vis = (t) => t.replace(/ /g, "␣"); // show spaces as a visible symbol
      n.pairs.forEach((p, i) => {
        const r = h("div", "pair");
        const note = h("div", "pair-note");
        const findInput = inputEl(p.find, (v) => { p.find = v; refreshNote(); changed(); }, { placeholder: "find (a space is fine)" });
        const refreshNote = () => {
          note.replaceChildren();
          // a find text that is padded with spaces (but isn't only spaces) is almost always a typo
          if (p.find.trim() !== "" && p.find !== p.find.trim()) {
            note.append(txt(`Looking for "${vis(p.find)}" — it includes a space (␣), so "${p.find.trim()}" on its own won't match. `));
            const fix = withText("button", "link-btn", "Remove the extra space");
            fix.type = "button";
            fix.addEventListener("click", () => {
              p.find = p.find.trim(); findInput.value = p.find; refreshNote(); changed();
            });
            note.append(fix);
          }
          note.hidden = note.childNodes.length === 0;
        };
        r.append(
          txt("replace"), findInput, txt("with"),
          inputEl(p.replace, (v) => { p.replace = v; changed(); }, { placeholder: "nothing = delete it" }),
        );
        if (n.pairs.length > 1) r.append(iconBtn("✕", "Remove", () => { n.pairs.splice(i, 1); redraw(); }));
        refreshNote();
        list.append(r, note);
      });
      b.append(list, linkBtn("+ Add another replacement", () => { n.pairs.push({ find: "", replace: "" }); redraw(); }),
        hint("Applied top to bottom. Leave 'with' empty to delete the text. Matching is case-sensitive, and spaces count."));
    },
    extract(n, b) {
      b.append(renderNode(n.source, "Text to search in"));
      b.append(fld("What do you want?", selectEl(EXTRACT_MODES, n.mode, (v) => { n.mode = v; redraw(); }), EXTRACT_HINTS[n.mode]));
      if (["before_text", "after_text", "between_text"].includes(n.mode)) {
        b.append(fld(n.mode === "between_text" ? "Starting at" : "The character / text",
          textWithChips(n, "text", "type it, or pick one below")));
      }
      if (n.mode === "between_text") b.append(fld("…and ending at", textWithChips(n, "text2", "type it, or pick one below")));
      if (["first_n", "last_n", "mid"].includes(n.mode)) {
        const r = h("div", "field-row");
        if (n.mode === "mid") r.append(fld("Starting at position", numInput(n, "start", 1)));
        r.append(fld("How many characters", numInput(n, "n", 0)));
        b.append(r);
      }
    },
    find(n, b) {
      b.append(renderNode(n.what, "Find this"));
      b.append(renderNode(n.within, "Inside this text"));
      b.append(check("Match upper/lower case exactly (FIND instead of SEARCH)", n.case_sensitive,
        (v) => { n.case_sensitive = v; changed(); }));
      b.append(hint("Gives the position of the text (1 = first character), or #VALUE! when it isn't there — wrap this in an 'If it fails' step to handle that."));
    },
    textfmt(n, b) {
      b.append(renderNode(n.source, "Value"));
      const mode = n.fmt_mode ?? n.fmt;
      b.append(fld("Format", selectEl(FORMATS, FORMATS.some((f) => f[0] === mode) ? mode : "custom", (v) => {
        n.fmt_mode = v;
        if (v !== "custom") n.fmt = v;
        redraw();
      })));
      if (!FORMATS.slice(0, -1).some((f) => f[0] === n.fmt) || n.fmt_mode === "custom") {
        b.append(fld("Custom format", inputEl(n.fmt, (v) => { n.fmt = v; changed(); }, { placeholder: "e.g. yyyy-mm or 0.0" })));
      }
    },

    calc(n, b) {
      b.append(renderNode(n.left, "First value"));
      b.append(fld("Operation", selectEl([["+", "+  add"], ["-", "−  subtract"], ["*", "×  multiply"], ["/", "÷  divide"], ["^", "^  power"]],
        n.op, (v) => { n.op = v; changed(); })));
      b.append(renderNode(n.right, "Second value"));
    },
    round(n, b) {
      b.append(renderNode(n.source, "Number"));
      b.append(fld("How", selectEl([["ROUND", "Round to nearest (ROUND)"], ["ROUNDUP", "Always up (ROUNDUP)"], ["ROUNDDOWN", "Always down (ROUNDDOWN)"]],
        n.fn, (v) => { n.fn = v; changed(); })));
      b.append(fld("Decimal places", numInput(n, "digits"), "Use a negative number to round to tens, hundreds …"));
    },
    rowagg(n, b) {
      b.append(fld("Function", selectEl([["SUM", "SUM — add up"], ["MAX", "MAX — highest"], ["MIN", "MIN — lowest"],
        ["AVERAGE", "AVERAGE"], ["COUNT", "COUNT — how many are numbers"]], n.fn, (v) => { n.fn = v; changed(); })));
      listOf(n.items, "Value", () => colNode(0), b);
    },

    edate(n, b) {
      b.append(renderNode(n.source, "Date"));
      b.append(renderNode(n.months, "Months to add (negative = subtract)"));
    },
    today(n, b) { b.append(hint("Today's date — it updates whenever the file is opened in Excel.")); },
    dateadd(n, b) {
      b.append(renderNode(n.source, "Date"));
      b.append(fld("Unit", selectEl(UNIT_OPTS, n.unit || "MONTHS", (v) => { n.unit = v; redraw(); })));
      b.append(renderNode(n.amount, AMOUNT_LABEL[n.unit || "MONTHS"]));
    },
    datediff(n, b) {
      b.append(renderNode(n.start, "From this date"));
      b.append(renderNode(n.end, "To this date"));
      b.append(fld("Difference in", selectEl(UNIT_OPTS, n.unit || "DAYS", (v) => { n.unit = v; changed(); })));
      const unitLabel = (UNIT_OPTS.find((u) => u[0] === (n.unit || "DAYS")) || [0, "units"])[1].toLowerCase();
      b.append(hint(`Whole ${unitLabel} from the first date to the second — negative if the first date is later.`));
    },

    lookup(n, b) {
      normalizeLookup(n);
      const cur = state.currentSheet;
      const cols = sheetCols(n.sheet);
      const hasKey = !!n.key;
      const first = n.cases[0];
      const simple = hasKey && n.cases.length === 1 && first.return_type === "col" && leafCount(first.where) === 0 && !n.match_pipe.length;
      const sum = h("div", "summary wiz-extra");
      B.updateSummary = () => { sum.textContent = describeLookup(n); };
      B.updateSummary();
      b.append(withText("div", "summary-title wiz-extra", "In plain words"), sum);

      // which of several matching rows?
      const orders = [["first", "The first match (top \u2192 bottom)"], ["last", "The last match (bottom \u2192 top)"],
                      ["nth", "A specific match: Nth from the top"], ["nth_last", "A specific match: Nth from the bottom"]];
      const needsOrder = n.cases.some((c) => c.return_type === "col" || c.return_type === "value");
      let allowed = orders;
      if (simple) allowed = n.method === "VLOOKUP" ? [] : (n.match || "exact") === "exact" ? orders : orders.slice(0, 2);
      if (!simple) n.match = "exact";
      if (!allowed.some((o) => o[0] === (n.search || "first"))) n.search = "first";

      const steps = [];
      steps.push({ key: "rows", short: "which rows", title: `Which rows of \u201c${cur}\u201d should get a value?`,
        help: "Rows that don't qualify are left empty.", build: (s2) => s2.append(renderThisRows(n, cur)) });

      steps.push({ key: "sheet", short: "where to look", title: "Where should it look?", help: "The sheet that holds the values you want. It can be this same sheet.",
        build: (s2) => {
          s2.append(selectEl(sheetOpts(), n.sheet, (v) => {
            const c = sheetCols(v);
            n.sheet = v; n.match_col = c[0];
            n.cases.forEach((cs) => { cs.return_col = c[1] || c[0]; resetLeafCols(cs.where, c[0]); });
            redraw();
          }));
        } });

      steps.push({ key: "link", short: "how they match", title: "How do the two sheets match?",
        help: "Usually a matching value such as a customer id. Messy values can be cleaned on either side first (trim, replace, last 5 characters\u2026).",
        build: (s2) => {
          s2.append(radioRow([[true, "Link the sheets by a matching value"], [false, "Don't link \u2014 choose rows using conditions only"]], hasKey, (v) => {
            if (v) {
              n.key_src = n.key_src || { col: curCols()[0], pipe: [] };
              n.match_col = n.match_col || sheetCols(n.sheet)[0];
              syncLink(n);
            } else { n.key = null; delete n.match_left; }
            redraw();
          }));
          if (hasKey && n.key_src) {
            s2.append(withText("div", "link-label", `In \u201c${cur}\u201d take:`));
            s2.append(renderPipe(curCols(), n.key_src, (st) => { syncLink(n); if (st) redraw(); else changed(); }));
            s2.append(withText("div", "link-label", `\u2026it must equal, in \u201c${n.sheet}\u201d:`));
            const mh = { col: n.match_col, pipe: n.match_pipe };
            s2.append(renderPipe(cols, mh, (st) => { n.match_col = mh.col; n.match_pipe = mh.pipe; syncLink(n); if (st) redraw(); else changed(); }));
          } else if (hasKey) {
            s2.append(renderNode(n.key, "Look for this value"));
            s2.append(fld("Match against column", colSelect(cols, n.match_col, (v) => { n.match_col = v; changed(); })));
          }
          if (simple) {                                        // classic lookup options, kept out of the way
            const adv = h("details", "advanced");
            adv.append(withText("summary", null, "Advanced: lookup method and matching mode"));
            adv.append(fld("Method", selectEl([
              ["XLOOKUP", "XLOOKUP (recommended, modern Excel)"],
              ["VLOOKUP", "VLOOKUP (return column must be right of match column)"],
              ["INDEX-MATCH", "INDEX-MATCH (works in every Excel version)"],
            ], n.method, (v) => {
              n.method = v;
              if (v === "VLOOKUP") { if (!["exact", "smaller"].includes(n.match)) n.match = "exact"; n.search = "first"; }
              if (v === "INDEX-MATCH") { if (n.match === "wildcard") n.match = "exact"; if (n.match !== "exact") n.search = "first"; }
              redraw();
            })));
            const MODES = {
              XLOOKUP: [["exact", "Exact match"], ["smaller", "Exact, or the next smaller value"],
                        ["larger", "Exact, or the next larger value"], ["wildcard", "Wildcard match (* and ?)"]],
              VLOOKUP: [["exact", "Exact match"], ["smaller", "Closest value that is not larger (list sorted low \u2192 high)"]],
              "INDEX-MATCH": [["exact", "Exact match"], ["smaller", "Closest value that is not larger (sorted low \u2192 high)"],
                              ["larger", "Closest value that is not smaller (sorted high \u2192 low)"]],
            }[n.method];
            adv.append(fld("How to match", selectEl(MODES, n.match || "exact", (v) => {
              n.match = v;
              if (v !== "exact" && (n.search === "nth" || n.search === "nth_last")) n.search = "first";
              if (n.method === "INDEX-MATCH" && v !== "exact") n.search = "first";
              redraw();
            })));
            if (n.method === "VLOOKUP") adv.append(hint("VLOOKUP always uses the first match."));
            s2.append(adv);
          }
        } });

      steps.push({ key: "result", short: "what comes back", title: "What should come back?",
        help: n.cases.length > 1 ? "Cases are tried from the top. The first one that finds a row wins."
          : "Optionally narrow the matching rows first, then choose what to take. You can add more cases: IF this, take that \u2014 ELSE IF \u2026",
        build: (s2) => {
          n.cases.forEach((cs, i) => s2.append(renderCase(n, cs, i, cols, hasKey)));
          s2.append(linkBtn("+ Add another case (ELSE IF \u2026)", () => { n.cases.push(newCase(cols)); redraw(); }));
        } });

      if (needsOrder && allowed.length > 1) {
        steps.push({ key: "order", short: "several matches", title: "If several rows qualify\u2026", help: "Choose which one to use.",
          build: (s2) => {
            s2.append(selectEl(allowed, n.search || "first", (v) => { n.search = v; redraw(); }));
            if (n.search === "nth" || n.search === "nth_last") {
              if (!n.nth) n.nth = litNode("2");
              s2.append(renderNode(n.nth, n.search === "nth" ? "Which match? (1 = first, 2 = second \u2026)" : "Which match from the bottom? (1 = last \u2026)"));
            }
          } });
      }

      steps.push({ key: "fallback", short: "if nothing is found", title: "If nothing is found\u2026", help: "What to show when no row qualifies.",
        build: (s2) => {
          s2.append(check("Use a fallback value", !!n.not_found, (on) => { n.not_found = on ? litNode("Not Found") : null; redraw(); }));
          if (n.not_found) s2.append(renderNode(n.not_found, "Otherwise use"));
          if (n.cases[n.cases.length - 1].return_type === "count") s2.append(hint("A count is 0 when nothing matches, so a fallback isn't needed."));
        } });

      b.append(wizard(steps));
    },
    agg(n, b) {      if (!n.where) n.where = n.criteria ? { join: "AND", items: n.criteria } : { join: "AND", items: [] };
      delete n.criteria;
      const cur = state.currentSheet, cols = sheetCols(n.sheet);
      const sum = h("div", "summary wiz-extra");
      B.updateSummary = () => { sum.textContent = describeAgg(n); };
      B.updateSummary();
      b.append(withText("div", "summary-title wiz-extra", "In plain words"), sum);

      b.append(wizard([
        { key: "rows", short: "which rows", title: `Which rows of \u201c${cur}\u201d should get a value?`, help: "Rows that don't qualify are left empty.",
          build: (s2) => s2.append(renderThisRows(n, cur)) },
        { key: "what", short: "what to work out", title: "What should it work out?", help: "Pick the calculation, the sheet to look at and the column to calculate on.",
          build: (s2) => {
            s2.append(fld("Calculate", selectEl([
              ["SUM", "Total (SUM / SUMIF)"], ["COUNT", "Count rows (COUNTIF)"], ["AVERAGE", "Average (AVERAGEIF)"],
              ["MAX", "Highest (MAXIFS)"], ["MIN", "Lowest (MINIFS)"],
            ], n.fn, (v) => { n.fn = v; redraw(); })));
            s2.append(fld("In sheet", selectEl(sheetOpts(), n.sheet, (v) => {
              const c = sheetCols(v);
              n.sheet = v; n.value_col = c[0];
              resetLeafCols(n.where, c[0]);
              redraw();
            })));
            s2.append(fld(n.fn === "COUNT" ? "Column to count (used only when there are no conditions)" : "Column to calculate on",
              colSelect(cols, n.value_col, (v) => { n.value_col = v; changed(); })));
          } },
        { key: "which", short: "which rows count", title: `Which rows of \u201c${n.sheet}\u201d count?`,
          help: "Compare a column with 'A column's value' of this row to link the sheets (e.g. customer id equals this row's customer id). Add more conditions to narrow it down; use an ANY group for OR.",
          build: (s2) => s2.append(renderFilter(cols, n.where, `rows of \u201c${n.sheet}\u201d where\u2026`, { emptyHint: "No conditions \u2014 every row of the sheet counts." })) },
      ]));
    },
  };

  // -------------------------------------------------------------- preview
  let timer = null, seq = 0;
  function changed() {
    if (B && B.updateSummary) B.updateSummary();
    clearTimeout(timer);
    timer = setTimeout(runPreview, 300);
  }
  function redraw() {
    const tree = $("#builderTree");
    if (!tree) return;
    B.updateSummary = null;
    B.wizard = false;
    tree.replaceChildren(renderNode(B.root, "Result", { test: true }));
    // In a wizard only the current question is on screen; summary, output and preview appear on the last step.
    const focus = B.wizard && !B.wizLast;
    tree.classList.toggle("wiz-mode", B.wizard);
    tree.classList.toggle("wiz-focus", focus);
    ["#builderTail", "#applyBtn"].forEach((sel) => { const el = $(sel); if (el) el.hidden = focus; });
    const t = $("#builderTitle");
    if (t && !t.dataset.fixed) t.textContent = B.root.type === PRESETS[B.op] ? (TITLES[B.op] || "Formula Builder") : "Formula Builder";
    changed();
  }
  function payload() {
    return {
      session_id: state.sessionId, sheet_name: state.currentSheet, operation: "EXPR",
      params: { expr: B.root }, output_column: "",
    };
  }
  async function runPreview() {
    const body = $("#previewBody");
    if (!body) return;
    const mine = ++seq;
    let data, ok;
    try {
      const res = await fetch(`${API_BASE}/api/operations/preview`, {
        method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload()),
      });
      ok = res.ok;
      data = await res.json();
    } catch (e) {
      ok = false; data = { detail: "Can't reach the server: " + e.message };
    }
    if (mine !== seq) return; // a newer edit superseded this one
    body.replaceChildren();
    if (!ok) {
      const d = data.detail;
      body.append(withText("div", "preview-error", typeof d === "string" ? d : "Something in the flow isn't filled in yet."));
      return;
    }
    if (data.formula) body.append(withText("code", "formula-preview", data.formula));
    const t = h("table", "preview-table");
    data.rows.slice(0, 5).forEach((r) => {
      const v = r.value === null || r.value === undefined || r.value === "" ? "(empty)" : String(r.value);
      const td = withText("td", v === "(empty)" ? "muted" : "", v);
      t.append(h("tr", null, [withText("th", null, "Row " + r.row), td]));
    });
    body.append(t);
  }

  // ---------------------------------------------------------------- open
  window.openBuilder = function (op) {
    const meta = currentSheetMeta();
    B = { root: makeNode(PRESETS[op] || "if"), op };
    openModal(`
      <h2 id="builderTitle">${esc(TITLES[op] || "Formula Builder")}</h2>
      <p class="modal-sub">${["LOOKUP", "CONDITIONAL_AGG"].includes(op) ? "Answer one question at a time — use Next and Back to move between them." : "Pick what the result should be. A Decision has a YES path and a NO path, and each path can hold another decision."}</p>
      <div id="builderTree"></div>
      <div id="builderTail">
        ${outputField(meta, DEFAULT_OUT[op] || "result")}
        <div class="preview">
          <div class="preview-title">Live preview <span>— nothing changes until you press Apply</span></div>
          <div id="previewBody" class="preview-body"></div>
        </div>
      </div>
      <div class="modal-actions">
        <button class="btn secondary" onclick="closeModal()">Cancel</button>
        <button class="btn primary" id="applyBtn">Apply</button>
      </div>
    `, (root) => {
      redraw();
      root.querySelector("#applyBtn").addEventListener("click", () => {
        submitOperation({
          session_id: state.sessionId, sheet_name: state.currentSheet, operation: "EXPR",
          label: TITLES[op] || "Formula", params: { expr: B.root },
          output_column: readOutput(root, DEFAULT_OUT[op] || "result"),
        });
      });
    }, true);
  };

  // Open the builder with a tree that already exists (e.g. translated from SQL).
  window.openBuilderWith = function (root, outputName, title) {
    window.openBuilder("BUILDER");
    replaceNode(B.root, JSON.parse(JSON.stringify(root)));
    const t = $("#builderTitle");
    if (t) t.dataset.fixed = title || "";
    redraw();
    if (title && t) t.textContent = title;
    const sel = $("#outputSel"), inp = $("#outputCol");
    if (sel && inp && outputName) {
      if (curCols().includes(outputName)) { sel.value = outputName; } else { sel.value = "__new__"; inp.value = outputName; }
      sel.dispatchEvent(new Event("change"));
    }
  };
})();
