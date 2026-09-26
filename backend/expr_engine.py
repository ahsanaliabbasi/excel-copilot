"""
expr_engine.py
--------------
Turns the dashboard's decision tree ("flow chart") into

  1. a REAL Excel formula for every row (written into the .xlsx), and
  2. the value that formula produces (computed here in Python so the UI can show
     results immediately, without Excel having to recalculate).

Every node type has two methods that must stay in sync:
    f_<type>(node, row_num) -> Excel formula text (no leading "=")
    v_<type>(node, i)       -> python value for data row i (0-based)

Node types
    col, lit, blank                      basic values
    if, iferr                            decisions / error handling (nestable)
    join, textfn, replace, extract,
    find, textfmt                        text
    calc, round, rowagg                  numbers
    edate, today                         dates
    lookup, agg                          XLOOKUP/VLOOKUP/INDEX-MATCH, SUMIF/COUNTIF/...

The python side follows Excel's own rules (blank cell -> 0 at the top level,
case-insensitive text compare, first match wins in lookups, round-half-up ...)
so the preview matches what Excel shows when the file is opened.
"""
from __future__ import annotations

import calendar
import datetime as dt
import re
from bisect import bisect_left, bisect_right
from decimal import Decimal, ROUND_DOWN, ROUND_HALF_UP, ROUND_UP

import pandas as pd

from formula_engine import col_letter, col_range_ref, quote_sheet

NA, VALUE, DIV0 = "#N/A", "#VALUE!", "#DIV/0!"
EPOCH = dt.date(1899, 12, 30)
DIGITS_ARR = "{0,1,2,3,4,5,6,7,8,9}"
LETTERS_ARR = "{" + ",".join(f'"{c}"' for c in "abcdefghijklmnopqrstuvwxyz") + "}"
ASCII_LETTERS = set("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ")


class XlError(Exception):
    """An Excel error value such as #N/A - propagates like it does in a sheet."""
    def __init__(self, code):
        super().__init__(code)
        self.code = code


# ---------------------------------------------------------------------------
# value helpers
# ---------------------------------------------------------------------------
def _py(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, "item"):
        return v.item()
    return v


def is_blank(v):
    return v is None or (isinstance(v, str) and v == "")


def is_num(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def tidy(x):
    return int(x) if isinstance(x, float) and x.is_integer() and abs(x) < 1e15 else x


def serial(d):
    if isinstance(d, dt.datetime):
        return tidy((d - dt.datetime(1899, 12, 30)).total_seconds() / 86400)
    return (d - EPOCH).days


def to_num(v):
    if is_blank(v):
        return 0
    if isinstance(v, bool):
        return int(v)
    if is_num(v):
        return v
    if isinstance(v, (dt.date, dt.datetime)):
        return serial(v)
    try:
        return tidy(float(str(v).strip().replace(",", "")))
    except ValueError:
        raise XlError(VALUE)


def to_text(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    if isinstance(v, (dt.date, dt.datetime)):
        return to_text(serial(v))          # Excel joins dates as serial numbers
    return str(v)


def to_date(v):
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if is_num(v):
        return EPOCH + dt.timedelta(days=int(v))
    if isinstance(v, str):
        s = v.strip()
        for f in ("%Y-%m-%d", "%m/%d/%Y", "%Y/%m/%d", "%d-%b-%Y", "%Y-%m-%d %H:%M:%S"):
            try:
                return dt.datetime.strptime(s, f).date()
            except ValueError:
                pass
    raise XlError(VALUE)      # (Excel treats a blank as 1900-01-00; we flag it instead)


def parse_lit(s):
    """Typed values: numbers become numbers, everything else stays text."""
    if isinstance(s, (int, float)) and not isinstance(s, bool):
        return s
    s = "" if s is None else str(s)
    if re.fullmatch(r"-?(0|[1-9]\d*)(\.\d+)?", s.strip()):
        f = float(s)
        return int(f) if "." not in s else f
    return s


def q(s):
    return '"' + str(s).replace('"', '""') + '"'


def lit_str(x):
    return q(x) if isinstance(x, str) else str(x)


def _key(v):
    if isinstance(v, bool):
        return (2, v)
    if is_num(v):
        return (0, float(v))
    if isinstance(v, (dt.date, dt.datetime)):
        return (0, float(serial(v)))
    return (1, str(v).lower())


def compare(l, r, op):
    """Excel comparison rules: blanks act like 0 / "", text is case-insensitive,
    numbers < text < logicals."""
    lb, rb = is_blank(l), is_blank(r)
    numlike = lambda x: is_num(x) or isinstance(x, (dt.date, dt.datetime))
    if lb and rb:
        l = r = ""
    elif lb:
        l = 0 if numlike(r) else ""
    elif rb:
        r = 0 if numlike(l) else ""
    a, b = _key(l), _key(r)
    return {"=": a == b, "<>": a != b, ">": a > b, "<": a < b, ">=": a >= b, "<=": a <= b}[op]


def xl_round(x, digits, mode):
    d = Decimal(str(to_num(x)))
    r = d.quantize(Decimal(1).scaleb(-digits), rounding=mode)
    return tidy(float(r)) if digits > 0 else int(r)


_DATE_TOK = re.compile(r"yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d", re.I)


def fmt_text(v, fmt):
    """Python version of Excel TEXT() for the common date / number formats."""
    if re.search(r"[ymd]", fmt, re.I) and not re.search(r"[#0]", fmt):
        d = to_date(v)
        names = {
            "yyyy": f"{d.year:04d}", "yy": f"{d.year % 100:02d}", "mmmm": d.strftime("%B"),
            "mmm": d.strftime("%b"), "mm": f"{d.month:02d}", "m": str(d.month),
            "dddd": d.strftime("%A"), "ddd": d.strftime("%a"), "dd": f"{d.day:02d}", "d": str(d.day),
        }
        return _DATE_TOK.sub(lambda m: names[m.group(0).lower()], fmt)
    m = re.search(r"[#0,]+(?:\.[0#]+)?%?", fmt)
    if not m:
        return fmt
    pat = m.group(0)
    pct = pat.endswith("%")
    ip, _, dp = pat.rstrip("%").partition(".")
    decimals, thousands, min_int = len(dp), "," in ip, ip.replace(",", "").count("0")
    x = Decimal(str(to_num(v))) * (100 if pct else 1)
    qv = x.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
    body = f"{abs(qv):,.{decimals}f}" if thousands else f"{abs(qv):.{decimals}f}"
    whole, dot, frac = body.partition(".")
    if not thousands:
        whole = whole.zfill(min_int)
    return fmt[:m.start()] + ("-" if qv < 0 else "") + whole + dot + frac + ("%" if pct else "") + fmt[m.end():]


def add_months(d, months):
    m = d.month - 1 + int(months)
    y = d.year + m // 12
    m = m % 12 + 1
    return dt.date(y, m, min(d.day, calendar.monthrange(y, m)[1]))


# ---------------------------------------------------------------------------
# the engine
# ---------------------------------------------------------------------------
# ---------------------------------------------------------------------------
# "where" conditions for lookups and totals: nestable AND / OR groups of
#   leaf  = {"col": <column of the searched sheet>, "op": ..., "value": <step evaluated on the current row>}
#   group = {"join": "AND" | "OR", "items": [leaf | group, ...]}
# The old flat list  n["criteria"]  still works (= an AND group).
# ---------------------------------------------------------------------------
def where_of(n):
    w = n.get("where")
    if w is None and n.get("criteria") is not None:
        w = {"join": "AND", "items": n["criteria"]}
    if w is not None and "items" in w and not w["items"]:
        return None
    return w


def is_group(x):
    return "items" in x


def where_leaves(w):
    if w is None:
        return
    if is_group(w):
        for it in w["items"]:
            yield from where_leaves(it)
    else:
        yield w


def where_is_classic(w):
    """A plain AND of simple conditions -> can use COUNTIFS / SUMIFS / MAXIFS ..."""
    if w is None:
        return True
    if any(l.get("left") for l in where_leaves(w)):              # a transformed column needs the array form
        return False
    return not is_group(w) or (w.get("join", "AND") == "AND" and all(not is_group(i) for i in w["items"]))


def where_conjuncts(w):
    """Conditions that ALL have to hold (used to narrow the rows to check with an index)."""
    if w is None:
        return
    if not is_group(w):
        yield w
    elif w.get("join", "AND") == "AND":
        for it in w["items"]:
            yield from where_conjuncts(it)


ORDER_OPS = (">", "<", ">=", "<=")


def search_is_nth(n):
    return n.get("search", "first") in ("nth", "nth_last")


class Engine:
    def __init__(self, ctx, all_ctx):
        self.ctx = ctx
        self.all = all_ctx
        self._lists = {}
        self._maps = {}
        self._views = {}
        self.cur = None            # while evaluating a condition on the searched sheet: that sheet (python side)
        self.arr = None            # while writing such a condition as an Excel formula: that sheet (columns become ranges)
        self.shared = ctx.get("_cache")              # {(sheet, column, version): [python values]} owned by the server

    # ---- sheet / column access -------------------------------------------
    def sheet(self, name):
        if not name or name == self.ctx["sheet_name"]:
            return self.ctx
        if name not in self.all:
            raise ValueError(f"Sheet '{name}' not found")
        return self.all[name]

    def idx(self, sh, col):
        if col not in sh["columns"]:
            raise ValueError(f"Column '{col}' not found in sheet '{sh['sheet_name']}'")
        return sh["columns"][col]

    def values(self, sh, col):
        key = (sh["sheet_name"], col)
        if key in self._lists:
            return self._lists[key]
        skey = (sh["sheet_name"], col, sh.get("ver"))
        if self.shared is not None and skey in self.shared:
            lst = self.shared[skey]
        else:
            self.idx(sh, col)
            series = sh["df"][col]
            if str(series.dtype).startswith("datetime"):
                lst = [_py(v) for v in series]
            else:
                raw = series.tolist()
                try:
                    lst = [None if v != v else v for v in raw]          # NaN -> None
                except (TypeError, ValueError):
                    lst = [_py(v) for v in raw]
            if self.shared is not None:
                self.shared[skey] = lst
        self._lists[key] = lst
        return lst

    def rng(self, sh, col):
        return col_range_ref(sh["sheet_name"], self.idx(sh, col), sh["header_row"], sh["last_row"])

    # ---- dispatch ---------------------------------------------------------
    def _only_view(self, n):
        tag = ("only", id(n))
        if tag not in self._views:
            self._views[tag] = {k: v for k, v in n.items() if k not in ("only_if", "only_else")}
        return self._views[tag]

    def f(self, n, row):
        oi = n.get("only_if")
        if oi and oi.get("items"):                            # IF(<this row meets>, <the step>, <otherwise: empty>)
            other = self.f(n["only_else"], row) if n.get("only_else") else '""'
            return f"IF({self.f_cond(oi, row)},{self.f(self._only_view(n), row)},{other})"
        fn = getattr(self, "f_" + str(n.get("type")), None)
        if not fn:
            raise ValueError(f"Unknown step type '{n.get('type')}'")
        return fn(n, row)

    def v(self, n, i):
        oi = n.get("only_if")
        if oi and oi.get("items"):
            if self.v_cond(oi, i):
                return self.v(self._only_view(n), i)
            return self.v(n["only_else"], i) if n.get("only_else") else ""
        return getattr(self, "v_" + n["type"])(n, i)

    # ---- basic values -----------------------------------------------------
    def f_col(self, n, row):
        if self.arr is not None:                              # inside a condition: the whole column of the searched sheet
            return self.rng(self.arr, n["name"])
        return f"{col_letter(self.idx(self.ctx, n['name']))}{row}"

    def v_col(self, n, i):
        return self.values(self.cur if self.cur is not None else self.ctx, n["name"])[i]

    def f_lit(self, n, row):
        return lit_str(parse_lit(n.get("value")))

    def v_lit(self, n, i):
        return parse_lit(n.get("value"))

    def f_blank(self, n, row):
        return '""'

    def v_blank(self, n, i):
        return ""

    # ---- conditions -------------------------------------------------------
    def f_cond(self, cond, row):
        parts = [self.f_cmp(it, row) for it in cond["items"]]
        if not parts:
            raise ValueError("Add at least one condition")
        return parts[0] if len(parts) == 1 else f"{cond.get('join', 'AND')}({','.join(parts)})"

    def f_cmp(self, it, row):
        op, L = it["op"], self.f(it["left"], row)
        if op == "blank":                       # empty, or nothing but spaces
            return f"(LEN(TRIM({L}))=0)"
        if op == "notblank":
            return f"(LEN(TRIM({L}))>0)"
        R = self.f(it["right"], row)
        if op in ("=", "<>", ">", "<", ">=", "<="):
            return f"({L}{op}{R})"
        if op == "contains":
            return f"ISNUMBER(SEARCH({R},{L}))"
        if op == "notcontains":
            return f"ISERROR(SEARCH({R},{L}))"
        if op == "starts":
            return f"(LEFT({L},LEN({R}))={R})"
        if op == "ends":
            return f"(RIGHT({L},LEN({R}))={R})"
        raise ValueError(f"Unknown comparison '{op}'")

    def v_cond(self, cond, i):
        res = [self.v_cmp(it, i) for it in cond["items"]]
        return any(res) if cond.get("join") == "OR" else all(res)

    def v_cmp(self, it, i):
        op, l = it["op"], self.v(it["left"], i)
        if op == "blank":
            return to_text(l).strip(" ") == ""
        if op == "notblank":
            return to_text(l).strip(" ") != ""
        r = self.v(it["right"], i)
        if op in ("=", "<>", ">", "<", ">=", "<="):
            return compare(l, r, op)
        a, b = to_text(l).lower(), to_text(r).lower()
        return {"contains": b in a, "notcontains": b not in a,
                "starts": a.startswith(b), "ends": a.endswith(b)}[op]

    # ---- decisions --------------------------------------------------------
    def f_if(self, n, row):
        return f"IF({self.f_cond(n['cond'], row)},{self.f(n['then'], row)},{self.f(n['else'], row)})"

    def v_if(self, n, i):
        return self.v(n["then"], i) if self.v_cond(n["cond"], i) else self.v(n["else"], i)

    def f_iferr(self, n, row):
        fn = "_xlfn.IFNA" if n.get("fn", "IFNA") == "IFNA" else "IFERROR"
        return f"{fn}({self.f(n['value'], row)},{self.f(n['fallback'], row)})"

    def v_iferr(self, n, i):
        try:
            return self.v(n["value"], i)
        except XlError as e:
            if n.get("fn", "IFNA") == "IFNA" and e.code != NA:
                raise
            return self.v(n["fallback"], i)

    # ---- text -------------------------------------------------------------
    def f_join(self, n, row):
        parts = [self.f(p, row) for p in n.get("parts", [])]
        if not parts:
            raise ValueError("Add at least one value to join")
        d = n.get("delimiter", "") or ""
        if n.get("skip_blank"):
            return f"_xlfn.TEXTJOIN({q(d)},TRUE,{','.join(parts)})"
        return (f"&{q(d)}&" if d else "&").join(parts)

    def v_join(self, n, i):
        d = n.get("delimiter", "") or ""
        texts = [to_text(self.v(p, i)) for p in n["parts"]]
        if n.get("skip_blank"):
            texts = [t for t in texts if t != ""]
        return d.join(texts)

    def f_textfn(self, n, row):
        if n["fn"] not in ("TRIM", "UPPER", "LOWER", "PROPER", "LEN", "VALUE"):
            raise ValueError(f"Unknown text function {n['fn']}")
        return f"{n['fn']}({self.f(n['source'], row)})"

    def v_textfn(self, n, i):
        fn, v = n["fn"], self.v(n["source"], i)
        if fn == "VALUE":
            if is_num(v):
                return v
            try:
                return tidy(float(to_text(v).strip().replace(",", "")))
            except ValueError:
                raise XlError(VALUE)
        s = to_text(v)
        if fn == "TRIM":
            return re.sub(" +", " ", s).strip(" ")
        if fn == "UPPER":
            return s.upper()
        if fn == "LOWER":
            return s.lower()
        if fn == "PROPER":
            return s.title()
        return len(s)

    def _pairs(self, n):
        pairs = [(p.get("find", ""), p.get("replace", "")) for p in n.get("pairs", [])]
        pairs = [p for p in pairs if p[0] != ""]
        if not pairs:
            raise ValueError("Add at least one replacement (the 'find' box can't be empty)")
        return pairs

    def f_replace(self, n, row):
        out = self.f(n["source"], row)
        for find, rep in self._pairs(n):
            out = f"SUBSTITUTE({out},{q(find)},{q(rep)})"
        return out

    def v_replace(self, n, i):
        s = to_text(self.v(n["source"], i))
        for find, rep in self._pairs(n):
            s = s.replace(find, rep)
        return s

    def _extract_need(self, n):
        mode = n["mode"]
        if mode in ("before_text", "after_text", "between_text") and not n.get("text"):
            raise ValueError("Type the text/character to look for")
        if mode == "between_text" and not n.get("text2"):
            raise ValueError("Type the ending text/character")
        return mode

    def f_extract(self, n, row):
        mode, s = self._extract_need(n), self.f(n["source"], row)
        c, c2 = q(n.get("text", "")), q(n.get("text2", ""))
        if mode == "before_text":
            return f"LEFT({s},FIND({c},{s}&{c})-1)"
        if mode == "after_text":
            if self.arr is not None:
                return f"MID({s},FIND({c},{s}&{c})+LEN({c}),LEN({s}))"
            return f'IFERROR(MID({s},FIND({c},{s})+LEN({c}),LEN({s})),"")'
        if mode == "between_text":
            st = f"FIND({c},{s})+LEN({c})"
            return f'IFERROR(MID({s},{st},FIND({c2},{s},{st})-({st})),"")'
        if mode in ("before_digit", "from_digit"):
            pos = f'MIN(FIND({DIGITS_ARR},{s}&"0123456789"))'
        elif mode in ("before_letter", "from_letter"):
            pos = f'MIN(SEARCH({LETTERS_ARR},{s}&"abcdefghijklmnopqrstuvwxyz"))'
        elif mode == "first_n":
            return f"LEFT({s},{int(n.get('n', 1))})"
        elif mode == "last_n":
            return f"RIGHT({s},{int(n.get('n', 1))})"
        elif mode == "mid":
            return f"MID({s},{int(n.get('start', 1))},{int(n.get('n', 1))})"
        else:
            raise ValueError(f"Unknown extract mode {mode}")
        return f"LEFT({s},{pos}-1)" if mode.startswith("before") else f"MID({s},{pos},LEN({s}))"

    def v_extract(self, n, i):
        mode, s = self._extract_need(n), to_text(self.v(n["source"], i))
        c, c2 = n.get("text", ""), n.get("text2", "")
        if mode == "before_text":
            k = s.find(c)
            return s if k < 0 else s[:k]
        if mode == "after_text":
            k = s.find(c)
            return "" if k < 0 else s[k + len(c):]
        if mode == "between_text":
            k = s.find(c)
            if k < 0:
                return ""
            j = s.find(c2, k + len(c))
            return "" if j < 0 else s[k + len(c):j]
        if mode in ("before_digit", "from_digit"):
            k = next((x for x, ch in enumerate(s) if ch in "0123456789"), len(s))
        elif mode in ("before_letter", "from_letter"):
            k = next((x for x, ch in enumerate(s) if ch in ASCII_LETTERS), len(s))
        elif mode == "first_n":
            return s[:int(n.get("n", 1))]
        elif mode == "last_n":
            k = int(n.get("n", 1))
            return s[-k:] if k > 0 else ""
        else:  # mid
            st, k = int(n.get("start", 1)), int(n.get("n", 1))
            return s[st - 1:st - 1 + k]
        return s[:k] if mode.startswith("before") else s[k:]

    def f_find(self, n, row):
        fn = "FIND" if n.get("case_sensitive") else "SEARCH"
        return f"{fn}({self.f(n['what'], row)},{self.f(n['within'], row)})"

    def v_find(self, n, i):
        what, within = to_text(self.v(n["what"], i)), to_text(self.v(n["within"], i))
        if not n.get("case_sensitive"):
            what, within = what.lower(), within.lower()
        k = within.find(what)
        if k < 0:
            raise XlError(VALUE)
        return k + 1

    def f_textfmt(self, n, row):
        return f"TEXT({self.f(n['source'], row)},{q(n.get('fmt', ''))})"

    def v_textfmt(self, n, i):
        v = self.v(n["source"], i)
        return to_text(v) if isinstance(v, str) and not is_num(v) else fmt_text(v, n.get("fmt", ""))

    # ---- numbers ----------------------------------------------------------
    def f_calc(self, n, row):
        if n["op"] not in ("+", "-", "*", "/", "^"):
            raise ValueError(f"Unknown operator {n['op']}")
        return f"({self.f(n['left'], row)}{n['op']}{self.f(n['right'], row)})"

    def v_calc(self, n, i):
        a, b, op = to_num(self.v(n["left"], i)), to_num(self.v(n["right"], i)), n["op"]
        if op == "/":
            if b == 0:
                raise XlError(DIV0)
            return tidy(a / b)
        return tidy({"+": lambda: a + b, "-": lambda: a - b, "*": lambda: a * b, "^": lambda: a ** b}[op]())

    def f_round(self, n, row):
        if n.get("fn", "ROUND") not in ("ROUND", "ROUNDUP", "ROUNDDOWN"):
            raise ValueError("Unknown rounding function")
        return f"{n.get('fn', 'ROUND')}({self.f(n['source'], row)},{int(n.get('digits', 0))})"

    def v_round(self, n, i):
        mode = {"ROUND": ROUND_HALF_UP, "ROUNDUP": ROUND_UP, "ROUNDDOWN": ROUND_DOWN}[n.get("fn", "ROUND")]
        return xl_round(self.v(n["source"], i), int(n.get("digits", 0)), mode)

    def f_rowagg(self, n, row):
        if n["fn"] not in ("SUM", "MAX", "MIN", "AVERAGE", "COUNT"):
            raise ValueError(f"Unknown function {n['fn']}")
        if not n.get("items"):
            raise ValueError("Add at least one value")
        return f"{n['fn']}({','.join(self.f(x, row) for x in n['items'])})"

    def v_rowagg(self, n, i):
        nums = [x for x in (self.v(it, i) for it in n["items"]) if is_num(x)]
        return self._reduce(n["fn"], nums)

    @staticmethod
    def _reduce(fn, nums):
        if fn == "COUNT":
            return len(nums)
        if fn == "SUM":
            return tidy(sum(nums))
        if fn == "AVERAGE":
            if not nums:
                raise XlError(DIV0)
            return tidy(sum(nums) / len(nums))
        if not nums:
            return 0
        return max(nums) if fn == "MAX" else min(nums)

    # ---- dates ------------------------------------------------------------
    def f_edate(self, n, row):
        return f"EDATE({self.f(n['source'], row)},{self.f(n['months'], row)})"

    def v_edate(self, n, i):
        return add_months(to_date(self.v(n["source"], i)), to_num(self.v(n["months"], i)))

    def f_today(self, n, row):
        return "TODAY()"

    def v_today(self, n, i):
        return dt.date.today()

    # ---- lookup -----------------------------------------------------------
    # n["match"]  : exact | smaller (exact or next smaller) | larger (exact or next larger) | wildcard
    # n["search"] : first (top to bottom) | last (bottom to top)
    def _lookup_opts(self, n):
        method, match = n.get("method", "XLOOKUP"), n.get("match", "exact")
        search = n.get("search", "first")
        last = search == "last"
        if search not in ("first", "last", "nth", "nth_last"):
            raise ValueError(f"Unknown search order '{search}'")
        if where_of(n) is not None or not n.get("key") or n.get("match_left"):
            if match != "exact":
                raise ValueError("Extra conditions only work together with an exact match.")
            if search in ("nth", "nth_last") and not n.get("nth"):
                raise ValueError("Say which match number you want (N).")
            return "INDEX-MATCH", "exact", last
        if search in ("nth", "nth_last"):
            if method == "VLOOKUP":
                raise ValueError("VLOOKUP can't pick the Nth match. Use XLOOKUP or INDEX-MATCH.")
            if match != "exact":
                raise ValueError("Picking the Nth match only works together with an exact match.")
            if not n.get("nth"):
                raise ValueError("Say which match number you want (N).")
        if match not in ("exact", "smaller", "larger", "wildcard"):
            raise ValueError(f"Unknown match mode '{match}'")
        if method == "VLOOKUP" and match in ("larger", "wildcard"):
            raise ValueError("VLOOKUP can only do an exact match or 'closest value that is not larger'. "
                             "Use XLOOKUP for the other modes.")
        if last and method == "VLOOKUP":
            raise ValueError("VLOOKUP always returns the first match. Use XLOOKUP or INDEX-MATCH to get the last one.")
        if last and method == "INDEX-MATCH" and match != "exact":
            raise ValueError("'Last match' with INDEX-MATCH only works together with an exact match.")
        return method, match, last

    def _lookup_view(self, n):
        """A lookup with exactly one plain case is the classic single lookup: reuse that code path unchanged."""
        cs = n.get("cases")
        if not cs:
            return n
        if id(n) not in self._views:
            view = n
            if len(cs) == 1 and cs[0].get("return_type", "col") == "col" and cs[0].get("return_col"):
                view = {k: v for k, v in n.items() if k != "cases"}
                view["return_col"], view["where"] = cs[0]["return_col"], cs[0].get("where")
            self._views[id(n)] = view
        return self._views[id(n)]

    def f_lookup(self, n, row):
        sh = self.sheet(n.get("sheet"))
        n = self._lookup_view(n)
        if n.get("cases"):
            return self._f_lookup_cases(n, sh, row)
        method, match, last = self._lookup_opts(n)
        nf = self.f(n["not_found"], row) if n.get("not_found") else None
        w = where_of(n)
        if w is not None or not n.get("key") or n.get("match_left"):
            return self._f_lookup_where(n, sh, w, row, nf)
        key = self.f(n["key"], row)
        m_rng, r_rng = self.rng(sh, n["match_col"]), self.rng(sh, n["return_col"])
        search = n.get("search", "first")
        if search in ("nth", "nth_last"):
            # AGGREGATE evaluates the array itself, so this works in any Excel from 2010 without Ctrl+Shift+Enter
            fn_no = 15 if search == "nth" else 14          # 15 = SMALL (from the top), 14 = LARGE (from the bottom)
            core = (f"INDEX({r_rng},_xlfn.AGGREGATE({fn_no},6,(ROW({m_rng})-{sh['header_row']})/({m_rng}={key}),"
                    f"{self.f(n['nth'], row)}))")
            return f"IFERROR({core},{nf})" if nf else core
        if method == "XLOOKUP":
            mm = {"exact": 0, "smaller": -1, "larger": 1, "wildcard": 2}[match]
            extra = ""
            if mm or last:
                extra = f",{nf or ''},{mm}" + (",-1" if last else "")
            elif nf:
                extra = f",{nf}"
            return f"_xlfn.XLOOKUP({key},{m_rng},{r_rng}{extra})"
        if method == "VLOOKUP":
            mi, ri = self.idx(sh, n["match_col"]), self.idx(sh, n["return_col"])
            if ri < mi:
                raise ValueError("VLOOKUP needs the return column to be to the RIGHT of the match "
                                 "column. Use XLOOKUP or INDEX-MATCH instead.")
            full = (f"{quote_sheet(sh['sheet_name'])}!${col_letter(mi)}${sh['header_row'] + 1}:"
                    f"${col_letter(ri)}${sh['last_row']}")
            core = f"VLOOKUP({key},{full},{ri - mi + 1},{'TRUE' if match == 'smaller' else 'FALSE'})"
        elif last:
            core = f"LOOKUP(2,1/({m_rng}={key}),{r_rng})"
        else:
            mt = {"exact": 0, "wildcard": 0, "smaller": 1, "larger": -1}[match]
            core = f"INDEX({r_rng},MATCH({key},{m_rng},{mt}))"
        return f"IFERROR({core},{nf})" if nf else core

    # ---- "where" conditions: Excel side --------------------------------------
    @staticmethod
    def _numeric_lit(node):
        return bool(node) and node.get("type") == "lit" and is_num(parse_lit(node.get("value")))

    ARRAY_OK = {"col", "lit", "blank", "textfn", "replace", "extract", "textfmt", "round", "calc"}

    def _check_array_ok(self, node):
        t = node.get("type")
        if t not in self.ARRAY_OK:
            raise ValueError("That step can't be used on a whole column inside a condition. "
                             "Use Trim, Upper/Lower, Length, Replace, First/Last N characters, Text before/after, Round or a calculation.")
        if t == "textfn" and node.get("fn") == "VALUE":
            raise ValueError("'Convert text to a number' can't be used inside a condition (it fails on text). Compare the text directly instead.")
        if t == "extract" and node.get("mode") in ("before_digit", "from_digit", "before_letter", "from_letter", "between_text"):
            raise ValueError("That way of cutting text can't be used inside a condition. Use First / Last N characters or Text before / after a character.")
        if t == "calc" and node.get("op") not in ("+", "-", "*"):
            raise ValueError("Only + - × can be used inside a condition.")
        for k in ("source", "left", "right"):
            if isinstance(node.get(k), dict):
                self._check_array_ok(node[k])

    def _array_expr(self, sh, node, row):
        """A column of `sh` run through some steps, as an Excel array expression (whole column at once)."""
        self._check_array_ok(node)
        prev, self.arr = self.arr, sh
        try:
            return self.f(node, row)
        finally:
            self.arr = prev

    def _expr_values(self, sh, key, node):
        """The same steps evaluated on every row of `sh` (cached)."""
        tag = ("exprvals", key, sh["sheet_name"])
        if tag not in self._maps:
            self._check_array_ok(node)
            out, prev = [], self.cur
            self.cur = sh
            try:
                for j in range(len(sh["df"])):
                    try:
                        out.append(self.v(node, j))
                    except XlError:
                        out.append("")
            finally:
                self.cur = prev
            self._maps[tag] = out
        return self._maps[tag]

    def _leaf_range(self, sh, c, row):
        return self._array_expr(sh, c["left"], row) if c.get("left") else self.rng(sh, c["col"])

    def _leaf_values(self, sh, c):
        return self._expr_values(sh, id(c), c["left"]) if c.get("left") else self.values(sh, c["col"])

    def _leaf_index(self, sh, c):
        tag = ("leafidx", id(c), sh["sheet_name"])
        if tag not in self._maps:
            mp = {}
            for j, x in enumerate(self._leaf_values(sh, c)):
                if not is_blank(x):
                    for k in self._eq_keys(x):
                        mp.setdefault(k, []).append(j)
            self._maps[tag] = mp
        return self._maps[tag]

    # ---- the column a lookup matches against (optionally transformed) ----
    def _match_values(self, sh, n):
        return self._expr_values(sh, ("m", id(n)), n["match_left"]) if n.get("match_left") else self.values(sh, n["match_col"])

    def _match_range(self, sh, n, row):
        return self._array_expr(sh, n["match_left"], row) if n.get("match_left") else self.rng(sh, n["match_col"])

    def _leaf_arr(self, sh, c, row):
        R, op = self._leaf_range(sh, c, row), c["op"]
        if op == "blank":
            return f"(LEN(TRIM({R}))=0)"
        if op == "notblank":
            return f"(LEN(TRIM({R}))>0)"
        V = self.f(c["value"], row)
        if op in ("=", "<>"):
            return f"({R}{op}{V})"
        if op in ORDER_OPS:                                  # text sorts after numbers in Excel, so only numbers can be "greater than 6000"
            core = f"({R}{op}{V})"
            return f"(ISNUMBER({R})*{core})" if self._numeric_lit(c.get("value")) else core
        return {"contains": f"ISNUMBER(SEARCH({V},{R}))", "notcontains": f"ISERROR(SEARCH({V},{R}))",
                "starts": f"(LEFT({R},LEN({V}))={V})", "ends": f"(RIGHT({R},LEN({V}))={V})"}[op]

    def _where_arr(self, sh, w, row):
        """Array expression that is TRUE for the rows of `sh` meeting the conditions (works without Ctrl+Shift+Enter
        inside SUMPRODUCT / AGGREGATE / LOOKUP / INDEX(...,0))."""
        if not is_group(w):
            return self._leaf_arr(sh, w, row)
        parts = [self._where_arr(sh, i, row) for i in w["items"]]
        if len(parts) == 1:
            return parts[0]
        return "(" + "*".join(parts) + ")" if w.get("join", "AND") == "AND" else "((" + "+".join(parts) + ")>0)"

    @staticmethod
    def _lookup_core(sh, cond, r_rng, search, nth_f):
        if search in ("nth", "nth_last"):
            no = 15 if search == "nth" else 14
            return f"INDEX({r_rng},_xlfn.AGGREGATE({no},6,(ROW({r_rng})-{sh['header_row']})/({cond}),{nth_f}))"
        if search == "last":
            return f"LOOKUP(2,1/({cond}),{r_rng})"
        return f"INDEX({r_rng},MATCH(1,INDEX(({cond})*1,0),0))"

    AGG_RETURNS = ("count", "agg")

    def _case_agg(self, n, c):
        """The totals step equivalent to a case that counts / sums the rows it qualifies (built once, so results are shared)."""
        tag = ("caseagg", id(c))
        if tag not in self._views:
            items = []
            if n.get("key"):
                items.append({"col": n["match_col"], "left": n.get("match_left"), "op": "=", "value": n["key"]})
            w = where_of(c)
            if w is not None:
                if is_group(w) and w.get("join", "AND") == "AND":
                    items.extend(w["items"])
                else:
                    items.append(w)
            fn = "COUNT" if c.get("return_type") == "count" else c.get("agg_fn", "SUM")
            self._views[tag] = {"type": "agg", "fn": fn, "sheet": n.get("sheet"), "value_col": c.get("return_col"),
                                "where": {"join": "AND", "items": items}}
        return self._views[tag]

    def _f_lookup_cases(self, n, sh, row):
        """IF case 1 -> value 1, ELSE IF case 2 -> value 2 ... ELSE fallback, written as nested IFERRORs."""
        nf = self.f(n["not_found"], row) if n.get("not_found") else None
        search = n.get("search", "first")
        if search not in ("first", "last", "nth", "nth_last"):
            raise ValueError(f"Unknown search order '{search}'")
        if search in ("nth", "nth_last") and not n.get("nth"):
            raise ValueError("Say which match number you want (N).")
        nth_f = self.f(n["nth"], row) if search in ("nth", "nth_last") else "1"
        key = None
        if n.get("key"):
            key = f"({self._match_range(sh, n, row)}={self.f(n['key'], row)})"
        cores = []
        for k, c in enumerate(n["cases"], start=1):
            w = where_of(c)
            parts = ([key] if key else []) + ([self._where_arr(sh, w, row)] if w is not None else [])
            if not parts:
                raise ValueError(f"Case {k} needs a value to match on or at least one condition")
            cond = "(" + "*".join(parts) + ")" if len(parts) > 1 else parts[0]
            if c.get("return_type") in self.AGG_RETURNS:
                if c["return_type"] == "agg" and not c.get("return_col"):
                    raise ValueError(f"Case {k}: pick the column to calculate on")
                cores.append(self.f_agg(self._case_agg(n, c), row))
            elif c.get("return_type", "col") == "col":
                if not c.get("return_col"):
                    raise ValueError(f"Case {k}: pick the column to take the value from")
                cores.append(self._lookup_core(sh, cond, self.rng(sh, c["return_col"]), search, nth_f))
            else:
                if not c.get("return_value"):
                    raise ValueError(f"Case {k}: give the value to use")
                cores.append(f"IF(SUMPRODUCT(--({cond}))>={nth_f},{self.f(c['return_value'], row)},NA())")
        always_answers = n["cases"][-1].get("return_type") == "count"       # a count is 0 rather than an error: no fallback needed
        out = f"IFERROR({cores[-1]},{nf})" if nf and not always_answers else cores[-1]
        for core in reversed(cores[:-1]):
            out = f"IFERROR({core},{out})"
        return out

    def _f_lookup_where(self, n, sh, w, row, nf):
        if not n.get("return_col"):
            raise ValueError("Pick the column to return")
        if not n.get("key") and w is None:
            raise ValueError("Give a value to match on, or at least one condition")
        r_rng = self.rng(sh, n["return_col"])
        parts = []
        if n.get("key"):
            parts.append(f"({self._match_range(sh, n, row)}={self.f(n['key'], row)})")
        if w is not None:
            parts.append(self._where_arr(sh, w, row))
        cond = "(" + "*".join(parts) + ")" if len(parts) > 1 else parts[0]
        search = n.get("search", "first")
        if search in ("nth", "nth_last"):
            no = 15 if search == "nth" else 14
            core = (f"INDEX({r_rng},_xlfn.AGGREGATE({no},6,(ROW({r_rng})-{sh['header_row']})/({cond}),"
                    f"{self.f(n['nth'], row)}))")
        elif search == "last":
            core = f"LOOKUP(2,1/({cond}),{r_rng})"
        else:
            core = f"INDEX({r_rng},MATCH(1,INDEX(({cond})*1,0),0))"
        return f"IFERROR({core},{nf})" if nf else core

    # ---- "where" conditions: Python side ---------------------------------------
    def _where_rvs(self, w, i, out=None):
        out = {} if out is None else out
        if is_group(w):
            for it in w["items"]:
                self._where_rvs(it, i, out)
        else:
            out[id(w)] = None if w["op"] in ("blank", "notblank") else self.v(w["value"], i)
        return out

    @staticmethod
    def _arr_match(x, op, rv, guard):
        if op == "blank":
            return to_text(x).strip(" ") == ""
        if op == "notblank":
            return to_text(x).strip(" ") != ""
        if op in ("=", "<>", ">", "<", ">=", "<="):
            if guard and op in ORDER_OPS and not (is_num(x) or isinstance(x, (dt.date, dt.datetime))):
                return False
            return compare(x, rv, op)
        a, b = to_text(x).lower(), to_text(rv).lower()
        return {"contains": b in a, "notcontains": b not in a, "starts": a.startswith(b), "ends": a.endswith(b)}[op]

    def _where_ok(self, sh, w, j, rvs, classic):
        if is_group(w):
            gen = (self._where_ok(sh, x, j, rvs, classic) for x in w["items"])
            return any(gen) if w.get("join", "AND") == "OR" else all(gen)
        x, op, rv = self._leaf_values(sh, w)[j], w["op"], rvs[id(w)]
        if classic:
            return self._crit_match(x, op, rv)
        return self._arr_match(x, op, rv, op in ORDER_OPS and self._numeric_lit(w.get("value")))

    def _where_candidates(self, sh, w, rvs, size):
        """A superset of the rows that can satisfy w: the shortest 'equals' index among the AND-ed conditions."""
        best = None
        for c in where_conjuncts(w):
            rv = rvs.get(id(c))
            if c["op"] == "=" and rv is not None and not is_blank(rv):
                found = self._leaf_index(sh, c).get(self._eq_probe(rv), [])
                if best is None or len(found) < len(best):
                    best = found
        return best if best is not None else range(size)

    def _lookup_pick(self, n, sh, w, i):
        """(row index or None, number of qualifying rows) for the key + conditions w, honouring first / last / Nth."""
        size = len(sh["df"])
        rvs = self._where_rvs(w, i) if w is not None else {}
        if n.get("key"):
            key = self.v(n["key"], i)
            if is_blank(key):                                # Excel: an empty key matches empty cells
                tag = ("blankidx", id(n))
                if tag not in self._maps:
                    self._maps[tag] = [j for j, k in enumerate(self._match_values(sh, n)) if is_blank(k)]
                cand = self._maps[tag]
            else:
                cand = self._match_index(sh, n).get(self._norm(key), [])
        else:
            cand = self._where_candidates(sh, w, rvs, size)
        hits = [j for j in cand if w is None or self._where_ok(sh, w, j, rvs, False)]
        search = n.get("search", "first")
        j = None
        if search in ("nth", "nth_last"):
            k = to_num(self.v(n["nth"], i))
            k = int(k) if k == int(k) else 0
            if 1 <= k <= len(hits):
                j = hits[k - 1] if search == "nth" else hits[-k]
        elif hits:
            j = hits[-1] if search == "last" else hits[0]
        return j, len(hits)

    def _v_lookup_where(self, n, sh, w, i):
        j, n_hits = self._lookup_pick(n, sh, w, i)
        if j is None and search_is_nth(n) and n_hits and not n.get("not_found"):
            raise XlError("#NUM!")
        if j is not None:
            return self.values(sh, n["return_col"])[j]
        if n.get("not_found"):
            return self.v(n["not_found"], i)
        raise XlError(NA)

    def _v_lookup_cases(self, n, sh, i):
        for c in n["cases"]:
            if c.get("return_type") in self.AGG_RETURNS:
                try:
                    return self.v_agg(self._case_agg(n, c), i)
                except XlError:                                  # e.g. AVERAGE of nothing -> like IFERROR, try the next case
                    continue
            j, _ = self._lookup_pick(n, sh, where_of(c), i)
            if j is not None:
                if c.get("return_type", "col") == "col":
                    return self.values(sh, c["return_col"])[j]
                return self.v(c["return_value"], i)
        if n.get("not_found"):
            return self.v(n["not_found"], i)
        raise XlError(NA)

    @staticmethod
    def _norm(k):
        if isinstance(k, (dt.date, dt.datetime)):
            return float(serial(k))
        return float(k) if is_num(k) else str(k).lower()

    @staticmethod
    def _wild(pattern):
        """Excel wildcards: * any run, ? any one character, ~ escapes the next character."""
        out, i = [], 0
        while i < len(pattern):
            ch = pattern[i]
            if ch == "~" and i + 1 < len(pattern):
                out.append(re.escape(pattern[i + 1]))
                i += 1
            elif ch == "*":
                out.append(".*")
            elif ch == "?":
                out.append(".")
            else:
                out.append(re.escape(ch))
            i += 1
        return re.compile("".join(out), re.I | re.S)

    def _match_index(self, sh, n):
        """{normalised key: [row indexes, top to bottom]} for the match column (built once)."""
        tag = ("idx", id(n), sh["sheet_name"], n["match_col"])
        if tag not in self._maps:
            mp = {}
            for j, k in enumerate(self._match_values(sh, n)):
                if not is_blank(k):
                    mp.setdefault(self._norm(k), []).append(j)
            self._maps[tag] = mp
        return self._maps[tag]

    def _all_matches(self, sh, n, key):
        """Every row whose match-column value equals the key, top to bottom."""
        if is_blank(key):
            return []
        return self._match_index(sh, n).get(self._norm(key), [])

    def _closest_struct(self, sh, n):
        """Sorted distinct values per type, each with its row indexes - lets 'next smaller / larger' use bisect."""
        tag = ("closest", id(n), sh["sheet_name"], n["match_col"])
        if tag not in self._maps:
            groups = {}
            for j, k in enumerate(self.values(sh, n["match_col"])):
                if not is_blank(k):
                    kk = _key(k)
                    groups.setdefault(kk[0], {}).setdefault(kk[1], []).append(j)
            self._maps[tag] = (groups, {rank: sorted(d) for rank, d in groups.items()})
        return self._maps[tag]

    def _lookup_index(self, n, sh, key, match, last):
        if is_blank(key):
            return None
        if match == "wildcard" and isinstance(key, str):
            memo = ("wild", id(n), key.lower(), last)
            if memo not in self._maps:
                keys = self.values(sh, n["match_col"])
                rx = self._wild(key)
                order = range(len(keys) - 1, -1, -1) if last else range(len(keys))
                self._maps[memo] = next((j for j in order if isinstance(keys[j], str) and rx.fullmatch(keys[j])), None)
            return self._maps[memo]
        if match in ("exact", "wildcard"):
            hits = self._match_index(sh, n).get(self._norm(key))
            return None if not hits else (hits[-1] if last else hits[0])
        # closest match: the largest value <= key ("smaller") or smallest value >= key ("larger")
        groups, ordered = self._closest_struct(sh, n)
        tk = _key(key)
        vals = ordered.get(tk[0])
        if not vals:
            return None
        if match == "smaller":
            p = bisect_right(vals, tk[1]) - 1
            if p < 0:
                return None
        else:
            p = bisect_left(vals, tk[1])
            if p >= len(vals):
                return None
        idxs = groups[tk[0]][vals[p]]
        return idxs[-1] if last else idxs[0]                     # ties keep the search order

    def v_lookup(self, n, i):
        sh = self.sheet(n.get("sheet"))
        n = self._lookup_view(n)
        if n.get("cases"):
            return self._v_lookup_cases(n, sh, i)
        _, match, last = self._lookup_opts(n)
        w = where_of(n)
        if w is not None or not n.get("key") or n.get("match_left"):
            return self._v_lookup_where(n, sh, w, i)
        search = n.get("search", "first")
        if search in ("nth", "nth_last"):
            hits = self._all_matches(sh, n, self.v(n["key"], i))
            k = to_num(self.v(n["nth"], i))
            k = int(k) if k == int(k) else 0
            ok = 1 <= k <= len(hits)
            j = (hits[k - 1] if search == "nth" else hits[-k]) if ok else None
            if j is None and not n.get("not_found"):
                raise XlError("#NUM!" if hits else NA)
        else:
            j = self._lookup_index(n, sh, self.v(n["key"], i), match, last)
        if j is not None:
            return self.values(sh, n["return_col"])[j]
        if n.get("not_found"):
            return self.v(n["not_found"], i)
        raise XlError(NA)

    # ---- conditional totals (SUMIF / COUNTIF / AVERAGEIF / MAXIFS ...) ----
    def _crit(self, c, row):
        op = c["op"]
        if op == "blank":
            return '""'
        if op == "notblank":
            return '"<>"'
        val = self.f(c["value"], row)
        if op == "=":
            return val
        if op in ("<>", ">", "<", ">=", "<="):
            return f'{q(op)}&{val}'
        return {"contains": f'"*"&{val}&"*"', "notcontains": f'"<>*"&{val}&"*"',
                "starts": f'{val}&"*"', "ends": f'"*"&{val}'}[op]

    def f_agg(self, n, row):
        fn, sh = n["fn"], self.sheet(n.get("sheet"))
        w = where_of(n)
        if fn not in ("SUM", "COUNT", "AVERAGE", "MAX", "MIN"):
            raise ValueError(f"Unknown function {fn}")
        if not where_is_classic(w):                        # OR / nested conditions -> array formulas
            cond = self._where_arr(sh, w, row)
            if fn == "COUNT":
                return f"SUMPRODUCT(--({cond}))"
            if not n.get("value_col"):
                raise ValueError("Pick the column to calculate on")
            vr = self.rng(sh, n["value_col"])
            if fn == "SUM":
                return f"SUMPRODUCT(--({cond}),{vr})"
            if fn == "AVERAGE":
                return f"SUMPRODUCT(--({cond}),{vr})/SUMPRODUCT(--({cond}),--ISNUMBER({vr}))"
            return f"IFERROR(_xlfn.AGGREGATE({14 if fn == 'MAX' else 15},6,{vr}/({cond}),1),0)"
        crits = [] if w is None else ([w] if not is_group(w) else w["items"])
        pairs = [(self.rng(sh, c["col"]), self._crit(c, row)) for c in crits]
        flat = ",".join(f"{r},{c}" for r, c in pairs)
        if fn == "COUNT":
            if not pairs:
                return f"COUNTA({self.rng(sh, n['value_col'])})"
            return f"COUNTIF({pairs[0][0]},{pairs[0][1]})" if len(pairs) == 1 else f"COUNTIFS({flat})"
        if not n.get("value_col"):
            raise ValueError("Pick the column to calculate on")
        vr = self.rng(sh, n["value_col"])
        if not pairs:
            return f"{fn}({vr})"
        if fn in ("MAX", "MIN"):
            return f"_xlfn.{fn}IFS({vr},{flat})"
        base = {"SUM": "SUM", "AVERAGE": "AVERAGE"}[fn]
        return f"{base}IF({pairs[0][0]},{pairs[0][1]},{vr})" if len(pairs) == 1 else f"{base}IFS({vr},{flat})"

    @staticmethod
    def _crit_match(x, op, rv):
        if op == "blank":
            return is_blank(x)
        if op == "notblank":
            return not is_blank(x)
        if is_blank(x):
            return op == "<>"
        if is_num(rv) and isinstance(x, str):
            try:
                x = float(x.replace(",", ""))
            except ValueError:
                pass
        if op in ("=", "<>", ">", "<", ">=", "<="):
            if _key(x)[0] != _key(rv)[0]:          # COUNTIF only compares like with like
                return op == "<>"
            return compare(x, rv, op)
        a, b = to_text(x).lower(), to_text(rv).lower()
        return {"contains": b in a, "notcontains": b not in a,
                "starts": a.startswith(b), "ends": a.endswith(b)}[op]

    @staticmethod
    def _eq_keys(x):
        """Index keys under which a cell can satisfy an 'equals' criterion (a superset - hits are re-checked)."""
        if isinstance(x, bool):
            return [("b", x)]
        if is_num(x):
            return [("n", float(x))]
        if isinstance(x, (dt.date, dt.datetime)):
            return [("n", float(serial(x)))]
        s_ = str(x)
        keys = [("s", s_.lower())]
        try:
            keys.append(("n", float(s_.replace(",", ""))))
        except ValueError:
            pass
        return keys

    @staticmethod
    def _eq_probe(rv):
        if isinstance(rv, bool):
            return ("b", rv)
        if is_num(rv):
            return ("n", float(rv))
        if isinstance(rv, (dt.date, dt.datetime)):
            return ("n", float(serial(rv)))
        return ("s", str(rv).lower())

    def _agg_index(self, sh, col):
        tag = ("aggidx", sh["sheet_name"], col)
        if tag not in self._maps:
            mp = {}
            for j, x in enumerate(self.values(sh, col)):
                if not is_blank(x):
                    for k in self._eq_keys(x):
                        mp.setdefault(k, []).append(j)
            self._maps[tag] = mp
        return self._maps[tag]

    def v_agg(self, n, i):
        fn, sh, w = n["fn"], self.sheet(n.get("sheet")), where_of(n)
        rvs = self._where_rvs(w, i) if w is not None else {}
        memo = ("agg", id(n), tuple(rvs[id(l)] for l in where_leaves(w)))       # rows with the same criteria share one answer
        if memo in self._maps:
            hit = self._maps[memo]
            if isinstance(hit, tuple) and hit and hit[0] == "__err__":
                raise XlError(hit[1])
            return hit
        try:
            result = self._agg_compute(n, sh, fn, w, rvs)
        except XlError as e:
            self._maps[memo] = ("__err__", e.code)
            raise
        self._maps[memo] = result
        return result

    def _agg_compute(self, n, sh, fn, w, rvs):
        size = len(sh["df"])
        classic = where_is_classic(w)
        if w is None:
            hits = None
        else:
            cand = self._where_candidates(sh, w, rvs, size)
            hits = [j for j in cand if self._where_ok(sh, w, j, rvs, classic)]
        first = next(where_leaves(w), None)
        vcol = n.get("value_col") or (first["col"] if first else None)
        vals = self.values(sh, vcol) if vcol else [None] * size
        if fn == "COUNT":
            return len(hits) if hits is not None else sum(1 for x in vals if not is_blank(x))
        pool = range(size) if hits is None else hits
        return self._reduce(fn, [vals[j] for j in pool if is_num(vals[j])])


def expr_columns(node):
    """Column names used by a step tree."""
    out = set()

    def walk(x):
        if isinstance(x, dict):
            if x.get("type") == "col" and x.get("name"):
                out.add(x["name"])
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)

    walk(node)
    return out


def referenced_columns(node, sheet_name):
    """Columns of the CURRENT sheet a tree reads (used to avoid circular refs
    when the result is written back into one of its own inputs)."""
    found = set()

    def walk(x):
        if isinstance(x, dict):
            if x.get("type") == "col" and x.get("name"):
                found.add(x["name"])
            if x.get("type") in ("agg", "lookup") and x.get("sheet") in (None, "", sheet_name):
                found.update([x.get("value_col"), x.get("match_col"), x.get("return_col")])
                found.update(expr_columns(x.get("match_left")))
                for w in [where_of(x)] + [where_of(c) for c in x.get("cases") or []]:
                    for l in where_leaves(w):
                        found.add(l.get("col"))
                        found.update(expr_columns(l.get("left")))
                for c in x.get("cases") or []:
                    found.add(c.get("return_col"))
            for k, v in x.items():
                if k == "match_left" or (k == "left" and "col" in x and "op" in x):
                    continue                                    # these columns belong to the searched sheet
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)

    walk(node)
    found.discard(None)
    return found


def dependency_info(node, sheet_name):
    """Which sheets a tree reads, and whether any step looks at OTHER rows (lookup / totals).
    Row-local trees can be recalculated for just the edited rows."""
    sheets, state = {sheet_name}, {"global": False}

    def walk(x):
        if isinstance(x, dict):
            if x.get("type") in ("agg", "lookup"):
                state["global"] = True
                sheets.add(x.get("sheet") or sheet_name)
            for v in x.values():
                walk(v)
        elif isinstance(x, list):
            for v in x:
                walk(v)

    walk(node)
    return {"global": state["global"], "sheets": sheets}


def op_expr(params, ctx, target_ctx=None, rows=None):
    """rows: optional list of 0-based data-row indexes to evaluate (values come back in that order)."""
    root = params["expr"]
    eng = Engine(ctx, ctx["_all"])
    eng.f(root, ctx["header_row"] + 1)            # validate the whole tree up front

    def formula_fn(row_num):
        return "=" + eng.f(root, row_num)

    out = []
    for i in (range(len(ctx["df"])) if rows is None else rows):
        try:
            val = eng.v(root, i)
        except XlError as e:
            val = e.code
        except ArithmeticError:
            val = VALUE
        out.append(0 if val is None else val)      # a bare reference to a blank cell shows 0
    return formula_fn, out
