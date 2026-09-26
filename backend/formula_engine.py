"""
formula_engine.py
------------------
This is the heart of the system: for every supported Excel function we define
TWO things:

  1. An "excel_formula" generator -> produces the REAL Excel formula string
     (e.g. "=XLOOKUP(A2,...)") that gets written into the .xlsx file so it
     works natively when opened in real Excel/Google Sheets.

  2. A "python_compute" function -> computes the same result using pandas,
     so the web UI can show live results immediately without needing Excel
     to recalculate.

Both must stay logically in sync. Add new functions by adding a new entry
to OPERATIONS.
"""
from __future__ import annotations
import re
import datetime as dt
import pandas as pd
from openpyxl.utils import get_column_letter, column_index_from_string


def col_letter(idx0: int) -> str:
    """0-indexed column number -> Excel column letter."""
    return get_column_letter(idx0 + 1)


def quote_sheet(name: str) -> str:
    """Excel requires single quotes around sheet names that contain spaces/dashes."""
    if re.match(r'^[A-Za-z_][A-Za-z0-9_]*$', name):
        return name
    return f"'{name}'"


def col_range_ref(sheet_name: str, col_idx0: int, header_row: int, last_row: int) -> str:
    letter = col_letter(col_idx0)
    return f"{quote_sheet(sheet_name)}!${letter}${header_row + 1}:${letter}${last_row}"


def cell_ref(sheet_name: str, col_idx0: int, row_num: int, same_sheet: bool) -> str:
    letter = col_letter(col_idx0)
    if same_sheet:
        return f"{letter}{row_num}"
    return f"{quote_sheet(sheet_name)}!{letter}{row_num}"


# ---------------------------------------------------------------------------
# Each builder receives:
#   params: dict from the frontend wizard
#   ctx: dict with sheet metadata needed to build refs
#      { sheet_name, header_row, last_row, columns: {colname: idx0}, df }
#   target_ctx: same shape, for a second sheet (lookups) - may be None
# Returns: (formula_fn(row_num) -> str, values: pd.Series)
# ---------------------------------------------------------------------------

def _get_col_idx(ctx, colname):
    if colname not in ctx["columns"]:
        raise ValueError(f"Column '{colname}' not found in sheet '{ctx['sheet_name']}'")
    return ctx["columns"][colname]


def op_concat(params, ctx, target_ctx=None):
    cols = params["columns"]                    # list of column names, in order
    delimiter = params.get("delimiter", "")      # e.g. " ", ", "
    idxs = [_get_col_idx(ctx, c) for c in cols]

    def formula_fn(row_num):
        parts = [cell_ref(ctx["sheet_name"], i, row_num, True) for i in idxs]
        if delimiter:
            joined = f'&"{delimiter}"&'.join(parts)
        else:
            joined = "&".join(parts)
        return f"={joined}"

    clean = ctx["df"][cols].apply(lambda s: s.map(lambda v: "" if pd.isna(v) else str(v)))
    series = clean.agg(delimiter.join, axis=1) if delimiter else clean.agg("".join, axis=1)
    return formula_fn, series


def op_lookup(params, ctx, target_ctx):
    """Covers XLOOKUP, VLOOKUP, INDEX-MATCH -- same intent, different formula syntax."""
    method = params.get("method", "XLOOKUP")     # XLOOKUP | VLOOKUP | INDEX-MATCH
    lookup_col = params["lookup_column"]           # column in CURRENT sheet to match on
    match_col = params["match_column"]             # column in TARGET sheet to match against
    return_col = params["return_column"]           # column in TARGET sheet to pull value from
    not_found = params.get("not_found_value", '"Not Found"')

    lookup_idx = _get_col_idx(ctx, lookup_col)
    match_idx = _get_col_idx(target_ctx, match_col)
    return_idx = _get_col_idx(target_ctx, return_col)

    t_sheet = target_ctx["sheet_name"]
    t_header = target_ctx["header_row"]
    t_last = target_ctx["last_row"]

    match_range = col_range_ref(t_sheet, match_idx, t_header, t_last)
    return_range = col_range_ref(t_sheet, return_idx, t_header, t_last)

    def formula_fn(row_num):
        lookup_cell = cell_ref(ctx["sheet_name"], lookup_idx, row_num, True)
        if method == "XLOOKUP":
            return f"=_xlfn.XLOOKUP({lookup_cell},{match_range},{return_range},{not_found})"
        elif method == "VLOOKUP":
            # VLOOKUP needs return column INDEX relative to match column, and a contiguous range
            match_letter = col_letter(match_idx)
            return_letter = col_letter(return_idx)
            match_ci = column_index_from_string(match_letter)
            return_ci = column_index_from_string(return_letter)
            if return_ci < match_ci:
                raise ValueError("VLOOKUP requires the return column to be to the RIGHT of "
                                  "the match column. Use XLOOKUP or INDEX-MATCH instead.")
            col_offset = return_ci - match_ci + 1
            full_range = (f"{quote_sheet(t_sheet)}!${match_letter}${t_header+1}:"
                           f"${col_letter(return_idx)}${t_last}")
            return f"=IFERROR(VLOOKUP({lookup_cell},{full_range},{col_offset},FALSE),{not_found})"
        else:  # INDEX-MATCH
            return (f"=IFERROR(INDEX({return_range},MATCH({lookup_cell},{match_range},0)),"
                    f"{not_found})")

    # python compute: build a mapping from target sheet
    tdf = target_ctx["df"]
    mapping = dict(zip(tdf[match_col], tdf[return_col]))
    nf = not_found.strip('"') if isinstance(not_found, str) else not_found
    series = ctx["df"][lookup_col].map(lambda v: mapping.get(v, nf))
    return formula_fn, series


def op_conditional_agg(params, ctx, target_ctx=None):
    """SUMIF / COUNTIF / MAXIF(=MAXIFS) / AVERAGEIF -> returns ONE value repeated,
    or per-row if criteria column is same row's group (common HR/report use case)."""
    func = params["agg_func"]              # SUM | COUNT | MAX | AVERAGE
    value_col = params.get("value_column")
    criteria_col = params["criteria_column"]
    operator = params.get("operator", "=")  # =, >, <, >=, <=, contains
    compare_value = params["compare_value"]

    crit_idx = _get_col_idx(ctx, criteria_col)
    crit_range = col_range_ref(ctx["sheet_name"], crit_idx, ctx["header_row"], ctx["last_row"])

    if operator == "contains":
        crit_expr = f'"*{compare_value}*"'
    elif operator == "=":
        crit_expr = f'"{compare_value}"' if isinstance(compare_value, str) else compare_value
    else:
        crit_expr = f'"{operator}{compare_value}"'

    if func == "COUNT":
        formula = f"=COUNTIF({crit_range},{crit_expr})"
        mask = _apply_operator(ctx["df"][criteria_col], operator, compare_value)
        value = int(mask.sum())
    else:
        value_idx = _get_col_idx(ctx, value_col)
        value_range = col_range_ref(ctx["sheet_name"], value_idx, ctx["header_row"], ctx["last_row"])
        fname = {"SUM": "SUMIF", "MAX": "MAXIFS", "AVERAGE": "AVERAGEIF"}[func]
        if fname == "MAXIFS":
            formula = f"=MAXIFS({value_range},{crit_range},{crit_expr})"
        else:
            formula = f"={fname}({crit_range},{crit_expr},{value_range})"
        mask = _apply_operator(ctx["df"][criteria_col], operator, compare_value)
        subset = ctx["df"].loc[mask, value_col]
        if func == "SUM":
            value = float(subset.sum())
        elif func == "MAX":
            value = float(subset.max()) if len(subset) else 0
        else:
            value = float(subset.mean()) if len(subset) else 0

    def formula_fn(row_num):
        return formula  # same aggregate formula for every row (or put in one summary cell)

    series = pd.Series([value] * len(ctx["df"]))
    return formula_fn, series


def _apply_operator(series, operator, value):
    if operator == "contains":
        return series.astype(str).str.contains(str(value), case=False, na=False)
    try:
        numeric = pd.to_numeric(series, errors="coerce")
        val = float(value)
        if operator == "=":
            return numeric == val
        if operator == ">":
            return numeric > val
        if operator == "<":
            return numeric < val
        if operator == ">=":
            return numeric >= val
        if operator == "<=":
            return numeric <= val
    except (ValueError, TypeError):
        pass
    if operator == "=":
        return series.astype(str) == str(value)
    return pd.Series([False] * len(series))


def op_if(params, ctx, target_ctx=None):
    column = params["condition_column"]
    operator = params.get("operator", "=")
    compare_value = params["compare_value"]
    true_value = params["true_value"]
    false_value = params["false_value"]
    wrap_ifna = params.get("wrap_ifna", False)

    col_idx = _get_col_idx(ctx, column)

    def excel_val(v):
        if isinstance(v, str) and not v.startswith("="):
            return f'"{v}"'
        return v

    def formula_fn(row_num):
        cell = cell_ref(ctx["sheet_name"], col_idx, row_num, True)
        op = ">=" if operator == ">=" else operator
        if operator == "contains":
            cond = f'ISNUMBER(SEARCH("{compare_value}",{cell}))'
        else:
            try:
                cmp_val = float(compare_value)
                cmp_val = int(cmp_val) if cmp_val.is_integer() else cmp_val
            except (TypeError, ValueError):
                cmp_val = f'"{compare_value}"'
            cond = f"{cell}{op}{cmp_val}"
        inner = f"=IF({cond},{excel_val(true_value)},{excel_val(false_value)})"
        if wrap_ifna:
            inner = f"=IFNA({inner[1:]},{excel_val(false_value)})"
        return inner

    mask = _apply_operator(ctx["df"][column], operator, compare_value)
    series = mask.map(lambda b: true_value if b else false_value)
    return formula_fn, series


def op_text_clean(params, ctx, target_ctx=None):
    action = params["action"]  # TRIM | LEFT | RIGHT | MID | SUBSTITUTE | LEN
    column = params["column"]
    col_idx = _get_col_idx(ctx, column)
    n = params.get("num_chars")
    start = params.get("start_pos")
    find_text = params.get("find_text")
    replace_text = params.get("replace_text", "")

    def formula_fn(row_num):
        cell = cell_ref(ctx["sheet_name"], col_idx, row_num, True)
        if action == "TRIM":
            return f"=TRIM({cell})"
        if action == "LEFT":
            return f"=LEFT({cell},{n})"
        if action == "RIGHT":
            return f"=RIGHT({cell},{n})"
        if action == "MID":
            return f"=MID({cell},{start},{n})"
        if action == "LEN":
            return f"=LEN({cell})"
        if action == "SUBSTITUTE":
            return f'=SUBSTITUTE({cell},"{find_text}","{replace_text}")'
        if action == "UPPER":
            return f"=UPPER({cell})"
        if action == "LOWER":
            return f"=LOWER({cell})"
        raise ValueError(f"Unknown text action {action}")

    s = ctx["df"][column].astype(str)
    if action == "TRIM":
        out = s.str.strip().str.replace(r"\s+", " ", regex=True)
    elif action == "LEFT":
        out = s.str[:n]
    elif action == "RIGHT":
        out = s.str[-n:]
    elif action == "MID":
        out = s.str[start - 1:start - 1 + n]
    elif action == "LEN":
        out = s.str.len()
    elif action == "SUBSTITUTE":
        out = s.str.replace(find_text, replace_text, regex=False)
    elif action == "UPPER":
        out = s.str.upper()
    elif action == "LOWER":
        out = s.str.lower()
    else:
        raise ValueError(f"Unknown text action {action}")
    return formula_fn, out


def op_round(params, ctx, target_ctx=None):
    column = params["column"]
    digits = params.get("digits", 0)
    col_idx = _get_col_idx(ctx, column)

    def formula_fn(row_num):
        cell = cell_ref(ctx["sheet_name"], col_idx, row_num, True)
        return f"=ROUND({cell},{digits})"

    series = pd.to_numeric(ctx["df"][column], errors="coerce").round(digits)
    return formula_fn, series


def op_date(params, ctx, target_ctx=None):
    action = params["action"]  # EDATE | TODAY
    column = params.get("column")
    months = params.get("months", 0)

    if action == "TODAY":
        def formula_fn(row_num):
            return "=TODAY()"
        series = pd.Series([dt.date.today()] * len(ctx["df"]))
        return formula_fn, series

    col_idx = _get_col_idx(ctx, column)

    def formula_fn(row_num):
        cell = cell_ref(ctx["sheet_name"], col_idx, row_num, True)
        return f"=EDATE({cell},{months})"

    def add_months(d):
        if pd.isna(d):
            return None
        if not isinstance(d, (dt.date, dt.datetime)):
            return None
        month = d.month - 1 + months
        year = d.year + month // 12
        month = month % 12 + 1
        day = min(d.day, 28)
        return dt.date(year, month, day)

    series = ctx["df"][column].map(add_months)
    return formula_fn, series


OPERATIONS = {
    "CONCAT": op_concat,
    "CONCATENATE": op_concat,
    "LOOKUP": op_lookup,            # method param distinguishes XLOOKUP/VLOOKUP/INDEX-MATCH
    "CONDITIONAL_AGG": op_conditional_agg,   # SUMIF/COUNTIF/MAXIF/AVERAGEIF
    "IF": op_if,
    "TEXT_CLEAN": op_text_clean,     # TRIM/LEFT/RIGHT/MID/SUBSTITUTE/LEN/UPPER/LOWER
    "ROUND": op_round,
    "DATE": op_date,                 # EDATE/TODAY
}
