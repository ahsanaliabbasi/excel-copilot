"""
sql_engine.py - use SQL on your sheets, and turn SQL into Excel formulas.

  * Every sheet is a table (use its real name in quotes, or the short name: "Data Set 2 - Vehicles" = data_set_2_vehicles).
    Columns are the sheet's headers.  Sheets are loaded into an in-memory SQLite database on demand.
  * run():        SELECT queries show their result; UPDATE queries are run on a copy and shown as a diff first.
                  Everything else (INSERT, DELETE, DROP, ATTACH ...) is refused, and the database is sandboxed.
  * translate():  the SQL patterns that map onto a live Excel formula become the same "step" trees the Formula
                  Builder uses (COUNTIFS / SUMIFS / MAXIFS / INDEX-MATCH ... with AND/OR conditions):
        SELECT COUNT(*) FROM vehicles v WHERE v.custid = this.custId AND v.year > 2010
        SELECT c.custId, COUNT(v.vehiclenum) AS vehicles
          FROM customers c LEFT JOIN vehicles v ON v.custid = c.custId GROUP BY c.custId
        SELECT v.model FROM vehicles v WHERE v.custid = this.custId AND v.year > 2010 LIMIT 1 OFFSET 1
    ("this" / "base" / "current" mean "the row being filled in".)
"""
from __future__ import annotations

import datetime as dt
import re
import sqlite3
import time

import pandas as pd
import sqlglot
from sqlglot import exp

OUTER_WORDS = {"this", "base", "current", "row", "outer"}
SQLITE_FILE_LIMIT_SECONDS = 30


class SqlError(Exception):
    """Something the user can fix; the message is shown as is."""


def ident(s):
    return '"' + str(s).replace('"', '""') + '"'


def slug(name):
    return re.sub(r"[^0-9a-zA-Z]+", "_", str(name)).strip("_").lower() or "sheet"


def sql_value(v):
    """A sheet value -> something SQLite stores (dates become ISO text so they sort and compare correctly)."""
    if v is None:
        return None
    try:
        if v != v:
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(v, (dt.datetime, pd.Timestamp)):
        return v.date().isoformat() if v.time() == dt.time(0) else v.isoformat(sep=" ")
    if isinstance(v, dt.date):
        return v.isoformat()
    if isinstance(v, bool):
        return int(v)
    if hasattr(v, "item"):
        v = v.item()
    return v if isinstance(v, (int, float, str, bytes)) else str(v)


def from_sql_value(v):
    """A value coming back from SQLite -> what goes in a cell (ISO dates become real dates again)."""
    if isinstance(v, str) and re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
        try:
            return dt.date.fromisoformat(v)
        except ValueError:
            return v
    return v


def _regexp(pattern, s):
    try:
        return 1 if s is not None and re.search(pattern, str(s)) else 0
    except re.error:
        return 0


def _proper(s):
    return None if s is None else str(s).title()


def parse(sql):
    text = (sql or "").strip().rstrip(";").strip()
    if not text:
        raise SqlError("Type a query first.")
    try:
        stmts = [s for s in sqlglot.parse(text, read="sqlite") if s is not None]
    except sqlglot.errors.ParseError as e:
        first = str(e).splitlines()[0]
        raise SqlError(f"That isn't valid SQL: {first}")
    if len(stmts) != 1:
        raise SqlError("Run one statement at a time.")
    return stmts[0]


class Workbench:
    """SQL access to a session's sheets. `get_df(name)` returns a sheet as a DataFrame."""

    def __init__(self, names, get_df):
        self.names, self.get_df = list(names), get_df
        self.slugs, used = {}, set()
        for n in self.names:
            s = base = slug(n)
            k = 2
            while s in used:
                s, k = f"{base}_{k}", k + 1
            used.add(s)
            self.slugs[s] = n
        self.slug_of = {v: k for k, v in self.slugs.items()}
        self.lower = {n.lower(): n for n in self.names}

    # ------------------------------------------------------------- schema
    def resolve_table(self, name):
        if name in self.names:
            return name
        return self.lower.get(name.lower()) or self.slugs.get(name.lower())

    def schema(self):
        out = []
        for n in self.names:
            df = self.get_df(n)
            out.append({"name": n, "alias": self.slug_of[n], "columns": [str(c) for c in df.columns], "rows": len(df),
                        "types": [self._kind(df.iloc[:, i]) for i in range(df.shape[1])]})
        return out

    @staticmethod
    def _kind(series):
        """'num' | 'date' | 'text' - used to suggest sensible examples."""
        if pd.api.types.is_numeric_dtype(series) and not pd.api.types.is_bool_dtype(series):
            return "num"
        if pd.api.types.is_datetime64_any_dtype(series):
            return "date"
        for v in series.head(50).tolist():
            if v is not None and v == v:
                return "date" if isinstance(v, (dt.date, dt.datetime)) else "text"
        return "text"

    def _tables_in(self, tree):
        cte = {c.alias.lower() for c in tree.find_all(exp.CTE)}
        found = []
        for t in tree.find_all(exp.Table):
            if not t.name or t.name.lower() in cte:
                continue
            real = self.resolve_table(t.name)
            if real is None:
                raise SqlError(f"There is no sheet called '{t.name}'. Sheets you can use: " +
                               ", ".join(f'"{n}"' if not re.fullmatch(r"\w+", n) else n for n in self.names))
            if real not in found:
                found.append(real)
        return found

    # ------------------------------------------------------------ database
    def _connect(self, tables):
        conn = sqlite3.connect(":memory:", check_same_thread=False, isolation_level=None)
        conn.create_function("regexp", 2, _regexp)
        conn.create_function("left", 2, lambda s, n: None if s is None else str(s)[:int(n)])
        conn.create_function("right", 2, lambda s, n: None if s is None else (str(s)[-int(n):] if int(n) > 0 else ""))
        conn.create_function("len", 1, lambda s: None if s is None else len(str(s)))
        conn.create_function("proper", 1, _proper)
        for real in tables:
            df = self.get_df(real)
            cols = [str(c) for c in df.columns]
            conn.execute(f"CREATE TABLE {ident(real)} ({', '.join(ident(c) for c in cols)})")
            if len(df) and cols:
                data = [[sql_value(v) for v in df.iloc[:, i].tolist()] for i in range(len(cols))]
                conn.executemany(f"INSERT INTO {ident(real)} VALUES ({','.join('?' * len(cols))})", zip(*data))
            short = self.slug_of[real]
            if short != real.lower() and short.lower() not in {t.lower() for t in tables}:
                conn.execute(f"CREATE VIEW {ident(short)} AS SELECT * FROM {ident(real)}")
        return conn

    @staticmethod
    def _sandbox(conn, allow_update=False):
        allowed = {sqlite3.SQLITE_SELECT, sqlite3.SQLITE_READ, sqlite3.SQLITE_FUNCTION, sqlite3.SQLITE_RECURSIVE}
        if allow_update:
            allowed.add(sqlite3.SQLITE_UPDATE)

        def auth(action, *_):
            return sqlite3.SQLITE_OK if action in allowed else sqlite3.SQLITE_DENY

        conn.set_authorizer(auth)
        deadline = time.time() + SQLITE_FILE_LIMIT_SECONDS
        conn.set_progress_handler(lambda: 1 if time.time() > deadline else 0, 20000)

    @staticmethod
    def _explain(e):
        msg = str(e)
        if "not authorized" in msg:
            return SqlError("That statement isn't allowed here. You can run SELECT queries and UPDATE statements.")
        if "interrupted" in msg:
            return SqlError(f"The query took longer than {SQLITE_FILE_LIMIT_SECONDS} seconds and was stopped.")
        return SqlError(msg[0].upper() + msg[1:] if msg else "The query failed.")

    # ---------------------------------------------------------------- run
    def _outer_qualifier(self, tree):
        """If the query refers to 'the row being filled in' (this.col), which qualifier does it use?"""
        own = {t.alias.lower() for t in tree.find_all(exp.Table) if t.alias} | {t.name.lower() for t in tree.find_all(exp.Table)}
        used = {c.table.lower() for c in tree.find_all(exp.Column) if c.table and c.table.lower() in OUTER_WORDS and c.table.lower() not in own}
        if len(used) > 1:
            raise SqlError("Use just one of this / base / current for the row being filled in.")
        return next(iter(used), None)

    def run(self, sql, base=None, limit=1000):
        tree = parse(sql)
        if isinstance(tree, exp.Update):
            return self.run_update(sql, tree)
        if not isinstance(tree, (exp.Select, exp.Union, exp.With, exp.Subquery)) and not tree.find(exp.Select):
            raise SqlError("You can run SELECT queries and UPDATE statements. Other statements (INSERT, DELETE, DROP ...) aren't supported.")
        if isinstance(tree, (exp.Insert, exp.Delete, exp.Create, exp.Drop, exp.Alter, exp.Command)):
            raise SqlError("You can run SELECT queries and UPDATE statements. Other statements (INSERT, DELETE, DROP ...) aren't supported.")
        tables = self._tables_in(tree)
        text, wrapped = (sql or "").strip().rstrip(";"), False
        outer = self._outer_qualifier(tree)
        if outer:                                             # "one value per row of the current sheet"
            if not base:
                raise SqlError("Open the sheet you want to fill in first, then use this.<column> in the query.")
            if base not in tables:
                tables.append(base)
            if len(getattr(tree, "expressions", [None])) != 1:
                raise SqlError("A query that uses this.<column> must return exactly one value per row (one column).")
            text = f'SELECT {ident(outer)}.rowid AS row_no, (\n{text}\n) AS result FROM {ident(base)} AS {ident(outer)}'
            wrapped = True
        conn = self._connect(tables)
        try:
            self._sandbox(conn)
            t0 = time.time()
            try:
                cur = conn.execute(text)
                cols = [d[0] for d in cur.description]
                rows = cur.fetchmany(limit + 1) if limit else cur.fetchall()
            except sqlite3.Error as e:
                raise self._explain(e)
            truncated = bool(limit) and len(rows) > limit
            total = len(rows)
            if truncated:
                rows = rows[:limit]
                try:
                    total = conn.execute(f"SELECT COUNT(*) FROM (\n{text}\n)").fetchone()[0]
                except sqlite3.Error:
                    total = None
            return {"kind": "select", "columns": cols, "rows": [list(r) for r in rows], "total": total,
                    "truncated": truncated, "ms": int((time.time() - t0) * 1000), "per_row": wrapped, "tables": tables}
        finally:
            conn.close()

    def run_update(self, sql, tree=None):
        """Run an UPDATE on a copy of the data and report exactly which cells would change."""
        tree = tree or parse(sql)
        if not isinstance(tree, exp.Update):
            raise SqlError("That isn't an UPDATE statement.")
        target = tree.this
        real = self.resolve_table(target.name) if isinstance(target, exp.Table) else None
        if real is None:
            raise SqlError("UPDATE needs one of your sheets, e.g. UPDATE \"Data Set 1\" SET state = UPPER(state) WHERE ...")
        if isinstance(target, exp.Table):                     # write to the real table, not a short-name view
            target.set("this", exp.to_identifier(real, quoted=True))
        others = [t for t in self._tables_in(tree) if t != real]
        conn = self._connect([real] + others)
        try:
            cols = [str(c) for c in self.get_df(real).columns]
            before = conn.execute(f"SELECT rowid, * FROM {ident(real)} ORDER BY rowid").fetchall()
            self._sandbox(conn, allow_update=True)
            t0 = time.time()
            try:
                cur = conn.execute(tree.sql(dialect="sqlite"))
                affected = cur.rowcount
            except sqlite3.Error as e:
                raise self._explain(e)
            conn.set_authorizer(None)
            after = conn.execute(f"SELECT rowid, * FROM {ident(real)} ORDER BY rowid").fetchall()
        finally:
            conn.close()
        changes = []
        for b, a in zip(before, after):
            if b != a:
                for j, (x, y) in enumerate(zip(b[1:], a[1:])):
                    if x != y:
                        changes.append({"row": b[0] - 1, "col": j, "column": cols[j], "before": x, "after": y})
        return {"kind": "update", "sheet": real, "rows_matched": affected, "cells_changed": len(changes),
                "rows_changed": len({c["row"] for c in changes}), "changes": changes, "columns": cols,
                "ms": int((time.time() - t0) * 1000)}

    # ---------------------------------------------------------- translate
    def translate(self, base, sql):
        """SQL -> {"outputs": [{"name", "node"}], "notes": [...]}  (raises SqlError with the reason when it can't)."""
        tree = parse(sql)
        if not isinstance(tree, exp.Select):
            raise SqlError("Only a single SELECT can become a formula.")
        for arg, label in (("having", "HAVING"), ("distinct", "DISTINCT"), ("with_", "WITH (CTEs)"), ("with", "WITH (CTEs)"),
                           ("qualify", "QUALIFY"), ("windows", "window functions")):
            if tree.args.get(arg):
                raise SqlError(f"{label} can't be turned into a formula. You can still run it and fill the result in as values.")
        if tree.find(exp.Window) or tree.find(exp.Union) or tree.find(exp.Subquery):
            raise SqlError("Sub-queries, UNION and window functions can't be turned into a formula. You can still run it and fill the result in as values.")
        frm = tree.args.get("from_") or tree.args.get("from")
        joins = tree.args.get("joins") or []
        if frm is None:
            raise SqlError("Add a FROM clause.")
        tables = [frm.this] + [j.this for j in joins]
        if not all(isinstance(t, exp.Table) for t in tables):
            raise SqlError("Sub-queries in FROM can't be turned into a formula.")
        if len(tables) > 2:
            raise SqlError("A formula can look at one other sheet at a time. Join just two sheets (or run it and fill the result in as values).")
        reals = []
        for t in tables:
            r = self.resolve_table(t.name)
            if r is None:
                raise SqlError(f"There is no sheet called '{t.name}'.")
            reals.append(r)
        notes = []
        cols_of = lambda real: {str(c).lower(): str(c) for c in self.get_df(real).columns}

        def ids_of(t, real):
            return ({t.alias.lower()} if t.alias else set()) | {real.lower(), slug(real)}

        if len(tables) == 1:                                   # correlated query: FROM other WHERE other.k = this.k
            form_a, tgt_t, tgt_real = True, tables[0], reals[0]
            base_t, base_real = None, base
            tgt_ids = ids_of(tgt_t, tgt_real)
            base_ids = set(OUTER_WORDS)
            if base != tgt_real:
                base_ids |= {base.lower(), slug(base)}
                base_ids -= tgt_ids
        else:
            form_a = False
            if reals[0] == base:
                base_t, tgt_t = tables[0], tables[1]
            elif reals[1] == base:
                base_t, tgt_t = tables[1], tables[0]
            else:
                raise SqlError(f"One of the joined sheets has to be the sheet you're filling in ('{base}').")
            base_real, tgt_real = base, self.resolve_table(tgt_t.name)
            if reals[0] == reals[1]:
                if not (base_t.alias and tgt_t.alias):
                    raise SqlError("Give the two copies of the sheet different aliases (e.g. FROM Sheet a JOIN Sheet b ...).")
                base_ids, tgt_ids = {base_t.alias.lower()}, {tgt_t.alias.lower()}
            else:
                base_ids, tgt_ids = ids_of(base_t, base_real), ids_of(tgt_t, tgt_real)
            side = joins[0].args.get("side")
            kind = joins[0].args.get("kind")
            if side in ("RIGHT", "FULL") or kind in ("CROSS",):
                raise SqlError("Only INNER and LEFT joins can become formulas.")
            notes.append("A formula is worked out for every row of the sheet, so rows with no match get 0 / empty "
                         "instead of being left out (as an INNER JOIN would).")
        bcols, tcols = cols_of(base_real), cols_of(tgt_real)

        def side_of(col):
            q, name = (col.table or "").lower(), col.name
            if q:
                if q in tgt_ids and q not in base_ids:
                    side = "t"
                elif q in base_ids and q not in tgt_ids:
                    side = "b"
                else:
                    raise SqlError(f"'{col.table}' could mean either sheet - give the sheets different aliases.")
            else:
                in_t, in_b = name.lower() in tcols, name.lower() in bcols
                if form_a:
                    side = "t" if in_t else "b" if in_b else None
                else:
                    if in_t and in_b:
                        raise SqlError(f"Column '{name}' is in both sheets - write it as alias.{name}.")
                    side = "t" if in_t else "b" if in_b else None
                if side is None:
                    raise SqlError(f"Neither sheet has a column called '{name}'.")
            table = tcols if side == "t" else bcols
            if name.lower() not in table:
                raise SqlError(f"The sheet '{tgt_real if side == 't' else base_real}' has no column '{name}'. "
                               f"Its columns: {', '.join(table.values())}")
            return side, table[name.lower()]

        def unwrap(e):
            while isinstance(e, exp.Paren):
                e = e.this
            return e

        def conjuncts(e):
            e = unwrap(e)
            if isinstance(e, exp.And):
                return conjuncts(e.left) + conjuncts(e.right)
            return [e]

        def disjuncts(e):
            e = unwrap(e)
            if isinstance(e, exp.Or):
                return disjuncts(e.left) + disjuncts(e.right)
            return [e]

        def value_node(e):
            e = unwrap(e)
            if isinstance(e, exp.Literal):
                return {"type": "lit", "value": str(e.this)}
            if isinstance(e, exp.Neg) and isinstance(unwrap(e.this), exp.Literal) and not unwrap(e.this).is_string:
                return {"type": "lit", "value": "-" + str(unwrap(e.this).this)}
            if isinstance(e, exp.Column):
                side, name = side_of(e)
                if side == "b":
                    return {"type": "col", "name": name}
                raise SqlError("Comparing two columns of the same sheet can't become a formula.")
            if isinstance(e, exp.Null):
                raise SqlError("Use IS NULL / IS NOT NULL to test for empty values.")
            raise SqlError(f"Only plain columns and fixed values can be compared in a condition: {e.sql()}")

        FLIP = {">": "<", "<": ">", ">=": "<=", "<=": ">=", "=": "=", "<>": "<>"}
        NEG = {">": "<=", "<": ">=", ">=": "<", "<=": ">", "=": "<>", "<>": "="}
        CMP = {exp.EQ: "=", exp.NEQ: "<>", exp.GT: ">", exp.GTE: ">=", exp.LT: "<", exp.LTE: "<="}

        def col_of(e, want):
            e = unwrap(e)
            if isinstance(e, exp.Column):
                side, name = side_of(e)
                if side == want:
                    return name
            return None

        def tgt_leaf(colname, op, value=None):
            return {"col": colname, "op": op, "value": value}

        def group(join, items):
            return items[0] if len(items) == 1 else {"join": join, "items": items}

        def to_where(e, negate=False):
            """A condition on the OTHER sheet -> leaf / group."""
            e = unwrap(e)
            if isinstance(e, exp.And):
                parts = [to_where(x, negate) for x in conjuncts(e)]
                return group("OR" if negate else "AND", parts)
            if isinstance(e, exp.Or):
                parts = [to_where(x, negate) for x in disjuncts(e)]
                return group("AND" if negate else "OR", parts)
            if isinstance(e, exp.Not):
                return to_where(e.this, not negate)
            for klass, op in CMP.items():
                if isinstance(e, klass):
                    l, r = unwrap(e.left), unwrap(e.right)
                    if col_of(l, "t") and not (isinstance(r, exp.Column) and side_of(r)[0] == "t"):
                        name, other, o = col_of(l, "t"), r, op
                    elif col_of(r, "t") and not (isinstance(l, exp.Column) and side_of(l)[0] == "t"):
                        name, other, o = col_of(r, "t"), l, FLIP[op]
                    else:
                        raise SqlError(f"Each condition needs one column of '{tgt_real}' compared with a value or a column of "
                                       f"'{base_real}': {e.sql()}")
                    return tgt_leaf(name, NEG[o] if negate else o, value_node(other))
            if isinstance(e, exp.In):
                name = col_of(e.this, "t")
                if not name or e.args.get("query") or not e.expressions:
                    raise SqlError(f"IN needs a column of '{tgt_real}' and a list of values: {e.sql()}")
                leaves = [tgt_leaf(name, "<>" if negate else "=", value_node(v)) for v in e.expressions]
                return group("AND" if negate else "OR", leaves)
            if isinstance(e, exp.Between):
                name = col_of(e.this, "t")
                if not name:
                    raise SqlError(f"BETWEEN needs a column of '{tgt_real}': {e.sql()}")
                lo, hi = value_node(e.args["low"]), value_node(e.args["high"])
                if negate:
                    return {"join": "OR", "items": [tgt_leaf(name, "<", lo), tgt_leaf(name, ">", hi)]}
                return {"join": "AND", "items": [tgt_leaf(name, ">=", lo), tgt_leaf(name, "<=", hi)]}
            if isinstance(e, exp.Is):
                name = col_of(e.this, "t")
                if not name or not isinstance(unwrap(e.expression), exp.Null):
                    raise SqlError(f"Only IS NULL / IS NOT NULL is supported after IS: {e.sql()}")
                return tgt_leaf(name, "notblank" if negate else "blank")
            if isinstance(e, (exp.Like, exp.ILike)):
                name, pat = col_of(e.this, "t"), unwrap(e.expression)
                if not name or not isinstance(pat, exp.Literal) or not pat.is_string:
                    raise SqlError(f"LIKE needs a column of '{tgt_real}' and a text pattern: {e.sql()}")
                neg = negate != bool(e.args.get("negate"))
                p = str(pat.this)
                core = p.strip("%")
                if "%" in core or "_" in core:
                    raise SqlError(f"Only patterns like 'abc%', '%abc' or '%abc%' can become formulas: {e.sql()}")
                starts, ends = p.startswith("%"), p.endswith("%")
                op = "contains" if starts and ends else "ends" if starts else "starts" if ends else "="
                if neg:
                    if op == "contains":
                        op = "notcontains"
                    elif op == "=":
                        op = "<>"
                    else:
                        raise SqlError(f"NOT LIKE with a starts-with / ends-with pattern can't become a formula: {e.sql()}")
                return tgt_leaf(name, op, {"type": "lit", "value": core})
            raise SqlError(f"This condition can't become a formula: {e.sql()}")

        def base_cond(e):
            """Conditions that only involve the row being filled in -> (join, [ {left, op, right} ])."""
            def one(x, negate=False):
                x = unwrap(x)
                if isinstance(x, exp.Not):
                    return one(x.this, not negate)
                for klass, op in CMP.items():
                    if isinstance(x, klass):
                        l, r = unwrap(x.left), unwrap(x.right)
                        if isinstance(l, exp.Column) and not isinstance(r, exp.Column):
                            name, val, o = side_of(l)[1], value_node_lit(r), op
                        elif isinstance(r, exp.Column) and not isinstance(l, exp.Column):
                            name, val, o = side_of(r)[1], value_node_lit(l), FLIP[op]
                        else:
                            raise SqlError(f"Unsupported condition on the current sheet: {x.sql()}")
                        return [{"left": {"type": "col", "name": name}, "op": NEG[o] if negate else o, "right": val}]
                if isinstance(x, exp.In) and isinstance(unwrap(x.this), exp.Column) and x.expressions:
                    name = side_of(unwrap(x.this))[1]
                    return [{"left": {"type": "col", "name": name}, "op": "<>" if negate else "=", "right": value_node_lit(v)} for v in x.expressions]
                if isinstance(x, exp.Is) and isinstance(unwrap(x.this), exp.Column):
                    name = side_of(unwrap(x.this))[1]
                    return [{"left": {"type": "col", "name": name}, "op": "notblank" if negate else "blank", "right": {"type": "lit", "value": ""}}]
                raise SqlError(f"Unsupported condition on the current sheet: {x.sql()}")

            def value_node_lit(v):
                v = unwrap(v)
                if isinstance(v, exp.Literal):
                    return {"type": "lit", "value": str(v.this)}
                if isinstance(v, exp.Neg) and isinstance(unwrap(v.this), exp.Literal):
                    return {"type": "lit", "value": "-" + str(unwrap(v.this).this)}
                raise SqlError(f"Only fixed values can be compared with the current sheet's columns here: {v.sql()}")

            e = unwrap(e)
            if isinstance(e, exp.Or):
                items = [i for d in disjuncts(e) for i in one(d)]
                return "OR", items
            return ("OR" if isinstance(unwrap(e), exp.In) else "AND"), one(e)

        # ---- split the conditions: ON + WHERE, at the top level
        raw = []
        for j in joins:
            if j.args.get("on") is not None:
                raw += conjuncts(j.args["on"])
        if tree.args.get("where") is not None:
            raw += conjuncts(tree.args["where"].this)
        tgt_items, base_conds = [], []
        for c in raw:
            sides = {side_of(col)[0] for col in c.find_all(exp.Column)}
            if not sides:
                continue
            if sides == {"b"}:
                base_conds.append(base_cond(c))
            else:
                tgt_items.append(to_where(c))
        if not form_a and not tgt_items:
            raise SqlError(f"Say how the sheets are linked, e.g. ON {tgt_t.alias or tgt_real}.custid = {base_t.alias or base_real}.custid")
        if tree.args.get("order") and tree.args.get("limit"):
            raise SqlError("ORDER BY together with LIMIT can't become a formula (a formula takes rows in sheet order). "
                           "Use the first / last / Nth option in the Formula Builder instead.")
        if tree.args.get("order"):
            notes.append("ORDER BY was ignored - a formula fills one value per row of the sheet.")
        if tree.args.get("group"):
            notes.append("GROUP BY was ignored - the value is worked out for each row of the sheet.")
        notes.append("Excel compares text without regard to upper/lower case.")

        limit = tree.args.get("limit")
        offset = tree.args.get("offset")
        nth = 1
        if limit is not None:
            try:
                lim = int(limit.expression.this)
                nth = (int(offset.expression.this) if offset is not None else 0) + 1
            except (ValueError, AttributeError):
                raise SqlError("LIMIT must be a number.")
            if lim != 1:
                raise SqlError("Only LIMIT 1 (optionally with OFFSET) can become a formula - a formula returns one value.")

        def wrap_base(node):
            for join, items in reversed(base_conds):
                node = {"type": "if", "cond": {"join": join, "items": items}, "then": node, "else": {"type": "blank"}}
            return node

        AGG = {exp.Count: "COUNT", exp.Sum: "SUM", exp.Avg: "AVERAGE", exp.Min: "MIN", exp.Max: "MAX"}
        outputs, taken = [], set()

        def unique(name):
            base_name, k = name, 2
            while name.lower() in taken:
                name, k = f"{base_name}_{k}", k + 1
            taken.add(name.lower())
            return name

        for item in tree.expressions:
            alias = item.alias if isinstance(item, exp.Alias) else None
            inner = item.this if isinstance(item, exp.Alias) else item
            if isinstance(inner, exp.Star) or (isinstance(inner, exp.Column) and (isinstance(inner.this, exp.Star) or side_of(inner)[0] == "b")):
                continue                                                # columns of the sheet itself are already there
            if type(inner) in AGG:
                fn, arg = AGG[type(inner)], inner.this
                if isinstance(arg, exp.Distinct):
                    raise SqlError("DISTINCT inside an aggregate can't become a formula.")
                items, value_col = list(tgt_items), None
                if isinstance(arg, exp.Star) or (isinstance(arg, exp.Literal) and fn == "COUNT"):
                    if fn != "COUNT":
                        raise SqlError(f"{fn}(*) isn't valid - name a column.")
                else:
                    value_col = col_of(arg, "t")
                    if not value_col:
                        raise SqlError(f"{fn}() needs a column of '{tgt_real}', e.g. {fn}({tgt_t.alias or tgt_real}.column): {inner.sql()}")
                    if fn == "COUNT":
                        items.append(tgt_leaf(value_col, "notblank"))
                if not items:
                    raise SqlError("COUNT(*) with no condition is just the number of rows - add a WHERE / ON condition.")
                node = {"type": "agg", "fn": fn, "sheet": tgt_real, "value_col": value_col,
                        "where": {"join": "AND", "items": items}}
                default = f"{fn.lower()}_{slug(value_col) if value_col else 'rows'}"
            elif isinstance(inner, exp.Column) and side_of(inner)[0] == "t":
                ret = side_of(inner)[1]
                search = "nth" if nth > 1 else "first"
                eqs = [i for i in tgt_items if not ("items" in i) and i["op"] == "=" and i["value"] and i["value"]["type"] == "col"]
                node = {"type": "lookup", "sheet": tgt_real, "return_col": ret, "search": search, "not_found": {"type": "blank"}}
                if nth > 1:
                    node["nth"] = {"type": "lit", "value": str(nth)}
                if len(tgt_items) == 1 and len(eqs) == 1:                # plain key lookup -> a normal XLOOKUP
                    node.update(method="XLOOKUP", match="exact", key=eqs[0]["value"], match_col=eqs[0]["col"])
                else:
                    node["where"] = {"join": "AND", "items": list(tgt_items)}
                default = ret
            else:
                raise SqlError("Each selected item has to be COUNT / SUM / AVG / MIN / MAX of one column, or one plain column "
                               f"of '{tgt_real}' - formulas can't be built from: {inner.sql()}")
            outputs.append({"name": unique(alias or default), "node": wrap_base(node)})
        if not outputs:
            raise SqlError(f"Select something to calculate, e.g. COUNT(*) or a column of '{tgt_real}'.")
        return {"outputs": outputs, "notes": notes, "base": base_real, "target": tgt_real}
