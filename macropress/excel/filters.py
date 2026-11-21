"""AutoFilter: the criteria a column can carry, and applying them to a sheet's rows."""

import re

from ..errors import VBANotImplemented
from ..values import (
    default_value,
    format_double,
    is_numeric_value,
    parse_number_text,
    to_str,
)
from .constants import MAX_ROW

# ---------------------------------------------------------------------------
# AutoFilter criteria
# ---------------------------------------------------------------------------


def wildcard_regex(pattern):
    out, i = [], 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "~" and i + 1 < len(pattern) and pattern[i + 1] in "*?~":
            out.append(re.escape(pattern[i + 1]))
            i += 2
            continue
        out.append(".*" if ch == "*" else "." if ch == "?" else re.escape(ch))
        i += 1
    return re.compile("".join(out), re.IGNORECASE | re.DOTALL)


class Criterion:
    def matches(self, sheet, r, c):
        raise NotImplementedError


class ValuesCriterion(Criterion):
    """A list of allowed texts, plus (optionally) blanks: what Excel saves for a ticked list."""

    def __init__(self, values, blank=False):
        self.values = {v.casefold() for v in values}
        self.blank = blank

    def matches(self, sheet, r, c):
        text = sheet.cell_text(r, c)
        if text == "":
            return self.blank
        return text.casefold() in self.values

    def describe(self):
        return {"values": sorted(self.values), "blank": self.blank}


class OpCriterion(Criterion):
    def __init__(self, op, operand):
        self.op, self.operand = op, operand

    def matches(self, sheet, r, c):
        text = sheet.cell_text(r, c)
        op, operand = self.op, self.operand
        if op in ("=", "<>"):
            if operand == "":
                hit = text == ""
            else:
                value = sheet.raw_value(r, c)
                num = parse_number_text(operand)
                if num is not None and is_numeric_value(value):
                    hit = float(value) == float(num)
                else:
                    hit = wildcard_regex(operand).fullmatch(text) is not None
            return hit if op == "=" else not hit
        if text == "":
            return False
        value = sheet.raw_value(r, c)
        num = parse_number_text(operand)
        if num is not None and is_numeric_value(value):
            a, b = float(value), float(num)
        else:
            a, b = text.casefold(), operand.casefold()
        return {">": a > b, "<": a < b, ">=": a >= b, "<=": a <= b}[op]

    def describe(self):
        return {"op": self.op, "operand": self.operand}


class BoolCriterion(Criterion):
    def __init__(self, kind, a, b):
        self.kind, self.a, self.b = kind, a, b

    def matches(self, sheet, r, c):
        x, y = self.a.matches(sheet, r, c), self.b.matches(sheet, r, c)
        return (x and y) if self.kind == "and" else (x or y)

    def describe(self):
        return {self.kind: [self.a.describe(), self.b.describe()]}


class UnsupportedCriterion(Criterion):
    def __init__(self, kind):
        self.kind = kind

    def matches(self, sheet, r, c):
        raise VBANotImplemented(f"AutoFilter criterion of type {self.kind}")

    def describe(self):
        return {"unsupported": self.kind}


OPS = (">=", "<=", "<>", ">", "<", "=")


def criterion_from_text(value):
    value = default_value(value)
    if is_numeric_value(value):
        return OpCriterion("=", format_double(float(value)))
    text = to_str(value)
    for op in OPS:
        if text.startswith(op):
            return OpCriterion(op, text[len(op) :])
    return OpCriterion("=", text)


class AutoFilterState:
    def __init__(self, r1, c1, r2, c2, criteria=None):
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2
        self.criteria = criteria or {}  # 1-based field -> Criterion

    @property
    def rect(self):
        return (self.r1, self.c1, self.r2, self.c2)


def apply_filter(sheet):
    af = sheet.autofilter
    app = sheet.book.app
    app.note("AF-MATCH-TEXT")
    for r in range(af.r1 + 1, af.r2 + 1):
        visible = all(crit.matches(sheet, r, af.c1 + f - 1) for f, crit in af.criteria.items())
        if visible:
            if r in sheet.hidden_rows and r in sheet.manual_hidden:
                app.note("AF-REFILTER-UNHIDES")
                sheet.manual_hidden.discard(r)
            sheet.hidden_rows.discard(r)
        else:
            sheet.hidden_rows.add(r)
    below = af.r2 + 1
    if below <= MAX_ROW and any(not sheet.is_empty(below, c) for c in range(af.c1, af.c2 + 1)):
        app.note("AF-RANGE-NOT-EXTENDED", f"{sheet.name}: data continues in row {below}")
    sheet.changed = True
