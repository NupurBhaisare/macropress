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


def _shown(sheet, r, c):
    """A cell's text as a filter matches it: without the spaces before and after it."""
    return sheet.cell_text(r, c).strip(" ")


class ValuesCriterion(Criterion):
    """A list of allowed texts, plus (optionally) blanks: what Excel saves for a ticked list."""

    def __init__(self, values, blank=False):
        self.values = {v.strip(" ").casefold() for v in values}
        self.blank = blank

    def matches(self, sheet, r, c):
        text = _shown(sheet, r, c)
        if text == "":
            return self.blank
        return text.casefold() in self.values

    def describe(self):
        return {"values": sorted(self.values), "blank": self.blank}


class OpCriterion(Criterion):
    """A comparison with one value. Text is matched without regard to case or to spaces before and after it, in the
    cell or in the criterion: "West" finds " west " too."""

    def __init__(self, op, operand):
        self.op, self.operand = op, operand

    def matches(self, sheet, r, c):
        value = sheet.raw_value(r, c)
        blank = value is None or (isinstance(value, str) and value == "")
        op, operand = self.op, self.operand.strip(" ")
        if operand == "" and op in ("=", "<>"):
            return blank == (op == "=")
        number = parse_number_text(operand)
        numbers = number is not None and is_numeric_value(value)
        if op == "=":  # "is": the cell shows this text, so 1250 formatted 1,250.00 is not "1250"
            return wildcard_regex(operand).fullmatch(_shown(sheet, r, c)) is not None
        if op == "<>":  # "is not": by value against a number, by what the cell shows against text
            return float(value) != number if numbers else wildcard_regex(operand).fullmatch(_shown(sheet, r, c)) is None
        if numbers:  # more or less than a number: only a number is, never text that reads as one
            a, b = float(value), number
        elif number is None and isinstance(value, str) and value != "":  # before or after a text: only text is
            a, b = _shown(sheet, r, c).casefold(), operand.casefold()
        else:
            return False
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
    below = af.r2
    while below < MAX_ROW and any(not sheet.is_empty(below + 1, c) for c in range(af.c1, af.c2 + 1)):
        below += 1  # rows of data directly below the range are the filter's too: Excel takes them in
    if below != af.r2:
        app.note("AF-RANGE-EXTENDED", f"{sheet.name}: down to row {below}")
        af.r2 = below
    for r in range(af.r1 + 1, af.r2 + 1):
        visible = all(crit.matches(sheet, r, af.c1 + f - 1) for f, crit in af.criteria.items())
        if visible:
            if r in sheet.hidden_rows and r in sheet.manual_hidden:
                app.note("AF-REFILTER-UNHIDES")
                sheet.manual_hidden.discard(r)
            sheet.hidden_rows.discard(r)
        else:
            sheet.hidden_rows.add(r)
    sheet.changed = True
