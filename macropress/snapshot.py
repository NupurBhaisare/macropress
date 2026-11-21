"""Snapshots of a workbook's state: what press compares before and after the buttons.

A snapshot holds every cell value (with its type), plus the view state: active sheet,
grouped sheets, each sheet's selection, hidden rows and AutoFilter.
"""

from .excel import PENDING, area_text, col_letters
from .formula import display_syntax
from .values import XLError


def _compress(nums):
    nums = sorted(nums)
    out, start, prev = [], None, None
    for x in nums:
        if start is None:
            start = prev = x
        elif x == prev + 1:
            prev = x
        else:
            out.append([start, prev])
            start = prev = x
    if start is not None:
        out.append([start, prev])
    return out


def _typed(value):
    if value is None:
        return None
    if value is PENDING:
        return ["?", "formula not evaluated"]
    if isinstance(value, bool):
        return ["b", value]
    if isinstance(value, XLError):
        return ["e", str(value)]
    if isinstance(value, (int, float)):
        return ["n", float(value)]
    return ["s", str(value)]


def take(book):
    book.recalc_all()
    sheets = {}
    view = {
        "active": book.active.name if book.active else None,
        "selected": [s.name for s in book.selected],
        "sheets": {},
    }
    for s in book.sheets:
        cells, formulas = {}, {}
        for r, row in s.rows.items():
            for c, cell in row.items():
                addr = f"{col_letters(c)}{r}"
                t = _typed(cell.value)
                if t is not None:
                    cells[addr] = t
                if cell.formula is not None:
                    formulas[addr] = display_syntax(cell.formula)
        sheets[s.name] = {"cells": cells, "formulas": formulas}
        af = s.autofilter
        view["sheets"][s.name] = {
            "selection": area_text(*s.sel, absolute=False),
            "active_cell": f"{col_letters(s.active_cell[1])}{s.active_cell[0]}",
            "hidden_rows": _compress(s.hidden_rows),
            "autofilter": None
            if af is None
            else {
                "ref": area_text(*af.rect, absolute=False),
                "criteria": {str(k): v.describe() for k, v in sorted(af.criteria.items())},
            },
        }
    return {"sheets": sheets, "view": view}
