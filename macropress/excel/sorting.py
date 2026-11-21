"""Sorting a range the way Excel does: Range.Sort and the Worksheet.Sort object."""

import functools

from .. import formula as fx
from ..errors import VBANotImplemented, VBARuntimeError
from ..values import (
    NOTHING,
    VBAObject,
    default_value,
    to_long,
)
from .addresses import area_text, shift_formula
from .constants import PENDING
from .range import Range


class SortKey:
    def __init__(self, key, order=None, sort_on=0, data_option=None):
        if not isinstance(key, Range):
            raise VBARuntimeError(1004, "The sort reference is not valid")
        self.key = key
        self.descending = order is not None and to_long(order) == 2
        if to_long(sort_on or 0) != 0:
            raise VBANotImplemented("sorting on cell or font colour, or icons")
        if data_option is not None and to_long(data_option) != 0:
            raise VBANotImplemented("sorting text as numbers (xlSortTextAsNumbers)")


class SortFields(VBAObject):
    vba_type_name = "SortFields"

    def __init__(self, sort):
        self.sort = sort

    def m_clear(self):
        self.sort.fields = []

    def m_add(self, key=None, sorton=None, order=None, customorder=None, dataoption=None):
        if customorder is not None:
            raise VBANotImplemented("a sort field with a custom order")
        field = SortKey(key, order, sorton, dataoption)
        self.sort.fields.append(field)
        return field

    def m_add2(self, key=None, sorton=None, order=None, customorder=None, dataoption=None, subfield=None):
        if subfield is not None:
            raise VBANotImplemented("a sort field with SubField")
        return self.m_add(key, sorton, order, customorder, dataoption)

    def get_count(self):
        return len(self.sort.fields)


class SortObject(VBAObject):
    """Worksheet.Sort: fields, range and options, kept on the sheet until changed (as in Excel)."""

    vba_type_name = "Sort"

    def __init__(self, sheet):
        self.sheet = sheet
        self.fields = []
        self.rng = None
        self.header = 2  # xlNo
        self.matchcase = False
        self.orientation = 1  # xlTopToBottom
        self.sortmethod = 1  # xlPinYin

    def get_sortfields(self):
        return SortFields(self)

    def m_setrange(self, rng):
        if not isinstance(rng, Range):
            raise VBARuntimeError(1004, "The sort reference is not valid")
        self.rng = rng

    def get_rng(self):
        return self.rng if self.rng is not None else NOTHING

    def get_header(self):
        return self.header

    def let_header(self, v):
        self.header = to_long(v)

    def get_matchcase(self):
        return self.matchcase

    def let_matchcase(self, v):
        self.matchcase = bool(default_value(v))

    def get_orientation(self):
        return self.orientation

    def let_orientation(self, v):
        self.orientation = to_long(v)

    def get_sortmethod(self):
        return self.sortmethod

    def let_sortmethod(self, v):
        self.sortmethod = to_long(v)

    def m_apply(self):
        if self.rng is None:
            raise VBARuntimeError(1004, "The sort reference is not valid")
        sort_range(self.rng.sheet, self.rng.rect, self.fields, self.header, self.matchcase, self.orientation)
        return True


def _sort_key(v, descending, match_case):
    """Excel's ascending order: numbers, text, logical, errors; blanks are always last, whichever the order."""
    if v is None or (isinstance(v, str) and v == ""):
        return (1, 0)
    rank = fx.type_rank(v)
    if rank == 0:
        k = (rank, float(v))
    elif rank == 1:
        k = (rank, fx.text_order_key(v, match_case))
    elif rank == 2:
        k = (rank, bool(v))
    else:
        k = (rank, str(v))
    return (0, _Desc(k) if descending else k)


@functools.total_ordering
class _Desc:
    __slots__ = ("k",)

    def __init__(self, k):
        self.k = k

    def __lt__(self, other):
        return other.k < self.k

    def __eq__(self, other):
        return self.k == other.k


def sort_range(sheet, rect, keys, header, match_case, orientation):
    """Sort the rows of `rect` (top to bottom) by `keys`, as Excel's Sort does."""
    app = sheet.book.app
    if orientation is not None and to_long(orientation) != 1:
        raise VBANotImplemented("sorting left to right")
    h = to_long(header) if header is not None else 2
    if h == 0:
        raise VBANotImplemented("Header:=xlGuess (Excel guesses whether the first row is a header)")
    if not keys:
        raise VBARuntimeError(1004, "The sort reference is not valid")
    r1, c1, r2, c2 = rect
    if h == 1:
        r1 += 1
    if r1 > r2:
        return
    for k in keys:
        if k.key.c1 < c1 or k.key.c1 > c2:
            raise VBARuntimeError(1004, "The sort reference is not valid")
    rows = list(range(r1, r2 + 1))
    hidden = [r for r in rows if r in sheet.hidden_rows]
    if hidden:  # hidden rows keep their place; only the visible rows are sorted among themselves
        app.note("SORT-HIDDEN-ROWS", f"{sheet.name}: {len(hidden)} hidden rows in {area_text(*rect)}")
        rows = [r for r in rows if r not in sheet.hidden_rows]
    app.note("SORT-TEXT-ORDER")
    match_case = bool(default_value(match_case)) if match_case is not None else False
    sheet.fresh()
    records = []
    for r in rows:
        row = sheet.rows.get(r, {})
        cells = {c: cell for c, cell in row.items() if c1 <= c <= c2}
        for k in keys:
            if k.key.c1 in cells and cells[k.key.c1].value is PENDING:
                sheet.raw_value(r, k.key.c1)  # raises: value unknown
        key = tuple(
            _sort_key(cells[k.key.c1].value if k.key.c1 in cells else None, k.descending, match_case) for k in keys
        )
        records.append((key, r, cells))
    ordered = sorted(records, key=lambda rec: rec[0])  # stable, like Excel's sort
    for r in rows:
        sheet.book.forget_formulas(sheet, r, c1, r, c2)
    moved = []
    for target, (_, source, cells) in zip(rows, ordered, strict=False):
        row = sheet.rows.setdefault(target, {})
        for c in [c for c in row if c1 <= c <= c2]:
            del row[c]
        for c, cell in cells.items():
            new = cell.copy()
            if cell.formula is not None:
                new.formula = shift_formula(cell.formula, target - source, 0)
                new.value = PENDING if target != source else cell.value
                moved.append((target, c, new.formula))
            row[c] = new
    for r in rows:
        if r in sheet.rows and not sheet.rows[r]:
            del sheet.rows[r]
    sheet.changed = True
    sheet.book.cells_changed(sheet, r1, c1, r2, c2)
    for r, c, text in moved:
        sheet.book.register_formula(sheet, r, c, text, False)
