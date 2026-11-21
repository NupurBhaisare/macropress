"""A1-style addresses: column letters, parsing and writing areas, moving a formula's references."""

import functools
import re

from .. import formula as fx
from ..errors import VBANotImplemented
from .constants import MAX_COL, MAX_ROW

# ---------------------------------------------------------------------------
# Addresses
# ---------------------------------------------------------------------------


def col_num(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + ord(ch) - 64
    return n


def col_letters(n):
    if not 1 <= n <= MAX_COL:
        raise ValueError(f"column {n} is off the sheet (columns run from 1 to {MAX_COL})")
    out = ""
    while n:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


_CELL = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d{1,7})$")
_COLS = re.compile(r"^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$")
_ROWS = re.compile(r"^\$?(\d{1,7}):\$?(\d{1,7})$")


def parse_area(text):
    """'A1', 'A1:B2', 'A:C', '4:10' -> (r1, c1, r2, c2), or None."""
    t = text.replace(" ", "")
    parts = t.split(":")
    if len(parts) == 2:
        m = _COLS.match(t)
        if m:
            return (1, col_num(m.group(1)), MAX_ROW, col_num(m.group(2)))
        m = _ROWS.match(t)
        if m:
            return (int(m.group(1)), 1, int(m.group(2)), MAX_COL)
        a, b = _CELL.match(parts[0]), _CELL.match(parts[1])
        if a and b:
            return (int(a.group(2)), col_num(a.group(1)), int(b.group(2)), col_num(b.group(1)))
        return None
    if len(parts) == 1:
        m = _CELL.match(t)
        if m:
            r, c = int(m.group(2)), col_num(m.group(1))
            return (r, c, r, c)
    return None


def area_text(r1, c1, r2, c2, absolute=True, row_absolute=None, column_absolute=None):
    dr = "$" if (absolute if row_absolute is None else row_absolute) else ""
    dc = "$" if (absolute if column_absolute is None else column_absolute) else ""
    if c1 == 1 and c2 == MAX_COL:
        return f"{dr}{r1}:{dr}{r2}"
    if r1 == 1 and r2 == MAX_ROW:
        return f"{dc}{col_letters(c1)}:{dc}{col_letters(c2)}"
    a = f"{dc}{col_letters(c1)}{dr}{r1}"
    if (r1, c1) == (r2, c2):
        return a
    return f"{a}:{dc}{col_letters(c2)}{dr}{r2}"


DATE_FORMAT_IDS = set(range(14, 23)) | {45, 46, 47}


def is_date_format(code):
    if not code or code in ("General", "@"):
        return False
    stripped = re.sub(r'"[^"]*"|\[[^\]]*\]|\\.|_.|\*.', "", code)
    first = stripped.split(";")[0]
    return bool(re.search(r"[dmyhs]", first, re.I)) and not re.search(r"[0#?]", first)


# One end of a reference's area: a cell (A1), a column (A) or a row (1), each part with or without its $.
_END = re.compile(r"(?:(\$?)([A-Za-z]{1,3}))?(?:(\$?)(\d+))?")


def shift_formula(text, drow, dcol):
    """The formula with its relative references moved by (drow, dcol), as Excel moves them when it fills, pastes or
    sorts a formula. A reference is what the formula parser reads as one, so one on another sheet (Data!A1, 'My
    sheet'!A1), a range, a column (A:C) or a row (2:5) moves like a cell, and nothing else does: a $ part, a string, a
    function, a [table column]. A reference moved off the sheet raises VBANotImplemented: Excel writes #REF! there."""
    if not drow and not dcol:
        return text
    return "".join(
        _move(piece, drow, dcol, text) if kind in fx.REFERENCE_KINDS else piece for kind, piece in _pieces(text)
    )


def has_relative_reference(text):
    """Whether a reference in the formula has a column or row without its $: Excel reads it from the cell using it."""
    for kind, piece in _pieces(text):
        if kind in fx.REFERENCE_KINDS:
            for end in piece.rpartition("!")[2].split(":"):
                cabs, col, rabs, row = _END.fullmatch(end).groups()
                if (col and not cabs) or (row and not rabs):
                    return True
    return False


@functools.lru_cache(maxsize=1024)
def _pieces(text):
    return tuple(fx.pieces(text))  # a fill, a paste or a shared formula moves one formula across many cells


def _move(ref, drow, dcol, formula):
    """One reference moved: the relative column of each end by dcol, its relative row by drow."""

    def off_sheet():
        return VBANotImplemented(
            f"moving ={formula} by {drow} rows and {dcol} columns: {ref} would leave the sheet (Excel writes #REF!)"
        )

    head, bang, area = ref.rpartition("!")
    ends = []
    for end in area.split(":"):
        cabs, col, rabs, row = (part or "" for part in _END.fullmatch(end).groups())
        if col and not cabs and dcol:
            n = col_num(col) + dcol
            if not 1 <= n <= MAX_COL:
                raise off_sheet()
            col = col_letters(n)
        if row and not rabs and drow:
            n = int(row) + drow
            if not 1 <= n <= MAX_ROW:
                raise off_sheet()
            row = str(n)
        ends.append(cabs + col + rabs + row)
    return head + bang + ":".join(ends)
