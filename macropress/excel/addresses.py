"""A1-style addresses: column letters, parsing and writing areas, moving a formula's references."""

import functools
import re

from .. import formula as fx
from ..errors import VBANotImplemented, VBARuntimeError
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


# ---------------------------------------------------------------------------
# R1C1 references
# ---------------------------------------------------------------------------


def formula_to_r1c1(text, r, c):
    """The formula (its text after the =) with its references written R1C1 style, as seen from cell (r, c): $A$1 is
    R1C1, A1 seen from C3 is R[-2]C[-2], and no offset at all is R or C alone. A whole column or row with one end
    shows once (C[-3], not C[-3]:C[-3]); a range of cells keeps both ends."""
    out = []
    for kind, piece in _pieces(text):
        if kind not in fx.REFERENCE_KINDS:
            out.append(piece)
            continue
        head, bang, area = piece.rpartition("!")
        ends = []
        for end in area.split(":"):
            cabs, col, rabs, row = _END.fullmatch(end).groups()
            rows = (f"R{row}" if rabs else _offset("R", int(row) - r)) if row else ""
            cols = (f"C{col_num(col)}" if cabs else _offset("C", col_num(col) - c)) if col else ""
            ends.append(rows + cols)
        if kind != "ref" and len(ends) == 2 and ends[0] == ends[1]:
            ends = ends[:1]
        out.append(head + bang + ":".join(ends))
    return "".join(out)


def _offset(letter, n):
    return letter if n == 0 else f"{letter}[{n}]"


_PART = r"(?:\[[-+]?\d+\]|\d+)?"
_ATOM = rf"R{_PART}C{_PART}|R{_PART}|C{_PART}"
_SHEET_NAME = r"(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!"
_R1C1_TOKEN = re.compile(
    rf"""(?P<str>"(?:[^"]|"")*")
    |(?P<ref>(?:{_SHEET_NAME})?(?:{_ATOM})(?::(?:{_ATOM}))?)(?![\w.(\[])
    |(?P<word>[A-Za-z_\\][\w.]*(?:\[(?:[^\[\]]|\[[^\]]*\])*\])*)
    |(?P<other>.)""",
    re.X | re.I | re.S,
)
_ATOM_PARTS = re.compile(r"(?i)(?:R(\[[-+]?\d+\]|\d+)?)?(?:(C)(\[[-+]?\d+\]|\d+)?)?")
_A1_CELL = re.compile(r"[A-Za-z]{1,3}\d+")


class R1C1Formula:
    """A formula written R1C1 style, read once and written out A1 style for each cell it is given to. A reference
    that runs off the sheet comes round from its other edge, as in Excel (R[-1]C in row 1 is row 1048576); a row
    past the sheet stays as written, a name that cannot exist, and a column past it is error 1004."""

    def __init__(self, text):
        self.text = text
        self.pieces = []  # text as it stands, or (sheet prefix, [atom, ...]) for a reference
        for m in _R1C1_TOKEN.finditer(text):
            kind, piece = m.lastgroup, m.group()
            if kind == "ref":
                head, bang, area = piece.rpartition("!")
                atoms = [self._atom(a) for a in area.split(":")]
                self.pieces.append(piece if None in atoms else (head + bang, atoms))
            elif kind == "word" and _A1_CELL.fullmatch(piece) or kind == "other" and piece == "$":
                raise VBANotImplemented(f"an A1 reference in a formula written R1C1 style (={text})")
            else:
                self.pieces.append(piece)

    @staticmethod
    def _atom(text):
        """(row, column) of one end: each None (not given), ("abs", n) or ("rel", offset); None for a row past the
        sheet, which Excel leaves as a name."""
        m = _ATOM_PARTS.fullmatch(text)
        row, has_c, col = m.group(1), m.group(2), m.group(3)
        has_r = text[:1] in "Rr"

        def part(given, present, limit):
            if not present:
                return None
            if given is None:
                return ("rel", 0)
            if given.startswith("["):
                n = int(given[1:-1])
                if abs(n) >= limit:
                    raise VBANotImplemented(f"the R1C1 offset {given}, more than the sheet holds")
                return ("rel", n)
            return ("abs", int(given))

        r, c = part(row, has_r, MAX_ROW), part(col, has_c, MAX_COL)
        if c is not None and c[0] == "abs" and not 1 <= c[1] <= MAX_COL:
            if c[1] == 0:
                raise VBANotImplemented(f"the R1C1 reference {text}")
            raise VBARuntimeError(1004, "Application-defined or object-defined error")
        if r is not None and r[0] == "abs" and not 1 <= r[1] <= MAX_ROW:
            return None
        return r, c

    def a1(self, r, c):
        """The formula as cell (r, c) has it, A1 style."""
        out = []
        for piece in self.pieces:
            if isinstance(piece, str):
                out.append(piece)
                continue
            prefix, atoms = piece
            kinds = {(a[0] is not None, a[1] is not None) for a in atoms}
            if len(atoms) == 2 and len(kinds) == 1 and kinds != {(True, True)}:  # rows or columns from one to another
                out.append(prefix + ":".join(self._line(a, r, c) for a in atoms))
            else:
                out.append(prefix + ":".join(self._end(a, r, c) for a in atoms))
        return "".join(out)

    @staticmethod
    def _where(part, here, limit):
        kind, n = part
        return n if kind == "abs" else (here - 1 + n) % limit + 1

    def _line(self, atom, r, c):
        """One end of a range of whole rows (2) or columns (B), $ where the end is absolute."""
        row, col = atom
        if row is not None:
            return ("$" if row[0] == "abs" else "") + str(self._where(row, r, MAX_ROW))
        return ("$" if col[0] == "abs" else "") + col_letters(self._where(col, c, MAX_COL))

    def _end(self, atom, r, c):
        row, col = atom
        if row is not None and col is not None:
            return (("$" if col[0] == "abs" else "") + col_letters(self._where(col, c, MAX_COL))
                    + ("$" if row[0] == "abs" else "") + str(self._where(row, r, MAX_ROW)))  # fmt: skip
        line = self._line(atom, r, c)
        return f"{line}:{line}"


def written_r1c1(text):
    """Whether a formula given to Formula is R1C1 style: Excel reads it so when it names a reference A1 style cannot
    (R1C1, RC[-1], R[2]), which no defined name may be."""
    return any(kind == "name" and re.fullmatch(_ATOM, piece, re.I) for kind, piece in _pieces(text))


# ---------------------------------------------------------------------------
# Cut and paste: references follow the cells they name
# ---------------------------------------------------------------------------


def _inside(a, b):
    return b[0] <= a[0] and a[2] <= b[2] and b[1] <= a[1] and a[3] <= b[3]


def _overlaps(a, b):
    return not (a[2] < b[0] or a[0] > b[2] or a[3] < b[1] or a[1] > b[3])


def _edge_moved_along(ref, cut, dr, dc):
    """Whether the cells cut hold a whole edge of the range and move along it: Excel then stretches or shrinks the
    range (SUM(A1:A3) with A3 cut to A5 becomes SUM(A1:A5)), which is not modelled."""
    a1, b1, a2, b2 = ref
    rows = [r for r in {a1, a2} if cut[0] <= r <= cut[2] and cut[1] <= b1 and b2 <= cut[3]]
    cols = [c for c in {b1, b2} if cut[1] <= c <= cut[3] and cut[0] <= a1 and a2 <= cut[2]]
    return (rows and dc == 0) or (cols and dr == 0)


def _moved_end(end, dr, dc):
    cabs, col, rabs, row = (part or "" for part in _END.fullmatch(end).groups())
    return cabs + (col_letters(col_num(col) + dc) if col else "") + rabs + (str(int(row) + dr) if row else "")


def moved_references(text, home, own, cut, dest, sheets):
    """The formula after Excel moves the cells `cut` (sheet, r1, c1, r2, c2) to `dest` (sheet, top row, left
    column): a reference wholly within the cells cut follows them, its $ kept; one wholly within the cells they land
    on, and not cut, is #REF!; one partly within either stays as it was. `own` is the sheet the formula's references
    with no sheet name are on, `home` the sheet it is on after the move; `sheets` finds a sheet by its name. Where
    Excel stretches or shrinks a range instead, or might, VBANotImplemented."""
    src, *s_rect = cut
    dst, tr, tc = dest
    dr, dc = tr - s_rect[0], tc - s_rect[1]
    landed = (tr, tc, tr + s_rect[2] - s_rect[0], tc + s_rect[3] - s_rect[1])
    out = []
    for kind, piece in _pieces(text):
        if kind not in fx.REFERENCE_KINDS:
            out.append(piece)
            continue
        name, area = fx.ref_parts(piece)
        target = own if name is None else sheets(name)
        rect = parse_area(area)
        moved = target is src and _inside(rect, s_rect)
        if moved:
            target, area = dst, ":".join(_moved_end(end, dr, dc) for end in area.split(":"))
        elif target is dst and _inside(rect, landed):
            if name is not None:
                raise VBANotImplemented(f"cutting onto cells another sheet's formula names ({piece})")
            out.append("#REF!")
            continue
        elif target is src and _overlaps(rect, s_rect):
            if src is not dst or _edge_moved_along(rect, s_rect, dr, dc):
                raise VBANotImplemented(f"a cut that moves part of the range {piece} a formula reads")
        if not moved and (name is not None or home is own):
            out.append(piece)
        elif name is not None or target is not home:
            out.append(f"{fx.sheet_ref(target.name)}!{area}")
        else:
            out.append(area)
    return "".join(out)
