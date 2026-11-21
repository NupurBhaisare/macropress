"""Worksheet: its cells, view state (selection, hidden rows), filter and sort state."""

import re

from ..errors import VBANotImplemented, VBARuntimeError
from ..values import (
    VBAObject,
    default_value,
    format_double,
    to_long,
    to_str,
)
from .addresses import _COLS, col_letters, col_num, parse_area
from .constants import MAX_COL, MAX_ROW, PENDING
from .range import Range
from .sorting import SortObject
from .storage import Cell

_COL = re.compile(r"^\$?([A-Za-z]{1,3})$")
# What Excel refuses in a sheet's name: one of :\/?*[], an apostrophe at either end, or "History".
_BAD_NAME = re.compile(r"[:\\/?*\[\]]|^'|'$|^history$", re.I)


class Sheet(VBAObject):
    vba_type_name = "Worksheet"

    def __init__(self, book, name, code_name=None, state="visible"):
        self.book = book
        self.name = name
        self.code_name = code_name
        self.state = state
        self.rows = {}  # row -> {col: Cell}
        self.hidden_rows = set()
        self.manual_hidden = set()  # hidden rows not explained by a filter
        self.hidden_cols = set()
        self.merged = []  # [(r1, c1, r2, c2)]
        self.autofilter = None
        self.sel = (1, 1, 1, 1)
        self.active_cell = (1, 1)
        self.dimension = None  # (r2, c2) from the file
        self.changed = False
        self.spills = {}  # (r, c) of a dynamic-array formula -> the rectangle it fills
        self.spill_owner = {}  # (r, c) of a spilled value -> (r, c) of its formula
        self.dynamic_cm = {}  # (r, c) -> the cell's cm attribute in the file (dynamic-array metadata)
        self.array_cells = {}  # (r, c) of an array formula entered with Ctrl+Shift+Enter -> the rectangle it fills
        self._sort = None

    # --- storage ---
    def fresh(self):
        """Calculate stale formulas before the sheet's values are read."""
        if self.book.dirty:
            self.book.refresh(self)
        return self

    def get_cell(self, r, c):
        if self.book.dirty:
            self.book.refresh_cell(self, r, c)
        row = self.rows.get(r)
        return row.get(c) if row else None

    def is_empty(self, r, c):
        cell = self.get_cell(r, c)
        return cell is None or (cell.value is None and cell.formula is None)

    def raw_value(self, r, c):
        cell = self.get_cell(r, c)
        if cell is None:
            return None
        if cell.value is PENDING:
            raise VBANotImplemented(
                f"reading {self.name}!{col_letters(c)}{r}: a formula whose value could not "
                f"be calculated (a function or reference the evaluator does not cover)"
            )
        return cell.value

    def cell_text(self, r, c):
        """The text the cell shows (General format only)."""
        cell = self.get_cell(r, c)
        if cell is None or cell.value is None:
            return ""
        v = self.raw_value(r, c)
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, str):
            return v
        if isinstance(v, (int, float)):
            if cell.numfmt not in (None, "General"):
                raise VBANotImplemented(
                    f"the displayed text of {self.name}!{col_letters(c)}{r} (number format {cell.numfmt!r})"
                )
            return format_double(float(v))
        return str(v)

    def put(self, r, c, value, formula=None, numfmt=None, dynamic=False):
        """Write a cell. numfmt is the format the value brings (a Date's), which Excel gives only a General cell: one
        with a number format keeps it, even emptied (a cell with no value or formula stays for its format)."""
        self.changed = True
        self.book.forget_formulas(self, r, c, r, c)
        old = self.rows.get(r, {}).get(c)
        numfmt = old.numfmt if old is not None and old.numfmt is not None else numfmt
        if value is None and formula is None and numfmt is None:
            row = self.rows.get(r)
            if row is not None:
                row.pop(c, None)
                if not row:
                    del self.rows[r]
        else:
            self.rows.setdefault(r, {})[c] = Cell(value, formula, numfmt)
            if formula is not None:
                self.book.register_formula(self, r, c, formula, dynamic)
        self.book.cells_changed(self, r, c, r, c)

    def clear_rect(self, r1, c1, r2, c2, formats=False):
        """Clear the cells' contents; their number formats stay, as with ClearContents, unless formats (Clear)."""
        self.changed = True
        self.book.forget_formulas(self, r1, c1, r2, c2)
        for r in [r for r in self.rows if r1 <= r <= r2]:
            row = self.rows[r]
            for c in [c for c in row if c1 <= c <= c2]:
                if row[c].numfmt is not None and not formats:
                    row[c] = Cell(None, None, row[c].numfmt)
                else:
                    del row[c]
            if not row:
                del self.rows[r]
        self.book.cells_changed(self, r1, c1, r2, c2)

    def check_merged(self, r1, c1, r2, c2):
        for m in self.merged:
            mr1, mc1, mr2, mc2 = m
            overlaps = not (mr2 < r1 or mr1 > r2 or mc2 < c1 or mc1 > c2)
            inside = r1 <= mr1 and mr2 <= r2 and c1 <= mc1 and mc2 <= c2
            if overlaps and not inside:
                raise VBARuntimeError(1004, "Cannot change part of a merged cell")

    def used_bounds(self):
        if self.book.dirty:
            # A stale formula's cell is used whatever its value, so only a stale dynamic array, which may spill, can
            # change the bounds. Calculating the others here would calculate, in the middle of a formula reading a
            # range (SUM(D5:D9)), the formulas that read that formula: a circular reference that is not one.
            self.book.refresh(self, spills_only=True)
        rows = [r for r, row in self.rows.items() if row]
        if not rows:
            return None
        cols = [c for row in self.rows.values() for c in row]
        r2, c2 = max(rows), max(cols)
        if self.dimension:
            r2, c2 = max(r2, self.dimension[0]), max(c2, self.dimension[1])
        return (min(rows), min(cols), r2, c2)

    def in_filter_mode(self):
        return self.autofilter is not None and bool(self.autofilter.criteria)

    # --- VBA members ---
    def get_name(self):
        return self.name

    def let_name(self, value):
        new = to_str(value)
        if not 1 <= len(new) <= 31 or _BAD_NAME.search(new):
            raise VBARuntimeError(1004, "You typed an invalid name for a sheet or chart")
        other = self.book.find_sheet(new)
        if other is not None and other is not self:
            raise VBARuntimeError(1004, "That name is already taken")
        if new != self.name:
            self.book.rename_sheet(self, new)
        self.changed = True

    def get_codename(self):
        return self.code_name or ""

    def get_index(self):
        return self.book.sheets.index(self) + 1

    def get_parent(self):
        return self.book

    def get_visible(self):
        return {"visible": -1, "hidden": 0, "veryHidden": 2}.get(self.state, -1)

    def get_range(self, cell1, cell2=None):
        if cell2 is None:
            if isinstance(cell1, Range):
                return Range(cell1.sheet, *cell1.rect)
            return self.range_from_text(to_str(cell1))
        a = cell1 if isinstance(cell1, Range) else self.range_from_text(to_str(cell1))
        b = cell2 if isinstance(cell2, Range) else self.range_from_text(to_str(cell2))
        if a.sheet is not b.sheet or a.sheet is not self:
            raise VBARuntimeError(1004, "Method 'Range' of object '_Worksheet' failed")
        return Range(self, min(a.r1, b.r1), min(a.c1, b.c1), max(a.r2, b.r2), max(a.c2, b.c2))

    def range_from_text(self, text):
        t = text.strip()
        sheet = self
        if "!" in t:
            sname, t = t.rsplit("!", 1)
            sname = sname.strip("'").replace("''", "'")
            sheet = self.book.find_sheet(sname)
            if sheet is None:
                raise VBARuntimeError(1004, f"Method 'Range' failed: no sheet {sname!r}")
        if "," in t:
            raise VBANotImplemented(f"multi-area range {text!r}")
        area = parse_area(t)
        if area is not None:
            return Range(sheet, *area)
        m = re.match(r"^([A-Za-z_][\w.]*)\[(.*)\]$", t)
        if m and m.group(1).lower() in self.book.tables:
            return self.table_range(self.book.tables[m.group(1).lower()], m.group(2))
        if t.lower() in self.book.tables:
            tbl = self.book.tables[t.lower()]
            return Range(
                self.book.find_sheet(tbl.sheet), tbl.r1 + tbl.header_rows, tbl.c1, tbl.r2 - tbl.totals_rows, tbl.c2
            )
        ref = self.book.resolve_name(t, sheet)
        if ref is not None:
            if ref.startswith("="):
                ref = ref[1:]
            return self.range_from_text(ref)
        raise VBARuntimeError(1004, f"Method 'Range' of object '_Global' failed ({text!r})")

    def table_range(self, tbl, spec):
        sheet = self.book.find_sheet(tbl.sheet)
        body_r1, body_r2 = tbl.r1 + tbl.header_rows, tbl.r2 - tbl.totals_rows
        s = spec.strip()
        if s.lower() == "#all":
            return Range(sheet, tbl.r1, tbl.c1, tbl.r2, tbl.c2)
        if s.lower() == "#headers":
            return Range(sheet, tbl.r1, tbl.c1, tbl.r1, tbl.c2)
        if s.lower() == "#data" or s == "":
            return Range(sheet, body_r1, tbl.c1, body_r2, tbl.c2)
        col = s.strip("[]")
        names = [c.lower() for c in tbl.columns]
        if col.lower() not in names:
            raise VBARuntimeError(1004, f"Table {tbl.name} has no column {col!r}")
        c = tbl.c1 + names.index(col.lower())
        return Range(sheet, body_r1, c, body_r2, c)

    def get_cells(self, rowindex=None, columnindex=None):
        if rowindex is None and columnindex is None:
            return Range(self, 1, 1, MAX_ROW, MAX_COL)
        return Range(self, 1, 1, MAX_ROW, MAX_COL).get_cells(rowindex, columnindex)

    def get_rows(self, rowindex=None):
        if rowindex is None:
            return Range(self, 1, 1, MAX_ROW, MAX_COL, view="rows")
        v = default_value(rowindex)
        if isinstance(v, str):
            area = parse_area(v if ":" in v else f"{v}:{v}")
            if area is None or area[1] != 1 or area[3] != MAX_COL:
                raise VBARuntimeError(13, "Type mismatch")
            return Range(self, area[0], 1, area[2], MAX_COL, view="rows")
        r = to_long(v)
        if not 1 <= r <= MAX_ROW:
            raise VBARuntimeError(1004, "Application-defined or object-defined error")
        return Range(self, r, 1, r, MAX_COL, view="rows")

    def get_columns(self, columnindex=None):
        if columnindex is None:
            return Range(self, 1, 1, MAX_ROW, MAX_COL, view="columns")
        v = default_value(columnindex)
        if isinstance(v, str):
            m = _COLS.match(v) or _COL.match(v)
            if not m:
                raise VBARuntimeError(13, "Type mismatch")
            c1 = col_num(m.group(1))
            c2 = col_num(m.group(2)) if m.lastindex == 2 else c1
            return Range(self, 1, c1, MAX_ROW, c2, view="columns")
        c = to_long(v)
        return Range(self, 1, c, MAX_ROW, c, view="columns")

    def m_select(self, replace=None):
        self.book.select_sheets([self], replace is None or bool(default_value(replace)))

    def m_activate(self):
        if self.state != "visible":
            raise VBARuntimeError(1004, "Activate method of Worksheet class failed")
        if self not in self.book.selected:
            self.book.selected = [self]
        self.book.active = self

    def get_usedrange(self):
        b = self.used_bounds()
        if b is None:
            return Range(self, 1, 1, 1, 1)
        return Range(self, *b)

    def get_autofiltermode(self):
        return self.autofilter is not None

    def let_autofiltermode(self, value):
        if not bool(default_value(value)) and self.autofilter is not None:
            self.remove_autofilter()

    def get_filtermode(self):
        return self.in_filter_mode()

    def m_showalldata(self):
        if self.autofilter is None:
            raise VBARuntimeError(1004, "ShowAllData method of Worksheet class failed")
        af = self.autofilter
        af.criteria = {}
        for r in range(af.r1 + 1, af.r2 + 1):
            self.hidden_rows.discard(r)
        self.changed = True

    def remove_autofilter(self):
        af = self.autofilter
        for r in range(af.r1 + 1, af.r2 + 1):
            self.hidden_rows.discard(r)
        self.autofilter = None
        self.changed = True

    def m_calculate(self):
        self.book.recalc_all()
        return True

    def get_sort(self):
        if self._sort is None:
            self._sort = SortObject(self)
        return self._sort
