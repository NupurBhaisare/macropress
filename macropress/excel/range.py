"""Range: reading and writing cells, moving around the sheet, copy and paste, filters and sorting."""

import datetime as dt
from decimal import ROUND_HALF_UP, Decimal

from .. import entry, formula as fx
from ..errors import VBANotImplemented, VBARuntimeError
from ..numfmt import (
    CURRENCY_VALUE,
    DATE,
    DATE_AND_TIME,
    TEXT,
    TIME,
    entered,
    is_currency_format,
    is_date_format,
    is_time_format,
    writes_a_currency_sign,
)
from ..values import (
    EMPTY,
    NULL,
    VBAArray,
    VBAObject,
    XLError,
    date_to_serial,
    default_value,
    format_double,
    is_currency,
    is_numeric_value,
    serial_to_date,
    to_currency,
    to_long,
    to_str,
    typed_number,
)
from .addresses import (
    R1C1Formula,
    area_text,
    col_num,
    formula_to_r1c1,
    moved_references,
    shift_formula,
    written_r1c1,
)
from .constants import MAX_COL, MAX_ROW, PENDING
from .filters import AutoFilterState, BoolCriterion, ValuesCriterion, apply_filter, criterion_from_text
from .storage import Cell

# ---------------------------------------------------------------------------
# Range
# ---------------------------------------------------------------------------


def _counting(start, at, period):
    """A line's values going on by 1 a step: `start` at position `at` of a block of `period` cells, repeated."""
    return lambda p: start + (p - at) // period if (p - at) % period == 0 else None


def _trend(ys):
    """The least-squares line through ys at 0, 1, 2...: what AutoFill continues numbers along."""
    n = len(ys)
    mean_x, mean_y = (n - 1) / 2, sum(ys) / n
    slope = sum((x - mean_x) * (y - mean_y) for x, y in enumerate(ys)) / sum((x - mean_x) ** 2 for x in range(n))
    return lambda p: mean_y + slope * (p - mean_x)


def _check_formula(text):
    """Excel refuses a formula with a number too large to type (1E+308) as it is entered: error 1004."""
    for kind, piece in fx.pieces(text):
        if kind == "num" and typed_number(piece) is None:
            raise VBARuntimeError(1004, "Application-defined or object-defined error")


class Clipboard:
    def __init__(self, sheet, rows, c1, c2, cut=False):
        self.sheet, self.rows, self.c1, self.c2, self.cut = sheet, rows, c1, c2, cut


# AutoFill's Type:= and what each does here; others are not modelled
FILL_DEFAULT, FILL_COPY, FILL_SERIES, FILL_FORMATS, FILL_VALUES, FILL_DAYS = 0, 1, 2, 3, 4, 5
# Names Excel counts on in a fill (its lists of months and weekdays, in the language of the machine)
_LISTED = {
    name.lower()
    for full in ("January February March April May June July August September October November December "
                 "Sunday Monday Tuesday Wednesday Thursday Friday Saturday").split()
    for name in (full, full[:3])
}  # fmt: skip


class Range(VBAObject):
    vba_type_name = "Range"

    def __init__(self, sheet, r1, c1, r2, c2, group=None, view="cells"):
        if r1 > r2:
            r1, r2 = r2, r1
        if c1 > c2:
            c1, c2 = c2, c1
        if r1 < 1 or c1 < 1 or r2 > MAX_ROW or c2 > MAX_COL:
            raise VBARuntimeError(1004, "Application-defined or object-defined error")
        self.sheet = sheet
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2
        self.group = group
        self.view = view

    @property
    def rect(self):
        return (self.r1, self.c1, self.r2, self.c2)

    @property
    def app(self):
        return self.sheet.book.app

    def __repr__(self):
        return f"Range({self.sheet.name}!{area_text(*self.rect)})"

    def _single_sheet_target(self, what):
        if self.group and len(self.group) > 1:
            raise VBANotImplemented(f"{what} on a selection while several sheets are grouped (unverified in Excel)")

    # --- values ---
    def _vba_value(self, r, c, value2=False):
        cell = self.sheet.get_cell(r, c)
        if cell is None or cell.value is None:
            return EMPTY
        v = self.sheet.raw_value(r, c)
        if not value2 and isinstance(v, float):  # Value reads a number by its format; Value2 as it is
            if is_date_format(cell.numfmt):
                return serial_to_date(v)
            if is_currency_format(cell.numfmt):
                return to_currency(v)
            if writes_a_currency_sign(cell.numfmt):
                self.app.note("CURRENCY-SIGN-IN-FORMAT", cell.numfmt)
        return v

    def _value(self, value2=False):
        if (self.r1, self.c1) == (self.r2, self.c2):
            return self._vba_value(self.r1, self.c1, value2)
        rows, cols = self.r2 - self.r1 + 1, self.c2 - self.c1 + 1
        if rows * cols > 5_000_000:
            raise VBANotImplemented("reading the value of a very large range")
        arr = VBAArray([(1, rows), (1, cols)])
        for r in range(rows):
            for c in range(cols):
                v = self._vba_value(self.r1 + r, self.c1 + c, value2)
                if v is not EMPTY:
                    arr.set([r + 1, c + 1], v)
        return arr

    def vba_default_get(self):
        return self._value()

    def get_value(self, rangevaluedatatype=None):
        return self._value()

    def get_value2(self):
        return self._value(value2=True)

    def vba_default_let(self, value):
        self.let_value(value)

    def let_value(self, value, rangevaluedatatype=None):
        self._assign(value, plain=False)

    def let_value2(self, value):
        """Value2 writes a Date or a Currency as the bare number: no date format, no rounding to cents."""
        self._assign(value, plain=True)

    def _assign(self, value, plain):
        self._single_sheet_target("writing values")
        self.sheet.check_merged(*self.rect)
        if (self.r2 - self.r1 + 1) * (self.c2 - self.c1 + 1) > 5_000_000:
            raise VBANotImplemented("writing one value into a very large range")
        areas = self._written_areas()
        if isinstance(value, VBAArray):
            self._guard_contents(areas)
            for r1, c1, r2, c2 in areas:
                self._write_array(value, r1, c1, r2, c2, plain)
            return
        v, refused = default_value(value), False
        for area in areas:
            if self.sheet.protects() and self.sheet.locks(*area) is not False:
                # a protected sheet: one value written over cells some of which are locked reaches the others
                self.app.note("PROTECTED-SHEET", f"{self.sheet.name}: {area_text(*area)}")
                refused = True
                for r in range(area[0], area[2] + 1):
                    for c in range(area[1], area[3] + 1):
                        if not self.sheet.locked_at(r, c):
                            self._write(v, r, c, r, c, plain)
            else:
                self._write(v, *area, plain)
        if refused:
            raise VBARuntimeError(1004, "The cell or chart you're trying to change is on a protected sheet")

    def _guard_contents(self, areas):
        """A change on a protected sheet to cells some of which are locked, and some not, is not modelled; to locked
        cells alone it is error 1004."""
        sheet = self.sheet
        if sheet.protects():
            if any(sheet.locks(*area) is None for area in areas):
                raise VBANotImplemented("changing locked and unlocked cells at once on a protected sheet (unverified)")
            sheet.guard("contents", areas)

    def _written_areas(self):
        """The parts of the range that a write or a clear reaches. On a filtered sheet these are the rows the filter
        shows, each run of them on its own (an array written to the range starts again in each); a range whose rows
        the filter all hides is reached whole."""
        sheet = self.sheet
        hidden = sheet.hidden_rows - sheet.manual_hidden if sheet.in_filter_mode() else ()
        hidden = sorted(r for r in hidden if self.r1 <= r <= self.r2)
        if not hidden or len(hidden) == self.r2 - self.r1 + 1:
            return [self.rect]
        self.app.note("FILTER-WRITES-VISIBLE")
        areas, start = [], self.r1
        for r in [*hidden, self.r2 + 1]:
            if r > start:
                areas.append((start, self.c1, r - 1, self.c2))
            start = r + 1
        return areas

    def _write(self, v, r1, c1, r2, c2, plain=False):
        """Write one value to every cell of the rectangle; relative references of a formula move with each cell."""
        if plain and isinstance(v, (dt.datetime, Decimal)):
            v = date_to_serial(v) if isinstance(v, dt.datetime) else float(v)
        cell = None
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                as_text = self._as_text(v, r, c)
                if as_text is not None:
                    self.sheet.put(r, c, as_text or None)
                    continue
                if cell is None:
                    cell = self._to_cell(v)
                stored, formula, fmt = cell
                if formula is not None:
                    formula = shift_formula(formula, r - self.r1, c - self.c1)
                self.sheet.put(r, c, stored, formula, fmt)

    def _write_array(self, arr, r1, c1, r2, c2, plain=False):
        if arr.bounds is None:
            raise VBARuntimeError(13, "Type mismatch")
        dims = arr.bounds
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                i, j = r - r1, c - c1
                try:
                    if len(dims) == 2:
                        v = arr.get([dims[0][0] + i, dims[1][0] + j])
                    else:
                        v = arr.get([dims[0][0] + j])
                except VBARuntimeError:
                    v = XLError("#N/A")
                self._write(default_value(v), r, c, r, c, plain)

    def _as_text(self, v, r, c):
        """What a cell formatted as Text holds after text or a date is written to it: the text as it is, never read
        as a number, a date or a formula. None for any other cell or value."""
        if not isinstance(v, (str, dt.datetime, Decimal)) or self.sheet.format_at(r, c) != TEXT:
            return None
        if is_currency(v):
            raise VBANotImplemented("writing a Currency to a cell formatted as Text")
        return v if isinstance(v, str) else to_str(v) if isinstance(v, dt.datetime) else None

    def _to_cell(self, v):
        """A VBA value written to a cell -> (stored value, formula, the number format the value brings)."""
        if v is EMPTY or v is NULL:
            return None, None, None
        if isinstance(v, (bool, XLError)):
            return v, None, None
        if is_currency(v):  # a Currency reaches the cell as an amount of money: rounded to cents, formatted as such
            return float(v.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)), None, CURRENCY_VALUE
        if is_numeric_value(v):
            return float(v), None, None
        if isinstance(v, dt.datetime):
            if v.date() == serial_to_date(0).date():
                return date_to_serial(v), None, TIME  # a time on its own
            return date_to_serial(v), None, DATE_AND_TIME if (v.hour, v.minute, v.second) != (0, 0, 0) else DATE
        if isinstance(v, str):
            if v == "":
                self.app.note("WRITE-EMPTY-STRING")
                return None, None, None
            value, formula, fmt = entry.read(v)
            if formula is not None:
                _check_formula(formula)
                formula = self._as_kept(formula, dynamic=False)
            return (PENDING if formula is not None else value), formula, fmt
        raise VBARuntimeError(13, "Type mismatch")

    def get_formula(self):
        """A formula as Excel keeps it (see _as_kept), or what a cell with no formula holds, as text: TRUE, #N/A, a
        date's number. A dynamic-array formula is shown as a formula that is not one reads it, without the @ that
        one takes as read (see formula.legacy_view)."""
        if (self.r1, self.c1) != (self.r2, self.c2):
            raise VBANotImplemented("Formula of a multi-cell range")
        cell = self.sheet.get_cell(self.r1, self.c1)
        if cell is None or (cell.formula is None and cell.value is None):
            return ""
        if cell.formula is not None:
            text = fx.display_syntax(cell.formula)
            if self.sheet.book.is_dynamic(self.sheet, self.r1, self.c1):
                text = fx.legacy_view(text)
                if text is None:
                    raise VBANotImplemented(f"Formula of ={fx.display_syntax(cell.formula)}: where its @ goes")
            return "=" + text
        v = self.sheet.raw_value(self.r1, self.c1)
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, float):
            return f"{v:.0f}" if 1e15 <= abs(v) < 1e21 else format_double(v)
        return str(v)

    def _write_formula(self, value, dynamic, r1c1=False):
        """A formula for each cell: its relative references move with each cell, as when Excel fills a range. One
        written R1C1 style (FormulaR1C1, or Formula given R1C1 text, which Excel reads too) is each cell's own."""
        text = default_value(value)
        if not isinstance(text, str) or not text.startswith("=") or len(text) == 1:
            return self._assign(value, plain=False)  # Formula = 5, a Date or text: the same as Value
        self._single_sheet_target("writing formulas")
        self.sheet.check_merged(*self.rect)
        if r1c1 or written_r1c1(text[1:]):
            r1c1_formula = R1C1Formula(text[1:])

            def at(r, c):
                return self._as_kept(r1c1_formula.a1(r, c), dynamic)
        else:
            _check_formula(text[1:])
            kept = self._as_kept(text[1:], dynamic)

            def at(r, c):
                return shift_formula(kept, r - self.r1, c - self.c1)

        _check_formula(at(self.r1, self.c1))
        areas = self._written_areas()
        self._guard_contents(areas)
        for r1, c1, r2, c2 in areas:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if self.sheet.format_at(r, c) == TEXT:
                        self.sheet.put(r, c, text)  # a cell formatted as Text keeps the formula as text
                    else:
                        self.sheet.put(r, c, PENDING, at(r, c), dynamic=dynamic)

    def let_formula(self, value):
        """Formula: a formula that returns one value (an array result gives its top-left value)."""
        self._write_formula(value, dynamic=False)

    def let_formula2(self, value):
        """Formula2: dynamic-array semantics; an array result spills into the cells below and to the right."""
        self._write_formula(value, dynamic=True)

    def get_formula2(self):
        """A formula as a dynamic-array formula reads it: one that is not one shows @ where it takes one value of a
        range or of what a function hands back (see formula.formula2_view)."""
        if (self.r1, self.c1) != (self.r2, self.c2):
            raise VBANotImplemented("Formula2 of a multi-cell range")
        cell = self.sheet.get_cell(self.r1, self.c1)
        if cell is None or cell.formula is None:
            return self.get_formula()
        text = fx.display_syntax(cell.formula)
        if not self.sheet.book.is_dynamic(self.sheet, self.r1, self.c1):
            text = fx.formula2_view(text, fx.ref_size(self.sheet.book, self.sheet))
            if text is None:
                raise VBANotImplemented(f"Formula2 of ={fx.display_syntax(cell.formula)}: where Excel puts its @")
        return "=" + text

    def _as_kept(self, formula, dynamic):
        """A formula (without its =) as Excel keeps what it is given (see formula.normalized); in one that is not a
        dynamic array, without the @ it takes as read."""
        text = fx.normalized(formula, self.sheet.book.find_sheet)
        if not dynamic and "@" in text:
            text = fx.legacy_view(text) or text
        return text

    def get_formular1c1(self):
        """The formula written R1C1 style, as seen from its cell; for a cell with no formula, what Formula gives."""
        if (self.r1, self.c1) != (self.r2, self.c2):
            raise VBANotImplemented("FormulaR1C1 of a multi-cell range")
        cell = self.sheet.get_cell(self.r1, self.c1)
        if cell is None or cell.formula is None:
            return self.get_formula()
        return "=" + formula_to_r1c1(fx.display_syntax(cell.formula), self.r1, self.c1)

    get_formula2r1c1 = get_formular1c1

    def let_formular1c1(self, value):
        self._write_formula(value, dynamic=False, r1c1=True)

    def let_formula2r1c1(self, value):
        self._write_formula(value, dynamic=True, r1c1=True)

    def get_text(self):
        if (self.r1, self.c1) != (self.r2, self.c2):
            raise VBANotImplemented("Text of a multi-cell range")
        return self.sheet.shown_text(self.r1, self.c1)

    def m_calculate(self):
        self.sheet.book.recalc_range(self.sheet, *self.rect)
        return True

    def get_numberformat(self):
        """The cells' number format; Null when they have more than one."""
        found = self.sheet.styles_in(*self.rect, 0)
        if len(found) > 1:
            return NULL
        code = found.pop() or "General"
        if code in (DATE, DATE_AND_TIME):
            self.app.note("DATE-FORMAT-NAME")
        return code

    def let_numberformat(self, value):
        """A number format for the cells, as Excel reads the code it is given (numfmt.entered). On a filtered sheet it
        reaches the rows the filter shows, as a write does."""
        self._single_sheet_target("setting NumberFormat")
        v = default_value(value)
        if isinstance(v, bool):  # "True" is no format Excel can read
            raise VBARuntimeError(1004, "Unable to set the NumberFormat property of the Range class")
        if isinstance(v, VBAArray):
            raise VBANotImplemented("NumberFormat set from an array")
        code = entered("" if v is EMPTY or v is NULL else to_str(v))
        self.sheet.guard("format")
        for area in self._written_areas():
            self.sheet.restyle(*area, numfmt=code)

    def get_locked(self):
        """Whether the cells are locked, as a protected sheet asks; Null when some are and some are not."""
        found = self.sheet.styles_in(*self.rect, 1)
        return NULL if len(found) > 1 else found.pop()

    def let_locked(self, value):
        self._single_sheet_target("setting Locked")
        v = default_value(value)
        if not isinstance(v, (bool, int, float)) or isinstance(v, VBAArray):
            raise VBARuntimeError(13, "Type mismatch")
        self.sheet.guard("lock")
        for area in self._written_areas():
            self.sheet.restyle(*area, locked=bool(v))

    # --- geometry ---
    def get_row(self):
        return self.r1

    def get_column(self):
        return self.c1

    def get_count(self):
        if self.view == "rows":
            return self.r2 - self.r1 + 1
        if self.view == "columns":
            return self.c2 - self.c1 + 1
        return (self.r2 - self.r1 + 1) * (self.c2 - self.c1 + 1)

    def get_worksheet(self):
        return self.sheet

    def get_parent(self):
        return self.sheet

    def get_address(self, rowabsolute=None, columnabsolute=None, referencestyle=None, external=None, relativeto=None):
        if referencestyle is not None and to_long(referencestyle) != 1:  # xlA1
            raise VBANotImplemented("Range.Address in R1C1 style")
        if external is not None and bool(default_value(external)):
            raise VBANotImplemented("Range.Address with External:=True")
        row_abs = rowabsolute is None or bool(default_value(rowabsolute))
        col_abs = columnabsolute is None or bool(default_value(columnabsolute))
        return area_text(*self.rect, row_absolute=row_abs, column_absolute=col_abs)

    def get_cells(self, rowindex=None, columnindex=None):
        if rowindex is None and columnindex is None:
            return Range(self.sheet, *self.rect)
        if columnindex is None:
            i = to_long(rowindex) - 1
            width = self.c2 - self.c1 + 1
            r, c = self.r1 + i // width, self.c1 + i % width
            return Range(self.sheet, r, c, r, c)
        cv = default_value(columnindex)
        col = col_num(cv) if isinstance(cv, str) else to_long(cv)
        r = self.r1 + to_long(rowindex) - 1
        c = self.c1 + col - 1
        return Range(self.sheet, r, c, r, c)

    def vba_index(self, args, kwargs):
        if self.view == "rows" and len(args) == 1:
            r = self.r1 + to_long(args[0]) - 1
            return Range(self.sheet, r, self.c1, r, self.c2)
        if self.view == "columns" and len(args) == 1:
            c = self.c1 + to_long(args[0]) - 1
            return Range(self.sheet, self.r1, c, self.r2, c)
        return self.get_cells(*args)

    get_item = get_cells

    def vba_iter(self):
        if self.view == "rows":
            return [Range(self.sheet, r, self.c1, r, self.c2) for r in range(self.r1, self.r2 + 1)]
        if self.view == "columns":
            return [Range(self.sheet, self.r1, c, self.r2, c) for c in range(self.c1, self.c2 + 1)]
        if self.get_count() > 1_000_000:
            raise VBANotImplemented("For Each over a very large range")
        return [Range(self.sheet, r, c, r, c) for r in range(self.r1, self.r2 + 1) for c in range(self.c1, self.c2 + 1)]

    def get_rows(self, rowindex=None):
        rows = Range(self.sheet, *self.rect, view="rows")
        return rows if rowindex is None else rows.vba_index([rowindex], {})

    def get_columns(self, columnindex=None):
        cols = Range(self.sheet, *self.rect, view="columns")
        return cols if columnindex is None else cols.vba_index([columnindex], {})

    def get_offset(self, rowoffset=0, columnoffset=0):
        dr, dc = to_long(rowoffset or 0), to_long(columnoffset or 0)
        return Range(self.sheet, self.r1 + dr, self.c1 + dc, self.r2 + dr, self.c2 + dc)

    def get_resize(self, rowsize=None, columnsize=None):
        rows = to_long(rowsize) if rowsize is not None else self.r2 - self.r1 + 1
        cols = to_long(columnsize) if columnsize is not None else self.c2 - self.c1 + 1
        if rows < 1 or cols < 1:
            raise VBARuntimeError(1004, "Application-defined or object-defined error")
        return Range(self.sheet, self.r1, self.c1, self.r1 + rows - 1, self.c1 + cols - 1)

    def get_entirerow(self):
        return Range(self.sheet, self.r1, 1, self.r2, MAX_COL, view="rows")

    def get_entirecolumn(self):
        return Range(self.sheet, 1, self.c1, MAX_ROW, self.c2, view="columns")

    def get_hidden(self):
        if self.c1 == 1 and self.c2 == MAX_COL:
            return all(r in self.sheet.hidden_rows for r in range(self.r1, self.r2 + 1))
        if self.r1 == 1 and self.r2 == MAX_ROW:
            return all(c in self.sheet.hidden_cols for c in range(self.c1, self.c2 + 1))
        raise VBARuntimeError(1004, "Unable to get the Hidden property of the Range class")

    def let_hidden(self, value):
        hide = bool(default_value(value))
        if self.c1 == 1 and self.c2 == MAX_COL:
            self.sheet.guard("rows")
            target, rng = self.sheet.hidden_rows, range(self.r1, self.r2 + 1)
        elif self.r1 == 1 and self.r2 == MAX_ROW:
            self.sheet.guard("columns")
            target, rng = self.sheet.hidden_cols, range(self.c1, self.c2 + 1)
        else:
            raise VBARuntimeError(1004, "Unable to set the Hidden property of the Range class")
        for i in rng:
            (target.add if hide else target.discard)(i)
            if target is self.sheet.hidden_rows:
                (self.sheet.manual_hidden.add if hide else self.sheet.manual_hidden.discard)(i)
        self.sheet.changed = True

    def get_end(self, direction):
        """Ctrl+arrow from the range's first cell: to the end of the filled cells it is in, or to the next filled
        cell. Hidden rows and columns are passed over as if they were not there."""
        d = to_long(direction)
        if d not in (-4121, -4162, -4161, -4159):  # xlDown, xlUp, xlToRight, xlToLeft
            raise VBARuntimeError(5, "Invalid procedure call or argument")
        sheet = self.sheet.fresh()
        r, c = self.r1, self.c1
        step = 1 if d in (-4121, -4161) else -1
        if d in (-4121, -4162):
            at, last, hidden = r, MAX_ROW if step == 1 else 1, sheet.hidden_rows
            filled_lines = [x for x in sheet.rows if not sheet.is_empty(x, c)]

            def empty(i):
                return sheet.is_empty(i, c)
        else:
            at, last, hidden = c, MAX_COL if step == 1 else 1, sheet.hidden_cols
            filled_lines = [x for x in (sheet.rows.get(r) or {}) if not sheet.is_empty(r, x)]

            def empty(i):
                return sheet.is_empty(r, i)

        def beyond(i):
            """The next line that is shown, or None at the sheet's edge."""
            i += step
            while i != last + step and i in hidden:
                i += step
            return None if i == last + step else i

        nxt = beyond(at)
        if nxt is not None and not empty(at) and not empty(nxt):
            while nxt is not None and not empty(nxt):
                at, nxt = nxt, beyond(nxt)
        else:
            ahead = [x for x in filled_lines if (x > at if step == 1 else x < at) and x not in hidden]
            at = (min if step == 1 else max)(ahead, default=last)
        return Range(sheet, at, c, at, c) if d in (-4121, -4162) else Range(sheet, r, at, r, at)

    def get_currentregion(self):
        sheet = self.sheet.fresh()
        r1, c1, r2, c2 = self.r1, self.c1, self.r1, self.c1

        def any_value(rr1, cc1, rr2, cc2):
            for r in range(max(rr1, 1), min(rr2, MAX_ROW) + 1):
                row = sheet.rows.get(r)
                if not row:
                    continue
                for c in row:
                    if max(cc1, 1) <= c <= min(cc2, MAX_COL) and not sheet.is_empty(r, c):
                        return True
            return False

        # Each side grows as far as it can before the next is tried, so a long table costs one pass down its rows, not
        # one pass per row; the region it ends at is the same.
        changed = True
        while changed:
            changed = False
            while r1 > 1 and any_value(r1 - 1, c1 - 1, r1 - 1, c2 + 1):
                r1, changed = r1 - 1, True
            while r2 < MAX_ROW and any_value(r2 + 1, c1 - 1, r2 + 1, c2 + 1):
                r2, changed = r2 + 1, True
            while c1 > 1 and any_value(r1 - 1, c1 - 1, r2 + 1, c1 - 1):
                c1, changed = c1 - 1, True
            while c2 < MAX_COL and any_value(r1 - 1, c2 + 1, r2 + 1, c2 + 1):
                c2, changed = c2 + 1, True
        return Range(sheet, r1, c1, r2, c2)

    # --- selection ---
    def m_select(self):
        book = self.sheet.book
        if self.sheet is not book.active:
            raise VBARuntimeError(1004, "Select method of Range class failed")
        targets = book.selected if self.sheet in book.selected else [self.sheet]
        if len(targets) > 1:
            self.app.note("GROUP-SELECT-RANGE")
        for s in targets:
            s.sel = self.rect
            s.active_cell = (self.r1, self.c1)
        return True

    def m_activate(self):
        book = self.sheet.book
        if self.sheet is not book.active:
            raise VBARuntimeError(1004, "Activate method of Range class failed")
        r1, c1, r2, c2 = self.sheet.sel
        if self.r1 == self.r2 and self.c1 == self.c2 and r1 <= self.r1 <= r2 and c1 <= self.c1 <= c2:
            self.sheet.active_cell = (self.r1, self.c1)
            return True
        return self.m_select()

    # --- editing ---
    def m_clearcontents(self):
        return self._clear(formats=False)

    def m_clear(self):
        self.app.note("FORMATS-NOT-MODELLED")
        return self._clear(formats=True)

    def _clear(self, formats):
        """ClearContents leaves the cells' number formats; Clear removes them too."""
        targets = [self.sheet]
        if self.group and len(self.group) > 1:
            targets = self.group
            self.app.note("GROUP-CLEARCONTENTS")
        for s in targets:
            s.check_merged(*self.rect)
        areas = self._written_areas() if len(targets) == 1 else [self.rect]
        for s in targets:
            s.guard("contents", areas)
        for s in targets:
            for area in areas:  # on a protected sheet Clear leaves unlocked cells' formats as they are
                s.clear_rect(*area, formats=formats and not s.protects())
        return True

    def m_delete(self, shift=None):
        raise VBANotImplemented("Range.Delete (shifting cells and every reference to them)")

    def m_insert(self, shift=None, copyorigin=None):
        raise VBANotImplemented("Range.Insert (shifting cells and every reference to them)")

    def m_autofit(self):
        self.app.note("FORMATS-NOT-MODELLED")
        return True

    def m_copy(self, destination=None):
        self._single_sheet_target("Copy")
        rows = list(range(self.r1, self.r2 + 1))
        hidden = [r for r in rows if r in self.sheet.hidden_rows]
        if hidden:
            if self.sheet.in_filter_mode():
                self.app.note("COPY-VISIBLE-IN-FILTER")
                rows = [r for r in rows if r not in self.sheet.hidden_rows]
            else:
                self.app.note("COPY-HIDDEN-NO-FILTER")
        clip = Clipboard(self.sheet, rows, self.c1, self.c2)
        if destination is not None:
            if not isinstance(destination, Range):
                raise VBARuntimeError(424, "Object required")
            destination._paste(clip, "all", True)
            return True
        self.app.clipboard = clip
        return True

    def m_cut(self, destination=None):
        """Cut: with a destination the cells move there now; without one they move when pasted (Worksheet.Paste)."""
        self._single_sheet_target("Cut")
        if destination is None:
            self.app.clipboard = Clipboard(self.sheet, list(range(self.r1, self.r2 + 1)), self.c1, self.c2, cut=True)
            return True
        if not isinstance(destination, Range):
            raise VBARuntimeError(424, "Object required")
        self._move(destination.sheet, destination.r1, destination.c1)
        self.app.clipboard = None
        return True

    def _move(self, dst, tr, tc):
        """Move the cells to dst, their top-left cell at (tr, tc): contents and formats go, the cells left behind are
        empty and General, and every formula and name in the workbook that reads them reads them where they are now
        (see addresses.moved_references). A filter does not keep its hidden rows back."""
        src, book = self.sheet, self.sheet.book
        r1, c1, r2, c2 = self.rect
        landed = (tr, tc, tr + r2 - r1, tc + c2 - c1)
        if landed[2] > MAX_ROW or landed[3] > MAX_COL:
            raise VBARuntimeError(1004, "The copy area and the paste area are not the same size")
        dr, dc = tr - r1, tc - c1
        whole = (c1, c2) == (1, MAX_COL) or (r1, r2) == (1, MAX_ROW)
        self._refuse_move(src, self.rect, dst, landed, whole)
        if not whole and (r2 - r1 + 1) * (c2 - c1 + 1) > 1_000_000:
            raise VBANotImplemented("cutting more than a million cells that are not whole rows or columns")
        src.fresh()
        dst.fresh()
        cells = {}  # what the cells cut hold, and their styles, before anything moves
        if whole:
            for r, row in src.rows.items():
                for c, cell in row.items():
                    if r1 <= r <= r2 and c1 <= c <= c2:
                        cells[(r, c)] = cell.copy()
        else:
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    cell = src.rows.get(r, {}).get(c)
                    cells[(r, c)] = cell.copy() if cell is not None else Cell(None, None, *src.inherited(r, c))
        cut, dest = (src, r1, c1, r2, c2), (dst, tr, tc)
        by_name = book.find_sheet
        for cell in cells.values():
            if cell.formula is not None:
                cell.formula = moved_references(cell.formula, dst, src, cut, dest, by_name)
        changed = []  # formulas elsewhere that read the cells cut, or the cells they land on
        for sheet in book.sheets:
            for r, row in sheet.rows.items():
                for c, cell in row.items():
                    if cell.formula is not None and not (sheet is src and r1 <= r <= r2 and c1 <= c <= c2):
                        text = moved_references(cell.formula, sheet, sheet, cut, dest, by_name)
                        if text != cell.formula:
                            changed.append((sheet, r, c, text))
        names = {
            key: [(scope, self._moved_name(ref, cut, dest, by_name)) for scope, ref in refs]
            for key, refs in book.names.items()
        }
        src.clear_rect(*self.rect, formats=True)
        dst.clear_rect(*landed, formats=False)
        for (r, c), cell in cells.items():
            dst._store(r + dr, c + dc, cell.value, cell.formula, *cell.style)
            if cell.formula is not None:
                book.register_formula(dst, r + dr, c + dc, cell.formula, False)
        for sheet, r, c, text in changed:
            sheet.rows[r][c].formula = text
            sheet.changed = True
            book.register_formula(sheet, r, c, text, False)
        book.names = names
        book.moves.append((src.name, self.rect, dst.name, landed))
        self.app.note("CUT-MOVES-REFERENCES", f"{src.name}!{area_text(*self.rect)} to {dst.name}!{area_text(*landed)}")
        dst.changed = True
        book.cells_changed(dst, *landed)

    @staticmethod
    def _refuse_move(src, s_rect, dst, landed, whole):
        """What a cut does that is not modelled: on a protected sheet, part of a merged area, a filter's range, a
        table, an array formula, and rows or columns that are hidden or formatted whole."""
        from .addresses import _overlaps  # noqa: PLC0415

        for sheet, rect in ((src, s_rect), (dst, landed)):
            if sheet.protection is not None and not sheet.ui_only:
                raise VBANotImplemented("cutting cells of a protected sheet (unverified)")
            sheet.check_merged(*rect)
            if sheet.autofilter is not None and _overlaps(rect, sheet.autofilter.rect):
                raise VBANotImplemented("cutting cells of a filter's range, or onto them")
            if any(_overlaps(rect, area) for area in (*sheet.spills.values(), *sheet.array_cells.values())):
                raise VBANotImplemented("cutting part of a dynamic array or array formula, or onto one")
            for t in sheet.book.tables.values():
                if t.sheet.lower() == sheet.name.lower() and _overlaps(rect, (t.r1, t.c1, t.r2, t.c2)):
                    raise VBANotImplemented("cutting cells of a table, or onto them")
            if whole and (
                any(rect[0] <= r <= rect[2] for r in (*sheet.hidden_rows, *sheet.row_formats))
                or any(rect[1] <= c <= rect[3] for c in (*sheet.hidden_cols, *sheet.col_formats, *sheet.col_unlocked))
            ):
                raise VBANotImplemented("cutting whole rows or columns that are hidden or formatted whole")

    @staticmethod
    def _moved_name(ref, cut, dest, by_name):
        text = ref[1:] if ref.startswith("=") else ref
        moved = moved_references(text, None, None, cut, dest, by_name)
        return ref if moved == text else ("=" if ref.startswith("=") else "") + moved

    def m_pastespecial(self, paste=None, operation=None, skipblanks=None, transpose=None):
        clip = self.app.clipboard
        if clip is None or clip.cut:  # cells cut are pasted whole, by Worksheet.Paste
            raise VBARuntimeError(1004, "PasteSpecial method of Range class failed")
        self._single_sheet_target("PasteSpecial")
        mode = -4104 if paste is None else to_long(paste)
        if operation is not None and to_long(operation) != -4142:
            raise VBANotImplemented("PasteSpecial with an Operation")
        if skipblanks is not None and bool(default_value(skipblanks)):
            raise VBANotImplemented("PasteSpecial SkipBlanks:=True")
        if transpose is not None and bool(default_value(transpose)):
            raise VBANotImplemented("PasteSpecial Transpose:=True")
        types = {  # Paste:= -> what it pastes, and whether the copied number formats come too
            -4104: ("all", True),  # xlPasteAll
            -4123: ("formulas", False),  # xlPasteFormulas: the destination keeps its number formats
            11: ("formulas", True),  # xlPasteFormulasAndNumberFormats
            -4163: ("values", False),  # xlPasteValues
            12: ("values", True),  # xlPasteValuesAndNumberFormats
            -4122: ("formats", True),  # xlPasteFormats
        }
        if mode == 8:  # xlPasteColumnWidths: nothing of the cells themselves
            self.app.note("FORMATS-NOT-MODELLED")
            return True
        if mode not in types:
            raise VBANotImplemented(f"PasteSpecial Paste:={mode}")
        self._paste(clip, *types[mode])
        return True

    def _paste_area(self, clip):
        """Where a paste lands: the copied block from this range's top-left cell, laid down as many times as the range
        holds it, down and across. A range a whole number of blocks high (or one cell high) and a whole number wide
        (or one cell wide) is filled with blocks; any other takes the block once. A block that would run off the
        sheet is Excel's error 1004."""
        height, width = len(clip.rows), clip.c2 - clip.c1 + 1
        rows, cols = self.r2 - self.r1 + 1, self.c2 - self.c1 + 1
        down = rows // height if rows % height == 0 else 1 if rows == 1 else None
        across = cols // width if cols % width == 0 else 1 if cols == 1 else None
        if down is None or across is None:
            down = across = 1
        r2, c2 = self.r1 + height * down - 1, self.c1 + width * across - 1
        if r2 > MAX_ROW or c2 > MAX_COL:
            raise VBARuntimeError(1004, "The copy area and the paste area are not the same size")
        return self.r1, self.c1, r2, c2

    def _paste_rows(self, clip, area):
        """The rows of the paste area the paste reaches. One cell copied over several on a filtered sheet goes to
        the rows the filter shows, as a write does; a block of cells lands whole, hidden rows and all."""
        r1, c1, r2, c2 = area
        sheet = self.sheet
        hidden = sheet.hidden_rows - sheet.manual_hidden if sheet.in_filter_mode() else ()
        hidden = [r for r in hidden if r1 <= r <= r2]
        if not hidden or (r1, c1) == (r2, c2):
            return range(r1, r2 + 1)
        if len(clip.rows) > 1 or clip.c1 != clip.c2:
            if (r2 - r1 + 1, c2 - c1 + 1) != (len(clip.rows), clip.c2 - clip.c1 + 1):
                raise VBANotImplemented("pasting copied cells over and over into a filtered range (unverified)")
            return range(r1, r2 + 1)
        self.app.note("FILTER-WRITES-VISIBLE")
        return [r for r in range(r1, r2 + 1) if r not in hidden]

    def _paste(self, clip, what, numfmts, area=None, anchor=None, rows=None, select=True):
        """Paste the copied cells. what: "all" (formulas, constants and formats), "formulas" (formulas and constants),
        "values" or "formats"; numfmts: whether the copied number formats come too, or the destination keeps its own.
        A copied cell brings its number format and lock even when it is empty, and when it has them from its row or
        column. Formats other than those are not modelled. The block is laid over `area` (the paste area by default)
        as often as it fits, the copy's first cell on `anchor` (the area's top-left by default) and on every whole
        block from it; `rows` (all of the area by default) are the rows it reaches."""
        dest, src = self.sheet, clip.sheet
        area = area or self._paste_area(clip)
        r1, c1, r2, c2 = area
        ar, ac = anchor or (r1, c1)
        rows = self._paste_rows(clip, area) if rows is None else rows
        if dest.protects():  # unlocked cells of a protected sheet take what is pasted, but not its formats
            dest.guard("contents", [(r, c1, r, c2) for r in rows])
            if what == "all":
                what, numfmts = "formulas", False
            elif what == "formats" or numfmts:
                raise VBANotImplemented("pasting number formats onto a protected sheet (unverified)")
        height, width = len(clip.rows), clip.c2 - clip.c1 + 1
        if what in ("all", "formats"):
            self.app.note("FORMATS-NOT-MODELLED")
        formats = numfmts or what in ("all", "formats")
        whole_rows = clip.c1 == 1 and clip.c2 == MAX_COL

        def across(offset):
            """The columns of the area where a copied column (its offset in the copy) lands."""
            c = c1 + (ac + offset - c1) % width
            return range(c, c2 + 1, width)

        src.fresh()
        copied = []  # the copied rows as they are now: a paste onto the copied cells must not read what it wrote
        for r in clip.rows:
            cells = {c: cell.copy() for c, cell in src.rows.get(r, {}).items() if clip.c1 <= c <= clip.c2}
            if what == "values" and any(cell.value is PENDING for cell in cells.values()):
                src.raw_value(r, next(c for c, cell in cells.items() if cell.value is PENDING))  # raises: unknown
            copied.append((r, cells, src.row_style(r)))
        # where a cell with no record still has a style on either sheet, the paste may have to give it one
        styled_src = [c - clip.c1 for c in {*src.col_formats, *src.col_unlocked} if clip.c1 <= c <= clip.c2]
        formatted = {c for o in styled_src for c in across(o)}
        formatted |= {c for c in {*dest.col_formats, *dest.col_unlocked} if c1 <= c <= c2}
        if whole_rows:
            self.app.note("PASTE-WHOLE-ROWS")
        pasted_formulas = []
        for dr in rows:
            src_r, cells, row_style = copied[(dr - ar) % height]
            if what != "formats":
                dest.book.forget_formulas(dest, dr, c1, dr, c2)
            old = dest.rows.get(dr, {})
            columns = {c for sc in cells for c in across(sc - clip.c1)} | {c for c in old if c1 <= c <= c2}
            if formats:
                columns |= formatted
                if (row_style is not None or dr in dest.row_formats) and not whole_rows:
                    columns |= set(range(c1, c2 + 1))
            had = {c: old[c].style if c in old else dest.inherited(dr, c) for c in columns}
            if formats and whole_rows:  # the row's own style comes with it
                dest.row_formats.pop(dr, None)
                dest.row_unlocked.discard(dr)
                if row_style is not None:
                    dest._set_row_style(dr, row_style)
            row = {c: cell for c, cell in old.items() if not c1 <= c <= c2}
            for c in columns:
                sc = clip.c1 + (c - ac) % width
                cell, prev = cells.get(sc), old.get(c)
                if not formats:
                    style = had[c]
                elif cell is not None:
                    style = cell.style
                else:
                    style = row_style if row_style is not None else src.col_style(sc)
                value = formula = None
                if what == "formats":
                    value, formula = (prev.value, prev.formula) if prev is not None else (None, None)
                elif cell is not None and what == "values":
                    value = cell.value
                elif cell is not None:
                    value, formula = cell.value, cell.formula
                    if formula is not None:
                        formula = shift_formula(formula, dr - src_r, c - sc)
                        value = PENDING if (dr - src_r or c - sc) else value
                        pasted_formulas.append((dr, c, formula))
                if value is not None or formula is not None or style != dest.inherited(dr, c):
                    row[c] = Cell(value, formula, *style)
            if row:
                dest.rows[dr] = row
            else:
                dest.rows.pop(dr, None)
            dest.changed = True
            if what != "formats":
                dest.book.cells_changed(dest, dr, c1, dr, c2)
        for dr, c, formula in pasted_formulas:
            dest.book.register_formula(dest, dr, c, formula, False)
        if select and dest is dest.book.active:
            dest.sel = area
            dest.active_cell = (r1, c1)

    # --- filling ---
    def _fill(self, step_r, step_c):
        """FillDown (1, 0), FillUp (-1, 0), FillRight (0, 1), FillLeft (0, -1): the range's first row (its last,
        its first column, its last) copied into the rest of it, formulas moved, formats too. A range one row deep
        fills from the row above it (and so on); none on the sheet's edge is error 1004. On a filtered sheet the
        rows the filter shows are filled."""
        self._single_sheet_target("a fill")
        r1, c1, r2, c2 = self.rect
        if step_r:
            edge = r1 if step_r == 1 else r2
            if r1 == r2:
                edge -= step_r
                if not 1 <= edge <= MAX_ROW:
                    raise VBARuntimeError(1004, "Fill method of Range class failed")
            else:
                r1, r2 = (r1 + 1, r2) if step_r == 1 else (r1, r2 - 1)
            source = (edge, c1, edge, c2)
        else:
            edge = c1 if step_c == 1 else c2
            if c1 == c2:
                edge -= step_c
                if not 1 <= edge <= MAX_COL:
                    raise VBARuntimeError(1004, "Fill method of Range class failed")
            else:
                c1, c2 = (c1 + 1, c2) if step_c == 1 else (c1, c2 - 1)
            source = (r1, edge, r2, edge)
        self.sheet.check_merged(r1, c1, r2, c2)
        clip = Clipboard(self.sheet, list(range(source[0], source[2] + 1)), source[1], source[3])
        area = (r1, c1, r2, c2)
        rows = [r for a in Range(self.sheet, *area)._written_areas() for r in range(a[0], a[2] + 1)]
        self._paste(clip, "all", True, area=area, anchor=source[:2], rows=rows, select=False)
        return True

    def m_filldown(self):
        return self._fill(1, 0)

    def m_fillup(self):
        return self._fill(-1, 0)

    def m_fillright(self):
        return self._fill(0, 1)

    def m_fillleft(self):
        return self._fill(0, -1)

    def m_autofill(self, destination=None, type=None):
        """AutoFill: the destination holds this range and runs on from it one way; the range is laid down over the
        rest, line by line (each column of a fill down, each row of a fill across). A line of formulas and blank
        cells repeats; so does any line with xlFillCopy. A cell filled alone is copied (a date counts on by days);
        a number alone in its line of a larger range counts on by 1; two numbers or more go on along their trend
        (least squares). Text that counts on (Item 1, Q1, months and weekdays) is not modelled."""
        if not isinstance(destination, Range):
            raise VBARuntimeError(1004, "AutoFill method of Range class failed")
        kind = FILL_DEFAULT if type is None else to_long(type)
        if kind not in (FILL_DEFAULT, FILL_COPY, FILL_SERIES, FILL_FORMATS, FILL_VALUES, FILL_DAYS):
            raise VBANotImplemented(f"AutoFill Type:={kind}")
        dest, sheet = destination, self.sheet
        self._single_sheet_target("AutoFill")
        r1, c1, r2, c2 = self.rect
        failed = VBARuntimeError(1004, "AutoFill method of Range class failed")
        if dest.sheet is not sheet or not (dest.r1 <= r1 and r2 <= dest.r2 and dest.c1 <= c1 and c2 <= dest.c2):
            raise failed
        down = (dest.c1, dest.c2) == (c1, c2) and (dest.r1 == r1) != (dest.r2 == r2)
        across = (dest.r1, dest.r2) == (r1, r2) and (dest.c1 == c1) != (dest.c2 == c2)
        if not (down or across):
            if dest.rect == self.rect or (dest.r1, dest.r2) != (r1, r2) and (dest.c1, dest.c2) != (c1, c2):
                if dest.rect == self.rect or not (dest.r1 == r1 and dest.c1 == c1):
                    raise failed
                raise VBANotImplemented("AutoFill into a range that runs on from the source both down and across")
            raise failed
        if down:
            area = (r2 + 1, c1, dest.r2, c2) if dest.r1 == r1 else (dest.r1, c1, r1 - 1, c2)
        else:
            area = (r1, c2 + 1, r2, dest.c2) if dest.c1 == c1 else (r1, dest.c1, r2, c1 - 1)
        hidden = sheet.hidden_rows - sheet.manual_hidden if sheet.in_filter_mode() else ()
        if any(area[0] <= r <= area[2] for r in hidden):
            raise VBANotImplemented("AutoFill into rows a filter hides (Excel fills nothing)")
        sheet.check_merged(*dest.rect)
        series = {} if kind in (FILL_COPY, FILL_FORMATS) else self._series(down, kind, area)
        what, numfmts = {FILL_FORMATS: ("formats", True), FILL_VALUES: ("formulas", False)}.get(kind, ("all", True))
        clip = Clipboard(sheet, list(range(r1, r2 + 1)), c1, c2)
        self._paste(clip, what, numfmts, area=area, anchor=(r1, c1), select=False)
        for (r, c), value in series.items():
            sheet.put(r, c, value)
        if series:
            self.app.note("AUTOFILL-SERIES")
        return True

    def _series(self, down, kind, area):
        """{cell of the area: value} for the lines of the source that count on rather than repeat."""
        sheet = self.sheet.fresh()
        r1, c1, r2, c2 = self.rect
        single = (r1, c1) == (r2, c2)
        lines = range(c1, c2 + 1) if down else range(r1, r2 + 1)
        out = {}
        for line in lines:
            cells = [(r, line) for r in range(r1, r2 + 1)] if down else [(line, c) for c in range(c1, c2 + 1)]
            kinds = [self._fill_kind(sheet, r, c) for r, c in cells]
            numbers = [i for i, k in enumerate(kinds) if k == "number"]
            if kind == FILL_SERIES:
                if len(cells) != 1 or kinds[0] != "number":
                    raise VBANotImplemented("AutoFill xlFillSeries of anything but one number to a line")
                values = _counting(sheet.raw_value(*cells[0]), 0, 1)
            elif kind == FILL_DAYS:
                if len(cells) != 1 or kinds[0] != "date":
                    raise VBANotImplemented("AutoFill xlFillDays of anything but one date to a line")
                values = _counting(sheet.raw_value(*cells[0]), 0, 1)
            elif set(kinds) <= {"formula", "blank", "text", "other"}:
                if "text" in kinds:
                    self.app.note("AUTOFILL-TEXT-COPIED")
                continue  # it repeats
            elif {"digits", "listed", "time"} & set(kinds):
                found = next(k for k in kinds if k in ("digits", "listed", "time"))
                raise VBANotImplemented(f"AutoFill of a series Excel makes of {'text' if found != 'time' else 'times'}")
            elif "date" in kinds:
                if not single:
                    raise VBANotImplemented("AutoFill of dates in a range of more than one cell")
                values = _counting(sheet.raw_value(*cells[0]), 0, 1)
            elif single:
                continue  # one number filled alone is copied
            elif len(numbers) == 1 and set(kinds) <= {"number", "formula", "blank"}:
                values = _counting(sheet.raw_value(*cells[numbers[0]]), numbers[0], len(cells))
            elif len(numbers) == len(cells):
                values = _trend([sheet.raw_value(*cell) for cell in cells])
            else:
                raise VBANotImplemented("AutoFill of a line that mixes numbers with text or with other numbers")
            positions = range(area[0], area[2] + 1) if down else range(area[1], area[3] + 1)
            start = r1 if down else c1
            for at in positions:
                value = values(at - start)
                if value is not None:
                    out[(at, line) if down else (line, at)] = value
        return out

    @staticmethod
    def _fill_kind(sheet, r, c):
        cell = sheet.get_cell(r, c)
        if cell is None or (cell.value is None and cell.formula is None):
            return "blank"
        if cell.formula is not None:
            return "formula"
        v = sheet.raw_value(r, c)
        if isinstance(v, (bool, XLError)):
            return "other"
        if isinstance(v, float):
            if is_date_format(cell.numfmt):
                return "date"
            return "time" if is_time_format(cell.numfmt) else "number"
        if any(ch.isdigit() for ch in v):
            return "digits"
        return "listed" if v.strip().lower() in _LISTED else "text"

    # --- filters ---
    def m_autofilter(
        self, field=None, criteria1=None, operator=None, criteria2=None, visibledropdown=None, subfield=None
    ):
        sheet = self.sheet
        sheet.guard("filter")
        af = sheet.autofilter
        if field is None and criteria1 is None and operator is None and criteria2 is None:
            if af is not None:
                sheet.remove_autofilter()
            else:
                sheet.autofilter = AutoFilterState(*self._filter_area())
                sheet.changed = True
            return True
        new = af is None
        if new:
            af = AutoFilterState(*self._filter_area())
        else:
            r1, c1, r2, c2 = self.rect
            overlaps = not (r2 < af.r1 or r1 > af.r2 or c2 < af.c1 or c1 > af.c2)
            if not overlaps:
                raise VBANotImplemented("AutoFilter on a range outside the sheet's existing AutoFilter (unverified)")
            if self.rect != af.rect:
                self.app.note(
                    "AF-EXISTING-RANGE",
                    f"{sheet.name}: code names {area_text(*self.rect)}, filter is {area_text(*af.rect)}",
                )
        f = to_long(field)
        if not 1 <= f <= af.c2 - af.c1 + 1:  # refused before the sheet is given a filter
            raise VBARuntimeError(1004, "AutoFilter method of Range class failed")
        criterion = None if criteria1 is None and operator is None else self._criterion(criteria1, operator, criteria2)
        sheet.autofilter = af
        if criterion is None:
            af.criteria.pop(f, None)
        else:
            af.criteria[f] = criterion
        apply_filter(sheet)
        return True

    def _filter_area(self):
        """What a new AutoFilter covers: this range; for one cell, its current region; for one row (a header row, or
        Rows(1)), the data below it, down its current region and across the columns of the row that hold data."""
        if (self.r1, self.c1) == (self.r2, self.c2):
            return self.get_currentregion().rect
        if self.r1 != self.r2:
            return self.rect
        row = self.sheet.fresh().rows.get(self.r1) or {}
        first = min((c for c in row if self.c1 <= c <= self.c2 and not self.sheet.is_empty(self.r1, c)), default=None)
        if first is None:
            return self.rect
        _, cc1, rr2, cc2 = Range(self.sheet, self.r1, first, self.r1, first).get_currentregion().rect
        self.app.note("AF-HEADER-ROW", f"{self.sheet.name}: {area_text(*self.rect)}")
        return (self.r1, max(self.c1, cc1), max(rr2, self.r1), min(self.c2, cc2))

    def _criterion(self, criteria1, operator, criteria2):
        op = to_long(operator) if operator is not None else None
        if op == 7:  # xlFilterValues
            values = default_value(criteria1)
            items = values.items_1d() if isinstance(values, VBAArray) else [values]
            texts = [to_str(v) for v in items]
            return ValuesCriterion([t for t in texts if t != ""], blank=any(t == "" for t in texts))
        if op not in (None, 1, 2):
            raise VBANotImplemented(f"AutoFilter Operator:={op}")
        first = criterion_from_text(criteria1) if criteria1 is not None else None
        if criteria2 is None:
            if first is None:
                raise VBARuntimeError(1004, "AutoFilter method of Range class failed")
            return first
        second = criterion_from_text(criteria2)
        return BoolCriterion("or" if op == 2 else "and", first, second)

    # --- not modelled yet ---
    def m_find(self, *args, **kwargs):
        raise VBANotImplemented("Range.Find")

    def m_sort(
        self,
        key1=None,
        order1=None,
        key2=None,
        type=None,
        order2=None,
        key3=None,
        order3=None,
        header=None,
        ordercustom=None,
        matchcase=None,
        orientation=None,
        sortmethod=None,
        dataoption1=None,
        dataoption2=None,
        dataoption3=None,
        subfield1=None,
        subfield2=None,
        subfield3=None,
    ):
        """Range.Sort: up to three keys, ascending unless Order:=xlDescending."""
        from .sorting import SortKey, sort_range  # sorting builds on Range, so it is imported here

        if ordercustom not in (None, 1) or type is not None:
            raise VBANotImplemented("Range.Sort with a custom order or Type")
        keys = []
        for key, order, option in (
            (key1, order1, dataoption1),
            (key2, order2, dataoption2),
            (key3, order3, dataoption3),
        ):
            if key is None:
                continue
            keys.append(
                SortKey(key if isinstance(key, Range) else self.sheet.range_from_text(to_str(key)), order, 0, option)
            )
        sort_range(self.sheet, self.rect, keys, header if header is not None else 2, matchcase, orientation)
        return True

    def get_specialcells(self, type=None, value=None):
        raise VBANotImplemented("Range.SpecialCells")
