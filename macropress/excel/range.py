"""Range: reading and writing cells, moving around the sheet, copy and paste, filters and sorting."""

import datetime as dt
import re

from .. import formula as fx
from ..errors import VBANotImplemented, VBARuntimeError
from ..values import (
    EMPTY,
    NULL,
    VBAArray,
    VBAObject,
    XLError,
    date_to_serial,
    default_value,
    is_numeric_value,
    parse_number_text,
    serial_to_date,
    to_long,
    to_str,
)
from .addresses import area_text, col_num, is_date_format, shift_formula
from .constants import MAX_COL, MAX_ROW, PENDING
from .filters import AutoFilterState, BoolCriterion, ValuesCriterion, apply_filter, criterion_from_text
from .storage import Cell

# ---------------------------------------------------------------------------
# Range
# ---------------------------------------------------------------------------


class Clipboard:
    def __init__(self, sheet, rows, c1, c2):
        self.sheet, self.rows, self.c1, self.c2 = sheet, rows, c1, c2


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
        if not value2 and isinstance(v, float) and is_date_format(cell.numfmt):
            return serial_to_date(v)
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
        self._single_sheet_target("writing values")
        self.sheet.check_merged(*self.rect)
        if isinstance(value, VBAArray):
            self._write_array(value)
            return
        stored, formula, fmt = self._to_cell(value)
        if (self.r2 - self.r1 + 1) * (self.c2 - self.c1 + 1) > 5_000_000:
            raise VBANotImplemented("writing one value into a very large range")
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                self.sheet.put(r, c, stored, formula, fmt)

    let_value2 = let_value

    def _write_array(self, arr):
        if arr.bounds is None:
            raise VBARuntimeError(13, "Type mismatch")
        dims = arr.bounds
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                i, j = r - self.r1, c - self.c1
                try:
                    if len(dims) == 2:
                        v = arr.get([dims[0][0] + i, dims[1][0] + j])
                    else:
                        v = arr.get([dims[0][0] + j])
                except VBARuntimeError:
                    v = XLError("#N/A")
                stored, formula, fmt = self._to_cell(v)
                self.sheet.put(r, c, stored, formula, fmt)

    def _to_cell(self, value):
        """A VBA value written to a cell -> (stored value, formula, number format)."""
        v = default_value(value)
        if v is EMPTY or v is NULL:
            return None, None, None
        if isinstance(v, bool):
            return v, None, None
        if isinstance(v, XLError):
            return v, None, None
        if is_numeric_value(v):
            return float(v), None, None
        if isinstance(v, dt.datetime):
            fmt = "m/d/yyyy h:mm" if (v.hour, v.minute, v.second) != (0, 0, 0) else "m/d/yyyy"
            return date_to_serial(v), None, fmt
        if isinstance(v, str):
            if v == "":
                self.app.note("WRITE-EMPTY-STRING")
                return None, None, None
            if v.startswith("="):
                return PENDING, v[1:], None
            if v.startswith("'"):
                return v[1:], None, None
            num = parse_number_text(v)
            if num is not None and not v.strip().lower().startswith("&h"):
                self.app.note("WRITE-TEXT-CONVERSION", f"{v!r} written as a number")
                return float(num), None, None
            if v.strip().upper() in ("TRUE", "FALSE"):
                self.app.note("WRITE-TEXT-CONVERSION", f"{v!r} written as a Boolean")
                return v.strip().upper() == "TRUE", None, None
            if (
                re.match(r"^\s*\d{1,4}[-/.]\d{1,2}([-/.]\d{1,4})?(\s+\d{1,2}:\d{2}(:\d{2})?)?\s*$", v)
                or re.match(r"^\s*\d{1,2}:\d{2}(:\d{2})?\s*([AaPp][Mm])?\s*$", v)
                or re.match(r"^\s*\d{1,2}[-\s][A-Za-z]{3,9}([-\s]\d{2,4})?\s*$", v)
            ):
                raise VBANotImplemented(f"writing {v!r} to a cell: Excel may turn it into a date")
            return v, None, None
        raise VBARuntimeError(13, "Type mismatch")

    def get_formula(self):
        if (self.r1, self.c1) != (self.r2, self.c2):
            raise VBANotImplemented("Formula of a multi-cell range")
        cell = self.sheet.get_cell(self.r1, self.c1)
        if cell is None:
            return ""
        if cell.formula is not None:
            return "=" + fx.display_syntax(cell.formula)
        return to_str(self._vba_value(self.r1, self.c1)) if cell.value is not None else ""

    def _write_formula(self, value, dynamic):
        self._single_sheet_target("writing formulas")
        text = to_str(value)
        for r in range(self.r1, self.r2 + 1):
            for c in range(self.c1, self.c2 + 1):
                if text.startswith("="):  # relative references move with each cell, as when Excel fills a range
                    self.sheet.put(r, c, PENDING, shift_formula(text[1:], r - self.r1, c - self.c1), dynamic=dynamic)
                else:
                    stored, formula, fmt = self._to_cell(text)
                    self.sheet.put(r, c, stored, formula, fmt)

    def let_formula(self, value):
        """Formula: a formula that returns one value (an array result gives its top-left value)."""
        self._write_formula(value, dynamic=False)

    def let_formula2(self, value):
        """Formula2: dynamic-array semantics; an array result spills into the cells below and to the right."""
        self._write_formula(value, dynamic=True)

    get_formula2 = get_formula

    def get_text(self):
        if (self.r1, self.c1) != (self.r2, self.c2):
            raise VBANotImplemented("Text of a multi-cell range")
        return self.sheet.cell_text(self.r1, self.c1)

    def get_numberformat(self):
        cell = self.sheet.get_cell(self.r1, self.c1)
        return (cell.numfmt if cell is not None else None) or "General"

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
            target, rng = self.sheet.hidden_rows, range(self.r1, self.r2 + 1)
        elif self.r1 == 1 and self.r2 == MAX_ROW:
            target, rng = self.sheet.hidden_cols, range(self.c1, self.c2 + 1)
        else:
            raise VBARuntimeError(1004, "Unable to set the Hidden property of the Range class")
        for i in rng:
            (target.add if hide else target.discard)(i)
            if target is self.sheet.hidden_rows:
                (self.sheet.manual_hidden.add if hide else self.sheet.manual_hidden.discard)(i)
        self.sheet.changed = True

    def get_end(self, direction):
        d = to_long(direction)
        sheet = self.sheet.fresh()
        r, c = self.r1, self.c1
        empty = sheet.is_empty
        if d == -4121:  # xlDown
            if not empty(r, c) and r < MAX_ROW and not empty(r + 1, c):
                while r < MAX_ROW and not empty(r + 1, c):
                    r += 1
            else:
                r = min((x for x in sheet.rows if x > r and not empty(x, c)), default=MAX_ROW)
            self._note_hidden_rows(self.r1, r)
            return Range(sheet, r, c, r, c)
        if d == -4162:  # xlUp
            if not empty(r, c) and r > 1 and not empty(r - 1, c):
                while r > 1 and not empty(r - 1, c):
                    r -= 1
            else:
                r = max((x for x in sheet.rows if x < r and not empty(x, c)), default=1)
            self._note_hidden_rows(r, self.r1)
            return Range(sheet, r, c, r, c)
        if d in (-4161, -4159):  # xlToRight, xlToLeft
            step = 1 if d == -4161 else -1
            limit = MAX_COL if step == 1 else 1
            if not empty(r, c) and c != limit and not empty(r, c + step):
                while c != limit and not empty(r, c + step):
                    c += step
            else:
                cols = [x for x in (sheet.rows.get(r) or {}) if (x > c if step == 1 else x < c) and not empty(r, x)]
                c = (min if step == 1 else max)(cols, default=limit)
            return Range(sheet, r, c, r, c)
        raise VBARuntimeError(5, "Invalid procedure call or argument")

    def _note_hidden_rows(self, a, b):
        lo, hi = min(a, b), max(a, b)
        if any(lo <= h <= hi for h in self.sheet.hidden_rows):
            self.app.note("END-IGNORES-HIDDEN")

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
        for s in targets:
            s.clear_rect(*self.rect, formats=formats)
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

    def m_pastespecial(self, paste=None, operation=None, skipblanks=None, transpose=None):
        clip = self.app.clipboard
        if clip is None:
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
            -4122: ("formats", False),  # xlPasteFormats
            8: ("formats", False),  # xlPasteColumnWidths
        }
        if mode not in types:
            raise VBANotImplemented(f"PasteSpecial Paste:={mode}")
        self._paste(clip, *types[mode])
        return True

    def _paste_area(self, clip):
        """Where a paste lands: the copied block, from this range's top-left cell. Excel pastes it there when this range
        is one cell or the block's size; a range of another size it fills by repeating the block, or refuses, which is
        not modelled. A block that would run off the sheet is Excel's error 1004."""
        height, width = len(clip.rows), clip.c2 - clip.c1 + 1
        size = (self.r2 - self.r1 + 1, self.c2 - self.c1 + 1)
        if size not in ((1, 1), (height, width)):
            raise VBANotImplemented(f"pasting {height}x{width} copied cells into {area_text(*self.rect)}, another size")
        r2, c2 = self.r1 + height - 1, self.c1 + width - 1
        if r2 > MAX_ROW or c2 > MAX_COL:
            raise VBARuntimeError(1004, "The copy area and the paste area are not the same size")
        return self.r1, self.c1, r2, c2

    def _paste(self, clip, what, numfmts):
        """Paste the copied cells. what: "all" (formulas, constants and formats), "formulas" (formulas and constants),
        "values" or "formats"; numfmts: whether the copied number formats come too, or the destination keeps its own.
        Formats other than number formats are not modelled."""
        dest = self.sheet
        r1, c1, r2, c2 = self._paste_area(clip)
        if what in ("all", "formats"):
            self.app.note("FORMATS-NOT-MODELLED")
        if what == "formats":
            return
        dcol = c1 - clip.c1
        whole_rows = clip.c1 == 1 and clip.c2 == MAX_COL
        snapshot = []
        clip.sheet.fresh()
        for r in clip.rows:
            row = clip.sheet.rows.get(r, {})
            cells = {}
            for c, cell in row.items():
                if clip.c1 <= c <= clip.c2:
                    if cell.value is PENDING and what == "values":
                        clip.sheet.raw_value(r, c)  # raises: value unknown
                    cells[c - clip.c1] = cell
            snapshot.append((r, cells))
        if whole_rows:
            self.app.note("PASTE-WHOLE-ROWS")
        pasted_formulas = []
        for i, (src_r, cells) in enumerate(snapshot):
            dr = r1 + i
            dest.book.forget_formulas(dest, dr, c1, dr, c2)
            if whole_rows:
                old = dest.rows.pop(dr, {})
                new_row = {}
            else:
                old = dict(dest.rows.get(dr, {}))  # the destination's cells as they were: their formats
                dest.clear_rect(dr, c1, dr, c2)
                new_row = dest.rows.setdefault(dr, {})
            for off, cell in cells.items():
                c = c1 + off
                prev = old.get(c) or new_row.get(c)
                numfmt = cell.numfmt if numfmts else (prev.numfmt if prev is not None else None)
                if what == "values":
                    if cell.value is None:
                        continue
                    new_row[c] = Cell(cell.value, None, numfmt)
                else:
                    formula = None
                    value = cell.value
                    if cell.formula is not None:
                        formula = shift_formula(cell.formula, dr - src_r, dcol)
                        value = PENDING if (dr - src_r or dcol) else cell.value
                        pasted_formulas.append((dr, c, formula))
                    new_row[c] = Cell(value, formula, numfmt)
            if new_row:
                dest.rows[dr] = new_row
            elif dr in dest.rows and not dest.rows[dr]:
                del dest.rows[dr]
            dest.changed = True
            dest.book.cells_changed(dest, dr, c1, dr, c2)
        for dr, c, formula in pasted_formulas:
            dest.book.register_formula(dest, dr, c, formula, False)
        if dest is dest.book.active:
            dest.sel = (r1, c1, r2, c2)
            dest.active_cell = (r1, c1)

    # --- filters ---
    def m_autofilter(
        self, field=None, criteria1=None, operator=None, criteria2=None, visibledropdown=None, subfield=None
    ):
        sheet = self.sheet
        af = sheet.autofilter
        if field is None and criteria1 is None and operator is None and criteria2 is None:
            if af is not None:
                sheet.remove_autofilter()
            else:
                sheet.autofilter = AutoFilterState(*self._filter_area())
                sheet.changed = True
            return True
        if af is None:
            af = sheet.autofilter = AutoFilterState(*self._filter_area())
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
        if not 1 <= f <= af.c2 - af.c1 + 1:
            raise VBARuntimeError(1004, "AutoFilter method of Range class failed")
        if criteria1 is None and operator is None:
            af.criteria.pop(f, None)
        else:
            af.criteria[f] = self._criterion(criteria1, operator, criteria2)
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
