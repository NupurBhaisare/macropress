"""Worksheet: its cells, view state (selection, hidden rows), filter and sort state."""

import re

from ..errors import VBANotImplemented, VBARuntimeError
from ..numfmt import after_entry, shown
from ..values import (
    VBAObject,
    default_value,
    format_double,
    to_long,
    to_str,
)
from .addresses import _COLS, col_letters, col_num, parse_area
from .constants import MAX_COL, MAX_ROW, PENDING
from .protection import ALLOW, Password, ProtectionObject, SheetProtection, check_password
from .range import Range
from .sorting import SortObject
from .storage import NORMAL, Cell

_COL = re.compile(r"^\$?([A-Za-z]{1,3})$")
# What Excel refuses in a sheet's name: one of :\/?*[], an apostrophe at either end, or "History".
_BAD_NAME = re.compile(r"[:\\/?*\[\]]|^'|'$|^history$", re.I)


def _as_read(new, read, text):
    """The protection the file had, where a macro protects the sheet again just as it was: the file keeps its own."""
    if read is None or (new.allow, new.objects, new.scenarios) != (read.allow, read.objects, read.scenarios):
        return new
    same = read.password.matches(text) if read.password is not None else not text
    return read if same else new


class Sheet(VBAObject):
    vba_type_name = "Worksheet"

    def __init__(self, book, name, code_name=None, state="visible"):
        self.book = book
        self.name = name
        self.code_name = code_name
        self.state = state
        self.rows = {}  # row -> {col: Cell}
        self.col_widths = []  # [(first column, last column, width)] of the columns given a width
        self.default_width = 9.14  # of the others: 8.43 characters and the cell's padding
        self.row_formats = {}  # row -> the number format of a row formatted as a whole (None: General)
        self.col_formats = {}  # col -> the number format of a column formatted as a whole, when not General
        self.row_unlocked = set()  # rows formatted as a whole whose cells are not locked
        self.col_unlocked = set()  # columns whose cells are not locked
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
        self.protection = None  # SheetProtection while the sheet is protected
        self.protection_read = None  # the protection the file had, kept as it was if a macro gives it back
        self.ui_only = False  # protected with UserInterfaceOnly: macros may change what a user may not
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

    def shown_text(self, r, c):
        """Range.Text: what the cell shows. Text shows whole. A number, TRUE or an error shows whole only where its
        column has room for it; where it has not, Excel rounds the number to fit or shows ####, which is not
        modelled: it turns on the column's width and its font."""
        text = self.cell_text(r, c)
        if text and not isinstance(self.raw_value(r, c), str) and len(text) > self.room(c):
            raise VBANotImplemented(
                f"the displayed text of {self.name}!{col_letters(c)}{r}: {text} may not fit its column as it stands"
            )
        return text

    def room(self, c):
        """How many characters of the default font a column shows."""
        width = next((w for first, last, w in self.col_widths if first <= c <= last), self.default_width)
        return int(width - 0.71) if c not in self.hidden_cols else 0  # a width in the file includes the cell's padding

    def cell_text(self, r, c):
        """The text the cell shows in a column wide enough for it: what a filter matches. A number is shown by its
        number format, for the formats numfmt.shown covers."""
        cell = self.get_cell(r, c)
        if cell is None or cell.value is None:
            return ""
        v = self.raw_value(r, c)
        if isinstance(v, bool):
            return "TRUE" if v else "FALSE"
        if isinstance(v, (int, float)):
            text = shown(float(v), cell.numfmt)
            if text is None:
                raise VBANotImplemented(
                    f"the displayed text of {self.name}!{col_letters(c)}{r}: "
                    f"{format_double(float(v))} in the number format {cell.numfmt or 'General'}"
                )
            return text
        return str(v)

    # --- formats: a cell's number format, and whether it is locked
    def inherited(self, r, c):
        """The style (number format, locked) of a cell the sheet holds no record of: its row's, or else its column's."""
        if r in self.row_formats:
            return self.row_formats[r], r not in self.row_unlocked
        return self.col_formats.get(c), c not in self.col_unlocked

    def inherited_format(self, r, c):
        """The number format of a cell the sheet holds no record of: its row's, or else its column's."""
        return self.inherited(r, c)[0]

    def style_at(self, r, c):
        cell = self.rows.get(r, {}).get(c)
        return cell.style if cell is not None else self.inherited(r, c)

    def format_at(self, r, c):
        """The cell's number format; None for General."""
        return self.style_at(r, c)[0]

    def locked_at(self, r, c):
        return self.style_at(r, c)[1]

    def col_style(self, c):
        """The style of a column: General and locked unless it is formatted whole."""
        return self.col_formats.get(c), c not in self.col_unlocked

    def row_style(self, r):
        """The style of a row formatted whole, or None for one that is not."""
        return (self.row_formats[r], r not in self.row_unlocked) if r in self.row_formats else None

    def _setcol_style(self, c, style):
        numfmt, locked = style
        if numfmt is None:
            self.col_formats.pop(c, None)
        else:
            self.col_formats[c] = numfmt
        (self.col_unlocked.discard if locked else self.col_unlocked.add)(c)

    def _set_row_style(self, r, style):
        self.row_formats[r] = style[0]
        (self.row_unlocked.discard if style[1] else self.row_unlocked.add)(r)

    def _store(self, r, c, value, formula, numfmt, locked=True):
        """Keep the cell, or no record of it when it is empty and formatted as its row or column is anyway."""
        if value is None and formula is None and (numfmt, locked) == self.inherited(r, c):
            row = self.rows.get(r)
            if row is not None:
                row.pop(c, None)
                if not row:
                    del self.rows[r]
        else:
            self.rows.setdefault(r, {})[c] = Cell(value, formula, numfmt, locked)

    def put(self, r, c, value, formula=None, numfmt=None, dynamic=False):
        """Write a cell. numfmt is the format the value brings (a Date's, the 0% of "5%"): a General cell takes it,
        and so does one with a format of Excel's own for something else (see numfmt.after_entry). Any other cell
        keeps its format, even emptied (a cell with no value or formula stays for its format)."""
        self.changed = True
        self.book.forget_formulas(self, r, c, r, c)
        have, locked = self.style_at(r, c)
        self._store(r, c, value, formula, after_entry(have, numfmt), locked)
        if formula is not None:
            self.book.register_formula(self, r, c, formula, dynamic)
        self.book.cells_changed(self, r, c, r, c)

    def clear_rect(self, r1, c1, r2, c2, formats=False):
        """Clear the cells' contents; their number formats stay, as with ClearContents, unless formats (Clear), which
        leaves the cells General and locked."""
        self.changed = True
        self.book.forget_formulas(self, r1, c1, r2, c2)
        for r in [r for r in self.rows if r1 <= r <= r2]:
            for c in [c for c in self.rows[r] if c1 <= c <= c2]:
                cell = self.rows[r][c]
                self._store(r, c, None, None, *(NORMAL if formats else cell.style))
        if formats:
            self._reset_formats(r1, c1, r2, c2)
        self.book.cells_changed(self, r1, c1, r2, c2)

    def _reset_formats(self, r1, c1, r2, c2):
        """Make the rectangle's cells General and locked where a format of their row or column would say otherwise."""
        if c1 == 1 and c2 == MAX_COL:
            for r in [r for r in self.row_formats if r1 <= r <= r2]:
                del self.row_formats[r]
                self.row_unlocked.discard(r)
        if r1 == 1 and r2 == MAX_ROW:
            for c in {c for c in (*self.col_formats, *self.col_unlocked) if c1 <= c <= c2}:
                self._setcol_style(c, NORMAL)
        rows = [r for r in self.row_formats if r1 <= r <= r2 and self.inherited(r, 1) != NORMAL]
        cols = {c for c in (*self.col_formats, *self.col_unlocked) if c1 <= c <= c2}
        if len(rows) * (c2 - c1 + 1) + len(cols) * (r2 - r1 + 1) > 1_000_000:
            raise VBANotImplemented("Clear on part of a great many formatted rows or columns")
        for r in rows:
            for c in range(c1, c2 + 1):
                self.rows.setdefault(r, {}).setdefault(c, Cell())
        for c in cols:
            for r in range(r1, r2 + 1):
                if r not in self.row_formats:
                    self.rows.setdefault(r, {}).setdefault(c, Cell())

    def restyle(self, r1, c1, r2, c2, **change):
        """Give the rectangle's cells a number format (numfmt=) or a lock (locked=), the rest of their styles kept, as
        Range.NumberFormat and Range.Locked do. Rows and columns formatted whole take the change themselves, so that
        a cell written there later has it; so do those of a range that runs to the sheet's last row or column, with
        their cells before the range kept as they were. Where a row's format would hide a column's, or the reverse,
        the cell is given a record of its own, as Excel gives it."""
        self.changed = True
        rows_whole, cols_whole = c1 == 1 and c2 == MAX_COL, r1 == 1 and r2 == MAX_ROW
        if rows_whole and cols_whole:
            self._restyle_sheet(change)
        elif r2 == MAX_ROW and c2 == MAX_COL:
            raise VBANotImplemented("formatting a range that runs to both the last row and the last column")
        elif c2 == MAX_COL:
            self._restyle_rows(r1, r2, c1, change)
        elif r2 == MAX_ROW:
            self._restyle_cols(c1, c2, r1, change)
        else:
            if (r2 - r1 + 1) * (c2 - c1 + 1) > 1_000_000:
                raise VBANotImplemented("formatting more than a million cells that are not whole rows or columns")
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    self._restyle_cell(r, c, change)

    @staticmethod
    def _changed(style, change):
        return change.get("numfmt", style[0]), change.get("locked", style[1])

    def _restyle_cell(self, r, c, change):
        cell = self.rows.get(r, {}).get(c)
        numfmt, locked = self._changed(self.style_at(r, c), change)
        if cell is not None and (cell.value is not None or cell.formula is not None):
            cell.numfmt, cell.locked = numfmt, locked
        else:  # an empty cell formatted as its row or column is anyway needs no record
            self._store(r, c, None, None, numfmt, locked)

    _MAX_RECORDS = 1_000_000

    def _count(self, made):
        if made > self._MAX_RECORDS:
            raise VBANotImplemented("formatting rows or columns whose cells must keep formats of their own: too many")

    def _keep(self, made):
        """Give each (r, c, style) a record of that style, where the cell has none and the style is not what it now
        takes from its row or column: what a cell keeps of its own."""
        self._count(len(made))
        for r, c, style in made:
            if c not in self.rows.get(r, {}) and style != self.inherited(r, c):
                self.rows.setdefault(r, {})[c] = Cell(None, None, *style)

    def _restyle_sheet(self, change):
        for row in self.rows.values():
            for cell in row.values():
                cell.numfmt, cell.locked = self._changed(cell.style, change)
        for r in self.row_formats:
            self._set_row_style(r, self._changed(self.inherited(r, 1), change))
        if "numfmt" in change:
            self.col_formats = (
                {} if change["numfmt"] is None else dict.fromkeys(range(1, MAX_COL + 1), change["numfmt"])
            )
        if "locked" in change:
            self.col_unlocked = set() if change["locked"] else set(range(1, MAX_COL + 1))

    def _restyle_rows(self, r1, r2, c1, change):
        styled = {c: self.col_style(c) for c in {*self.col_formats, *self.col_unlocked}}
        plain_after = self._changed(NORMAL, change)
        # In a row not formatted whole before, a cell of a formatted column keeps its column's style, the change made,
        # and a cell before the range keeps its column's style, each where the row's new style would say otherwise.
        plain_made = [(c, self._changed(st, change)) for c, st in styled.items() if c >= c1]
        plain_made += [(c, styled.get(c, NORMAL)) for c in range(1, c1)]
        plain_made = [(c, st) for c, st in plain_made if st != plain_after]
        made = []
        for r in range(r1, r2 + 1):
            if r in self.row_formats:
                before = self.inherited(r, 1)
                after = self._changed(before, change)
                made += [(r, c, before) for c in range(1, c1)] if after != before else []
            else:
                after = plain_after
                made += [(r, c, st) for c, st in plain_made]
            self._count(len(made))
            for c, cell in (self.rows.get(r) or {}).items():
                if c >= c1:
                    cell.numfmt, cell.locked = self._changed(cell.style, change)
            self._set_row_style(r, after)
        self._keep(made)

    def _restyle_cols(self, c1, c2, r1, change):
        cols = {c: self.col_style(c) for c in range(c1, c2 + 1)}
        moved = [c for c, st in cols.items() if self._changed(st, change) != st]
        # the cells before the range, in rows not formatted whole, keep their columns' styles as they were
        plain = [r for r in range(1, r1) if r not in self.row_formats] if moved else []
        self._count(len(plain) * len(moved))
        made = [(r, c, cols[c]) for r in plain for c in moved]
        # a row formatted whole hides the columns' style: its cells in the range keep the row's, the change made
        for r in [r for r in self.row_formats if r >= r1]:
            before = self.inherited(r, 1)
            after = self._changed(before, change)
            if after != before:
                made += [(r, c, after) for c in cols]
            self._count(len(made))
        for r, row in self.rows.items():
            if r >= r1:
                for c, cell in row.items():
                    if c1 <= c <= c2:
                        cell.numfmt, cell.locked = self._changed(cell.style, change)
        for c, st in cols.items():
            self._setcol_style(c, self._changed(st, change))
        self._keep(made)

    def styles_in(self, r1, c1, r2, c2, part):
        """The distinct number formats (part 0) or locks (part 1) of the rectangle's cells, stopping at two."""
        found = set()
        col_styles = {c: self.col_style(c)[part] for c in {*self.col_formats, *self.col_unlocked} if c1 <= c <= c2}
        plain_cols = (c2 - c1 + 1) - len(col_styles)
        rows = sorted({r for r in (*self.rows, *self.row_formats) if r1 <= r <= r2})
        for r in rows:
            row = self.rows.get(r) or {}
            cells = [cell for c, cell in row.items() if c1 <= c <= c2]
            found.update(cell.style[part] for cell in cells)
            if len(cells) < c2 - c1 + 1:  # cells with no record take the row's style, or their columns'
                if r in self.row_formats:
                    found.add(self.inherited(r, 1)[part])
                else:
                    found.update(v for c, v in col_styles.items() if c not in row)
                    if plain_cols - sum(1 for c in row if c1 <= c <= c2 and c not in col_styles) > 0:
                        found.add(NORMAL[part])
            if len(found) > 1:
                return found
        if (r2 - r1 + 1) > len(rows):  # rows with neither records nor a format of their own: the columns' styles
            found.update(col_styles.values())
            if plain_cols:
                found.add(NORMAL[part])
        return found

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
        # The cells the sheet holds, and the rows and columns formatted whole: a formatted column counts as used
        # (on its first row, if nothing else says which), as a formatted row does.
        rows = [r for r, row in self.rows.items() if row] + list(self.row_formats)
        cols = [c for row in self.rows.values() for c in row] + list({*self.col_formats, *self.col_unlocked})
        if not rows and not cols:
            return None
        r1, c1, r2, c2 = min(rows, default=1), min(cols, default=1), max(rows, default=1), max(cols, default=1)
        if self.dimension:
            r2, c2 = max(r2, self.dimension[0]), max(c2, self.dimension[1])
        return (r1, c1, r2, c2)

    def in_filter_mode(self):
        return self.autofilter is not None and bool(self.autofilter.criteria)

    # --- protection ---
    def protects(self):
        """Whether the sheet's protection holds macros back: it is protected, and not with UserInterfaceOnly."""
        return self.protection is not None and not self.ui_only

    def locks(self, r1, c1, r2, c2):
        """Whether the rectangle's cells are locked: True (all), False (none) or None (some are)."""
        found = self.styles_in(r1, c1, r2, c2, 1)
        return found.pop() if len(found) == 1 else None

    def guard(self, what, rects=()):
        """Error 1004 where the sheet's protection refuses a macro: "contents" (a write, clear or paste) of any locked
        cell of rects, "format" (a number format, unless AllowFormattingCells), "lock" (Locked), "rows" or "columns"
        (hiding them, unless AllowFormattingRows or AllowFormattingColumns), "sort" (unless AllowSorting, and never
        of locked cells), "filter" (the AutoFilter method, whatever AllowFiltering says)."""
        if not self.protects():
            return
        self.book.app.note("PROTECTED-SHEET", f"{self.name}: {what}")
        allow = self.protection.allow
        refused = {
            "contents": lambda: any(self.locks(*rect) is not False for rect in rects),
            "format": lambda: "allowformattingcells" not in allow,
            "lock": lambda: True,
            "rows": lambda: "allowformattingrows" not in allow,
            "columns": lambda: "allowformattingcolumns" not in allow,
            "sort": lambda: "allowsorting" not in allow or any(self.locks(*rect) is not False for rect in rects),
            "filter": lambda: True,
        }[what]()
        if refused:
            raise VBARuntimeError(1004, f"The cell or chart you're trying to change is on a protected sheet ({what})")

    def m_protect(
        self, password=None, drawingobjects=None, contents=None, scenarios=None, userinterfaceonly=None,
        allowformattingcells=None, allowformattingcolumns=None, allowformattingrows=None, allowinsertingcolumns=None,
        allowinsertingrows=None, allowinsertinghyperlinks=None, allowdeletingcolumns=None, allowdeletingrows=None,
        allowsorting=None, allowfiltering=None, allowusingpivottables=None,
    ):  # fmt: skip
        """Protect the sheet. Options left out are off; a sheet protected again takes the options given, and keeps its
        password. UserInterfaceOnly, once given, holds for the session, as it does in Excel."""
        given = locals()
        if contents is not None and not bool(default_value(contents)):
            raise VBANotImplemented("Protect Contents:=False (a sheet whose cells are not protected)")
        text = to_str(password) if password is not None else ""
        on = lambda v, default: default if v is None else bool(default_value(v))  # noqa: E731
        allow = {option for option in ALLOW if on(given[option], False)}
        new = SheetProtection(
            Password(text=text) if text else None, allow, on(drawingobjects, True), on(scenarios, True)
        )
        new = _as_read(new, self.protection_read, text)
        old = self.protection
        if old is not None:
            same_options = (old.allow, old.objects, old.scenarios) == (new.allow, new.objects, new.scenarios)
            same_password = old.password.matches(text) if old.password is not None else not text
            if same_options:
                new = old  # nothing changes: not the password either
            elif not same_password:
                raise VBANotImplemented("protecting a protected sheet again with other options and another password")
            else:
                new.password = old.password
        if userinterfaceonly is not None:
            if not bool(default_value(userinterfaceonly)) and self.ui_only:
                raise VBANotImplemented("Protect UserInterfaceOnly:=False after UserInterfaceOnly:=True (unverified)")
            self.ui_only = self.ui_only or bool(default_value(userinterfaceonly))
        self.protection = new
        self.changed = True

    def m_unprotect(self, password=None):
        p = self.protection
        if p is None:
            return
        if password is None and p.password is not None:  # Excel asks for it
            password = self.book.app.ask(f"the password to unprotect the sheet {self.name}")
        if p.password is not None:
            check_password(p.password, to_str(password))
        self.protection = None
        self.changed = True

    def get_protectcontents(self):
        return self.protection is not None

    def get_protectdrawingobjects(self):
        return self.protection is not None and self.protection.objects

    def get_protectscenarios(self):
        return self.protection is not None and self.protection.scenarios

    def get_protectionmode(self):
        return self.ui_only

    def get_protection(self):
        return ProtectionObject(self)

    # --- VBA members ---
    def get_name(self):
        return self.name

    def let_name(self, value):
        self.book.guard_structure("rename a sheet")
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

    def let_visible(self, value):
        """Show or hide the sheet: True, -1 or 1 shows it, False or 0 hides it, 2 (xlSheetVeryHidden) hides it from
        the user's own Unhide. The last sheet shown cannot be hidden. When the active sheet is hidden the next one
        shown to its right is active, else the one to its left."""
        v = default_value(value)
        if isinstance(v, str) or not isinstance(v, (bool, int, float)):
            raise VBARuntimeError(1004, "Unable to set the Visible property of the Worksheet class")
        state = (
            ("visible" if v else "hidden")
            if isinstance(v, bool)
            else {-1: "visible", 1: "visible", 0: "hidden", 2: "veryHidden"}.get(to_long(v))
        )
        if state is None:
            raise VBARuntimeError(1004, "Unable to set the Visible property of the Worksheet class")
        if state == self.state:
            return
        book = self.book
        book.guard_structure("show or hide a sheet")
        shown = [s for s in book.sheets if s.state == "visible" and s is not self]
        if state != "visible" and not shown:
            raise VBARuntimeError(1004, "Unable to set the Visible property of the Worksheet class")
        self.state = state
        if state != "visible":
            if book.active is self:
                i = book.sheets.index(self)
                after = [s for s in book.sheets[i + 1 :] if s.state == "visible"]
                book.active = after[0] if after else [s for s in book.sheets[:i] if s.state == "visible"][-1]
                book.selected = [book.active]
            elif self in book.selected:
                book.selected.remove(self)

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

    def m_paste(self, destination=None, link=None):
        """Worksheet.Paste: what was copied, into the destination or else the sheet's selection, as Range.Copy pastes
        it; what was cut moves there."""
        if link is not None and bool(default_value(link)):
            raise VBANotImplemented("Worksheet.Paste Link:=True")
        app = self.book.app
        clip = app.clipboard
        if clip is None:
            raise VBARuntimeError(1004, "Paste method of Worksheet class failed")
        if destination is None:
            if self is not self.book.active:
                raise VBARuntimeError(1004, "Paste method of Worksheet class failed")
            destination = Range(self, *self.sel)
        elif not isinstance(destination, Range):
            raise VBARuntimeError(424, "Object required")
        if clip.cut:
            Range(clip.sheet, clip.rows[0], clip.c1, clip.rows[-1], clip.c2)._move(
                destination.sheet, destination.r1, destination.c1
            )
            app.clipboard = None
            if destination.sheet is self.book.active:
                h, w = clip.rows[-1] - clip.rows[0], clip.c2 - clip.c1
                destination.sheet.sel = (destination.r1, destination.c1, destination.r1 + h, destination.c1 + w)
                destination.sheet.active_cell = (destination.r1, destination.c1)
            return True
        destination._paste(clip, "all", True)
        return True

    def get_usedrange(self):
        b = self.used_bounds()
        if b is None:
            return Range(self, 1, 1, 1, 1)
        return Range(self, *b)

    def get_autofiltermode(self):
        return self.autofilter is not None

    def let_autofiltermode(self, value):
        if not bool(default_value(value)) and self.autofilter is not None:
            if self.protects() and "allowfiltering" not in self.protection.allow:
                raise VBANotImplemented("removing the AutoFilter of a protected sheet that does not allow filtering")
            self.remove_autofilter()

    def get_filtermode(self):
        return self.in_filter_mode()

    def m_showalldata(self):
        if not self.in_filter_mode() or self.protects():  # nothing is filtered out, or the sheet is protected
            raise VBARuntimeError(1004, "ShowAllData method of Worksheet class failed")
        af = self.autofilter
        af.criteria = {}
        for r in range(af.r1 + 1, af.r2 + 1):
            self.hidden_rows.discard(r)
        self.changed = True
        self.book.rows_shown_changed(self)

    def remove_autofilter(self):
        af = self.autofilter
        for r in range(af.r1 + 1, af.r2 + 1):
            self.hidden_rows.discard(r)
        self.autofilter = None
        self.changed = True
        self.book.rows_shown_changed(self)

    def m_calculate(self):
        self.book.recalc_all()
        return True

    def get_sort(self):
        if self._sort is None:
            self._sort = SortObject(self)
        return self._sort
