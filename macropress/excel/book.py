"""Workbook: its sheets and names, and keeping its formulas calculated as cells change."""

import math

from .. import formula as fx
from ..errors import VBANotImplemented, VBARuntimeError
from ..values import (
    VBAArray,
    VBAObject,
    XLError,
    default_value,
    is_numeric_value,
    to_long,
    to_str,
)
from .addresses import col_letters, has_relative_reference
from .constants import MAX_COL, MAX_ROW, PENDING
from .protection import BookProtection, Password, check_password
from .sheet import Sheet
from .storage import Cell

# ---------------------------------------------------------------------------
# Workbook and sheets
# ---------------------------------------------------------------------------

VOLATILE = {"NOW", "TODAY", "RAND", "RANDBETWEEN", "RANDARRAY", "OFFSET", "INDIRECT", "CELL", "INFO"}


def _cell_result(v):
    """What a formula cell holds for a result: a blank shows as 0, numbers are doubles, and a number too big for one
    is #NUM!, as in Excel."""
    if v is None:
        return 0.0
    if isinstance(v, int) and not isinstance(v, bool):
        return float(v)
    if isinstance(v, float) and not math.isfinite(v):
        return XLError("#NUM!")
    return v


def _holds_something(cell):
    return cell is not None and (cell.value is not None or cell.formula is not None)


class FormulaInfo:
    __slots__ = ("sheet", "r", "c", "text", "dynamic", "ast", "precedents", "error", "volatile")

    def __init__(self, sheet, r, c, text, dynamic):
        self.sheet, self.r, self.c, self.text, self.dynamic = sheet, r, c, text, dynamic
        self.ast, self.precedents, self.error, self.volatile = None, None, None, False


class Readers:
    """The formulas that read cells of one sheet, kept so that the ones a change reaches are found without passing
    the rest. A reference of a few rows is filed under each of its rows; a taller one of a few columns under each of
    its columns; one both tall and wide apart. Only formulas whose values are current are kept here: one already
    waiting to be calculated has no further to go."""

    SPAN = 32

    def __init__(self):
        self.by_row, self.by_col, self.wide = {}, {}, set()

    def _places(self, r1, c1, r2, c2):
        if r2 - r1 < self.SPAN:
            return self.by_row, range(r1, r2 + 1)
        if c2 - c1 < self.SPAN:
            return self.by_col, range(c1, c2 + 1)
        return None, ()

    def add(self, key, rect):
        index, places = self._places(*rect)
        if index is None:
            self.wide.add(key)
        for i in places:
            index.setdefault(i, set()).add(key)

    def remove(self, key, rect):
        index, places = self._places(*rect)
        self.wide.discard(key)
        for i in places:
            keys = index.get(i)
            if keys is not None:
                keys.discard(key)
                if not keys:
                    del index[i]

    def near(self, r1, c1, r2, c2):
        """The formulas that may read a cell of the rectangle: every one that does, and some that do not."""
        found = set(self.wide)
        for index, lo, hi in ((self.by_row, r1, r2), (self.by_col, c1, c2)):
            if hi - lo < len(index):
                for i in range(lo, hi + 1):
                    found.update(index.get(i, ()))
            else:
                for i, keys in index.items():
                    if lo <= i <= hi:
                        found.update(keys)
        return found


class Book(VBAObject):
    vba_type_name = "Workbook"

    def __init__(self, name="Book1"):
        self.name = name
        self.sheets = []
        self.names = {}  # lower name -> list of (scope sheet name or None, refers_to)
        self.tables = {}  # lower name -> Table
        self.date1904 = False
        self.manual = False  # the calculation mode the file was saved in
        self.protection = None  # BookProtection while the workbook is protected
        self.protection_read = None  # the protection the file had, kept as it was if a macro gives it back
        self.seed = b""  # what a new password's salt is made from: the file's own bytes, so a run can be repeated
        self.moves = []  # cells cut and pasted: (sheet name, rect, sheet name, rect), for the save to check
        self.calc_on_save = True
        self.active = None
        self.selected = []
        self.app = None
        self.now = None  # Now() for NOW()/TODAY() in formulas
        self.formula_index = None  # (sheet, r, c) -> FormulaInfo, built when first needed
        self.dirty = set()  # formulas to calculate before their values are read: their inputs changed
        self.stale = set()  # formulas whose inputs changed under manual calculation: they show their old values
        self._dirty_on, self._dirty_spills = {}, {}  # sheet -> its dirty formulas; its dirty dynamic arrays
        self._by_sheet, self._readers, self._always = {}, {}, set()
        self._evaluating = set()
        self._spill_writer = None
        self._calculating = False

    @classmethod
    def blank(cls, sheet_names=("Sheet1",)):
        book = cls()
        for i, name in enumerate(sheet_names):
            book.sheets.append(Sheet(book, name, f"Sheet{i + 1}"))
        book.active = book.sheets[0]
        book.selected = [book.sheets[0]]
        return book

    # --- keeping formulas current: Excel's automatic calculation, done when a value is next read ---
    def _formulas(self):
        """Every formula in the book, parsed once, with the cells it reads."""
        if self.formula_index is None:
            self.formula_index, self._by_sheet, self._readers, self._always = {}, {}, {}, set()
            for sheet in self.sheets:
                for r, row in sheet.rows.items():
                    for c, cell in row.items():
                        if cell.formula is not None:  # one saved with no value is calculated when it is read
                            dirty = cell.value is PENDING
                            self._register(sheet, r, c, cell.formula, (r, c) in sheet.spills, dirty=dirty)
        return self.formula_index

    def _register(self, sheet, r, c, text, dynamic, dirty):
        key = (sheet, r, c)
        self._unregister(key)
        info = FormulaInfo(sheet, r, c, text, dynamic)
        try:
            if (r, c) in sheet.array_cells and sheet.array_cells[(r, c)] != (r, c, r, c):
                raise VBANotImplemented("an array formula entered over several cells")
            info.ast = fx.parse(text)
            info.precedents = self._precedents(sheet, fx.references(info.ast))
            info.volatile = bool(VOLATILE & fx.function_names(info.ast))
        except VBANotImplemented as err:  # not evaluated: it keeps its saved value until a cell it reads changes
            info.error, info.precedents = str(err), self._named_precedents(sheet, text)
        self.formula_index[key] = info
        self._by_sheet.setdefault(sheet, set()).add(key)
        if dirty:
            self._mark(key)
        if key not in self.dirty and key not in self.stale:
            self._watch(key)
        return info

    def _precedents(self, sheet, nodes):
        """The rectangles a formula reads, from its ref and name nodes; one that is #REF! or #NAME? reads nothing."""
        ev = fx.Evaluator(self, sheet, self._read_for_formula)
        rects = []
        for node in nodes:
            try:
                rects.append(ev.rect_of(node))
            except fx._Err:
                pass  # a #REF! or #NAME? reference reads no cells
        return rects

    def _named_precedents(self, sheet, text):
        """For a formula the evaluator cannot read: the rectangles its text names, or None when it can read cells it
        does not name (through INDIRECT or OFFSET, from another workbook, or a [table column] of its own row)."""
        nodes = []
        for kind, piece in fx.pieces(text):
            if kind in fx.REFERENCE_KINDS or kind == "name":
                nodes.append(("ref" if kind in fx.REFERENCE_KINDS else "name", piece))
            elif (kind == "func" and fx.bare_name(piece) in VOLATILE) or (kind is None and piece.startswith("[")):
                return None
        try:
            return self._precedents(sheet, nodes)
        except VBANotImplemented:
            return None

    def _watch(self, key):
        """The formula's value is current: from now on a change to a cell it reads makes it stale."""
        info = self.formula_index[key]
        if info.precedents is None or info.volatile:
            self._always.add(key)  # reads cells we cannot name, or recalculates on every change
        for p in info.precedents or ():
            self._readers.setdefault(p[0], Readers()).add(key, p[1:])

    def _unwatch(self, key):
        info = self.formula_index[key]
        self._always.discard(key)
        for p in info.precedents or ():
            if p[0] in self._readers:
                self._readers[p[0]].remove(key, p[1:])

    def _mark(self, key):
        self.dirty.add(key)
        self._dirty_on.setdefault(key[0], set()).add(key)
        if self.formula_index[key].dynamic:
            self._dirty_spills.setdefault(key[0], set()).add(key)

    def _unmark(self, key):
        self.dirty.discard(key)
        self.stale.discard(key)
        for index in (self._dirty_on, self._dirty_spills):
            if key[0] in index:
                index[key[0]].discard(key)

    def _unregister(self, key):
        if key not in self.formula_index:
            return
        self._unwatch(key)
        del self.formula_index[key]
        self._by_sheet.get(key[0], set()).discard(key)

    def _manual(self):
        return self.app is not None and self.app.settings.get("calculation") == -4135

    def is_dynamic(self, sheet, r, c):
        """Whether the cell's formula is a dynamic-array formula (Formula2), whose ranges are taken whole."""
        info = self._formulas().get((sheet, r, c))
        return info is not None and info.dynamic

    def register_formula(self, sheet, r, c, text, dynamic):
        self._formulas()
        info = self._register(sheet, r, c, text, dynamic, dirty=True)
        if info.error and self.app is not None:
            self.app.note("FORMULA-NOT-EVALUATED", f"{sheet.name}!{col_letters(c)}{r}: {info.error}")
        if self._manual():
            self.evaluate((sheet, r, c))  # entered under manual calculation: calculated once, from what cells show

    def forget_formulas(self, sheet, r1, c1, r2, c2):
        """Formulas in the rectangle are being overwritten or cleared; a spill goes with its formula."""
        if self.formula_index is None:
            return
        keys = self._by_sheet.get(sheet)
        if not keys:
            return
        if (r2 - r1 + 1) * (c2 - c1 + 1) <= len(keys):
            found = [(sheet, r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if (sheet, r, c) in keys]
        else:
            found = [k for k in keys if r1 <= k[1] <= r2 and c1 <= k[2] <= c2]
        for key in found:
            self._unmark(key)
            self._unregister(key)
            sheet.array_cells.pop((key[1], key[2]), None)
            spill = sheet.spills.pop((key[1], key[2]), None)
            if spill:
                self._drop_spill(sheet, (key[1], key[2]))
                self.cells_changed(sheet, *spill)

    def _drop_spill(self, sheet, anchor):
        """The cells a dynamic array filled are empty again; each keeps its own format, as it had it."""
        for r, c in [k for k, a in sheet.spill_owner.items() if a == anchor]:
            del sheet.spill_owner[(r, c)]
            cell = sheet.rows.get(r, {}).get(c)
            if cell is not None:
                sheet._store(r, c, None, None, *cell.style)

    def cells_changed(self, sheet, r1, c1, r2, c2):
        """Cells of `sheet` in the rectangle changed: the formulas that read them (and so on) are now stale."""
        if self._spill_writer is None and sheet.spill_owner:
            owned = sheet.spill_owner
            if (r2 - r1 + 1) * (c2 - c1 + 1) <= len(owned):
                cells = [(r, c) for r in range(r1, r2 + 1) for c in range(c1, c2 + 1) if (r, c) in owned]
            else:
                cells = [k for k in owned if r1 <= k[0] <= r2 and c1 <= k[1] <= c2]
            for cell in cells:
                anchor = owned.pop(cell)  # written over: the formula must spill again (#SPILL!)
                if (sheet,) + anchor in self.formula_index:
                    self._stale((sheet,) + anchor)
        if not self._formulas():
            return
        todo = [(sheet, r1, c1, r2, c2)]
        while todo:
            s, a1, b1, a2, b2 = todo.pop()
            hits = list(self._always)
            if s in self._readers:
                for key in self._readers[s].near(a1, b1, a2, b2):
                    if any(
                        p[0] is s and not (p[3] < a1 or p[1] > a2 or p[4] < b1 or p[2] > b2)
                        for p in self.formula_index[key].precedents
                    ):
                        hits.append(key)
            for key in hits:
                if key != self._spill_writer and key not in self.dirty and key not in self.stale:
                    todo.append(self._stale(key))

    def _stale(self, key):
        """A formula's inputs changed. It is calculated when next read; under manual calculation it goes on showing
        the value it has until the next calculation. Returns the cells its own value fills."""
        s, r, c = key
        self._unwatch(key)
        cell = s.rows.get(r, {}).get(c)
        if self._manual() and not self._calculating and cell is not None and cell.value is not PENDING:
            self.stale.add(key)
            self.app.note("CALC-MANUAL")
        else:
            self._mark(key)
        return (s,) + (s.spills.get((r, c)) or (r, c, r, c))

    def refresh(self, sheet, spills_only=False):
        """Before cells of `sheet` are read: calculate its stale formulas (what Excel shows by then). With spills_only,
        only its stale dynamic arrays: their spills are what decides which cells the sheet uses."""
        for key in list((self._dirty_spills if spills_only else self._dirty_on).get(sheet, ())):
            if key not in self._evaluating:
                self.evaluate(key)

    def refresh_cell(self, sheet, r, c):
        """Before one cell is read: calculate its formula if stale, or the stale array formula whose value it shows,
        and any stale dynamic array that may fill it."""
        key = (sheet, *sheet.spill_owner.get((r, c), (r, c)))
        if key in self.dirty and key not in self._evaluating:
            self.evaluate(key)
        for key in list(self._dirty_spills.get(sheet, ())):
            if key[1] <= r and key[2] <= c and key not in self._evaluating:
                self.evaluate(key)

    def recalc_all(self):
        """Calculate every formula whose inputs changed: Calculate, a save, or leaving manual calculation."""
        for key in list(self.stale):
            self._mark(key)
        self.stale.clear()
        budget = 20 * len(self.formula_index or ()) + 1000
        self._calculating = True
        try:
            while self.dirty:
                budget -= 1
                if budget < 0:
                    raise VBANotImplemented("formulas that keep changing each other (a circular reference)")
                self.evaluate(next(iter(self.dirty)))
        finally:
            self._calculating = False

    def recalc_range(self, sheet, r1, c1, r2, c2):
        """Range.Calculate: the rectangle's formulas are calculated from what the cells they read show now."""
        self._formulas()
        for key in [k for k in self._by_sheet.get(sheet, ()) if r1 <= k[1] <= r2 and c1 <= k[2] <= c2]:
            self._unwatch(key)
            self.stale.discard(key)
            self._mark(key)
            self.evaluate(key)

    def evaluate(self, key):
        if key not in self.dirty:
            return
        if key in self._evaluating:
            raise VBANotImplemented(f"a circular reference through {key[0].name}!{col_letters(key[2])}{key[1]}")
        info = self.formula_index[key]
        self._evaluating.add(key)
        try:
            if info.error is not None:
                result = PENDING
            else:
                try:
                    sheet, r, c = key
                    arrays = info.dynamic or (r, c) in sheet.array_cells  # Formula2, or Ctrl+Shift+Enter
                    result = fx.Evaluator(self, sheet, self._read_for_formula, (r, c), arrays).run(info.ast)
                except VBANotImplemented as err:
                    result = PENDING
                    if self.app is not None:
                        self.app.note("FORMULA-NOT-EVALUATED", f"{key[0].name}!{col_letters(key[2])}{key[1]}: {err}")
        finally:
            self._evaluating.discard(key)
        self._unmark(key)
        self._store(key, info, result)
        if key in self.formula_index and key not in self.dirty:
            self._watch(key)

    def _read_for_formula(self, sheet, r, c):
        if (sheet, r, c) in self._evaluating:
            raise VBANotImplemented(f"a circular reference through {sheet.name}!{col_letters(c)}{r}")
        if self.dirty:
            self.refresh_cell(sheet, r, c)
        cell = sheet.rows.get(r, {}).get(c)
        if cell is None:
            return None
        if cell.value is PENDING:
            raise VBANotImplemented(f"{sheet.name}!{col_letters(c)}{r} holds a formula that could not be calculated")
        return cell.value

    def _store(self, key, info, result):
        sheet, r, c = key
        cell = sheet.rows.get(r, {}).get(c)
        if cell is None:
            return
        sheet.changed = True
        if self.app is not None:
            self.app.note("FORMULAS-RECALCULATED")
        old = cell.value
        if result is PENDING:
            cell.value = PENDING
            for k, a in sheet.spill_owner.items():  # its old spilled values are unknown too
                if a == (r, c):
                    sheet.rows[k[0]][k[1]].value = PENDING
        elif isinstance(result, fx.Array) and info.dynamic:
            self._spill(key, cell, result)
            return
        else:
            cell.value = _cell_result(self._one_value(sheet, r, c, result))
        if cell.value is not old and not (cell.value == old and type(cell.value) is type(old)):
            self.cells_changed(sheet, r, c, r, c)  # a formula that read the old value, entered while it was stale

    def _one_value(self, sheet, r, c, result):
        """What a formula that is not a dynamic array shows of a result of several cells: of a range, the cell in the
        formula's own row or column (#VALUE! when the range passes it by); of an array, the first value. An array
        formula entered the old way, with Ctrl+Shift+Enter, shows the first value of either."""
        if not isinstance(result, fx.Array):
            return result
        if result.rect is not None and result.extent != (1, 1) and (r, c) not in sheet.array_cells:
            _, r1, c1, r2, c2 = result.rect
            if c1 == c2 and r1 <= r <= r2:
                return result.cell(r - r1, 0)
            if r1 == r2 and c1 <= c <= c2:
                return result.cell(0, c - c1)
            return XLError("#VALUE!")
        return result.cell(0, 0)

    def _spill(self, key, cell, arr):
        """A dynamic-array result fills the cells below and to the right, or shows #SPILL! if one is in use."""
        sheet, r, c = key
        old = sheet.spills.get((r, c))
        self._drop_spill(sheet, (r, c))
        rows, cols = arr.extent  # the whole of it: a range's blank cells spill too, as 0
        new = (r, c, r + rows - 1, c + cols - 1)
        if rows == 0 or cols == 0:
            cell.value, new = XLError("#CALC!"), (r, c, r, c)
        elif (
            new[2] > MAX_ROW
            or new[3] > MAX_COL
            or any(  # a cell that holds something is in the way; one that only has a format is not
                (rr, cc) != (r, c) and _holds_something(sheet.rows.get(rr, {}).get(cc))
                for rr in range(new[0], new[2] + 1)
                for cc in range(new[1], new[3] + 1)
            )
        ):
            cell.value, new = XLError("#SPILL!"), (r, c, r, c)
        else:
            for i, values in enumerate(arr.full()):
                for j, v in enumerate(values):
                    if v is None and self.app is not None:
                        self.app.note("SPILL-BLANK-AS-ZERO")
                    if (i, j) == (0, 0):
                        cell.value = _cell_result(v)
                    else:  # the cell keeps its format and lock
                        style = sheet.style_at(r + i, c + j)
                        sheet.rows.setdefault(r + i, {})[c + j] = Cell(_cell_result(v), None, *style)
                        sheet.spill_owner[(r + i, c + j)] = (r, c)
        sheet.spills[(r, c)] = new
        self._spill_writer = key
        try:  # formulas reading cells the spill now fills, or no longer fills, are stale
            for rect in {old, new} - {None}:
                self.cells_changed(sheet, *rect)
        finally:
            self._spill_writer = None

    def find_sheet(self, name):
        key = name.lower()
        for s in self.sheets:
            if s.name.lower() == key:
                return s
        return None

    def sheet_by_codename(self, code_name):
        for s in self.sheets:
            if s.code_name and s.code_name.lower() == code_name.lower():
                return s
        return None

    def sheet_item(self, index):
        v = default_value(index)
        if isinstance(v, VBAArray):
            return SheetGroup(self, [self.sheet_item(x) for x in v.items_1d()])
        if isinstance(v, str):
            s = self.find_sheet(v)
            if s is None:
                raise VBARuntimeError(9, f"Subscript out of range (no sheet named {v!r})")
            return s
        if is_numeric_value(v):
            i = to_long(v)
            if not 1 <= i <= len(self.sheets):
                raise VBARuntimeError(9, "Subscript out of range")
            return self.sheets[i - 1]
        raise VBARuntimeError(13, "Type mismatch")

    def get_sheets(self, index=None):
        coll = SheetsCollection(self)
        return coll if index is None else self.sheet_item(index)

    get_worksheets = get_sheets

    def get_name(self):
        return self.name

    def get_activesheet(self):
        return self.active

    def get_names(self, index=None):
        coll = NamesCollection(self)
        return coll if index is None else coll.item(index)

    def m_save(self):
        return True

    def m_protect(self, password=None, structure=None, windows=None):
        """Protect the workbook's structure: its sheets cannot be added, deleted, renamed, hidden or shown."""
        if windows is not None and bool(default_value(windows)):
            raise VBANotImplemented("Workbook.Protect Windows:=True")
        if self.protection is not None:
            raise VBANotImplemented("protecting a protected workbook again (unverified)")
        text = to_str(password) if password is not None else ""
        structure = structure is None or bool(default_value(structure))
        read = self.protection_read
        if read is not None and (read.structure, read.windows) == (structure, False):
            if read.password.matches(text, "workbook") if read.password is not None else not text:
                self.protection = read  # protected again just as the file had it: the file keeps its own
                return
        self.protection = BookProtection(Password(text=text) if text else None, structure) if structure else None

    def m_unprotect(self, password=None):
        p = self.protection
        if p is None:
            return
        if password is None and p.password is not None:  # Excel asks for it
            password = self.app.ask("the password to unprotect the workbook")
        if p.password is not None:
            check_password(p.password, to_str(password), "workbook")
        self.protection = None

    def get_protectstructure(self):
        return self.protection is not None and self.protection.structure

    def get_protectwindows(self):
        return self.protection is not None and self.protection.windows

    def guard_structure(self, what):
        """Error 1004 when the workbook's structure is protected: its sheets cannot be renamed, hidden or shown."""
        if self.protection is not None and self.protection.structure:
            raise VBARuntimeError(1004, f"Unable to {what}: the workbook is protected")

    def m_activate(self):
        return True

    # --- selection state ---
    def select_sheets(self, sheets, replace=True):
        for s in sheets:
            if s.state != "visible":
                raise VBARuntimeError(1004, f"Select method of Worksheet class failed (sheet {s.name!r} is hidden)")
        if replace:
            self.selected = list(dict.fromkeys(sheets))
            if len(sheets) > 1:
                self.app.note("GROUP-SELECT-ACTIVE")
            self.active = sheets[0]
        else:
            for s in sheets:
                if s not in self.selected:
                    self.selected.append(s)
            self.app.note("GROUP-EXTEND-KEEPS-ACTIVE")

    def rename_sheet(self, sheet, new):
        """The sheet takes a new name, and the formulas and defined names that refer to it take it too, as in Excel."""
        renamed = {sheet.name.lower(): new}
        old = sheet.name
        sheet.name = new
        for s in self.sheets:
            for r, row in s.rows.items():
                for c, cell in row.items():
                    if cell.formula is not None:
                        text = fx.rename_sheets(cell.formula, renamed)
                        if text != cell.formula:
                            cell.formula = text
                            s.changed = True
                            if self.formula_index is not None:  # parsed again: its references name the new sheet
                                self._register(s, r, c, text, (r, c) in s.spills, dirty=False)
        self.names = {
            key: [
                (new if scope is not None and scope.lower() == old.lower() else scope, fx.rename_sheets(ref, renamed))
                for scope, ref in refs
            ]  # fmt: skip
            for key, refs in self.names.items()
        }

    def resolve_name(self, name, sheet):
        refs = self.names.get(name.lower())
        if not refs:
            return None
        local = [
            r for scope, r in refs if scope is not None and sheet is not None and scope.lower() == sheet.name.lower()
        ]
        glob = [r for scope, r in refs if scope is None]
        ref = (local or glob or [None])[0]
        if ref is not None and has_relative_reference(ref.lstrip("=")):
            raise VBANotImplemented(f"the name {name} = {ref}: a relative reference, read from the cell that uses it")
        return ref


class SheetsCollection(VBAObject):
    vba_type_name = "Sheets"

    def __init__(self, book):
        self.book = book

    def vba_index(self, args, kwargs):
        return self.book.sheet_item(args[0])

    def get_item(self, index):
        return self.book.sheet_item(index)

    def get_count(self):
        return len(self.book.sheets)

    def vba_iter(self):
        return list(self.book.sheets)

    def m_select(self, replace=None):
        self.book.select_sheets(self.book.sheets, replace is None or bool(default_value(replace)))

    def m_add(self, *args, **kwargs):
        raise VBANotImplemented("Sheets.Add")


class SheetGroup(VBAObject):
    vba_type_name = "Sheets"

    def __init__(self, book, sheets):
        self.book, self.sheets = book, sheets

    def m_select(self, replace=None):
        self.book.select_sheets(self.sheets, replace is None or bool(default_value(replace)))

    def get_count(self):
        return len(self.sheets)

    def vba_index(self, args, kwargs):
        i = to_long(args[0])
        if not 1 <= i <= len(self.sheets):
            raise VBARuntimeError(9, "Subscript out of range")
        return self.sheets[i - 1]

    def vba_iter(self):
        return list(self.sheets)


class NamesCollection(VBAObject):
    vba_type_name = "Names"

    def __init__(self, book):
        self.book = book

    def item(self, index):
        name = to_str(index)
        if self.book.resolve_name(name, self.book.active) is None:
            raise VBARuntimeError(1004, f"Name {name!r} not found")
        return NameObject(self.book, name)

    def vba_index(self, args, kwargs):
        return self.item(args[0])


class NameObject(VBAObject):
    vba_type_name = "Name"

    def __init__(self, book, name):
        self.book, self.name = book, name

    def get_refersto(self):
        return "=" + self.book.resolve_name(self.name, self.book.active)

    def get_referstorange(self):
        return self.book.active.get_range(self.name)

    def get_name(self):
        return self.name
