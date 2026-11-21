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


class FormulaInfo:
    __slots__ = ("sheet", "r", "c", "text", "dynamic", "ast", "precedents", "error", "volatile")

    def __init__(self, sheet, r, c, text, dynamic):
        self.sheet, self.r, self.c, self.text, self.dynamic = sheet, r, c, text, dynamic
        self.ast, self.precedents, self.error, self.volatile = None, None, None, False


class Book(VBAObject):
    vba_type_name = "Workbook"

    def __init__(self, name="Book1"):
        self.name = name
        self.sheets = []
        self.names = {}  # lower name -> list of (scope sheet name or None, refers_to)
        self.tables = {}  # lower name -> Table
        self.date1904 = False
        self.active = None
        self.selected = []
        self.app = None
        self.now = None  # Now() for NOW()/TODAY() in formulas
        self.formula_index = None  # (sheet, r, c) -> FormulaInfo, built when first needed
        self.dirty = set()  # formulas whose inputs changed since they were last calculated
        self._by_sheet, self._readers, self._always = {}, {}, set()
        self._evaluating = set()
        self._spill_writer = None

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
            info.ast = fx.parse(text)
            info.precedents = self._precedents(sheet, fx.references(info.ast))
            info.volatile = bool(VOLATILE & fx.function_names(info.ast))
        except VBANotImplemented as err:  # not evaluated: it keeps its saved value until a cell it reads changes
            info.error, info.precedents = str(err), self._named_precedents(sheet, text)
        self.formula_index[key] = info
        self._by_sheet.setdefault(sheet, set()).add(key)
        if info.precedents is None or info.volatile:
            self._always.add(key)  # reads cells we cannot name, or recalculates on every change
        for p in info.precedents or ():
            self._readers.setdefault(p[0], {}).setdefault(key, []).append(p[1:])
        if dirty:
            self.dirty.add(key)
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

    def _unregister(self, key):
        info = self.formula_index.pop(key, None)
        if info is None:
            return
        self._by_sheet.get(key[0], set()).discard(key)
        self._always.discard(key)
        for p in info.precedents or ():
            self._readers.get(p[0], {}).pop(key, None)

    def register_formula(self, sheet, r, c, text, dynamic):
        self._formulas()
        info = self._register(sheet, r, c, text, dynamic, dirty=True)
        if info.error and self.app is not None:
            self.app.note("FORMULA-NOT-EVALUATED", f"{sheet.name}!{col_letters(c)}{r}: {info.error}")

    def forget_formulas(self, sheet, r1, c1, r2, c2):
        """Formulas in the rectangle are being overwritten or cleared; a spill goes with its formula."""
        if self.formula_index is None:
            return
        keys = self._by_sheet.get(sheet)
        if not keys:
            return
        if (r1, c1) == (r2, c2):
            found = [(sheet, r1, c1)] if (sheet, r1, c1) in keys else []
        else:
            found = [k for k in keys if r1 <= k[1] <= r2 and c1 <= k[2] <= c2]
        for key in found:
            self._unregister(key)
            self.dirty.discard(key)
            spill = sheet.spills.pop((key[1], key[2]), None)
            if spill:
                self._drop_spill(sheet, (key[1], key[2]))
                self.cells_changed(sheet, *spill)

    def _drop_spill(self, sheet, anchor):
        for cell in [k for k, a in sheet.spill_owner.items() if a == anchor]:
            del sheet.spill_owner[cell]
            row = sheet.rows.get(cell[0])
            if row is not None:
                row.pop(cell[1], None)
                if not row:
                    del sheet.rows[cell[0]]

    def cells_changed(self, sheet, r1, c1, r2, c2):
        """Cells of `sheet` in the rectangle changed: the formulas that read them (and so on) are now stale."""
        if self._spill_writer is None and sheet.spill_owner:
            for cell in [k for k in sheet.spill_owner if r1 <= k[0] <= r2 and c1 <= k[1] <= c2]:
                anchor = sheet.spill_owner.pop(cell)  # written over: the formula must spill again (#SPILL!)
                self._stale((sheet,) + anchor)
        if not self._formulas():
            return
        todo = [(sheet, r1, c1, r2, c2)]
        while todo:
            s, a1, b1, a2, b2 = todo.pop()
            hits = [k for k in self._always if k not in self.dirty]
            for key, rects in self._readers.get(s, {}).items():
                if (
                    key not in self.dirty
                    and key != self._spill_writer
                    and any(not (p[2] < a1 or p[0] > a2 or p[3] < b1 or p[1] > b2) for p in rects)
                ):
                    hits.append(key)
            for key in hits:
                if key not in self.dirty:
                    todo.append(self._stale(key))

    def _stale(self, key):
        self.dirty.add(key)
        s, r, c = key
        return (s,) + (s.spills.get((r, c)) or (r, c, r, c))

    def refresh(self, sheet, spills_only=False):
        """Before cells of `sheet` are read: calculate its stale formulas (what Excel shows by then). With spills_only,
        only its stale dynamic arrays: their spills are what decides which cells the sheet uses."""
        manual = self.app is not None and self.app.settings.get("calculation") == -4135
        for key in [k for k in self.dirty if k[0] is sheet and k not in self._evaluating]:
            if spills_only and not self.formula_index[key].dynamic:
                continue
            if (
                manual
                and sheet.rows.get(key[1], {}).get(key[2]) is not None
                and sheet.rows[key[1]][key[2]].value is not PENDING
            ):
                self.app.note("CALC-MANUAL")  # manual calculation: Excel shows the old value
                continue
            self.evaluate(key)

    def recalc_all(self):
        budget = 20 * len(self.formula_index or ()) + 1000
        while self.dirty:
            budget -= 1
            if budget < 0:
                raise VBANotImplemented("formulas that keep changing each other (a circular reference)")
            self.evaluate(next(iter(self.dirty)))

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
                    result = fx.Evaluator(self, key[0], self._read_for_formula).run(info.ast)
                except VBANotImplemented as err:
                    result = PENDING
                    if self.app is not None:
                        self.app.note("FORMULA-NOT-EVALUATED", f"{key[0].name}!{col_letters(key[2])}{key[1]}: {err}")
        finally:
            self._evaluating.discard(key)
        self.dirty.discard(key)
        self._store(key, info, result)

    def _read_for_formula(self, sheet, r, c):
        if (sheet, r, c) in self._evaluating:
            raise VBANotImplemented(f"a circular reference through {sheet.name}!{col_letters(c)}{r}")
        for key in [k for k in self.dirty if k[0] is sheet and k not in self._evaluating]:
            if (key[1], key[2]) == (r, c) or (self.formula_index[key].dynamic and r >= key[1] and c >= key[2]):
                self.evaluate(key)
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
        if result is PENDING:
            cell.value = PENDING
            for k, a in sheet.spill_owner.items():  # its old spilled values are unknown too
                if a == (r, c):
                    sheet.rows[k[0]][k[1]].value = PENDING
            return
        if isinstance(result, fx.Array):
            if info.dynamic:
                self._spill(key, cell, result)
                return
            result = result.rows[0][0] if result.rows and result.rows[0] else None
        cell.value = _cell_result(result)

    def _spill(self, key, cell, arr):
        """A dynamic-array result fills the cells below and to the right, or shows #SPILL! if one is in use."""
        sheet, r, c = key
        old = sheet.spills.get((r, c))
        self._drop_spill(sheet, (r, c))
        rows, cols = arr.shape
        new = (r, c, r + rows - 1, c + cols - 1)
        if rows == 0 or cols == 0:
            cell.value, new = XLError("#CALC!"), (r, c, r, c)
        elif (
            new[2] > MAX_ROW
            or new[3] > MAX_COL
            or any(
                (rr, cc) != (r, c) and sheet.rows.get(rr, {}).get(cc) is not None
                for rr in range(new[0], new[2] + 1)
                for cc in range(new[1], new[3] + 1)
            )
        ):
            cell.value, new = XLError("#SPILL!"), (r, c, r, c)
        else:
            for i, values in enumerate(arr.rows):
                for j, v in enumerate(values):
                    if v is None and self.app is not None:
                        self.app.note("SPILL-BLANK-AS-ZERO")
                    if (i, j) == (0, 0):
                        cell.value = _cell_result(v)
                    else:
                        sheet.rows.setdefault(r + i, {})[c + j] = Cell(_cell_result(v), None, None)
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
