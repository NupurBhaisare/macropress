"""Application and the globals VBA code sees (ActiveSheet, Range, Cells, Workbooks, ...)."""

from ..errors import VBANotImplemented, VBARuntimeError
from ..values import (
    NOTHING,
    VBAObject,
    default_value,
    is_numeric_value,
    to_date,
    to_long,
)
from .constants import XL_CONSTANTS
from .functions import WorksheetFunction
from .range import Range

# ---------------------------------------------------------------------------
# Application and globals
# ---------------------------------------------------------------------------


class Application(VBAObject):
    vba_type_name = "Application"

    def __init__(self, book, ledger, clock):
        self.book = book
        book.app = self
        self.ledger = ledger
        self.clock = clock
        self.clipboard = None
        self.constants = XL_CONSTANTS
        self.settings = {
            "screenupdating": True,
            "displayalerts": True,
            "enableevents": True,
            "calculation": -4105,
            "statusbar": False,
            "displaystatusbar": True,
        }
        self.wsf = WorksheetFunction(self)
        self._globals = {}

    def note(self, key, detail=None):
        self.ledger.note(key, detail)

    def globals_for(self, module):
        """The unqualified names a module's code sees: Range, Cells, Rows and Columns are its own sheet's in a sheet's
        module, the active sheet's anywhere else."""
        sheet = None
        if module is not None and module._kind == "document":
            sheet = self.book.sheet_by_codename(module._name)
        if sheet not in self._globals:
            self._globals[sheet] = Globals(self, sheet)
        return self._globals[sheet]

    # members
    def get_application(self):
        return self

    def get_activeworkbook(self):
        return self.book

    def get_thisworkbook(self):
        return self.book

    def get_workbooks(self, index=None):
        coll = WorkbooksCollection(self.book)
        return coll if index is None else coll.item(index)

    def get_activesheet(self):
        return self.book.active

    def get_activecell(self):
        s = self.book.active
        r, c = s.active_cell
        return Range(s, r, c, r, c)

    def get_selection(self):
        s = self.book.active
        r1, c1, r2, c2 = s.sel
        group = list(self.book.selected) if len(self.book.selected) > 1 and s in self.book.selected else None
        return Range(s, r1, c1, r2, c2, group=group)

    def get_sheets(self, index=None):
        return self.book.get_sheets(index)

    def get_worksheets(self, index=None):
        return self.book.get_worksheets(index)

    def get_range(self, cell1, cell2=None):
        return self.book.active.get_range(cell1, cell2)

    def get_cells(self, rowindex=None, columnindex=None):
        return self.book.active.get_cells(rowindex, columnindex)

    def get_rows(self, rowindex=None):
        return self.book.active.get_rows(rowindex)

    def get_columns(self, columnindex=None):
        return self.book.active.get_columns(columnindex)

    def get_worksheetfunction(self):
        return self.wsf

    def get_cutcopymode(self):
        return 1 if self.clipboard is not None else False

    def let_cutcopymode(self, value):
        v = default_value(value)
        if v is False or v == 0:
            self.clipboard = None

    def _setting(name):
        def getter(self):
            return self.settings[name]

        def setter(self, value):
            self.settings[name] = default_value(value)

        return getter, setter

    get_screenupdating, let_screenupdating = _setting("screenupdating")
    get_displayalerts, let_displayalerts = _setting("displayalerts")
    get_enableevents, let_enableevents = _setting("enableevents")
    get_calculation, let_calculation = _setting("calculation")
    get_statusbar, let_statusbar = _setting("statusbar")
    get_displaystatusbar, let_displaystatusbar = _setting("displaystatusbar")

    def m_intersect(self, *ranges):
        rs = [r for r in ranges if r is not None]
        if not rs or any(not isinstance(r, Range) for r in rs):
            raise VBARuntimeError(13, "Type mismatch")
        sheet = rs[0].sheet
        r1, c1, r2, c2 = rs[0].rect
        for r in rs[1:]:
            if r.sheet is not sheet:
                raise VBARuntimeError(1004, "Method 'Intersect' of object '_Application' failed")
            r1, c1 = max(r1, r.r1), max(c1, r.c1)
            r2, c2 = min(r2, r.r2), min(c2, r.c2)
        if r1 > r2 or c1 > c2:
            return NOTHING
        return Range(sheet, r1, c1, r2, c2)

    def m_union(self, *ranges):
        raise VBANotImplemented("Application.Union (multi-area ranges)")

    def m_calculate(self):
        self.book.recalc_all()
        return True

    def m_wait(self, time):
        self.clock.wait_until(to_date(time))
        return True

    def m_run(self, macro, *args):
        raise VBANotImplemented("Application.Run (calling a macro by its name)")


class Globals(VBAObject):
    """Names usable without a qualifier: Range, Cells, ActiveSheet, Selection, ...

    In a sheet's code module, Range/Cells/Rows/Columns refer to that sheet.
    """

    vba_type_name = "Global"

    def __init__(self, app, sheet=None):
        self.app = app
        self.sheet = sheet

    def _target(self):
        return self.sheet if self.sheet is not None else self.app.book.active

    def get_application(self):
        return self.app

    def get_activeworkbook(self):
        return self.app.book

    def get_thisworkbook(self):
        return self.app.book

    def get_workbooks(self, index=None):
        return self.app.get_workbooks(index)

    def get_activesheet(self):
        return self.app.book.active

    def get_activecell(self):
        return self.app.get_activecell()

    def get_selection(self):
        return self.app.get_selection()

    def get_sheets(self, index=None):
        return self.app.book.get_sheets(index)

    def get_worksheets(self, index=None):
        return self.app.book.get_worksheets(index)

    def get_range(self, cell1, cell2=None):
        return self._target().get_range(cell1, cell2)

    def get_cells(self, rowindex=None, columnindex=None):
        return self._target().get_cells(rowindex, columnindex)

    def get_rows(self, rowindex=None):
        return self._target().get_rows(rowindex)

    def get_columns(self, columnindex=None):
        return self._target().get_columns(columnindex)

    def get_worksheetfunction(self):
        return self.app.wsf

    def m_intersect(self, *ranges):
        return self.app.m_intersect(*ranges)

    def m_union(self, *ranges):
        return self.app.m_union(*ranges)

    def m_calculate(self):
        return self.app.m_calculate()


class WorkbooksCollection(VBAObject):
    vba_type_name = "Workbooks"

    def __init__(self, book):
        self.book = book

    def item(self, index):
        v = default_value(index)
        if (is_numeric_value(v) and to_long(v) == 1) or (isinstance(v, str) and v.lower() == self.book.name.lower()):
            return self.book
        raise VBANotImplemented("other open workbooks")

    def vba_index(self, args, kwargs):
        return self.item(args[0])

    def get_count(self):
        return 1

    def m_open(self, *args, **kwargs):
        raise VBANotImplemented("Workbooks.Open (other files)")

    def m_add(self, *args, **kwargs):
        raise VBANotImplemented("Workbooks.Add")
