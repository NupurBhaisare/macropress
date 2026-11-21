"""A model of the Excel objects VBA code uses: Application, Workbook, Worksheet, Range.

Cell values follow Excel's storage: numbers are floats, text is str, booleans are bool,
errors are XLError. Anything not modelled raises VBANotImplemented so a run never
continues on a guess.
"""

from .addresses import DATE_FORMAT_IDS, area_text, col_letters, col_num, is_date_format, parse_area, shift_formula
from .application import Application, Globals, WorkbooksCollection
from .book import VOLATILE, Book, FormulaInfo, NameObject, NamesCollection, SheetGroup, SheetsCollection
from .constants import MAX_COL, MAX_ROW, PENDING, XL_CONSTANTS
from .filters import (
    OPS,
    AutoFilterState,
    BoolCriterion,
    Criterion,
    OpCriterion,
    UnsupportedCriterion,
    ValuesCriterion,
    apply_filter,
    criterion_from_text,
    wildcard_regex,
)
from .functions import WorksheetFunction
from .range import Clipboard, Range
from .sheet import Sheet
from .sorting import SortFields, SortKey, SortObject, sort_range
from .storage import Cell, Table

__all__ = [
    "MAX_ROW",
    "MAX_COL",
    "PENDING",
    "XL_CONSTANTS",
    "col_num",
    "col_letters",
    "parse_area",
    "area_text",
    "DATE_FORMAT_IDS",
    "is_date_format",
    "shift_formula",
    "Cell",
    "Table",
    "wildcard_regex",
    "Criterion",
    "ValuesCriterion",
    "OpCriterion",
    "BoolCriterion",
    "UnsupportedCriterion",
    "OPS",
    "criterion_from_text",
    "AutoFilterState",
    "apply_filter",
    "Clipboard",
    "Range",
    "SortKey",
    "SortFields",
    "SortObject",
    "sort_range",
    "Sheet",
    "VOLATILE",
    "FormulaInfo",
    "Book",
    "SheetsCollection",
    "SheetGroup",
    "NamesCollection",
    "NameObject",
    "WorksheetFunction",
    "Application",
    "Globals",
    "WorkbooksCollection",
]
