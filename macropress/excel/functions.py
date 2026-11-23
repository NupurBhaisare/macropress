"""Application.WorksheetFunction: the worksheet functions VBA calls directly."""

import datetime as dt

from .. import entry, formula as fx
from ..errors import VBANotImplemented, VBARuntimeError
from ..values import (
    EMPTY,
    VBAArray,
    VBAObject,
    XLError,
    date_to_serial,
    default_value,
    is_numeric_value,
    to_long,
    to_str,
    type_mismatch,
)
from .constants import MAX_COL, MAX_ROW
from .filters import wildcard_regex
from .range import Range

# ---------------------------------------------------------------------------
# WorksheetFunction
# ---------------------------------------------------------------------------


class WorksheetFunction(VBAObject):
    vba_type_name = "WorksheetFunction"

    def __init__(self, app):
        self.app = app

    def _fail(self, name):
        return VBARuntimeError(1004, f"Unable to get the {name} property of the WorksheetFunction class")

    def m_find(self, find_text, within_text, start_num=None):
        a, b = to_str(find_text), to_str(within_text)
        st = 1 if start_num is None else to_long(start_num)
        if st < 1 or st > len(b) + (1 if a == "" else 0):
            raise self._fail("Find")
        pos = b.find(a, st - 1)
        if pos < 0:
            raise self._fail("Find")
        return float(pos + 1)

    def m_search(self, find_text, within_text, start_num=None):
        a, b = to_str(find_text), to_str(within_text)
        st = 1 if start_num is None else to_long(start_num)
        if st < 1 or st > len(b):
            raise self._fail("Search")
        m = wildcard_regex(a).search(b, st - 1)
        if m is None:
            raise self._fail("Search")
        return float(m.start() + 1)

    def _cells(self, arg):
        """The values a worksheet function reads, as the sheet holds them: a date is its serial number, and so is a
        VBA Date passed in."""
        if isinstance(arg, Range):
            if arg.get_count() > 5_000_000:
                raise VBANotImplemented("worksheet function over a very large range")
            values = [arg._vba_value(r, c, True) for r in range(arg.r1, arg.r2 + 1) for c in range(arg.c1, arg.c2 + 1)]
        elif isinstance(arg, VBAArray):
            values = arg.iter_all()
        else:
            values = [default_value(arg)]
        return [date_to_serial(v) if isinstance(v, dt.datetime) else v for v in values]

    def _numbers(self, name, args):
        """The numbers Sum, Max and Min read. In a range or an array, text, TRUE/FALSE and blanks are passed over. A
        value passed on its own counts as one written in a formula does: text that reads as a number, TRUE as 1.
        An error among them, or text that is no number, is the function's error: run-time error 1004 to VBA."""
        nums = []
        for a in args:
            alone = not isinstance(a, (Range, VBAArray))
            for v in self._cells(a):
                if isinstance(v, XLError):
                    raise self._fail(name)
                if is_numeric_value(v):
                    nums.append(float(v))
                elif alone and isinstance(v, bool):
                    nums.append(float(v))
                elif alone and isinstance(v, str):
                    number = entry.value_of(v)
                    if number is None:
                        raise self._fail(name)
                    nums.append(number)
        return nums

    def m_sum(self, *args):
        try:
            return fx.total(self._numbers("Sum", args))
        except fx._Err:
            raise self._fail("Sum") from None

    def m_max(self, *args):
        nums = self._numbers("Max", args)
        return max(nums) if nums else 0.0

    def m_min(self, *args):
        nums = self._numbers("Min", args)
        return min(nums) if nums else 0.0

    def m_counta(self, *args):
        return float(sum(1 for a in args for v in self._cells(a) if v is not EMPTY))

    # --- CountIf, SumIf and the rest: formula.criterion has the rules ---
    def _area(self, arg, extent=None):
        """A range argument: its cells as a formula reads them. Only a Range will do (error 424, as for an array).
        With extent, SumIf's sum range: its first cell and on, over the criteria range's size."""
        if not isinstance(arg, Range):
            raise VBARuntimeError(424, "Object required")
        r2, c2 = (arg.r2, arg.c2) if extent is None else (arg.r1 + extent[0] - 1, arg.c1 + extent[1] - 1)
        if r2 > MAX_ROW or c2 > MAX_COL:
            raise VBANotImplemented("a sum range that runs past the edge of the sheet")
        return arg.sheet.book.formula_cells(arg.sheet, arg.r1, arg.c1, r2, c2)

    def _criterion(self, value):
        """A criterion from VBA, as a formula is given one: a date is its serial number and Empty a blank (the number
        0); a Range of one cell gives its value, of more is a type mismatch."""
        if isinstance(value, Range) and value.get_count() != 1:
            raise type_mismatch()
        v = default_value(value)
        if v is EMPTY:
            return None
        if isinstance(v, dt.datetime):
            return date_to_serial(v)
        if isinstance(v, (bool, str)):  # an error (CVErr) is an XLError, a str
            return v
        if is_numeric_value(v):
            return float(v)
        raise VBANotImplemented(f"a criterion of type {type(v).__name__}")

    def _pairs(self, args):
        if len(args) % 2:
            raise VBANotImplemented("a criteria range without its criterion")
        return [self._area(a) for a in args[0::2]], [self._criterion(c) for c in args[1::2]]

    def _one_criterion(self, range_, criteria, sum_range):
        area = self._area(range_)
        target = area if sum_range is None else self._area(sum_range, area.extent)
        return fx.where([area], [self._criterion(criteria)], target)[1]

    def _found(self, name, work):
        """What work() gives; an error value it comes to is run-time error 1004."""
        try:
            return work()
        except fx._Err:
            raise self._fail(name) from None

    def m_countif(self, range_, criteria):
        return self._found("CountIf", lambda: float(fx.where([self._area(range_)], [self._criterion(criteria)])[0]))

    def m_countifs(self, *args):
        return self._found("CountIfs", lambda: float(fx.where(*self._pairs(args))[0]))

    def m_sumif(self, range_, criteria, sum_range=None):
        return self._found("SumIf", lambda: fx.sum_found(self._one_criterion(range_, criteria, sum_range)))

    def m_averageif(self, range_, criteria, average_range=None):
        return self._found("AverageIf", lambda: fx.average_found(self._one_criterion(range_, criteria, average_range)))

    def m_sumifs(self, sum_range, *args):
        return self._found("SumIfs", lambda: fx.sum_found(fx.where(*self._pairs(args), self._area(sum_range))[1]))

    def m_averageifs(self, average_range, *args):
        return self._found(
            "AverageIfs", lambda: fx.average_found(fx.where(*self._pairs(args), self._area(average_range))[1])
        )

    def m_maxifs(self, max_range, *args):
        return self._found("MaxIfs", lambda: fx.max_found(fx.where(*self._pairs(args), self._area(max_range))[1]))

    def m_minifs(self, min_range, *args):
        return self._found("MinIfs", lambda: fx.min_found(fx.where(*self._pairs(args), self._area(min_range))[1]))

    # --- VLookup, Match and the rest: the formula functions, given VBA's values ---
    def _looked_in(self, arg):
        """What a lookup looks in: a Range's cells as a formula reads them, still the cells they are, or a VBA array,
        one of a single dimension being a row."""
        if isinstance(arg, Range):
            return self._area(arg)
        if isinstance(arg, VBAArray) and arg.bounds is not None and len(arg.bounds) in (1, 2):
            if len(arg.bounds) == 1:
                ((lo, hi),) = arg.bounds
                return fx.Array([[_in_formula(arg.get([i])) for i in range(lo, hi + 1)]])
            (r1, r2), (c1, c2) = arg.bounds
            return fx.Array([[_in_formula(arg.get([r, c])) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)])
        raise VBANotImplemented(f"a lookup in a {type(default_value(arg)).__name__}")

    def _looked_for(self, value):
        """The value a lookup looks for: a Range of one cell gives the cell's value; a VBA Date stays one (see
        _lookup)."""
        if isinstance(value, Range):
            if value.get_count() != 1:
                raise VBANotImplemented("a Range of several cells looked up from VBA")
            return self._area(value).cell(0, 0)
        v = default_value(value)
        return v if isinstance(v, dt.datetime) else _in_formula(v)

    def _lookup(self, name, fn, key, *args, not_found=None, searched=None):
        """fn, a formula function, given VBA's values (a trailing None is an argument left out, another one left
        empty): its result, cells as a Range; an error it comes to is error 1004. VBA hands a Date over as text
        written by the regional settings, which finds no number: nothing at all where the values searched (the first
        array, unless told) hold no text."""
        if isinstance(key, dt.datetime):
            searched = searched or next(a for a in args if isinstance(a, fx.Array))
            if any(isinstance(v, str) and not isinstance(v, XLError) for v in searched.values()):
                raise VBANotImplemented(f"{name} of a Date, which VBA hands over as text written by regional settings")
            if not_found is not None:
                return not_found
            raise self._fail(name)
        given = list(args)
        while given and given[-1] is None:
            given.pop()
        nodes = [("const", key)] + [("missing",) if a is None else ("const", a) for a in given]
        try:
            v = fn(fx.Evaluator(None, None, None), nodes)
        except fx._Err:
            raise self._fail(name) from None
        if isinstance(v, fx.Array):
            if v.extent != (1, 1):
                raise VBANotImplemented(f"{name} handing several cells or values back to VBA")
            if v.rect is not None:
                sheet, r, c = v.rect[:3]
                return Range(sheet, r, c, r, c)
            v = v.cell(0, 0)
        if isinstance(v, XLError):
            raise self._fail(name)
        if v is None:
            raise VBANotImplemented(f"{name} finding a blank cell, from VBA")
        return v

    def m_vlookup(self, lookup_value, table_array, col_index_num, range_lookup=None):
        table = self._looked_in(table_array)
        return self._lookup("VLookup", fx.f_vlookup, self._looked_for(lookup_value), table, _in_formula(col_index_num),
                            _in_formula(range_lookup), searched=fx.Array([row[:1] for row in table.rows]))  # fmt: skip

    def m_hlookup(self, lookup_value, table_array, row_index_num, range_lookup=None):
        table = self._looked_in(table_array)
        return self._lookup("HLookup", fx.f_hlookup, self._looked_for(lookup_value), table, _in_formula(row_index_num),
                            _in_formula(range_lookup), searched=fx.Array(table.rows[:1]))  # fmt: skip

    def m_match(self, lookup_value, lookup_array, match_type=None):
        return self._lookup("Match", fx.f_match, self._looked_for(lookup_value), self._looked_in(lookup_array),
                            _in_formula(match_type))  # fmt: skip

    def m_xmatch(self, lookup_value, lookup_array, match_mode=None, search_mode=None):
        return self._lookup("XMatch", fx.f_xmatch, self._looked_for(lookup_value), self._looked_in(lookup_array),
                            _in_formula(match_mode), _in_formula(search_mode))  # fmt: skip

    def m_lookup(self, lookup_value, lookup_vector, result_vector=None):
        result = None if result_vector is None else self._looked_in(result_vector)
        return self._lookup(
            "Lookup", fx.f_lookup, self._looked_for(lookup_value), self._looked_in(lookup_vector), result
        )

    def m_xlookup(self, lookup_value, lookup_array, return_array, if_not_found=None, match_mode=None, search_mode=None):
        missing = None if if_not_found is None else default_value(if_not_found)
        return self._lookup("XLookup", fx.f_xlookup, self._looked_for(lookup_value), self._looked_in(lookup_array),
                            self._looked_in(return_array), None if missing is None else _in_formula(missing),
                            _in_formula(match_mode), _in_formula(search_mode), not_found=missing)  # fmt: skip

    def m_index(self, array, row_num, column_num=None):
        return self._lookup("Index", lambda ev, nodes: fx.f_index(ev, [nodes[1], nodes[0], *nodes[2:]]),
                            _in_formula(row_num), self._looked_in(array), _in_formula(column_num))  # fmt: skip

    def m_choose(self, index_num, *values):
        k = _in_formula(index_num)
        try:
            k = fx.num(k)
        except fx._Err:
            raise self._fail("Choose") from None
        if not 1 <= k < len(values) + 1:
            raise self._fail("Choose")
        return values[int(k) - 1]


def _in_formula(v):
    """A VBA value as a formula sees it: Empty is blank, a date its serial number, a Currency a Double."""
    v = default_value(v)
    if v is EMPTY or v is None:
        return None
    if isinstance(v, dt.datetime):
        return date_to_serial(v)
    if isinstance(v, (bool, str)):
        return v
    if is_numeric_value(v):
        return float(v)
    raise VBANotImplemented(f"a {type(v).__name__} given to a worksheet function")
