"""Application.WorksheetFunction: the worksheet functions VBA calls directly."""

import datetime as dt

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
)
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
        """The numbers Sum, Max and Min read: text, TRUE/FALSE and blanks are skipped, and an error among them is the
        function's error, which VBA sees as run-time error 1004."""
        nums = []
        for a in args:
            for v in self._cells(a):
                if isinstance(v, XLError):
                    raise self._fail(name)
                if is_numeric_value(v):
                    nums.append(float(v))
        return nums

    def m_xlookup(self, lookup_value, lookup_array, return_array, if_not_found=None, match_mode=None, search_mode=None):
        if (match_mode not in (None, 0) and to_long(match_mode) != 0) or (
            search_mode not in (None, 1) and to_long(search_mode) != 1
        ):
            raise VBANotImplemented("XLookup with match_mode or search_mode")
        if isinstance(lookup_value, VBAArray) or (isinstance(lookup_value, Range) and lookup_value.get_count() != 1):
            # Excel looks up each value of an array and returns an array of results
            raise VBANotImplemented("XLookup with an array or a range of several cells as its lookup value")
        keys = self._cells(lookup_array)
        values = self._cells(return_array)
        target = self._cells(lookup_value)[0]
        for k, v in zip(keys, values, strict=False):
            if _xl_equal(k, target):
                if isinstance(v, XLError):
                    raise self._fail("XLookup")
                return v
        if if_not_found is not None:
            return default_value(if_not_found)
        raise self._fail("XLookup")

    def m_sum(self, *args):
        return float(sum(self._numbers("Sum", args)))

    def m_max(self, *args):
        nums = self._numbers("Max", args)
        return max(nums) if nums else 0.0

    def m_min(self, *args):
        nums = self._numbers("Min", args)
        return min(nums) if nums else 0.0

    def m_counta(self, *args):
        return float(sum(1 for a in args for v in self._cells(a) if v is not EMPTY))


def _xl_equal(a, b):
    if isinstance(a, str) and isinstance(b, str):
        return a.casefold() == b.casefold()
    if is_numeric_value(a) and is_numeric_value(b):
        return float(a) == float(b)
    if isinstance(a, bool) and isinstance(b, bool):
        return a == b
    return False
