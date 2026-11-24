"""Application.WorksheetFunction: the worksheet functions VBA calls directly, as WorksheetFunction.X and as
Application.X, and Evaluate."""

import datetime as dt
import re

from .. import entry, formula as fx
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
    to_long,
    to_str,
    type_mismatch,
)
from .constants import MAX_COL, MAX_ROW
from .filters import wildcard_regex
from .range import Range

_VALUE = XLError("#VALUE!")

# The functions VBA cannot reach by name, recorded from Excel: error 438. VBA has its own for most (Len, Abs, Year);
# IF, IFS and SWITCH it has none of. Application's own Left, Value, Rows and Columns answer before a function would.
NOT_FROM_VBA = set(
    """ABS CHAR CODE COLUMN COLUMNS CONCATENATE DATE DATEDIF DAY ERROR.TYPE EXACT EXP HOUR IF IFS INDIRECT INT ISBLANK
    ISREF LEFT LEN LOWER MID MINUTE MOD MONTH N NA NOT NOW OFFSET RIGHT ROW ROWS SECOND SIGN SQRT SWITCH T TIME TODAY
    TRUNC TYPE UPPER VALUE YEAR""".split()
)
APPLICATION_OWN = {"LEFT", "VALUE", "ROWS", "COLUMNS"}
# Whose text arguments early-bound VBA hands over its own way: a range or an array there is a type mismatch (error 13)
# early-bound and joined late-bound, and several texts come back as the last of them early-bound (Excel for Mac:
# TextJoin("-", True, "a", "", "b") is "b") and joined late-bound.
_TEXT_ARGUMENTS = {"TEXTJOIN", "CONCAT"}
# Arguments only true or false matter in, (function, position from 0): TRUE is the same there as VBA's -1.
_TRUTHS = {("TEXTJOIN", 1), ("DAYS360", 2)}
_PLAIN_NUMBER = re.compile(r"^\s*[-+]?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?\s*$")


class WorksheetFunctionError(VBARuntimeError):
    """A worksheet function's error value: error 1004 to VBA through WorksheetFunction, the value itself through
    Application, which is what IsError(Application.Match(...)) relies on."""

    def __init__(self, name, err=_VALUE):
        super().__init__(1004, f"Unable to get the {name} property of the WorksheetFunction class")
        self.err = err


# ---------------------------------------------------------------------------
# WorksheetFunction
# ---------------------------------------------------------------------------


class WorksheetFunction(VBAObject):
    vba_type_name = "WorksheetFunction"

    def __init__(self, app):
        self.app = app

    def _fail(self, name, err=_VALUE):
        return WorksheetFunctionError(name, err)

    def _own(self, name):
        key = name.lower()
        return hasattr(self, "m_" + key) or hasattr(self, "get_" + key)

    def vba_member(self, name, args=(), kwargs=None):
        return self.call(name, list(args), kwargs or {}, hand_back=False)

    def call(self, name, args, kwargs, hand_back):
        """A worksheet function VBA names: WorksheetFunction.X, or Application.X with hand_back, where an error comes
        back as its value instead of error 1004. A function written here for VBA answers first; any other Excel lets
        VBA reach is the formula's own, given VBA's values (see _node)."""
        if self._own(name):
            fewest, most, step, _wants = fx._ARGUMENTS.get(name.upper(), (0, 255, 1, ""))
            given = len(args) - next((i for i, a in enumerate(reversed(args)) if a is not None), len(args))
            if not fewest <= given <= most or (given - fewest) % step:
                raise VBARuntimeError(1004, "Invalid number of arguments.")
            try:
                return super().vba_member(name, args, kwargs)
            except WorksheetFunctionError as e:
                if hand_back:
                    return e.err
                raise
        fname = name.upper().replace("_", ".")
        holder = "Application" if hand_back else "WorksheetFunction"
        if fname in NOT_FROM_VBA:
            raise VBARuntimeError(438, "Object doesn't support this property or method")
        if fname not in fx.FUNCTIONS:
            raise VBANotImplemented(f"{holder}.{name} is not modelled")
        if kwargs:
            raise VBANotImplemented(f"{holder}.{name} given named arguments")
        while args and args[-1] is None:  # optional arguments left out at the end
            args.pop()
        fewest, most, step, _wants = fx._ARGUMENTS.get(fname, (None, None, 1, ""))
        if fewest is None:
            raise VBANotImplemented(f"{holder}.{name}: how many arguments Excel takes is not known")
        if not fewest <= len(args) <= most or (len(args) - fewest) % step:
            raise VBARuntimeError(1004, "Invalid number of arguments.")  # late-bound; early-bound VBA would not compile
        early = not hand_back
        if early and fname in _TEXT_ARGUMENTS and len(args) > (3 if fname == "TEXTJOIN" else 1):
            raise VBANotImplemented(
                f"WorksheetFunction.{name} given several texts, which Excel early-bound joins apart"
            )
        nodes = [self._node(fname, i, a, early, holder) for i, a in enumerate(args)]
        takes = fx.TAKES.get(fname, "r")
        one_value = [takes[min(i, len(takes) - 1)] in "vwo" for i in range(len(args))]
        dates = any(one and isinstance(default_value(a), dt.datetime) for one, a in zip(one_value, args, strict=True))
        book = self.app.book
        try:
            v = book._evaluator(book.active, None, True).call(fname, nodes)
        except fx._Err as e:
            v = e.err
        except ZeroDivisionError:
            v = XLError("#DIV/0!")
        if early:
            texts = any(one and n[0] == "const" and _is_text(n[1]) for one, n in zip(one_value, nodes, strict=True))
            if isinstance(v, XLError) and texts:  # error 13 first, from a parameter declared as a number?
                raise VBANotImplemented(f"WorksheetFunction.{name} failing on text, which VBA may refuse as a type")
            if isinstance(v, str) and dates:  # the date written out as text by VBA, for a parameter declared so?
                raise VBANotImplemented(f"WorksheetFunction.{name} given a date where it reads text")
        return self._result(name, v, hand_back)

    def _node(self, fname, i, a, early, holder):
        """An argument as the formula function takes it. A value Excel would read the same however VBA hands it over
        is given; one that depends on it is not answered: early-bound WorksheetFunction.X turns an argument into the
        type its parameter is declared as first (TRUE into -1, a range of several cells into error 13), late-bound
        and Application.X do not (TRUE is 1, a range is gone through value by value), and which a call is cannot be
        told here."""
        takes = fx.TAKES.get(fname, "r")
        one_value = takes[min(i, len(takes) - 1)] in "vwo"
        label = f"{holder}.{fname.title().replace('.', '_')}"
        if a is None:
            return ("missing",)
        if isinstance(a, Range):
            cells = self._area(a)
            if one_value and cells.extent == (1, 1):
                return ("const", cells.cell(0, 0))
            if early and (one_value or fname in _TEXT_ARGUMENTS):
                raise VBANotImplemented(f"{label} given a range of several cells for one value")
            return ("const", cells)
        if isinstance(a, VBAArray):
            if early and (one_value or fname in _TEXT_ARGUMENTS):
                raise VBANotImplemented(f"{label} given an array for one value")
            if fx._ARGUMENTS.get(fname, (0, 0, 1, ""))[3][i : i + 1] == "r":
                raise VBANotImplemented(f"{label} given an array where a formula must have a reference")
            return ("const", self._looked_in(a))
        v = default_value(a)
        if v is EMPTY:
            return ("missing",)  # as an argument left empty in a formula: 0, and counted by COUNTA
        if v is NULL:
            if one_value:
                raise VBANotImplemented(f"{label} given Null")
            return ("missing",)
        if isinstance(v, bool) and (fname, i) in _TRUTHS:
            return ("const", v)
        if early and one_value and (isinstance(v, (bool, XLError)) or isinstance(v, str) and _numberish(v)):
            raise VBANotImplemented(f"{label} given {v!r}, which early-bound VBA may turn into another type first")
        return ("const", _in_formula(v))

    def _result(self, name, v, hand_back):
        """What comes back to VBA: cells a function hands back as a Range, an array as a VBA array from 1 (one row of
        one dimension), an error as error 1004 (its value through Application), a blank as Empty."""
        if isinstance(v, fx.Array):
            if v.rect is not None:
                return Range(*v.rect)
            return vba_array(v)
        if isinstance(v, XLError):
            if hand_back:
                return v
            raise WorksheetFunctionError(name, v)
        return EMPTY if v is None else v

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
        except fx._Err as e:
            raise self._fail(name, e.err) from None

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
            raise self._fail(name, XLError("#N/A"))
        given = list(args)
        while given and given[-1] is None:
            given.pop()
        nodes = [("const", key)] + [("missing",) if a is None else ("const", a) for a in given]
        try:
            v = fn(fx.Evaluator(None, None, None), nodes)
        except fx._Err as e:
            raise self._fail(name, e.err) from None
        if isinstance(v, fx.Array):
            return self._result(name, v, False)
        if isinstance(v, XLError):
            raise self._fail(name, v)
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
        r, c = _in_formula(row_num), _in_formula(column_num)
        v = self._lookup("Index", lambda ev, nodes: fx.f_index(ev, [nodes[1], nodes[0], *nodes[2:]]), r,
                         self._looked_in(array), c)  # fmt: skip
        if not isinstance(array, Range) and 0.0 in (r, c) and not isinstance(v, (VBAArray, Range)):
            return VBAArray.from_list([v], lo=1)  # all of a row or column that is one value: an array still
        return v

    def m_transpose(self, arg):
        """Transpose: rows into columns, of a range's values as Value2 has them (a date a number) or of an array's as
        they are; a row comes back in one dimension, VBA's array of one dimension being a row."""
        if isinstance(arg, Range):
            arg = arg.get_value2()
            if not isinstance(arg, VBAArray):
                arg = VBAArray.from_list([arg], lo=1)
        if isinstance(arg, VBAArray) and arg.bounds is not None and len(arg.bounds) in (1, 2):
            if len(arg.bounds) == 1:
                rows = [arg.items_1d()]
            else:
                (r1, r2), (c1, c2) = arg.bounds
                rows = [[arg.get([r, c]) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
        else:
            raise VBANotImplemented(f"Transpose of a {type(default_value(arg)).__name__}")
        columns = [[_transposed(v) for v in col] for col in zip(*rows, strict=True)]
        return vba_array(fx.Array(columns), empty=None)

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


def _transposed(v):
    """A value as Transpose hands it back: a number a Double (Transpose(Array(1, 2)) holds Doubles)."""
    if is_numeric_value(v) and not isinstance(v, bool):
        return float(v)
    return v


def _is_text(v):
    return isinstance(v, str) and not isinstance(v, XLError)


def _numberish(s):
    """Text Excel reads as a number that VBA's own conversion may read otherwise ("5%", "$5", a date): anything but a
    plain decimal number."""
    if _PLAIN_NUMBER.match(s):
        return False
    try:
        return entry.value_of(s) is not None
    except VBANotImplemented:
        return True


def vba_array(a, empty=EMPTY):
    """An array of a formula's as VBA gets it: from 1, one row in one dimension, more in two; a blank is `empty`
    (None to keep the values as they are)."""
    rows = a.full()

    def value(v):
        return empty if v is None and empty is not None else v

    if len(rows) == 1:
        return VBAArray.from_list([value(v) for v in rows[0]], lo=1)
    out = VBAArray([(1, len(rows)), (1, len(rows[0]))])
    for i, row in enumerate(rows, 1):
        for j, v in enumerate(row, 1):
            out.set([i, j], value(v))
    return out


def evaluate(sheet, text):
    """Evaluate("..."), and [...]: a formula's text worked out against `sheet` as Excel works out one entered with
    Ctrl+Shift+Enter (its ranges whole). A reference, or cells a function hands back, is a Range; an array a VBA
    array; an error its value. Text that is no formula, or longer than 255 characters, is #VALUE!; a formula
    MacroPress cannot read stops the run."""
    s = to_str(default_value(text))
    if len(s) > 255:
        return _VALUE
    s = s.strip()
    if s.startswith("="):
        s = s[1:].strip()
    if not s or _plainly_not_a_formula(s):
        return _VALUE
    book = sheet.book
    tree = fx.parse(fx.normalized(s, book.find_sheet))
    node = tree
    while node[0] == "paren":
        node = node[1]
    ev = book._evaluator(sheet, None, True)
    if node[0] in ("ref", "name", "range"):
        try:
            return Range(*ev.rect_of(node))
        except fx._Err as e:
            return e.err
    v = ev.run(tree)
    if isinstance(v, fx.Array):
        return Range(*v.rect) if v.rect is not None else vba_array(v)
    return EMPTY if v is None else v


def _plainly_not_a_formula(s):
    """Text no formula can be: ending on an operator, a comma or an open bracket, or with brackets or quotes that do
    not close."""
    if s[-1] in "+-*/^&=<>,(:" or s.count('"') % 2:
        return True
    depth = 0
    for ch in re.sub(r'"[^"]*"', "", s):
        depth += {"(": 1, ")": -1}.get(ch, 0)
        if depth < 0:
            return True
    return depth != 0
