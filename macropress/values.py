"""VBA values: Variant semantics, conversions, operators and arrays.

Representation
    Empty, Null, Nothing  -> the EMPTY, NULL, NOTHING singletons
    Boolean               -> bool
    Byte/Integer/Long     -> int
    Single/Double         -> float
    Currency/Decimal      -> decimal.Decimal
    Date                  -> datetime.datetime
    String                -> str
    Error (CVErr)         -> XLError
    arrays                -> VBAArray
    objects               -> VBAObject subclasses (Excel model)
"""

import datetime as dt
import inspect
import math
import re
from decimal import ROUND_HALF_EVEN, Decimal, InvalidOperation

from .errors import VBANotImplemented, VBAObjectError, VBARuntimeError


class _Special:
    __slots__ = ("name",)

    def __init__(self, name):
        object.__setattr__(self, "name", name)

    def __repr__(self):
        return self.name

    def _no_object(self):
        if self.name == "Nothing":
            return VBAObjectError(91, "Object variable or With block variable not set")
        return VBAObjectError(424, "Object required")

    def __getattr__(self, member):  # x.Value where x holds no object: VBA's error, which On Error can handle
        if member.startswith("__"):
            raise AttributeError(member)
        raise self._no_object()

    def __setattr__(self, member, value):
        raise self._no_object()


EMPTY = _Special("Empty")
NULL = _Special("Null")
NOTHING = _Special("Nothing")

EPOCH = dt.datetime(1899, 12, 30)

_note = None  # set by the running session: _note(key, detail) records a behaviour the model relied on


def set_note(fn):
    global _note
    _note = fn


def note(key, detail=None):
    if _note is not None:
        _note(key, detail)


class XLError(str):
    """An Excel error value such as #N/A (VBA: a Variant of subtype Error)."""


def type_mismatch():
    return VBARuntimeError(13, "Type mismatch")


# ---------------------------------------------------------------------------
# Objects
# ---------------------------------------------------------------------------

_SIG_CACHE = {}


def _accepts_args(fn):
    key = getattr(fn, "__func__", fn)
    if key not in _SIG_CACHE:
        params = [p for p in inspect.signature(fn).parameters.values()]
        _SIG_CACHE[key] = bool(params)
    return _SIG_CACHE[key]


class VBAObject:
    """Base for objects VBA code can use: members are looked up case-insensitively.

    A subclass exposes
        m_<name>(...)        a method
        get_<name>(...)      a property (parameterised if it takes arguments)
        let_<name>(v, ...)   assignment to a property
    VBA named arguments arrive as lower-case keyword arguments. A missing optional
    argument arrives as None.
    """

    vba_type_name = "Object"

    def vba_member(self, name, args=(), kwargs=None):
        kwargs = kwargs or {}
        key = name.lower()
        fn = getattr(self, "m_" + key, None) or getattr(self, "get_" + key, None)
        if fn is None:
            raise VBANotImplemented(f"{self.vba_type_name}.{name} is not modelled")
        if (args or kwargs) and not _accepts_args(fn):
            return index_value(fn(), args, kwargs)
        return call_bound(fn, f"{self.vba_type_name}.{name}", args, kwargs)

    def vba_has_member(self, name):
        key = name.lower()
        return any(hasattr(self, p + key) for p in ("m_", "get_", "let_"))

    def vba_let_member(self, name, args, kwargs, value):
        key = name.lower()
        fn = getattr(self, "let_" + key, None)
        if fn is not None:
            return call_bound(fn, f"{self.vba_type_name}.{name}", (value,) + tuple(args), kwargs or {})
        if hasattr(self, "get_" + key) or hasattr(self, "m_" + key):
            # e.g. ws.Cells(1, 2) = 5: assign to the default member of what the member returns
            target = self.vba_member(name, args, kwargs)
            if isinstance(target, VBAObject):
                return target.vba_default_let(value)
            # a property the model reads but cannot set (Range.NumberFormat): an error On Error could skip would
            # leave the property as it was
            raise VBANotImplemented(f"setting {self.vba_type_name}.{name} is not modelled")
        raise VBANotImplemented(f"assignment to {self.vba_type_name}.{name} is not modelled")

    def vba_default_get(self):
        raise VBARuntimeError(438, f"{self.vba_type_name} has no default property")

    def vba_default_let(self, value):
        raise VBARuntimeError(438, f"{self.vba_type_name} has no default property")

    def vba_index(self, args, kwargs):
        raise VBARuntimeError(438, f"{self.vba_type_name} cannot be indexed")

    def vba_iter(self):
        raise VBARuntimeError(451, f"{self.vba_type_name} is not a collection")


def call_bound(fn, label, args, kwargs):
    try:
        inspect.signature(fn).bind(*args, **kwargs)
    except TypeError as err:
        unknown = [k for k in kwargs if k not in inspect.signature(fn).parameters]
        if unknown:
            raise VBANotImplemented(f"named argument {unknown[0]!r} of {label} is not modelled") from err
        raise VBARuntimeError(450, f"Wrong number of arguments for {label}") from err
    return fn(*args, **kwargs)


def index_value(value, args, kwargs=None):
    """value(args): array element or an object's default member with arguments."""
    if isinstance(value, VBAArray):
        if kwargs:
            raise VBARuntimeError(13, "Type mismatch")
        return value.get([to_long(default_value(a)) for a in args])
    if isinstance(value, VBAObject):
        return value.vba_index(args, kwargs or {})
    raise type_mismatch()


def default_value(v):
    """Let-coercion: an object used as a value gives its default property (Range -> Value)."""
    while isinstance(v, VBAObject):
        v = v.vba_default_get()
    return v


def is_object(v):
    return isinstance(v, VBAObject) or v is NOTHING


# ---------------------------------------------------------------------------
# Arrays
# ---------------------------------------------------------------------------


class VBAArray:
    """A VBA array with explicit bounds per dimension."""

    def __init__(self, bounds=None, elem_type="variant", values=None):
        self.elem_type = elem_type
        self.bounds = bounds  # list of (lo, hi) or None for a dynamic array not yet sized
        self.data = {}
        if values is not None:  # 1-D from a Python list
            lo = bounds[0][0] if bounds else 0
            self.bounds = [(lo, lo + len(values) - 1)]
            for i, v in enumerate(values):
                self.data[(lo + i,)] = v

    @classmethod
    def from_list(cls, values, lo=0):
        return cls([(lo, lo + len(values) - 1)], "variant", values)

    def _key(self, idx):
        if self.bounds is None:
            raise VBARuntimeError(9, "Subscript out of range")
        if len(idx) != len(self.bounds):
            raise VBARuntimeError(9, "Subscript out of range")
        for i, (lo, hi) in zip(idx, self.bounds, strict=False):
            if i < lo or i > hi:
                raise VBARuntimeError(9, "Subscript out of range")
        return tuple(idx)

    def get(self, idx):
        key = self._key(idx)
        if key in self.data:
            return self.data[key]
        return default_for_type(self.elem_type)

    def set(self, idx, value):
        self.data[self._key(idx)] = value

    def lbound(self, dim=1):
        if self.bounds is None or not 1 <= dim <= len(self.bounds):
            raise VBARuntimeError(9, "Subscript out of range")
        return self.bounds[dim - 1][0]

    def ubound(self, dim=1):
        if self.bounds is None or not 1 <= dim <= len(self.bounds):
            raise VBARuntimeError(9, "Subscript out of range")
        return self.bounds[dim - 1][1]

    def items_1d(self):
        if self.bounds is None:
            return []
        if len(self.bounds) != 1:
            raise VBANotImplemented("this operation on a multi-dimensional array")
        lo, hi = self.bounds[0]
        return [self.get([i]) for i in range(lo, hi + 1)]

    def iter_all(self):
        """Elements in VBA's For Each order (first index varies fastest)."""
        if self.bounds is None:
            return []
        ranges = [range(lo, hi + 1) for lo, hi in self.bounds]
        out = []

        def walk(prefix, dims):
            if not dims:
                out.append(self.get(list(reversed(prefix))))
                return
            for i in dims[0]:
                walk(prefix + [i], dims[1:])

        walk([], list(reversed(ranges)))
        return out

    def redim(self, bounds, preserve):
        if preserve and self.bounds is not None:
            if len(bounds) != len(self.bounds) or bounds[:-1] != self.bounds[:-1]:
                raise VBARuntimeError(9, "Subscript out of range")
            old = self.data
            self.bounds = bounds
            self.data = {k: v for k, v in old.items() if self._fits(k)}
        else:
            self.bounds = bounds
            self.data = {}

    def _fits(self, key):
        return all(lo <= i <= hi for i, (lo, hi) in zip(key, self.bounds, strict=False))

    def copy(self):
        new = VBAArray(list(self.bounds) if self.bounds else None, self.elem_type)
        new.data = dict(self.data)
        return new


# ---------------------------------------------------------------------------
# Types
# ---------------------------------------------------------------------------

NUMERIC_TYPES = {"byte", "integer", "long", "longlong", "longptr", "single", "double", "currency", "decimal"}
OBJECT_TYPES = {"object", "range", "worksheet", "workbook", "sheets", "collection", "dictionary"}


def normalise_type(type_name):
    if not type_name:
        return "variant"
    t = type_name.lower()
    if "." in t:  # Excel.Range, Scripting.Dictionary
        t = t.rsplit(".", 1)[1]
    return t


def default_for_type(t):
    t = normalise_type(t)
    if t in ("byte", "integer", "long", "longlong", "longptr"):
        return 0
    if t in ("single", "double"):
        return 0.0
    if t in ("currency", "decimal"):
        return Decimal(0)
    if t == "string":
        return ""
    if t == "boolean":
        return False
    if t == "date":
        return EPOCH
    if t == "variant":
        return EMPTY
    return NOTHING  # any object type


def is_object_type(t):
    t = normalise_type(t)
    return t not in NUMERIC_TYPES and t not in ("string", "boolean", "date", "variant")


def coerce(value, type_name):
    """Assign `value` to a variable declared As `type_name` (Let semantics)."""
    t = normalise_type(type_name)
    if t == "variant":
        return value
    if is_object_type(t):
        raise VBARuntimeError(91, "Object variable or With block variable not set")
    if value is NULL:
        raise VBARuntimeError(94, "Invalid use of Null")
    if t == "string":
        return to_str(value)
    if t == "boolean":
        return to_bool(value)
    if t == "date":
        return to_date(value)
    if t == "byte":
        return _int_range(to_long(value), 0, 255)
    if t == "integer":
        return _int_range(to_long(value), -32768, 32767)
    if t in ("long", "longptr"):
        return to_long(value)
    if t == "longlong":
        return bankers_int(to_number(value))
    if t in ("single", "double"):
        return float(to_number(value))
    if t == "currency":
        return to_currency(value)
    if t == "decimal":
        return to_decimal(value)
    raise VBANotImplemented(f"variables of type {type_name}")


def _int_range(n, lo, hi):
    if n < lo or n > hi:
        raise VBARuntimeError(6, "Overflow")
    return n


# ---------------------------------------------------------------------------
# Conversions
# ---------------------------------------------------------------------------

_NUMBER = r"(\d[\d,]*\.?\d*|\.\d+)([eEdD][-+]?\d+)?"
# A number, with its sign in front ("- 5"), behind ("5-") or as parentheses ("(5)").
_NUM_TEXT = re.compile(rf"^\s*(?:([-+]?)\s*{_NUMBER}|{_NUMBER}\s*(-)|(\()\s*{_NUMBER}\s*\))\s*$")
_HEX_TEXT = re.compile(r"^\s*&([hH][0-9a-fA-F]+|[oO][0-7]+)&?\s*$")


def parse_number_text(s):
    """VBA's string-to-number conversion (assumes '.' decimal and ',' thousands separators): the number, a Double, or
    None. A currency sign is not modelled: VBA reads the one of the machine's regional settings, and no other."""
    m = _NUM_TEXT.match(s)
    if m:
        sign, body, exp, body2, exp2, minus, paren, body3, exp3 = m.groups()
        body, exp = body or body2 or body3, exp or exp2 or exp3
        negative = sign == "-" or bool(minus) or bool(paren)
        text = body.replace(",", "") + (exp.replace("d", "e").replace("D", "e") if exp else "")
        try:
            value = float(text)
        except ValueError:
            return None
        if not math.isfinite(value):
            return None
        return -value if negative else value
    m = _HEX_TEXT.match(s)
    if m:
        digits = m.group(1)
        return float(int(digits[1:], 16 if digits[0] in "hH" else 8))
    if s.count("$") == 1 and _NUM_TEXT.match(s.replace("$", " ")):
        raise VBANotImplemented(
            f"reading {s!r} as a number: whether VBA takes a currency sign depends on the machine's regional settings"
        )
    return None


def typed_number(text):
    """A number as Excel reads it from digits typed in a cell or written in a formula: its first 15 significant
    digits, the rest dropped (1234567890123456789 is 1234567890123450000). None when it is too large for a number:
    the largest that can be typed is 9.99999999999999E+307."""
    try:
        sign, digits, exponent = Decimal(text).as_tuple()
    except InvalidOperation:
        return None
    digits = list(digits)
    while len(digits) > 1 and digits[0] == 0:
        digits.pop(0)
    if len(digits) > 15:
        exponent += len(digits) - 15
        digits = digits[:15]
    try:
        value = float(Decimal((sign, tuple(digits), exponent)))
    except (OverflowError, InvalidOperation):
        return None
    return value if abs(value) < 1e308 else None


def to_number(v):
    """Numeric value for arithmetic: int, float or Decimal."""
    v = default_value(v)
    if isinstance(v, bool):
        return -1 if v else 0
    if isinstance(v, (int, float, Decimal)):
        return v
    if v is EMPTY:
        return 0
    if isinstance(v, dt.datetime):
        return date_to_serial(v)
    if isinstance(v, XLError):
        raise type_mismatch()
    if isinstance(v, str):
        n = parse_number_text(v)
        if n is None:  # "True" is no number either: CDbl("True") is a type mismatch, though CBool("True") is True
            raise type_mismatch()
        return n
    if v is NULL:
        raise VBARuntimeError(94, "Invalid use of Null")
    raise type_mismatch()


def bankers_int(x):
    """Round to the nearest integer, halves to even (CInt, CLng, implicit conversion)."""
    if isinstance(x, int):
        return x
    if isinstance(x, Decimal):
        return int(x.quantize(Decimal(1), rounding=ROUND_HALF_EVEN))
    if math.isnan(x) or math.isinf(x):
        raise VBARuntimeError(6, "Overflow")
    return int(round(x))


def to_long(v):
    n = bankers_int(to_number(v))
    return _int_range(n, -2147483648, 2147483647)


def hex_value(n, long=False):
    """An &H or &O number as VBA types it: one that fits 16 bits is an Integer (&HFFFF is -1) unless it ends in &, a
    bigger one a Long (&HFFFFFFFF is -1)."""
    if not long and n <= 0xFFFF:
        return n - 0x10000 if n > 0x7FFF else n
    if n <= 0xFFFFFFFF:
        return n - 0x100000000 if n > 0x7FFFFFFF else n
    raise VBARuntimeError(6, "Overflow")


def to_double(v):
    return float(to_number(v))


CURRENCY_MAX = Decimal("922337203685477.5807")


def _currency(exact):
    """An exact value as a Currency: four decimals, a half going to the even digit; error 6 past its range."""
    try:
        value = exact.quantize(Decimal("0.0001"), rounding=ROUND_HALF_EVEN)
    except InvalidOperation:  # far too many digits to hold
        raise VBARuntimeError(6, "Overflow") from None
    if abs(value) > CURRENCY_MAX:
        raise VBARuntimeError(6, "Overflow")
    return value


def to_currency(v):
    """A Currency: the number, as the Double holds it, rounded to four decimals."""
    n = to_number(v)
    return _currency(n if isinstance(n, Decimal) else Decimal(n))


def is_currency(v):
    """A Currency value: a Decimal that CDec did not make."""
    return isinstance(v, Decimal) and not isinstance(v, CDecValue)


def like_decimal(original, result):
    """The result of an operation on one Decimal, as the kind the original was: CDec's stays CDec's."""
    return CDecValue(result) if isinstance(original, CDecValue) else result


def to_decimal(v):
    n = to_number(v)
    if isinstance(n, Decimal):
        return n
    try:
        return Decimal(repr(n)) if isinstance(n, float) else Decimal(n)
    except InvalidOperation:
        raise type_mismatch() from None


def to_bool(v):
    v = default_value(v)
    if isinstance(v, bool):
        return v
    if v is EMPTY:
        return False
    if v is NULL:
        raise VBARuntimeError(94, "Invalid use of Null")
    if isinstance(v, (int, float, Decimal)):
        return v != 0
    if isinstance(v, dt.datetime):
        return date_to_serial(v) != 0
    if isinstance(v, str):
        s = v.strip().lower()
        if s == "true":
            return True
        if s == "false":
            return False
        n = parse_number_text(v)
        if n is None:
            raise type_mismatch()
        return n != 0
    raise type_mismatch()


def condition_true(v):
    """The test of If/While/Do: Null counts as False."""
    v = default_value(v)
    if v is NULL:
        return False
    return to_bool(v)


def date_to_serial(d):
    delta = d - EPOCH
    return delta.days + delta.seconds / 86400 + delta.microseconds / 86400e6


def serial_to_date(x):
    x = float(x)
    if not -657434 <= x < 2958466:  # 1 January 100 to 31 December 9999
        raise VBARuntimeError(6, "Overflow")
    return EPOCH + dt.timedelta(days=x)


MONTH_NAMES = (
    "january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november",
    "december",
)  # fmt: skip
_TIME_TEXT = re.compile(r"^(\d{1,2})(?::(\d{1,2}))?(?::(\d{1,2}))?\s*([AaPp][Mm])?$")
_TIME_AT_END = re.compile(r"(?:\d{1,2}:\d{1,2}(?::\d{1,2})?\s*(?:[AaPp][Mm])?|\d{1,2}\s*[AaPp][Mm])\s*$")


def two_digit_year(year):
    """The year VBA and Excel read in a year of two digits: 0 to 29 are 2000 to 2029, 30 to 99 are 1930 to 1999."""
    return year + (2000 if year < 30 else 1900) if 0 <= year <= 99 else year


def _time_of(text):
    m = _TIME_TEXT.match(text.strip())
    if m is None or (m.group(2) is None and m.group(4) is None):
        return None
    hour, minute, second, half = int(m.group(1)), int(m.group(2) or 0), int(m.group(3) or 0), (m.group(4) or "").lower()
    if half:
        if not 1 <= hour <= 12:
            return None
        hour = hour % 12 + (12 if half == "pm" else 0)
    return (hour, minute, second) if hour < 24 and minute < 60 and second < 60 else None


def _date_of(text):
    """(year, month, day) of a date's text, None when it is not one. A day and month in digits are read month first,
    as US regional settings read them (VBA reads them as the machine's settings say), unless only the other order
    is a date: 13/5/2024 is 13 May."""
    if re.sub(r"\d+|[A-Za-z]+\.?|[\s,/\-]", "", text):
        return None
    tokens = re.findall(r"\d+|[A-Za-z]+", text)
    words, numbers = [t for t in tokens if t[0].isalpha()], [t for t in tokens if t[0].isdigit()]
    if len(words) > 1 or not numbers:
        return None
    if words:  # 5 Jan 2024, Jan 5, 2024, 5-Jan-24, January 2024
        word = words[0].lower()
        month = next((i for i, m in enumerate(MONTH_NAMES, 1) if word == m or word == m[:3]), None)
        if month is None:
            return None
        if len(numbers) == 2:
            first, second = numbers
            day, year = (second, first) if len(first) == 4 else (first, second)
            return two_digit_year(int(year)) if len(year) <= 2 else int(year), month, int(day)
        if len(numbers) == 1 and (len(numbers[0]) == 4 or int(numbers[0]) > 31):
            return int(numbers[0]), month, 1
        raise VBANotImplemented(f"the date {text.strip()!r}: it has no year, and VBA takes the current one")
    if len(numbers) == 3:
        a, b, c = numbers
        if len(a) == 4 or int(a) > 31:  # 2024-01-05
            return int(a), int(b), int(c)
        month, day = int(a), int(b)
        if month > 12 >= day:
            month, day = day, month
        elif day <= 12 and month != day:
            note("DATE-TEXT-ORDER", text.strip())
        return two_digit_year(int(c)) if len(c) <= 2 else int(c), month, day
    if len(numbers) == 2:
        raise VBANotImplemented(f"the date {text.strip()!r}: it has no year, and VBA takes the current one")
    return None


def parse_date_text(text):
    """The date and time a text holds, as VBA reads it with US regional settings; None when it holds neither."""
    text = text.strip()
    time_match = _TIME_AT_END.search(text)
    time_text = time_match.group(0) if time_match else ""
    date_text = text[: time_match.start()].strip() if time_match else text
    time = _time_of(time_text) if time_text else (0, 0, 0)
    if time is None or not (date_text or time_text):
        return None
    if not date_text:
        return EPOCH.replace(hour=time[0], minute=time[1], second=time[2])
    date = _date_of(date_text)
    if date is None:
        return None
    try:
        return dt.datetime(*date, *time)
    except ValueError:
        return None


def to_date(v):
    v = default_value(v)
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, (int, float, Decimal)) and not isinstance(v, bool):
        return serial_to_date(v)
    if v is EMPTY:
        return EPOCH
    if isinstance(v, str):
        n = parse_number_text(v)
        if n is not None:
            try:
                return serial_to_date(n)
            except VBARuntimeError:
                raise type_mismatch() from None
        d = parse_date_text(v)
        if d is not None:
            return d
    raise type_mismatch()


def format_double(x):
    """CStr of a Double: up to 15 significant digits (assumes '.' as decimal separator)."""
    if x == 0:
        return "0"
    if float(x).is_integer() and abs(x) < 1e15:
        return str(int(x))
    s = f"{x:.15G}"
    if "E" in s:
        mant, exp = s.split("E")
        if "." in mant:
            mant = mant.rstrip("0").rstrip(".")
        sign = "-" if exp.startswith("-") else "+"
        return f"{mant}E{sign}{exp.lstrip('+-').zfill(2)}"
    if "." in s:
        s = s.rstrip("0").rstrip(".")
    return s


def format_date(d):
    """CStr of a Date in the US short format (VBA uses the machine's regional settings): the date, the time, or both.
    A time on its own is one with no date (day 0)."""
    note("DATE-TEXT-US")
    has_date = d.date() != EPOCH.date()
    has_time = (d.hour, d.minute, d.second) != (0, 0, 0) or not has_date
    date_part = f"{d.month}/{d.day}/{d.year}"
    hour12 = d.hour % 12 or 12
    time_part = f"{hour12}:{d.minute:02d}:{d.second:02d} {'AM' if d.hour < 12 else 'PM'}"
    if has_date and has_time:
        return f"{date_part} {time_part}"
    if has_time:
        return time_part
    return date_part


def to_str(v):
    v = default_value(v)
    if isinstance(v, XLError):
        raise type_mismatch()
    if isinstance(v, str):
        return str(v)
    if isinstance(v, bool):
        return "True" if v else "False"
    if isinstance(v, int):
        return str(v)
    if isinstance(v, float):
        return format_double(v)
    if isinstance(v, Decimal):
        s = format(v.normalize(), "f")
        return s
    if isinstance(v, dt.datetime):
        return format_date(v)
    if v is EMPTY:
        return ""
    if v is NULL:
        raise VBARuntimeError(94, "Invalid use of Null")
    raise type_mismatch()


def is_numeric_value(v):
    return isinstance(v, (int, float, Decimal)) and not isinstance(v, bool)


# ---------------------------------------------------------------------------
# Operators
# ---------------------------------------------------------------------------


def _num_pair(a, b):
    return to_number(a), to_number(b)


def _numeric_result(x):
    return x


def _date_kind(v):
    return isinstance(v, dt.datetime)


def op_add(a, b):
    a, b = default_value(a), default_value(b)
    if a is NULL or b is NULL:
        return NULL
    if isinstance(a, str) and isinstance(b, str):
        return a + b
    if isinstance(a, str) and b is EMPTY:
        return a
    if isinstance(b, str) and a is EMPTY:
        return b
    if _date_kind(a) or _date_kind(b):
        return serial_to_date(float(to_number(a)) + float(to_number(b)))
    return _sum(*_num_pair(a, b), 1)


def op_sub(a, b):
    a, b = default_value(a), default_value(b)
    if a is NULL or b is NULL:
        return NULL
    if _date_kind(a) and _date_kind(b):
        return date_to_serial(a) - date_to_serial(b)
    if _date_kind(a):
        return serial_to_date(date_to_serial(a) - float(to_number(b)))
    return _sum(*_num_pair(a, b), -1)


def _sum(x, y, sign):
    """x + y, or x - y for sign -1. A Currency with any other number gives a Currency: the other is made one first
    (to four decimals), so CCur(1.5) + 1.00005 is 2.5001."""
    if isinstance(x, CDecValue) or isinstance(y, CDecValue):
        return CDecValue(to_decimal(x) + sign * to_decimal(y))
    if isinstance(x, Decimal) or isinstance(y, Decimal):
        return _currency(to_currency(x) + sign * to_currency(y))
    if isinstance(x, int) and isinstance(y, int):
        return x + sign * y
    return _finite(float(x) + sign * float(y))


def op_mul(a, b):
    """A Currency times a whole number or another Currency is a Currency; times a Double it is a Double."""
    a, b = default_value(a), default_value(b)
    if a is NULL or b is NULL:
        return NULL
    x, y = _num_pair(a, b)
    if isinstance(x, CDecValue) or isinstance(y, CDecValue):
        return CDecValue(to_decimal(x) * to_decimal(y))
    if isinstance(x, int) and isinstance(y, int):
        return x * y
    if not isinstance(x, float) and not isinstance(y, float):
        return _currency(x * y)
    return _finite(float(x) * float(y))


def _finite(x):
    """A Double result: where Python's float would run to inf, VBA stops with Overflow."""
    if not math.isfinite(x):
        raise VBARuntimeError(6, "Overflow")
    return x


def op_div(a, b):
    a, b = default_value(a), default_value(b)
    if a is NULL or b is NULL:
        return NULL
    x, y = _num_pair(a, b)
    if y == 0:
        if x == 0:
            raise VBARuntimeError(6, "Overflow")
        raise VBARuntimeError(11, "Division by zero")
    if isinstance(x, CDecValue) or isinstance(y, CDecValue):
        return CDecValue(to_decimal(x) / to_decimal(y))
    return _finite(float(x) / float(y))  # a Currency divided, or divided by, gives a Double


def op_intdiv(a, b):
    a, b = default_value(a), default_value(b)
    if a is NULL or b is NULL:
        return NULL
    x, y = to_long(a), to_long(b)
    if y == 0:
        raise VBARuntimeError(11, "Division by zero")
    q = abs(x) // abs(y)
    return q if (x >= 0) == (y >= 0) else -q


def op_mod(a, b):
    a, b = default_value(a), default_value(b)
    if a is NULL or b is NULL:
        return NULL
    x, y = to_long(a), to_long(b)
    if y == 0:
        raise VBARuntimeError(11, "Division by zero")
    r = abs(x) % abs(y)
    return r if x >= 0 else -r


def op_pow(a, b):
    a, b = default_value(a), default_value(b)
    if a is NULL or b is NULL:
        return NULL
    x, y = float(to_number(a)), float(to_number(b))
    try:
        r = x**y
    except ZeroDivisionError:
        raise VBARuntimeError(5, "Invalid procedure call or argument") from None
    except OverflowError:
        raise VBARuntimeError(6, "Overflow") from None
    if isinstance(r, complex):
        raise VBARuntimeError(5, "Invalid procedure call or argument")
    return _finite(r)


def op_concat(a, b):
    a, b = default_value(a), default_value(b)
    if a is NULL and b is NULL:
        return NULL
    return ("" if a is NULL else to_str(a)) + ("" if b is NULL else to_str(b))


def op_neg(a):
    a = default_value(a)
    if a is NULL:
        return NULL
    if _date_kind(a):
        return serial_to_date(-date_to_serial(a))
    n = to_number(a)
    return like_decimal(n, -n)


def _logical_operands(a, b):
    a, b = default_value(a), default_value(b)
    return a, b


def _as_int_bits(v):
    if isinstance(v, bool):
        return -1 if v else 0
    return to_long(v)


def op_not(a):
    a = default_value(a)
    if a is NULL:
        return NULL
    if isinstance(a, bool):
        return not a
    if a is EMPTY:
        return -1
    if isinstance(a, str):
        try:
            return not to_bool(a) if a.strip().lower() in ("true", "false") else ~to_long(a)
        except VBARuntimeError:
            raise
    return ~_as_int_bits(a)


def _logical(a, b, fn_bool, fn_int):
    a, b = _logical_operands(a, b)
    if a is NULL or b is NULL:
        return NULL
    if isinstance(a, bool) and isinstance(b, bool):
        return fn_bool(a, b)
    if isinstance(a, str) and a.strip().lower() in ("true", "false"):
        a = to_bool(a)
    if isinstance(b, str) and b.strip().lower() in ("true", "false"):
        b = to_bool(b)
    if isinstance(a, bool) and isinstance(b, bool):
        return fn_bool(a, b)
    return fn_int(_as_int_bits(a), _as_int_bits(b))


def op_and(a, b):
    return _logical(a, b, lambda x, y: x and y, lambda x, y: x & y)


def op_or(a, b):
    return _logical(a, b, lambda x, y: x or y, lambda x, y: x | y)


def op_xor(a, b):
    return _logical(a, b, lambda x, y: x != y, lambda x, y: x ^ y)


def op_eqv(a, b):
    return _logical(a, b, lambda x, y: x == y, lambda x, y: ~(x ^ y))


def op_imp(a, b):
    return _logical(a, b, lambda x, y: (not x) or y, lambda x, y: (~x) | y)


def compare(a, b, text_compare=False, kinds="vv"):
    """Three-way comparison: -1, 0, 1, or NULL.

    When a number meets a string, VBA goes by how each side is declared, which `kinds` gives for the two sides:
    "n" a numeric type, "s" a String, "v" a Variant.
      Variant and Variant     the number is smaller, whatever the string holds
      Variant and a String    both are compared as text
      a number and a Variant  as numbers if the Variant's text is a number, else the number is smaller
      a number and a String   as numbers; error 13 if the text is not a number
    """
    a, b = default_value(a), default_value(b)
    if a is NULL or b is NULL:
        return NULL
    if isinstance(a, XLError) or isinstance(b, XLError):
        raise type_mismatch()
    if a is EMPTY and b is EMPTY:
        return 0
    if a is EMPTY:
        a = "" if isinstance(b, str) else 0
    if b is EMPTY:
        b = "" if isinstance(a, str) else 0
    if isinstance(a, str) and isinstance(b, str):
        x, y = (a.casefold(), b.casefold()) if text_compare else (a, b)
        return (x > y) - (x < y)
    a_num = not isinstance(a, str)
    b_num = not isinstance(b, str)
    if a_num and b_num:
        x, y = to_number(a), to_number(b)
        if isinstance(x, CDecValue) or isinstance(y, CDecValue):
            x, y = to_decimal(x), to_decimal(y)
        elif isinstance(x, Decimal) or isinstance(y, Decimal):
            x, y = to_currency(x), to_currency(y)  # a Currency against a Double: the Double to four decimals
        return (x > y) - (x < y)
    # a number against a string
    text_kind, number_kind = (kinds[1], kinds[0]) if a_num else (kinds[0], kinds[1])
    if text_kind == "s" and number_kind != "n":
        x, y = (to_str(a), b) if a_num else (a, to_str(b))
        if text_compare:
            x, y = x.casefold(), y.casefold()
        return (x > y) - (x < y)
    if number_kind == "n":
        n = parse_number_text(b if a_num else a)
        if n is not None:
            x, y = (to_number(a), n) if a_num else (n, to_number(b))
            return (x > y) - (x < y)
        if text_kind == "s":
            raise type_mismatch()
    return -1 if a_num else 1


def like(text, pattern, text_compare=False):
    s = to_str(text)
    p = to_str(pattern)
    out = []
    i = 0
    while i < len(p):
        ch = p[i]
        if ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        elif ch == "#":
            out.append(r"\d")
        elif ch == "[":
            j = p.find("]", i + 1)
            if j < 0:
                raise VBARuntimeError(93, "Invalid pattern string")
            body = p[i + 1 : j]
            neg = body.startswith("!")
            if neg:
                body = body[1:]
            out.append("[" + ("^" if neg else "") + body.replace("\\", "\\\\") + "]")
            i = j
        else:
            out.append(re.escape(ch))
        i += 1
    flags = re.DOTALL | (re.IGNORECASE if text_compare else 0)
    return re.fullmatch("".join(out), s, flags) is not None


def vba_round(x, digits=0):
    """VBA Round(): the number times 10^digits, as a Double, rounded to a whole number with halves to even.
    So Round(2.5) is 2, Round(2.675, 2) is 2.68 (267.5 exactly) and Round(1.015, 2) is 1.01 (101.49999999999999)."""
    n = to_number(x)
    digits = int(digits)
    if digits < 0:
        raise VBARuntimeError(5, "Invalid procedure call or argument")
    if isinstance(n, int):
        return n
    if isinstance(n, Decimal):
        return like_decimal(n, n.quantize(Decimal(1).scaleb(-digits), rounding=ROUND_HALF_EVEN))
    scaled = n * 10**digits
    if not math.isfinite(scaled):
        return n
    return round(scaled) / 10**digits


class CDecValue(Decimal):
    """What CDec gives: a Variant of subtype Decimal. Any other Decimal here is a Currency."""


def type_name(v):
    v_ = v
    if isinstance(v_, VBAObject):
        return v_.vba_type_name
    if v_ is EMPTY:
        return "Empty"
    if v_ is NULL:
        return "Null"
    if v_ is NOTHING:
        return "Nothing"
    if isinstance(v_, bool):
        return "Boolean"
    if isinstance(v_, int):
        return "Long" if not -32768 <= v_ <= 32767 else "Integer"
    if isinstance(v_, float):
        return "Double"
    if isinstance(v_, Decimal):
        return "Decimal" if isinstance(v_, CDecValue) else "Currency"
    if isinstance(v_, dt.datetime):
        return "Date"
    if isinstance(v_, XLError):
        return "Error"
    if isinstance(v_, str):
        return "String"
    if isinstance(v_, VBAArray):
        return "Variant()"
    return type(v_).__name__
