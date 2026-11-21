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

from .errors import VBANotImplemented, VBARuntimeError


class _Special:
    __slots__ = ("name",)

    def __init__(self, name):
        self.name = name

    def __repr__(self):
        return self.name


EMPTY = _Special("Empty")
NULL = _Special("Null")
NOTHING = _Special("Nothing")

EPOCH = dt.datetime(1899, 12, 30)


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
    if t in ("currency", "decimal"):
        return (
            to_decimal(value).quantize(Decimal("0.0001"), rounding=ROUND_HALF_EVEN)
            if t == "currency"
            else to_decimal(value)
        )
    raise VBANotImplemented(f"variables of type {type_name}")


def _int_range(n, lo, hi):
    if n < lo or n > hi:
        raise VBARuntimeError(6, "Overflow")
    return n


# ---------------------------------------------------------------------------
# Conversions
# ---------------------------------------------------------------------------

_NUM_TEXT = re.compile(r"^\s*([-+]?)\s*(\d[\d,]*\.?\d*|\.\d+)([eEdD][-+]?\d+)?\s*$")
_HEX_TEXT = re.compile(r"^\s*&[hH]([0-9a-fA-F]+)&?\s*$")


def parse_number_text(s):
    """VBA's string-to-number conversion (assumes '.' decimal and ',' thousands separators)."""
    m = _NUM_TEXT.match(s)
    if m:
        sign, body, exp = m.groups()
        body = body.replace(",", "")
        text = sign + body + (exp.replace("d", "e").replace("D", "e") if exp else "")
        try:
            value = float(text)
        except ValueError:
            return None
        return value
    m = _HEX_TEXT.match(s)
    if m:
        return int(m.group(1), 16)
    if s.strip().lower() in ("true", "false"):
        return None
    return None


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
        if n is None:
            if v.strip().lower() == "true":
                return -1
            if v.strip().lower() == "false":
                return 0
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
    return EPOCH + dt.timedelta(days=float(x))


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
            return serial_to_date(n)
        for fmt in (
            "%m/%d/%Y %H:%M:%S",
            "%m/%d/%Y %H:%M",
            "%m/%d/%Y",
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d",
            "%H:%M:%S",
            "%H:%M",
        ):
            try:
                d = dt.datetime.strptime(v.strip(), fmt)
                if fmt.startswith("%H"):
                    d = EPOCH.replace(hour=d.hour, minute=d.minute, second=d.second)
                return d
            except ValueError:
                pass
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
    """CStr of a Date in the US short format (VBA uses the machine's locale)."""
    has_time = (d.hour, d.minute, d.second) != (0, 0, 0)
    has_date = d.date() != EPOCH.date()
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
        return serial_to_date(to_number(a) + to_number(b))
    x, y = _num_pair(a, b)
    return _mix(x, y, x + y if not _mixed_decimal(x, y) else Decimal(str(x)) + Decimal(str(y)))


def op_sub(a, b):
    a, b = default_value(a), default_value(b)
    if a is NULL or b is NULL:
        return NULL
    if _date_kind(a) and _date_kind(b):
        return date_to_serial(a) - date_to_serial(b)
    if _date_kind(a):
        return serial_to_date(date_to_serial(a) - to_number(b))
    x, y = _num_pair(a, b)
    return _mix(x, y, x - y if not _mixed_decimal(x, y) else Decimal(str(x)) - Decimal(str(y)))


def op_mul(a, b):
    a, b = default_value(a), default_value(b)
    if a is NULL or b is NULL:
        return NULL
    x, y = _num_pair(a, b)
    return _mix(x, y, x * y if not _mixed_decimal(x, y) else Decimal(str(x)) * Decimal(str(y)))


def _mixed_decimal(x, y):
    return isinstance(x, Decimal) != isinstance(y, Decimal) and (isinstance(x, Decimal) or isinstance(y, Decimal))


def _mix(x, y, result):
    if isinstance(x, Decimal) or isinstance(y, Decimal):
        return result if isinstance(result, Decimal) else Decimal(str(result))
    if isinstance(x, int) and isinstance(y, int):
        return int(result)
    return _finite(float(result))


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
    if isinstance(x, Decimal) or isinstance(y, Decimal):
        return to_decimal(x) / to_decimal(y)
    return _finite(float(x) / float(y))


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
    return -n


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


def compare(a, b, text_compare=False, on_num_vs_text=None):
    """Three-way Variant comparison: -1, 0, 1, or NULL.

    on_num_vs_text is called when a number meets a non-numeric string (recorded as an
    assumption: VBA treats the number as smaller, without an error).
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
        if isinstance(x, Decimal) or isinstance(y, Decimal):
            x, y = to_decimal(x), to_decimal(y)
        return (x > y) - (x < y)
    # number against string
    s = b if a_num else a
    n = parse_number_text(s)
    if n is not None:
        x, y = (to_number(a), n) if a_num else (n, to_number(b))
        return (x > y) - (x < y)
    if on_num_vs_text:
        on_num_vs_text()
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
    """VBA Round(): halves to even, on the decimal form of the number."""
    n = to_number(x)
    d = to_decimal(n)
    q = Decimal(1).scaleb(-int(digits))
    r = d.quantize(q, rounding=ROUND_HALF_EVEN)
    if isinstance(n, int) and digits >= 0:
        return int(r)
    if isinstance(n, Decimal):
        return r
    return float(r)


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
        return "Decimal"
    if isinstance(v_, dt.datetime):
        return "Date"
    if isinstance(v_, XLError):
        return "Error"
    if isinstance(v_, str):
        return "String"
    if isinstance(v_, VBAArray):
        return "Variant()"
    return type(v_).__name__
