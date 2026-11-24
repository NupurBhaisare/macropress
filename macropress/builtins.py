"""VBA's built-in functions, constants, and the Err and Debug objects."""

import calendar
import datetime as dt
import inspect
import math
import re
from decimal import ROUND_HALF_UP, Decimal

from .errors import VBANotImplemented, VBARuntimeError
from .values import (
    EMPTY,
    EPOCH,
    NOTHING,
    NULL,
    CDecValue,
    VBAArray,
    VBAObject,
    XLError,
    _Special,
    bankers_int,
    compare,
    default_value,
    hex_value,
    is_numeric_value,
    like_decimal,
    parse_number_text,
    to_bool,
    to_currency,
    to_date,
    to_decimal,
    to_double,
    to_long,
    to_number,
    to_str,
    two_digit_year,
    type_name,
    vba_round,
)

MISSING = _Special("Missing")


class ErrObject(VBAObject):
    vba_type_name = "ErrObject"

    def __init__(self):
        self.number = 0
        self.description = ""
        self.source = ""

    def set(self, number, description):
        self.number, self.description = number, description

    def clear(self):
        self.number, self.description, self.source = 0, "", ""

    def vba_default_get(self):
        return self.number

    def get_number(self):
        return self.number

    def let_number(self, v):
        self.number = to_long(v)

    def get_description(self):
        return self.description

    def let_description(self, v):
        self.description = to_str(v)

    def get_source(self):
        return self.source

    def m_clear(self):
        self.clear()

    def m_raise(self, number, source=None, description=None, helpfile=None, helpcontext=None):
        raise VBARuntimeError(
            to_long(number),
            to_str(description) if description is not None else "Application-defined or object-defined error",
        )


class DebugObject(VBAObject):
    vba_type_name = "Debug"

    def __init__(self, ctx):
        self.ctx = ctx

    def m_print(self, *args):
        self.ctx.messages.append(("Debug.Print", " ".join(error_text(a) or to_str(a) for a in args)))

    def m_assert(self, cond):
        return None


VB_CONSTANTS = {
    "true": True,
    "false": False,
    "empty": EMPTY,
    "null": NULL,
    "nothing": NOTHING,
    "vbcrlf": "\r\n",
    "vbcr": "\r",
    "vblf": "\n",
    "vbnewline": "\r\n",
    "vbtab": "\t",
    "vbnullstring": "",
    "vbnullchar": "\0",
    "vbback": "\b",
    "vbformfeed": "\f",
    "vbverticaltab": "\v",
    "vbok": 1,
    "vbcancel": 2,
    "vbabort": 3,
    "vbretry": 4,
    "vbignore": 5,
    "vbyes": 6,
    "vbno": 7,
    "vbokonly": 0,
    "vbokcancel": 1,
    "vbabortretryignore": 2,
    "vbyesnocancel": 3,
    "vbyesno": 4,
    "vbretrycancel": 5,
    "vbcritical": 16,
    "vbquestion": 32,
    "vbexclamation": 48,
    "vbinformation": 64,
    "vbdefaultbutton1": 0,
    "vbdefaultbutton2": 256,
    "vbdefaultbutton3": 512,
    "vbbinarycompare": 0,
    "vbtextcompare": 1,
    "vbempty": 0,
    "vbnull": 1,
    "vbinteger": 2,
    "vblong": 3,
    "vbsingle": 4,
    "vbdouble": 5,
    "vbcurrency": 6,
    "vbdate": 7,
    "vbstring": 8,
    "vbobject": 9,
    "vberror": 10,
    "vbboolean": 11,
    "vbvariant": 12,
    "vbdecimal": 14,
    "vbbyte": 17,
    "vbarray": 8192,
    "vbsunday": 1,
    "vbmonday": 2,
    "vbtuesday": 3,
    "vbwednesday": 4,
    "vbthursday": 5,
    "vbfriday": 6,
    "vbsaturday": 7,
    "vbusesystemdayofweek": 0,
    "vbuppercase": 1,
    "vblowercase": 2,
    "vbpropercase": 3,
    "vbobjecterror": -2147221504,
    "vbdatabasecompare": 2,
    "vbdefaultbutton4": 768,
    "vbapplicationmodal": 0,
    "vbsystemmodal": 4096,
    "vbmsgboxhelpbutton": 16384,
    "vbmsgboxsetforeground": 65536,
    "vbmsgboxright": 524288,
    "vbmsgboxrtlreading": 1048576,
    "vbusesystem": 0,
    "vbfirstjan1": 1,
    "vbfirstfourdays": 2,
    "vbfirstfullweek": 3,
    "vbgeneraldate": 0,
    "vblongdate": 1,
    "vbshortdate": 2,
    "vblongtime": 3,
    "vbshorttime": 4,
    "vbtrue": -1,
    "vbfalse": 0,
    "vbusedefault": -2,
    "vbblack": 0,
    "vbred": 255,
    "vbgreen": 65280,
    "vbyellow": 65535,
    "vbblue": 16711680,
    "vbmagenta": 16711935,
    "vbcyan": 16776960,
    "vbwhite": 16777215,
    "vbnormal": 0,
    "vbreadonly": 1,
    "vbhidden": 2,
    "vbsystem": 4,
    "vbvolume": 8,
    "vbdirectory": 16,
    "vbarchive": 32,
    "vbwide": 4,
    "vbnarrow": 8,
    "vbunicode": 64,
    "vbfromunicode": 128,
    "vbmethod": 1,
    "vbget": 2,
    "vblet": 4,
    "vbset": 8,
    "vbdataobject": 13,
    "vbuserdefinedtype": 36,
    "vblonglong": 20,
}

ERROR_CODES = {
    2000: "#NULL!",
    2007: "#DIV/0!",
    2015: "#VALUE!",
    2023: "#REF!",
    2029: "#NAME?",
    2036: "#NUM!",
    2042: "#N/A",
    2045: "#SPILL!",
    2050: "#CALC!",
}


def _null(*vals):
    return any(default_value(v) is NULL for v in vals)


def _compare_mode(ctx, frame, compare_arg):
    if compare_arg is None:
        return ctx.text_compare(frame)
    return to_long(compare_arg) == 1


# --- strings ---


def b_len(ctx, frame, x):
    x = default_value(x)
    if x is NULL:
        return NULL
    return len(to_str(x))


def b_left(ctx, frame, s, length):
    if _null(s):
        return NULL
    n = to_long(length)
    if n < 0:
        raise VBARuntimeError(5, "Invalid procedure call or argument")
    return to_str(s)[:n]


def b_right(ctx, frame, s, length):
    if _null(s):
        return NULL
    n = to_long(length)
    if n < 0:
        raise VBARuntimeError(5, "Invalid procedure call or argument")
    text = to_str(s)
    return text[max(len(text) - n, 0) :]


def b_mid(ctx, frame, s, start, length=None):
    if _null(s):
        return NULL
    st = to_long(start)
    if st < 1:
        raise VBARuntimeError(5, "Invalid procedure call or argument")
    text = to_str(s)
    if length is None:
        return text[st - 1 :]
    ln = to_long(length)
    if ln < 0:
        raise VBARuntimeError(5, "Invalid procedure call or argument")
    return text[st - 1 : st - 1 + ln]


def b_trim(ctx, frame, s):
    return NULL if _null(s) else to_str(s).strip(" ")


def b_ltrim(ctx, frame, s):
    return NULL if _null(s) else to_str(s).lstrip(" ")


def b_rtrim(ctx, frame, s):
    return NULL if _null(s) else to_str(s).rstrip(" ")


def b_ucase(ctx, frame, s):
    return NULL if _null(s) else to_str(s).upper()


def b_lcase(ctx, frame, s):
    return NULL if _null(s) else to_str(s).lower()


def b_instr(ctx, frame, *args, compare_=None):
    # InStr(string1, string2) or InStr(start, string1, string2[, compare])
    if len(args) >= 3:
        start, s1, s2 = to_long(args[0]), args[1], args[2]
        cmp_ = args[3] if len(args) > 3 else compare_
    elif len(args) == 2:
        start, s1, s2, cmp_ = 1, args[0], args[1], compare_
    else:
        raise VBARuntimeError(450, "Wrong number of arguments for InStr")
    if start < 1:
        raise VBARuntimeError(5, "Invalid procedure call or argument")
    if _null(s1, s2):
        return NULL
    a, b = to_str(s1), to_str(s2)
    if start > len(a):
        return 0
    if b == "":
        return start
    if _compare_mode(ctx, frame, cmp_):
        a, b = a.casefold(), b.casefold()
    return a.find(b, start - 1) + 1


def b_instrrev(ctx, frame, s1, s2, start=-1, compare_=None):
    if _null(s1, s2):
        return NULL
    a, b = to_str(s1), to_str(s2)
    st = to_long(start)
    if st == -1:
        st = len(a)
    if b == "":
        return st
    if _compare_mode(ctx, frame, compare_):
        a, b = a.casefold(), b.casefold()
    return a.rfind(b, 0, st) + 1


def b_replace(ctx, frame, expression, find, replace, start=1, count=-1, compare_=None):
    text, f, r = to_str(expression), to_str(find), to_str(replace)
    st, cnt = to_long(start), to_long(count)
    text = text[st - 1 :]
    if f == "":
        return text
    if _compare_mode(ctx, frame, compare_):
        out, i, done = [], 0, 0
        low, lf = text.casefold(), f.casefold()
        while True:
            j = low.find(lf, i)
            if j < 0 or (cnt >= 0 and done >= cnt):
                out.append(text[i:])
                break
            out.append(text[i:j] + r)
            i = j + len(f)
            done += 1
        return "".join(out)
    return text.replace(f, r) if cnt < 0 else text.replace(f, r, cnt)


def b_split(ctx, frame, expression, delimiter=" ", limit=-1, compare_=None):
    text = to_str(expression)
    if text == "":
        return VBAArray([(0, -1)], "string")
    d = to_str(delimiter)
    lim = to_long(limit)
    parts = [text] if d == "" else text.split(d, lim - 1 if lim > 0 else -1)
    arr = VBAArray.from_list(parts)
    arr.elem_type = "string"
    return arr


def b_join(ctx, frame, arr, delimiter=" "):
    if not isinstance(arr, VBAArray):
        raise VBARuntimeError(13, "Type mismatch")
    return to_str(delimiter).join(to_str(v) for v in arr.items_1d())


def b_strcomp(ctx, frame, a, b, compare_=None):
    if _null(a, b):
        return NULL
    return compare(to_str(a), to_str(b), _compare_mode(ctx, frame, compare_))


def b_strreverse(ctx, frame, s):
    return to_str(s)[::-1]


def b_space(ctx, frame, n):
    return " " * to_long(n)


def b_string(ctx, frame, n, ch):
    ch = default_value(ch)
    c = chr(to_long(ch)) if is_numeric_value(ch) else to_str(ch)[:1]
    return c * to_long(n)


def b_chr(ctx, frame, code):
    n = to_long(code)
    if 128 <= n <= 159:  # Windows-1252, as VBA's Chr: Chr(146) is a right quote, not a control character
        try:
            return bytes([n]).decode("cp1252")
        except UnicodeDecodeError:
            pass
    return chr(n)


def b_asc(ctx, frame, s):
    text = to_str(s)
    if not text:
        raise VBARuntimeError(5, "Invalid procedure call or argument")
    if ord(text[0]) < 128:
        return ord(text[0])
    try:  # the character's Windows-1252 code, as VBA's Asc; one outside it reads as "?"
        return text[0].encode("cp1252")[0]
    except UnicodeEncodeError:
        return 63


# The number Val reads, in VBA's own grammar (D is an exponent too): "inf", "nan" or "1_000" are not numbers in it.
_VAL_NUMBER = re.compile(r"[+-]?(?:\d+\.?\d*|\.\d+)(?:[eEdD][+-]?\d+)?")


def b_val(ctx, frame, s):
    """The number the text starts with, read past spaces, tabs and line feeds; &H and &O start hex and octal."""
    text = "".join(ch for ch in to_str(s) if ch not in " \t\n\r")
    radix = {"&h": 16, "&o": 8}.get(text[:2].lower())
    if radix:
        digits = re.match(r"[0-9a-fA-F]*" if radix == 16 else r"[0-7]*", text[2:]).group()
        return float(hex_value(int(digits, radix))) if digits else 0.0
    m = _VAL_NUMBER.match(text)
    x = float(m.group().translate(str.maketrans("dD", "ee"))) if m else 0.0
    if math.isinf(x):
        raise VBARuntimeError(6, "Overflow")
    return x


def b_str(ctx, frame, x):
    x = default_value(x)
    if x is NULL:
        return NULL
    s = to_str(x)
    if not is_numeric_value(x):
        return s
    if s.startswith(("0.", "-0.")):  # Str writes a fraction without its zero: " .5", "-.25"
        s = s.replace("0.", ".", 1)
    return " " + s if to_number(x) >= 0 else s


# --- conversions ---


def b_cstr(ctx, frame, x):
    return error_text(x) or to_str(x)


def error_text(x):
    """How CStr and Debug.Print write an error value ("Error 2042"); None for anything else."""
    v = default_value(x)
    if not isinstance(v, XLError):
        return None
    return f"Error {next((code for code, text in ERROR_CODES.items() if text == v), 2015)}"


def b_cint(ctx, frame, x):
    n = bankers_int(to_number(x))
    if not -32768 <= n <= 32767:
        raise VBARuntimeError(6, "Overflow")
    return n


def b_clng(ctx, frame, x):
    return to_long(x)


def b_cdbl(ctx, frame, x):
    return to_double(x)


def b_csng(ctx, frame, x):
    return float(to_number(x))


def b_cdec(ctx, frame, x):
    ctx.ledger.note("CDEC-WINDOWS-ONLY")
    return CDecValue(to_decimal(x))


def b_ccur(ctx, frame, x):
    return to_currency(x)


def b_cbool(ctx, frame, x):
    return to_bool(x)


def b_cdate(ctx, frame, x):
    return to_date(x)


def b_cvar(ctx, frame, x):
    return default_value(x)


def b_cbyte(ctx, frame, x):
    n = to_long(x)
    if not 0 <= n <= 255:
        raise VBARuntimeError(6, "Overflow")
    return n


def b_int(ctx, frame, x):
    v = to_number(x)
    if isinstance(v, int):
        return v
    r = math.floor(v)
    return float(r) if isinstance(v, float) else like_decimal(v, Decimal(r))


def b_fix(ctx, frame, x):
    v = to_number(x)
    if isinstance(v, int):
        return v
    r = math.trunc(v)
    return float(r) if isinstance(v, float) else like_decimal(v, Decimal(r))


def b_abs(ctx, frame, x):
    v = to_number(x)
    return like_decimal(v, abs(v))


def b_sgn(ctx, frame, x):
    v = to_number(x)
    return (v > 0) - (v < 0)


def b_sqr(ctx, frame, x):
    v = to_double(x)
    if v < 0:
        raise VBARuntimeError(5, "Invalid procedure call or argument")
    return math.sqrt(v)


def b_exp(ctx, frame, x):
    return math.exp(to_double(x))


def b_log(ctx, frame, x):
    v = to_double(x)
    if v <= 0:
        raise VBARuntimeError(5, "Invalid procedure call or argument")
    return math.log(v)


def b_round(ctx, frame, x, digits=0):
    ctx.ledger.note("ROUND-HALF-EVEN")
    return vba_round(default_value(x), to_long(digits))


# --- dates ---


def b_now(ctx, frame):
    return ctx.clock.read().replace(microsecond=0)


def b_date(ctx, frame):
    t = ctx.clock.read()
    return dt.datetime(t.year, t.month, t.day)


def b_time(ctx, frame):
    t = ctx.clock.read()
    return EPOCH.replace(hour=t.hour, minute=t.minute, second=t.second)


def b_timer(ctx, frame):
    t = ctx.clock.read()
    return t.hour * 3600 + t.minute * 60 + t.second + t.microsecond / 1e6


def b_year(ctx, frame, d):
    return to_date(d).year


def b_month(ctx, frame, d):
    return to_date(d).month


def b_day(ctx, frame, d):
    return to_date(d).day


def b_hour(ctx, frame, d):
    return to_date(d).hour


def b_minute(ctx, frame, d):
    return to_date(d).minute


def b_second(ctx, frame, d):
    return to_date(d).second


def b_weekday(ctx, frame, d, firstdayofweek=1):
    wd = (to_date(d).isoweekday() % 7) + 1  # Sunday = 1
    first = to_long(firstdayofweek) or 1
    return (wd - first) % 7 + 1


def b_dateserial(ctx, frame, year, month, day):
    y, m, d_ = two_digit_year(to_long(year)), to_long(month), to_long(day)
    y += (m - 1) // 12
    m = (m - 1) % 12 + 1
    return dt.datetime(y, m, 1) + dt.timedelta(days=d_ - 1)


def b_timeserial(ctx, frame, hour, minute, second):
    return EPOCH + dt.timedelta(hours=to_long(hour), minutes=to_long(minute), seconds=to_long(second))


def b_dateadd(ctx, frame, interval, number, date):
    unit = to_str(interval).lower()
    d = to_date(date)
    k = to_long(number)
    if unit in ("d", "y", "w"):
        return d + dt.timedelta(days=k)
    if unit == "ww":
        return d + dt.timedelta(weeks=k)
    if unit == "h":
        return d + dt.timedelta(hours=k)
    if unit == "n":
        return d + dt.timedelta(minutes=k)
    if unit == "s":
        return d + dt.timedelta(seconds=k)
    if unit in ("m", "q", "yyyy"):
        months = k * {"m": 1, "q": 3, "yyyy": 12}[unit]
        total = d.month - 1 + months
        y, m = d.year + total // 12, total % 12 + 1
        day = min(d.day, calendar.monthrange(y, m)[1])
        return d.replace(year=y, month=m, day=day)
    raise VBANotImplemented(f"DateAdd interval {unit!r}")


def b_datediff(ctx, frame, interval, date1, date2, firstdayofweek=None, firstweekofyear=None):
    unit = to_str(interval).lower()
    a, b = to_date(date1), to_date(date2)
    if unit in ("d", "y"):
        return (b.date() - a.date()).days
    if unit == "yyyy":
        return b.year - a.year
    if unit == "m":
        return (b.year - a.year) * 12 + b.month - a.month
    if unit in ("h", "n", "s"):  # the boundaries crossed, as VBA counts: 10:59 to 11:01 is one hour
        size = {"h": 3600, "n": 60, "s": 1}[unit]
        return math.floor((b - EPOCH).total_seconds() / size) - math.floor((a - EPOCH).total_seconds() / size)
    raise VBANotImplemented(f"DateDiff interval {unit!r}")


def b_datevalue(ctx, frame, s):
    d = to_date(s)
    return dt.datetime(d.year, d.month, d.day)


def b_timevalue(ctx, frame, s):
    d = to_date(s)
    return EPOCH.replace(hour=d.hour, minute=d.minute, second=d.second)


def b_format(ctx, frame, expression, fmt=None, firstdayofweek=None, firstweekofyear=None):
    v = default_value(expression)
    if fmt is None:
        return "" if v is EMPTY else to_str(v)
    f = to_str(fmt)
    low = f.lower()
    date_formats = {
        "yyyy-mm-dd": "%Y-%m-%d",
        "dd/mm/yyyy": "%d/%m/%Y",
        "mm/dd/yyyy": "%m/%d/%Y",
        "yyyymmdd": "%Y%m%d",
        "hh:mm:ss": "%H:%M:%S",
        "hh:mm": "%H:%M",
        "hh:nn:ss": "%H:%M:%S",
        "yyyy-mm-dd hh:mm:ss": "%Y-%m-%d %H:%M:%S",
        "yyyy-mm-dd hh:nn:ss": "%Y-%m-%d %H:%M:%S",
    }
    if low in date_formats:
        return to_date(v).strftime(date_formats[low])
    if low in ("0", "0.0", "0.00", "0.000", "0.0000", "#,##0", "#,##0.00", "fixed", "standard"):
        x = to_double(v)
        decimals = {"fixed": 2, "standard": 2}.get(low, len(low.split(".")[1]) if "." in low else 0)
        q = float(to_decimal(x).quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP))
        text = f"{q:,.{decimals}f}" if "," in low or low == "standard" else f"{q:.{decimals}f}"
        return text
    if low in ("@", ""):
        return to_str(v)
    raise VBANotImplemented(f"Format with format string {f!r}")


# --- information ---


def b_isempty(ctx, frame, x):
    return default_value(x) is EMPTY


def b_isnull(ctx, frame, x):
    return default_value(x) is NULL


def b_isnumeric(ctx, frame, x):
    v = default_value(x)
    if v is EMPTY:
        return True
    if isinstance(v, bool) or is_numeric_value(v):
        return True
    if isinstance(v, str):
        return parse_number_text(v) is not None
    return False


def b_isdate(ctx, frame, x):
    v = default_value(x)
    if isinstance(v, dt.datetime):
        return True
    if isinstance(v, str):
        try:
            to_date(v)
            return parse_number_text(v) is None
        except VBARuntimeError:
            return False
    return False


def b_isarray(ctx, frame, x):
    return isinstance(x, VBAArray)


def b_isobject(ctx, frame, x):
    return isinstance(x, VBAObject) or x is NOTHING


def b_iserror(ctx, frame, x):
    return isinstance(default_value(x), XLError)


def b_ismissing(ctx, frame, x):
    return x is MISSING


def b_typename(ctx, frame, x):
    return type_name(x)


def b_vartype(ctx, frame, x):
    names = {
        "Empty": 0,
        "Null": 1,
        "Integer": 2,
        "Long": 3,
        "Double": 5,
        "Currency": 6,
        "Decimal": 14,
        "Date": 7,
        "String": 8,
        "Error": 10,
        "Boolean": 11,
    }
    if isinstance(x, VBAArray):
        return 8204
    if isinstance(x, VBAObject) or x is NOTHING:
        return 9
    return names.get(type_name(x), 12)


# --- arrays ---


def b_array(ctx, frame, *items):
    """Array(...): from 0, or from 1 in a module with Option Base 1."""
    return VBAArray.from_list(list(items), lo=frame._module._base if frame is not None else 0)


def b_ubound(ctx, frame, arr, dimension=1):
    if not isinstance(arr, VBAArray):
        raise VBARuntimeError(13, "Type mismatch")
    return arr.ubound(to_long(dimension))


def b_lbound(ctx, frame, arr, dimension=1):
    if not isinstance(arr, VBAArray):
        raise VBARuntimeError(13, "Type mismatch")
    return arr.lbound(to_long(dimension))


def b_iif(ctx, frame, expression, truepart, falsepart):
    return truepart if to_bool(expression) else falsepart


def b_choose(ctx, frame, index, *choices):
    i = to_long(index)
    if 1 <= i <= len(choices):
        return choices[i - 1]
    return NULL


def b_switch(ctx, frame, *pairs):
    """The value after the first condition that is true; Null when none is."""
    if len(pairs) % 2:
        raise VBARuntimeError(5, "Invalid procedure call or argument")
    for condition, value in zip(pairs[::2], pairs[1::2], strict=True):
        if default_value(condition) is not NULL and to_bool(condition):
            return value
    return NULL


def b_cverr(ctx, frame, number):
    code = to_long(number)
    return XLError(ERROR_CODES.get(code, f"#ERR{code}"))


def b_rgb(ctx, frame, red, green, blue):
    return to_long(red) + 256 * to_long(green) + 65536 * to_long(blue)


# --- interaction ---


MSGBOX_ANSWERS = {"ok": 1, "cancel": 2, "abort": 3, "retry": 4, "ignore": 5, "yes": 6, "no": 7}
# The buttons each kind of MsgBox shows: vbOKOnly, vbOKCancel, vbAbortRetryIgnore, vbYesNoCancel, vbYesNo, vbRetryCancel
MSGBOX_BUTTONS = [("ok",), ("ok", "cancel"), ("abort", "retry", "ignore"), ("yes", "no", "cancel"), ("yes", "no"),
                  ("retry", "cancel")]  # fmt: skip


def b_msgbox(ctx, frame, prompt, buttons=0, title=None, helpfile=None, context=None):
    kind = to_long(buttons) & 7
    ctx.messages.append(("MsgBox", to_str(prompt)))
    if kind == 0:
        ctx.ledger.note("MSGBOX-NO-UI")
        return 1
    question = to_str(prompt)[:60]
    if not ctx.answers:
        raise VBANotImplemented(f"MsgBox asks the user a question ({question!r}); pass an answer (yes, no, ok, ...)")
    answer = to_str(ctx.answers.pop(0)).strip().lower()
    offered = MSGBOX_BUTTONS[kind] if kind < len(MSGBOX_BUTTONS) else ()
    by_number = {str(MSGBOX_ANSWERS[name]): name for name in offered}
    answer = by_number.get(answer, answer)
    if answer not in offered:  # not a button of this box: a mistake in the answers, not an error of the macro
        raise VBANotImplemented(
            f"the answer {answer!r} to the MsgBox question {question!r}: its buttons are {', '.join(offered)}"
        )
    return MSGBOX_ANSWERS[answer]


def b_inputbox(ctx, frame, prompt, title=None, default=None, xpos=None, ypos=None, helpfile=None, context=None):
    ctx.messages.append(("InputBox", to_str(prompt)))
    if ctx.answers:
        return to_str(ctx.answers.pop(0))
    raise VBANotImplemented(f"InputBox asks the user for input ({to_str(prompt)[:60]!r}); pass an answer")


def b_environ(ctx, frame, name):
    key = to_str(name)
    if key in ctx.environ:
        return ctx.environ[key]
    raise VBANotImplemented(f"Environ({key!r}) depends on the machine; pass a value")


def b_doevents(ctx, frame):
    return 0


def b_randomize(ctx, frame, seed=None):
    return EMPTY


def b_rnd(ctx, frame, number=None):
    raise VBANotImplemented("Rnd: random numbers cannot be reproduced")


_FUNCS = {name[2:]: fn for name, fn in dict(globals()).items() if name.startswith("b_")}
_FUNCS.update({"cstr": b_cstr, "instrrev": b_instrrev})


def _wrap(fn, name):
    signature = inspect.signature(fn)
    params = list(signature.parameters.values())[2:]

    def call(ctx, frame, args, kwargs):
        # VBA names the argument Compare; the Python parameter is compare_
        kwargs = {("compare_" if k == "compare" else k): v for k, v in (kwargs or {}).items()}
        if None in args:  # an argument left out between commas, MsgBox "Done", , "Title": its default
            args = list(args)
            for i, p in enumerate(params[: len(args)]):
                if args[i] is None and p.kind is p.POSITIONAL_OR_KEYWORD:
                    if p.default is p.empty:
                        raise VBARuntimeError(449, f"Argument not optional for {name}")
                    args[i] = p.default
        try:
            signature.bind(ctx, frame, *args, **kwargs)
        except TypeError as err:
            raise VBARuntimeError(450, f"Wrong number of arguments for {name}") from err
        return fn(ctx, frame, *args, **kwargs)

    return call


BUILTINS = {name: _wrap(fn, name) for name, fn in _FUNCS.items()}
# The builtins that go by the calling module's options, so they get its frame: those with a Compare argument (left
# out, it is the module's Option Compare), and Array (Option Base).
TEXT_COMPARE = {name for name, fn in _FUNCS.items() if "compare_" in inspect.signature(fn).parameters}
FRAME_AWARE = TEXT_COMPARE | {"array"}
