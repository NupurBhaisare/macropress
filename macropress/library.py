"""Excel's worksheet functions past the core ones formula.py has: rounding and arithmetic, dates, text, information,
statistics, finance, and SUBTOTAL and AGGREGATE. Each is written to Excel's recorded answers (tests/parity, modules
TRound, TWDate, TWText, TInfo, TFin and TSub); where a probe left Excel's answer open, the function stops with
VBANotImplemented rather than guess, and the formula is left for Excel to calculate.

The functions register themselves in formula.FUNCTIONS, TAKES and the table of arguments Excel checks as a formula is
entered."""

import bisect
import datetime as dt
import math
from decimal import ROUND_DOWN, ROUND_UP, Decimal, localcontext

from . import entry, formula as fx
from .errors import VBANotImplemented
from .formula import Array, XLError, _Err, finite, num, text, truth
from .numfmt import BUILTIN_FORMATS, is_date_format, is_time_format
from .values import fifteen


def _fail(code):
    raise _Err(XLError(code))


def _whole(v):
    """A count, a number of digits or a period, as Excel takes one: the number with its fraction dropped."""
    return math.trunc(num(v))


def _arg(ev, args, i, default):
    """Argument i, or default when it is left out or left empty."""
    if i >= len(args) or args[i][0] == "missing":
        return default
    return ev.ev(args[i])


# ---------------------------------------------------------------------------------------------------------- rounding
# ROUNDUP, ROUNDDOWN, CEILING, FLOOR and MROUND work on the number to 15 digits, as ROUND does (ROUNDDOWN(4.35*100,0)
# is 435, though the product falls short of it); QUOTIENT and MOD on the number as it is (QUOTIENT(4.35*100,1) is 434).


def _round_to(x, digits, mode):
    x = fifteen(x)
    if x == 0:
        return 0.0
    with localcontext() as ctx:
        ctx.prec = 1000
        return finite(float(Decimal(repr(x)).quantize(Decimal(1).scaleb(-max(min(digits, 400), -400)), rounding=mode)))


def f_roundup(ev, args):
    return _round_to(num(ev.ev(args[0])), _whole(ev.ev(args[1])), ROUND_UP)


def f_rounddown(ev, args):
    return _round_to(num(ev.ev(args[0])), _whole(ev.ev(args[1])), ROUND_DOWN)


def f_trunc(ev, args):
    return _round_to(num(ev.ev(args[0])), _whole(_arg(ev, args, 1, 0.0)), ROUND_DOWN)


def _multiple(x, step, up):
    """x rounded up or down to a multiple of step, from the quotient to 15 digits (FLOOR(0.3,0.1) is 0.3)."""
    q = fifteen(x / step)
    return finite((math.ceil(q) if up else math.floor(q)) * step)


def f_ceiling(ev, args):
    x, step = num(ev.ev(args[0])), num(ev.ev(args[1]))
    if step == 0:
        return 0.0
    if x > 0 > step:
        _fail("#NUM!")
    return _multiple(x, step, True)  # a negative number with a positive step goes up, toward 0


def f_floor(ev, args):
    x, step = num(ev.ev(args[0])), num(ev.ev(args[1]))
    if step == 0:
        return 0.0 if x == 0 else _fail("#DIV/0!")
    if x > 0 > step:
        _fail("#NUM!")
    return _multiple(x, step, False)


def f_ceiling_math(ev, args):
    """CEILING.MATH: the sign of the step does not count; a negative number goes toward 0 unless mode is not 0."""
    x, step = num(ev.ev(args[0])), abs(num(_arg(ev, args, 1, 1.0)))
    if step == 0:
        return 0.0
    away = x < 0 and num(_arg(ev, args, 2, 0.0)) != 0
    return _multiple(x, step, not away)


def f_floor_math(ev, args):
    """FLOOR.MATH: the sign of the step does not count; a negative number goes away from 0 unless mode is not 0."""
    x, step = num(ev.ev(args[0])), abs(num(_arg(ev, args, 1, 1.0)))
    if step == 0:
        return 0.0
    toward = x < 0 and num(_arg(ev, args, 2, 0.0)) != 0
    return _multiple(x, step, toward)


def f_ceiling_precise(ev, args):
    x, step = num(ev.ev(args[0])), abs(num(_arg(ev, args, 1, 1.0)))
    return 0.0 if step == 0 else _multiple(x, step, True)


def f_floor_precise(ev, args):
    x, step = num(ev.ev(args[0])), abs(num(_arg(ev, args, 1, 1.0)))
    return 0.0 if step == 0 else _multiple(x, step, False)


def f_mround(ev, args):
    """MROUND: the nearest multiple, half away from 0; a number and a multiple of other signs are #NUM!. A quotient a
    hair below a half counts as the half when it is within the 14th decimal place of it: MROUND(1.15,0.1) is 1.2
    (11.4999999999999982 tenths), MROUND(6.05,0.1) is 6 (60.4999999999999929)."""
    x, step = _atp(ev.ev(args[0])), _atp(ev.ev(args[1]))
    if step == 0 or x == 0:
        return 0.0
    if (x < 0) != (step < 0):
        _fail("#NUM!")
    q = round(abs(x / step), 14)
    return finite(math.floor(q + 0.5) * step)


# MOD gives #NUM! once the quotient is too large for its remainder to mean much: checked at 1.1E+12 (allowed) and
# 1.2E+12 (#NUM!); between the two MacroPress stops.
_MOD_ALLOWED, _MOD_REFUSED = 1.1e12, 1.2e12


def f_mod(ev, args):
    n, d = num(ev.ev(args[0])), num(ev.ev(args[1]))
    if d == 0:
        _fail("#DIV/0!")
    q = n / d
    if abs(q) >= _MOD_REFUSED:
        _fail("#NUM!")
    if abs(q) > _MOD_ALLOWED:
        raise VBANotImplemented(f"MOD({n!r},{d!r}): whether Excel finds a quotient of {q:.4g} too large is not known")
    r = math.fmod(n, d)  # the remainder itself, not n - d*INT(n/d): MOD(1E+300,1E+290) is 9.99999907769267E+289
    return r + d if r and (r < 0) != (d < 0) else r


def f_quotient(ev, args):
    n, d = _atp(ev.ev(args[0])), _atp(ev.ev(args[1]))
    if d == 0:
        _fail("#DIV/0!")
    return float(math.trunc(n / d))


def f_sign(ev, args):
    x = num(ev.ev(args[0]))
    return float((x > 0) - (x < 0))


def f_power(ev, args):
    return fx.power(num(ev.ev(args[0])), num(ev.ev(args[1])))


def f_sqrt(ev, args):
    x = num(ev.ev(args[0]))
    return math.sqrt(x) if x >= 0 else _fail("#NUM!")


def f_exp(ev, args):
    try:
        return finite(math.exp(num(ev.ev(args[0]))))
    except OverflowError:
        return _fail("#NUM!")


def f_ln(ev, args):
    x = num(ev.ev(args[0]))
    return math.log(x) if x > 0 else _fail("#NUM!")


def f_log(ev, args):
    x, base = num(ev.ev(args[0])), num(_arg(ev, args, 1, 10.0))
    if x <= 0 or base <= 0:
        _fail("#NUM!")
    if base == 1:
        _fail("#DIV/0!")
    return math.log10(x) if base == 10 else math.log(x) / math.log(base)


def f_log10(ev, args):
    x = num(ev.ev(args[0]))
    return math.log10(x) if x > 0 else _fail("#NUM!")


def f_pi(ev, args):
    return math.pi


def f_even(ev, args):
    x = fifteen(num(ev.ev(args[0])))
    return math.copysign(math.ceil(abs(x) / 2) * 2, x) if x else 0.0


def f_odd(ev, args):
    x = fifteen(num(ev.ev(args[0])))
    return math.copysign(math.ceil((abs(x) - 1) / 2) * 2 + 1, x if x else 1.0)


def f_product(ev, args):
    numbers = fx._numbers(ev, args, "PRODUCT")
    if not numbers:
        return 0.0
    out = 1.0
    for x, n in numbers:
        out *= x**n
    return finite(out)


# ------------------------------------------------------------------------------------------------------------- dates
# A date is its serial number, days since the 0th of January 1900. Excel's 1900 has a 29th of February, serial 60, so
# the serials before it are one day off the calendar's: DATE, YEAR, MONTH, DAY and WEEKDAY keep that day, as Excel
# does; the functions that count months (EOMONTH, EDATE) go by the calendar, in which February 1900 has 28 days.

_EPOCH = dt.date(1899, 12, 30)
_LAST = 2958465  # 31 December 9999


def _calendar(serial):
    """(year, month, day) of a whole serial number as DATE and DAY have it: serial 0 is the 0th of January 1900, 60
    the 29th of February."""
    if serial == 0:
        return 1900, 1, 0
    if serial == 60:
        return 1900, 2, 29
    d = _EPOCH + dt.timedelta(days=serial + (1 if serial < 60 else 0))
    return d.year, d.month, d.day


def _first_of(year, month):
    """The serial number of the first of a month (negative before 1900)."""
    s = (dt.date(year, month, 1) - _EPOCH).days
    return s - 1 if (year, month) < (1900, 3) else s


def _serial(d):
    """The serial number of a calendar date; #NUM! before 1900 or after 9999."""
    s = (d - _EPOCH).days
    if s < 61:
        s -= 1
    if not 1 <= s <= _LAST:
        _fail("#NUM!")
    return float(s)


def _date_of(serial):
    """The calendar date of a whole serial number, for the functions that count days and months. The 0th of January
    and the 29th of February 1900 are not dates there: what Excel makes of them is not known."""
    if serial in (0, 60):
        raise VBANotImplemented(f"the serial number {serial}, which is not a calendar date, in a date function")
    return _EPOCH + dt.timedelta(days=serial + (1 if serial < 60 else 0))


def _seconds(x):
    """A date and time as whole seconds, to the nearest: DAY(59.99999999) is 29, as the 0.86 of a second left rounds up
    to the next day."""
    return math.floor(x * 86400 + 0.5)


def _day(v):
    """The day a date falls on, as YEAR, MONTH, DAY and WEEKDAY take it: to the nearest second."""
    x = num(v)
    if x < 0:
        _fail("#NUM!")
    serial = _seconds(x) // 86400
    if serial > _LAST:
        _fail("#NUM!")
    return serial


def _atp(v):
    """A number for the functions that came from the Analysis ToolPak (EOMONTH, MROUND, ISEVEN and the rest), which
    take text that reads as a number but not TRUE or FALSE."""
    if isinstance(v, bool):
        _fail("#VALUE!")
    return num(v)


def _date_arg(v, atp=True):
    """A date given to a function that counts days or months: its fraction dropped, #NUM! if negative. Within half
    a second of midnight, where DAY rounds up to the next day, what these functions do is not known."""
    x = _atp(v) if atp else num(v)
    if x < 0 or x > _LAST + 1:
        _fail("#NUM!")
    serial = math.floor(x)
    if x - serial >= 1 - 0.5 / 86400:
        raise VBANotImplemented(f"the date {x!r}, within half a second of the next day")
    if serial > _LAST:
        _fail("#NUM!")
    return serial


def f_date(ev, args):
    y, m, d = (math.floor(num(ev.ev(a))) for a in args)
    if not 0 <= y <= 9999:
        _fail("#NUM!")
    if y < 1900:
        y += 1900
    y, m = y + (m - 1) // 12, (m - 1) % 12 + 1
    if not 1 <= y <= 9999:
        _fail("#NUM!")
    serial = _first_of(y, m) + d - 1
    if not 0 <= serial <= _LAST:
        _fail("#NUM!")
    return float(serial)


def f_year(ev, args):
    return float(_calendar(_day(ev.ev(args[0])))[0])


def f_month(ev, args):
    return float(_calendar(_day(ev.ev(args[0])))[1])


def f_day(ev, args):
    return float(_calendar(_day(ev.ev(args[0])))[2])


# WEEKDAY's return types: the day its week starts on (0 Sunday to 6 Saturday), and the number that day gets.
_WEEK_STARTS = {1: (0, 1), 2: (1, 1), 3: (1, 0), 11: (1, 1), 12: (2, 1), 13: (3, 1), 14: (4, 1), 15: (5, 1), 16: (6, 1),
                17: (0, 1)}  # fmt: skip


def f_weekday(ev, args):
    serial = _day(ev.ev(args[0]))
    kind = _whole(_arg(ev, args, 1, 1.0))
    if kind not in _WEEK_STARTS:
        _fail("#NUM!")
    start, first = _WEEK_STARTS[kind]
    return float((serial + 6 - start) % 7 + first)  # serial 1 is a Sunday


def _months_on(serial, months):
    """(year, month) months after the month of a serial number, as EOMONTH and EDATE count."""
    if serial == 0:
        year, month = 1900, 1
    else:
        d = _date_of(serial)
        year, month = d.year, d.month
    total = year * 12 + month - 1 + months
    return total // 12, total % 12 + 1


def _month_days(year, month):
    return ((dt.date(year + month // 12, month % 12 + 1, 1)) - dt.date(year, month, 1)).days


def f_eomonth(ev, args):
    serial, months = _date_arg(ev.ev(args[0])), math.trunc(_atp(ev.ev(args[1])))
    year, month = _months_on(serial, months)
    if not 1900 <= year <= 9999:
        _fail("#NUM!")
    return _serial(dt.date(year, month, _month_days(year, month)))


def f_edate(ev, args):
    serial, months = _date_arg(ev.ev(args[0])), math.trunc(_atp(ev.ev(args[1])))
    if serial == 0:
        raise VBANotImplemented("EDATE of the 0th of January 1900")
    day = _date_of(serial).day
    year, month = _months_on(serial, months)
    if not 1900 <= year <= 9999:
        _fail("#NUM!")
    return _serial(dt.date(year, month, min(day, _month_days(year, month))))


def _days_on(year, month, day):
    """The date of a day that may run past the end of its month (the 31st of February is the 3rd of March)."""
    return dt.date(year, month, 1) + dt.timedelta(days=day - 1)


def f_datedif(ev, args):
    """DATEDIF, the units as Excel counts them, "md" included: past the end of the month before the end date it goes
    negative (31 January to 1 March 2015 is -2)."""
    a, b = _date_arg(ev.ev(args[0]), atp=False), _date_arg(ev.ev(args[1]), atp=False)
    unit = text(ev.ev(args[2])).upper()
    if a > b or unit not in ("Y", "M", "D", "MD", "YM", "YD"):
        _fail("#NUM!")
    if unit == "D":
        return float(b - a)
    d1, d2 = _date_of(a), _date_of(b)
    months = (d2.year - d1.year) * 12 + d2.month - d1.month - (d2.day < d1.day)
    if unit == "Y":
        return float(d2.year - d1.year - ((d2.month, d2.day) < (d1.month, d1.day)))
    if unit == "M":
        return float(months)
    if unit == "YM":
        return float(months % 12)
    if unit == "MD":
        if d1.day <= d2.day:
            return float(d2.day - d1.day)
        year, month = (d2.year, d2.month - 1) if d2.month > 1 else (d2.year - 1, 12)
        return float((d2 - _days_on(year, month, d1.day)).days)
    if (d2.month, d2.day) >= (d1.month, d1.day):  # "yd": the start moved into the end's year, or the end into the
        return float((d2 - _days_on(d2.year, d1.month, d1.day)).days)  # year after the start's
    return float((_days_on(d1.year + 1, d2.month, d2.day) - d1).days)


def f_days(ev, args):
    end, start = num(ev.ev(args[0])), num(ev.ev(args[1]))
    if end < 0 or start < 0:
        _fail("#NUM!")
    whole = math.floor(end) - math.floor(start)
    if whole != math.trunc(end - start):  # DAYS(45296.1,45290.9): the days between, or the dates' difference
        raise VBANotImplemented("DAYS of two dates with times of day that make the days between them unclear")
    return float(whole)


def _iso_week(serial):
    """The ISO week of a serial number: the week of its Thursday, the days counted as Excel counts them (serial 1 a
    Sunday, so ISOWEEKNUM(1) is 52)."""
    thursday = serial - (serial + 5) % 7 + 3
    d = _EPOCH + dt.timedelta(days=thursday + (1 if thursday < 60 else 0))
    return float((d.timetuple().tm_yday - 1) // 7 + 1)


def f_isoweeknum(ev, args):
    return _iso_week(_day(ev.ev(args[0])))


# WEEKNUM's return types: the day its weeks start on (0 Sunday); 21 is the ISO week.
_WEEKNUM_STARTS = {1: 0, 2: 1, 11: 1, 12: 2, 13: 3, 14: 4, 15: 5, 16: 6, 17: 0}


def f_weeknum(ev, args):
    serial = _date_arg(ev.ev(args[0]))
    kind = math.trunc(_atp(_arg(ev, args, 1, 1.0)))
    if kind == 21:
        return _iso_week(serial)
    if kind not in _WEEKNUM_STARTS:
        _fail("#NUM!")
    year = _calendar(serial)[0]
    jan1 = _first_of(year, 1)
    offset = ((jan1 + 6) % 7 - _WEEKNUM_STARTS[kind]) % 7
    return float((serial - jan1 + offset) // 7 + 1)


# --- working days: the weekends NETWORKDAYS.INTL and WORKDAY.INTL take by number, their days 0 Monday to 6 Sunday
_WEEKENDS = {1: (5, 6), 2: (6, 0), 3: (0, 1), 4: (1, 2), 5: (2, 3), 6: (3, 4), 7: (4, 5), 11: (6,), 12: (0,), 13: (1,),
             14: (2,), 15: (3,), 16: (4,), 17: (5,)}  # fmt: skip


def _weekend(v):
    """The days off, 0 Monday to 6 Sunday: a weekend's number or seven 0s and 1s from Monday."""
    if v is None:
        return frozenset((5, 6))
    if isinstance(v, str):
        if len(v) != 7 or set(v) - {"0", "1"} or v == "1111111":
            _fail("#VALUE!")
        return frozenset(i for i, ch in enumerate(v) if ch == "1")
    kind = _whole(v)
    if kind not in _WEEKENDS:
        _fail("#NUM!")
    return frozenset(_WEEKENDS[kind])


def _holidays(ev, node):
    """The holidays given, each its whole serial number; blank cells are none."""
    if node is None or node[0] == "missing":
        return []
    v = ev.ev(node)
    out = set()
    for x in v.values() if isinstance(v, Array) else [v]:
        if x is None:
            continue
        if isinstance(x, XLError):
            raise _Err(x)
        if isinstance(x, (str, bool)):
            raise VBANotImplemented(f"a holiday given as {x!r}")
        if x < 0:
            _fail("#NUM!")
        out.add(math.floor(x))
    return sorted(out)


def _workday(serial, off):
    return (serial + 5) % 7 not in off  # serial 1 is a Sunday: (1 + 5) % 7 == 6


def _count_workdays(a, b, off, holidays):
    """The working days from a to b, both counted, a <= b."""
    full, rest = divmod(b - a + 1, 7)
    count = full * (7 - len(off)) + sum(1 for k in range(rest) if _workday(a + full * 7 + k, off))
    lo, hi = bisect.bisect_left(holidays, a), bisect.bisect_right(holidays, b)
    return count - sum(1 for h in holidays[lo:hi] if _workday(h, off))


def f_networkdays(ev, args):
    a, b = _date_arg(ev.ev(args[0])), _date_arg(ev.ev(args[1]))
    off, holidays = _weekend(None), _holidays(ev, args[2] if len(args) > 2 else None)
    return float(_count_workdays(a, b, off, holidays) if a <= b else -_count_workdays(b, a, off, holidays))


def f_networkdays_intl(ev, args):
    a, b = _date_arg(ev.ev(args[0])), _date_arg(ev.ev(args[1]))
    off = _weekend(_arg(ev, args, 2, None))
    holidays = _holidays(ev, args[3] if len(args) > 3 else None)
    return float(_count_workdays(a, b, off, holidays) if a <= b else -_count_workdays(b, a, off, holidays))


def _step_workdays(start, days, off, holidays):
    """The day `days` working days after start (before it, for a negative count); start itself on 0 of them."""
    step, left, d = (1 if days > 0 else -1), abs(days), start
    if left > _LAST:
        _fail("#NUM!")
    per_week = 7 - len(off)
    while left:
        nxt = d + 7 * step
        lo, hi = (d + 1, nxt) if step > 0 else (nxt, d - 1)
        next_holiday = holidays[bisect.bisect_left(holidays, lo) :][:1]
        if left > per_week and not any(lo <= h <= hi for h in next_holiday):
            d, left = nxt, left - per_week
        else:
            d += step
            if _workday(d, off) and not _is_holiday(d, holidays):
                left -= 1
        if not 0 <= d <= _LAST:
            _fail("#NUM!")
    return float(d)


def _is_holiday(d, holidays):
    i = bisect.bisect_left(holidays, d)
    return i < len(holidays) and holidays[i] == d


def f_workday(ev, args):
    start, days = _date_arg(ev.ev(args[0])), math.trunc(_atp(ev.ev(args[1])))
    return _step_workdays(start, days, _weekend(None), _holidays(ev, args[2] if len(args) > 2 else None))


def f_workday_intl(ev, args):
    start, days = _date_arg(ev.ev(args[0])), math.trunc(_atp(ev.ev(args[1])))
    off = _weekend(_arg(ev, args, 2, None))
    return _step_workdays(start, days, off, _holidays(ev, args[3] if len(args) > 3 else None))


# --- counting in years
def _last_of_month(d):
    return d.day == _month_days(d.year, d.month)


def _is_leap(year):
    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)


def _feb29_between(d1, d2):
    """Whether actual/actual counts a 366-day year for dates less than a year apart (as Excel decides it)."""
    if _is_leap(d1.year) and (d1.year == d2.year or d1.month <= 2):
        return True
    if _is_leap(d2.year):
        return d2.month > 2 or (d2.month == 2 and d2.day == 29)
    return False


def _thirty_360(d1, d2, day1, day2):
    return ((d2.year - d1.year) * 360 + (d2.month - d1.month) * 30 + day2 - day1) / 360


def f_yearfrac(ev, args):
    """YEARFRAC on Excel's five bases: 0 US 30/360, 1 actual/actual, 2 actual/360, 3 actual/365, 4 European 30/360."""
    a, b = _date_arg(ev.ev(args[0])), _date_arg(ev.ev(args[1]))
    basis = math.trunc(_atp(_arg(ev, args, 2, 0.0)))
    if basis not in range(5):
        _fail("#NUM!")
    a, b = min(a, b), max(a, b)
    if basis == 2:
        return (b - a) / 360
    if basis == 3:
        return (b - a) / 365
    d1, d2 = _date_of(a), _date_of(b)
    if basis == 4:
        return _thirty_360(d1, d2, min(d1.day, 30), min(d2.day, 30))
    if basis == 0:
        day1, day2 = d1.day, d2.day
        if day1 == 31 and day2 == 31:
            day1 = day2 = 30
        elif day1 == 31:
            day1 = 30
        elif day1 == 30 and day2 == 31:
            day2 = 30
        elif d1.month == 2 and _last_of_month(d1):
            day1 = 30
            if d2.month == 2 and _last_of_month(d2):
                day2 = 30
        return _thirty_360(d1, d2, day1, day2)
    over_a_year = d1.year != d2.year and (d1.year + 1 != d2.year or (d1.month, d1.day) < (d2.month, d2.day))
    if over_a_year:
        years = range(d1.year, d2.year + 1)
        return (b - a) / (sum(366 if _is_leap(y) else 365 for y in years) / len(years))
    return (b - a) / (366 if _feb29_between(d1, d2) else 365)


def f_days360(ev, args):
    """DAYS360: by the US rules unless method is TRUE (the European ones)."""
    a, b = _date_arg(ev.ev(args[0]), atp=False), _date_arg(ev.ev(args[1]), atp=False)
    european = truth(_arg(ev, args, 2, False))
    d1, d2 = _date_of(a), _date_of(b)
    day1, day2 = min(d1.day, 30), min(d2.day, 30)
    year2, month2 = d2.year, d2.month
    if not european:
        if _last_of_month(d1):
            day1 = 30
        if d2.day == 31:
            if day1 < 30:
                year2, month2, day2 = (d2.year, d2.month + 1, 1) if d2.month < 12 else (d2.year + 1, 1, 1)
            else:
                day2 = 30
    return float((year2 - d1.year) * 360 + (month2 - d1.month) * 30 + day2 - day1)


# --- times
def f_time(ev, args):
    parts = [num(ev.ev(a)) for a in args]
    if any(p != math.floor(p) for p in parts):
        raise VBANotImplemented("TIME of hours, minutes or seconds with a fraction")
    h, m, s = (int(p) for p in parts)
    total = h * 3600 + m * 60 + s
    if h > 32767 or total < 0:
        _fail("#NUM!")
    return (total % 86400) / 86400


def _clock(v):
    x = num(v)
    if x < 0:
        _fail("#NUM!")
    return _seconds(x) % 86400


def f_hour(ev, args):
    return float(_clock(ev.ev(args[0])) // 3600)


def f_minute(ev, args):
    return float(_clock(ev.ev(args[0])) // 60 % 60)


def f_second(ev, args):
    return float(_clock(ev.ev(args[0])) % 60)


# -------------------------------------------------------------------------------------------------------------- text
_TEXT_LIMIT = 32767  # characters in a cell


def _start(ev, args, i, within):
    """FIND's and SEARCH's start: 1 to one past the text's end, else #VALUE!."""
    start = _whole(_arg(ev, args, i, 1.0))
    if not 1 <= start <= len(within) + 1:
        _fail("#VALUE!")
    return start - 1


def f_find(ev, args):
    find, within = text(ev.ev(args[0])), text(ev.ev(args[1]))
    pos = within.find(find, _start(ev, args, 2, within))
    return float(pos + 1) if pos >= 0 else _fail("#VALUE!")


def f_search(ev, args):
    from .excel.filters import wildcard_regex  # noqa: PLC0415 - the excel package imports this module

    find, within = text(ev.ev(args[0])), text(ev.ev(args[1]))
    m = wildcard_regex(find).search(within, _start(ev, args, 2, within))
    return float(m.start() + 1) if m else _fail("#VALUE!")


def f_substitute(ev, args):
    """SUBSTITUTE: every one, or the nth, where the search for the nth goes on from the character after each found
    (SUBSTITUTE("aaa","aa","b",2) is "ab")."""
    s, old, new = text(ev.ev(args[0])), text(ev.ev(args[1])), text(ev.ev(args[2]))
    if len(args) < 4 or args[3][0] == "missing":
        out = s.replace(old, new) if old else s
    else:
        n = _whole(ev.ev(args[3]))
        if n < 1:
            _fail("#VALUE!")
        pos = -1
        for _ in range(n):
            pos = s.find(old, pos + 1) if old else -1
            if pos < 0:
                return s
        out = s[:pos] + new + s[pos + len(old) :]
    return out if len(out) <= _TEXT_LIMIT else _fail("#VALUE!")


def f_replace(ev, args):
    s = text(ev.ev(args[0]))
    start, n = _whole(ev.ev(args[1])), _whole(ev.ev(args[2]))
    if start < 1 or n < 0:
        _fail("#VALUE!")
    out = s[: start - 1] + text(ev.ev(args[3])) + s[start - 1 + n :]
    return out if len(out) <= _TEXT_LIMIT else _fail("#VALUE!")


def f_rept(ev, args):
    s, n = text(ev.ev(args[0])), _whole(ev.ev(args[1]))
    if n < 0 or len(s) * n > _TEXT_LIMIT:
        _fail("#VALUE!")
    return s * n


def f_exact(ev, args):
    return text(ev.ev(args[0])) == text(ev.ev(args[1]))


def f_proper(ev, args):
    """PROPER: a letter after anything but a letter in capitals, the rest small (it's is It'S, 2nd is 2Nd)."""
    out, after_letter = [], False
    for ch in text(ev.ev(args[0])):
        if ch.isalpha():
            changed = ch.lower() if after_letter else ch.upper()
            out.append(changed if len(changed) == 1 else ch)
            after_letter = True
        else:
            out.append(ch)
            after_letter = False
    return "".join(out)


def _flat_values(v):
    """A value, or an array's values row by row, blanks among them."""
    if isinstance(v, Array):
        return [x for row in v.full() for x in row]
    return [v]


def f_textjoin(ev, args):
    delimiters = [text(d) if d is not None else "" for d in _flat_values(ev.ev(args[0]))]
    skip_empty = truth(ev.operand(args[1]))
    parts = []
    for a in args[2:]:
        if a[0] == "missing":
            values = [None]
        else:
            v = ev.ev(a)
            values = _flat_values(v)
        for x in values:
            if isinstance(x, XLError):
                raise _Err(x)
            s = text(x)
            if s == "" and skip_empty:
                continue
            parts.append(s)
    out = [parts[0]] if parts else []
    for i, s in enumerate(parts[1:]):
        out += [delimiters[i % len(delimiters)], s]
    joined = "".join(out)
    return joined if len(joined) <= _TEXT_LIMIT else _fail("#VALUE!")


# CHAR and CODE: Windows reads codes 128 to 255 as Windows-1252; Excel for Mac as Mac Roman. MacroPress follows Windows.
def f_char(ev, args):
    n = _whole(ev.ev(args[0]))
    if not 1 <= n <= 255:
        _fail("#VALUE!")
    try:
        return bytes([n]).decode("cp1252")
    except UnicodeDecodeError:
        return chr(n)  # the five codes Windows-1252 leaves out


def f_code(ev, args):
    s = text(ev.ev(args[0]))
    if not s:
        _fail("#VALUE!")
    try:
        return float(s[0].encode("cp1252")[0])
    except UnicodeEncodeError:
        raise VBANotImplemented(f"CODE of {s[0]!r}, which Windows-1252 does not have") from None


def f_unichar(ev, args):
    n = _whole(ev.ev(args[0]))
    if not 1 <= n <= 0x10FFFF or 0xD800 <= n <= 0xDFFF:
        _fail("#VALUE!")
    if n > 0xFFFF:
        raise VBANotImplemented("UNICHAR past the first 65536 characters, which Excel keeps as two")
    return chr(n)


def f_unicode(ev, args):
    s = text(ev.ev(args[0]))
    return float(ord(s[0])) if s else _fail("#VALUE!")


def f_clean(ev, args):
    """CLEAN: the characters 0 to 31 taken out; 127 and the rest stay."""
    return "".join(ch for ch in text(ev.ev(args[0])) if ord(ch) >= 32)


def _first(v):
    """N's and T's argument: a range gives its first cell, wherever the formula is."""
    return v.cell(0, 0) if isinstance(v, Array) else v


def f_n(ev, args):
    v = _first(ev.ev(args[0]))
    if isinstance(v, XLError):
        raise _Err(v)
    if isinstance(v, bool):
        return float(v)
    return float(v) if isinstance(v, (int, float)) else 0.0


def f_t(ev, args):
    v = _first(ev.ev(args[0]))
    if isinstance(v, XLError):
        raise _Err(v)
    return v if isinstance(v, str) else ""


def f_value(ev, args):
    v = ev.ev(args[0])
    if isinstance(v, XLError):
        raise _Err(v)
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return float(v)
    if not isinstance(v, str) or v == "":
        raise VBANotImplemented(f"VALUE of {v!r}")
    number = entry.value_of(v)
    return number if number is not None else _fail("#VALUE!")


# ----------------------------------------------------------------------------------------------- logic, information
def f_ifs(ev, args):
    """IFS: the value after the first condition that holds (a reference stays the cells it names); none is #N/A."""
    for i in range(0, len(args) - 1, 2):
        if truth(ev.ev(args[i])):
            return None if args[i + 1][0] == "missing" else ev.keep(args[i + 1])
    return _fail("#N/A")


def f_switch(ev, args):
    """SWITCH: the result after the first value equal to the expression (as = compares), else the default, else #N/A."""
    v = ev.ev(args[0])
    if isinstance(v, XLError):
        raise _Err(v)
    pairs = len(args) - 1
    for i in range(1, pairs - pairs % 2, 2):
        if fx.compare(v, ev.ev(args[i])) == 0:
            return ev.keep(args[i + 1])
    return ev.keep(args[-1]) if pairs % 2 else _fail("#N/A")


def f_ifna(ev, args):
    v = ev.ev(args[0])
    return ev.ev(args[1]) if v == XLError("#N/A") else v


def _value(ev, args):
    return ev.ev(args[0])


def f_iserror(ev, args):
    return isinstance(_value(ev, args), XLError)


def f_iserr(ev, args):
    v = _value(ev, args)
    return isinstance(v, XLError) and v != XLError("#N/A")


def f_isna(ev, args):
    return _value(ev, args) == XLError("#N/A")


def f_islogical(ev, args):
    return isinstance(_value(ev, args), bool)


def f_isnontext(ev, args):
    v = _value(ev, args)
    return not isinstance(v, str) or isinstance(v, XLError)


def f_iseven(ev, args):
    return math.trunc(_atp(_value(ev, args))) % 2 == 0


def f_isodd(ev, args):
    return math.trunc(_atp(_value(ev, args))) % 2 == 1


def f_isref(ev, args):
    node = args[0]
    while node[0] == "paren":
        node = node[1]
    try:
        if node[0] in ("ref", "name", "range"):
            ev.rect_of(node)
            return True
        if node[0] == "func" and node[1] in fx._GIVES_CELLS:
            v = ev.keep(node)
            return isinstance(v, Array) and v.rect is not None
    except _Err:
        return False
    return False


def f_na(ev, args):
    return XLError("#N/A")


_ERROR_TYPES = {"#NULL!": 1, "#DIV/0!": 2, "#VALUE!": 3, "#REF!": 4, "#NAME?": 5, "#NUM!": 6, "#N/A": 7, "#SPILL!": 9,
                "#CALC!": 14}  # fmt: skip


def f_error_type(ev, args):
    v = _value(ev, args)
    if isinstance(v, XLError) and str(v) in _ERROR_TYPES:
        return float(_ERROR_TYPES[str(v)])
    return _fail("#N/A")


def f_type(ev, args):
    """TYPE: 1 a number (or blank), 2 text, 4 TRUE or FALSE, 16 an error, 64 an array. A range in a formula that is
    not a dynamic array gives one cell, as for a function that wants one value."""
    try:
        v = ev.operand(args[0])
    except _Err:
        return 16.0
    if isinstance(v, Array):
        if v.extent != (1, 1):
            return 64.0
        v = v.cell(0, 0)
    if isinstance(v, XLError):
        return 16.0
    if isinstance(v, bool):
        return 4.0
    return 2.0 if isinstance(v, str) else 1.0


# -------------------------------------------------------------------------------------------------------- statistics
def _expanded(numbers):
    """[(x, n)] written out, x n times."""
    total = sum(n for _, n in numbers)
    if total > Array.LIMIT:
        raise VBANotImplemented(f"{total} numbers, more than this sorts one by one")
    return [x for x, n in numbers for _ in range(n)]


def _numbers_of(ev, args, name, whole=False):
    """The numbers among the arguments as SUM reads them; with whole, an expression is worked out as an array whatever
    the formula (MODE and IRR take theirs so)."""
    if whole:
        args = [a if a[0] in ("ref", "name", "range", "missing") else ("const", ev.whole(a)) for a in args]
    return _expanded(fx._numbers(ev, args, name))


def _mean(xs):
    return fx.total(xs) / len(xs)


def _variance(xs, sample):
    """The variance as Excel's digits show it: from the sum and the sum of squares (VAR.S({0.1,0.2,0.3}) is
    0.00999999999999999), unless those cancel to nothing, where it is from the mean (VAR({100000001,100000002,
    100000003}) is 1). Where the two part by more than a last digit, which Excel gives is not known."""
    n = len(xs)
    if n < (2 if sample else 1):
        _fail("#DIV/0!")
    total, squares = 0.0, 0.0
    for x in xs:
        total += x
        squares += x * x
    by_sums = squares - total * total / n
    mean, by_mean = total / n, 0.0
    for x in xs:
        by_mean += (x - mean) ** 2
    if by_sums > 0 and abs(by_sums - by_mean) > 1e-12 * by_mean:
        raise VBANotImplemented("a variance whose sum of squares all but cancels: Excel's digits are not known")
    return (by_sums if by_sums > 0 else by_mean) / (n - 1 if sample else n)


def _median(xs):
    if not xs:
        _fail("#NUM!")
    xs = sorted(xs)
    half = len(xs) // 2
    return xs[half] if len(xs) % 2 else (xs[half - 1] + xs[half]) / 2


def _mode(xs):
    """The value most often there, the first to reach that count if several do; #N/A where none is there twice."""
    counts = {}
    for x in xs:
        counts[x] = counts.get(x, 0) + 1
    best = max(counts.values(), default=0)
    if best < 2:
        _fail("#N/A")
    return next(x for x in xs if counts[x] == best)


def _kth(xs, k, largest):
    """The kth largest or smallest; a k with a fraction is rounded up (LARGE(...,1.5) is the 2nd)."""
    k = math.ceil(k)
    if not 1 <= k <= len(xs):
        _fail("#NUM!")
    return sorted(xs, reverse=largest)[k - 1]


def _percentile(xs, p, exclusive):
    xs = sorted(xs)
    n = len(xs)
    if not xs or not 0 <= p <= 1 or (exclusive and not 0 < p < 1):
        _fail("#NUM!")
    if exclusive:
        pos = p * (n + 1) - 1
        if pos < 0 or pos > n - 1:
            _fail("#NUM!")
    else:
        pos = p * (n - 1)
    lo = math.floor(pos)
    return xs[lo] if lo + 1 >= n else xs[lo] + (pos - lo) * (xs[lo + 1] - xs[lo])


def _quartile(xs, q, exclusive):
    q = math.trunc(q)
    if not (0 < q < 4 if exclusive else 0 <= q <= 4):
        _fail("#NUM!")
    return _percentile(xs, q / 4, exclusive)


def f_median(ev, args):
    return _median(_numbers_of(ev, args, "MEDIAN"))


def f_mode(ev, args):
    return _mode(_numbers_of(ev, args, "MODE", whole=True))


def f_stdev_s(ev, args):
    return math.sqrt(_variance(_numbers_of(ev, args, "STDEV"), True))


def f_stdev_p(ev, args):
    return math.sqrt(_variance(_numbers_of(ev, args, "STDEVP"), False))


def f_var_s(ev, args):
    return _variance(_numbers_of(ev, args, "VAR"), True)


def f_var_p(ev, args):
    return _variance(_numbers_of(ev, args, "VARP"), False)


def f_large(ev, args):
    return _kth(_numbers_of(ev, args[:1], "LARGE"), num(ev.ev(args[1])), True)


def f_small(ev, args):
    return _kth(_numbers_of(ev, args[:1], "SMALL"), num(ev.ev(args[1])), False)


def f_percentile_inc(ev, args):
    return _percentile(_numbers_of(ev, args[:1], "PERCENTILE"), num(ev.ev(args[1])), False)


def f_percentile_exc(ev, args):
    return _percentile(_numbers_of(ev, args[:1], "PERCENTILE.EXC"), num(ev.ev(args[1])), True)


def f_quartile_inc(ev, args):
    return _quartile(_numbers_of(ev, args[:1], "QUARTILE"), num(ev.ev(args[1])), False)


def f_quartile_exc(ev, args):
    return _quartile(_numbers_of(ev, args[:1], "QUARTILE.EXC"), num(ev.ev(args[1])), True)


def _rank(ev, args):
    x = num(ev.ev(args[0]))
    cells = list(ev.keep(args[1]).values())
    if any(isinstance(v, XLError) for v in cells):
        raise VBANotImplemented("RANK over a range holding an error")
    numbers = [v for v in cells if isinstance(v, (int, float)) and not isinstance(v, bool)]
    ascending = num(_arg(ev, args, 2, 0.0)) != 0
    same = sum(1 for v in numbers if v == x)
    if not same:
        _fail("#N/A")
    return 1 + sum(1 for v in numbers if (v < x if ascending else v > x)), same


def f_rank(ev, args):
    return float(_rank(ev, args)[0])


def f_rank_avg(ev, args):
    first, same = _rank(ev, args)
    return first + (same - 1) / 2


def f_countblank(ev, args):
    v = ev.keep(args[0])
    rows, cols = v.extent
    held = sum(n for x, n in v.counted() if x is not None and x != "")
    return float(rows * cols - held)


# ----------------------------------------------------------------------------------------------------------- finance
# The closed forms of the time value of money, with a rate of 0 apart; type 1 is payment at the start of each period.


def _power(base, n):
    """base ** n as Excel works out the time value of money: a whole power by squaring and multiplying (the last digits
    of PV(0.05/12,360,-1073.64) are Excel's only so), one with a fraction as pow does."""
    try:
        if n != math.floor(n):
            return finite(base**n)
        result, square, k = 1.0, base, abs(int(n))
        while k:
            if k & 1:
                result *= square
            square *= square
            k >>= 1
        return finite(1 / result if n < 0 else result)
    except (OverflowError, ZeroDivisionError):
        return _fail("#NUM!")
    except TypeError:  # a negative number to a fractional power is complex
        return _fail("#NUM!")


def _growth(rate, nper):
    return _power(1 + rate, nper)


def _pay_type(ev, args, i):
    return 1.0 if num(_arg(ev, args, i, 0.0)) != 0 else 0.0


def _pmt(rate, nper, pv, fv, kind):
    if rate == 0:
        return -(pv + fv) / nper if nper else _fail("#NUM!")
    if nper == 0:
        _fail("#NUM!")
    q = _growth(rate, nper)
    if q == 1:  # a rate too small to show in (1 + rate) ** nper: as for no rate (PMT(1E-300,10,1000) is -100)
        return -(pv + fv) / nper
    if 1 + rate * kind == 0:
        _fail("#NUM!")
    return -(fv + pv * q) * rate / ((1 + rate * kind) * (q - 1))


def _fv(rate, nper, pmt, pv, kind):
    if rate == 0:
        return -(pv + pmt * nper)
    q = _growth(rate, nper)
    return -(pv * q + pmt * (1 + rate * kind) * (q - 1) / rate)


def f_pmt(ev, args):
    rate, nper, pv = (num(ev.ev(a)) for a in args[:3])
    if rate <= -1:
        _fail("#NUM!")
    return finite(_pmt(rate, nper, pv, num(_arg(ev, args, 3, 0.0)), _pay_type(ev, args, 4)))


def _ipmt(rate, per, nper, pv, fv, kind):
    if not 1 <= per <= nper:
        _fail("#NUM!")
    pmt = _pmt(rate, nper, pv, fv, kind)
    if per == 1:
        interest = 0.0 if kind else -pv
    elif kind:
        interest = _fv(rate, per - 2, pmt, pv, 1) - pmt
    else:
        interest = _fv(rate, per - 1, pmt, pv, 0)
    return interest * rate, pmt


def f_ipmt(ev, args):
    rate, per, nper, pv = (num(ev.ev(a)) for a in args[:4])
    return finite(_ipmt(rate, per, nper, pv, num(_arg(ev, args, 4, 0.0)), _pay_type(ev, args, 5))[0])


def f_ppmt(ev, args):
    rate, per, nper, pv = (num(ev.ev(a)) for a in args[:4])
    interest, pmt = _ipmt(rate, per, nper, pv, num(_arg(ev, args, 4, 0.0)), _pay_type(ev, args, 5))
    return finite(pmt - interest)


def f_pv(ev, args):
    rate, nper, pmt = (num(ev.ev(a)) for a in args[:3])
    fv, kind = num(_arg(ev, args, 3, 0.0)), _pay_type(ev, args, 4)
    if rate == 0:
        return finite(-(fv + pmt * nper))
    if rate == -1:
        _fail("#DIV/0!")
    q = _growth(rate, nper)
    return finite(-(fv + pmt * (1 + rate * kind) * (q - 1) / rate) / q)


def f_fv(ev, args):
    rate, nper, pmt = (num(ev.ev(a)) for a in args[:3])
    return finite(_fv(rate, nper, pmt, num(_arg(ev, args, 3, 0.0)), _pay_type(ev, args, 4)))


def f_nper(ev, args):
    rate, pmt, pv = (num(ev.ev(a)) for a in args[:3])
    fv, kind = num(_arg(ev, args, 3, 0.0)), _pay_type(ev, args, 4)
    if rate == 0:
        return -(pv + fv) / pmt if pmt else _fail("#DIV/0!")
    if 1 + rate == 1:
        raise VBANotImplemented("NPER at a rate too small to change 1 + rate")
    a, b = pmt * (1 + rate * kind) - fv * rate, pmt * (1 + rate * kind) + pv * rate
    if b == 0 or a / b <= 0 or rate <= -1:
        _fail("#NUM!")
    return finite(math.log(a / b) / math.log(1 + rate))


def f_npv(ev, args):
    rate = num(ev.ev(args[0]))
    if rate == -1:
        _fail("#DIV/0!")
    total = 0.0
    for i, x in enumerate(_numbers_of(ev, args[1:], "NPV"), 1):
        total += x * _power(1 + rate, -i)
    return finite(total)


def _cells_and_values(v):
    return [x for row in v.full() for x in row] if isinstance(v, Array) else [v]


def f_xnpv(ev, args):
    rate = num(ev.ev(args[0]))
    values, dates = _cells_and_values(ev.ev(args[1])), _cells_and_values(ev.ev(args[2]))
    if rate <= 0 or len(values) != len(dates):
        _fail("#NUM!")
    for x in values + dates:
        if isinstance(x, XLError):
            raise _Err(x)
        if x is None:
            raise VBANotImplemented("XNPV with a blank among its values or dates")
    if any(not isinstance(x, (int, float)) or isinstance(x, bool) for x in values):
        _fail("#NUM!")
    days = [_date_arg(d, atp=True) for d in dates]
    if any(d < days[0] for d in days):
        _fail("#NUM!")
    return finite(sum(x / (1 + rate) ** ((d - days[0]) / 365) for x, d in zip(values, days, strict=True)))


def f_irr(ev, args):
    """IRR: #NUM! for cash flows that never change sign, which have no rate. Any other is found by trial, and Excel
    stops within its tolerance of the rate, where it stops depending on its search (IRR({-50,-50,30,30,60}) is
    0.06892884559521573, the rate 0.0689288456029159): not answered."""
    flows = _numbers_of(ev, args[:1], "IRR", whole=True)
    if not (any(x > 0 for x in flows) and any(x < 0 for x in flows)):
        _fail("#NUM!")
    return f_found_by_trial(ev, args)


def f_found_by_trial(ev, args):
    raise VBANotImplemented(
        "a rate Excel finds by trial and stops within its tolerance of, where the digits a cell shows depend on its "
        "search (IRR, XIRR, RATE)"
    )


def f_sln(ev, args):
    cost, salvage, life = (num(ev.ev(a)) for a in args)
    return (cost - salvage) / life if life else _fail("#DIV/0!")


def f_syd(ev, args):
    cost, salvage, life, per = (num(ev.ev(a)) for a in args)
    if life <= 0 or not 0 < per <= life:
        _fail("#NUM!")
    return (cost - salvage) * (life - per + 1) * 2 / (life * (life + 1))


def f_db(ev, args):
    """DB: fixed-declining balance, its rate rounded to three places; a first year of `month` months, and a last
    period for the months left."""
    cost, salvage, life, period = (num(ev.ev(a)) for a in args[:4])
    month = num(_arg(ev, args, 4, 12.0))
    if cost < 0 or salvage < 0 or life <= 0 or period <= 0 or not 1 <= month <= 12:
        _fail("#NUM!")
    if period > (life + 1 if month < 12 else life):
        _fail("#NUM!")
    if period != math.floor(period):
        raise VBANotImplemented("DB of a period with a fraction")
    if cost == 0:
        raise VBANotImplemented("DB of a cost of 0")
    rate = fx.round_half_up(fifteen(1 - (salvage / cost) ** (1 / life)), 3)
    total = 0.0
    for p in range(1, int(period) + 1):
        if p == 1:
            dep = cost * rate * month / 12
        elif p == life + 1:
            dep = (cost - total) * rate * (12 - month) / 12
        else:
            dep = (cost - total) * rate
        total += dep
    return dep


def f_ddb(ev, args):
    """DDB: double-declining (or `factor`) balance, never below the salvage value: the value at the period's start less
    the value at its end, each the cost declined by a power of (1 - factor/life)."""
    cost, salvage, life, period = (num(ev.ev(a)) for a in args[:4])
    factor = num(_arg(ev, args, 4, 2.0))
    if cost < 0 or salvage < 0 or life <= 0 or period <= 0 or factor <= 0 or period > life:
        _fail("#NUM!")
    if period != math.floor(period):
        raise VBANotImplemented("DDB of a period with a fraction")
    rate = factor / life
    if rate >= 1:
        rate, before = 1.0, cost if period == 1 else 0.0
    else:
        before = cost * _power(1 - rate, period - 1)
    after = cost * _power(1 - rate, period)
    return max(before - salvage if after < salvage else before - after, 0.0)


# ---------------------------------------------------------------------------------------------- SUBTOTAL, AGGREGATE
# Both pass over cells whose formulas call SUBTOTAL or AGGREGATE anywhere (AGGREGATE unless asked not to). SUBTOTAL's
# 1 to 11 pass over every hidden row while the sheet is filtered (FilterMode), the rows the filter hid and any other,
# and none while it is not; 101 to 111 pass over hidden rows always. Hidden columns count.


def _reduce(kind, xs):
    """AGGREGATE's functions 1 to 13 (SUBTOTAL's 1 to 11) over the values of cells, errors already passed over or not."""
    if kind == 3:
        return float(sum(1 for x in xs if x is not None))
    for x in xs:
        if isinstance(x, XLError):
            raise _Err(x)
    numbers = [x for x in xs if isinstance(x, (int, float)) and not isinstance(x, bool)]
    if kind == 2:
        return float(len(numbers))
    if kind == 1:
        return _mean(numbers) if numbers else _fail("#DIV/0!")
    if kind == 4:
        return max(numbers, default=0.0)
    if kind == 5:
        return min(numbers, default=0.0)
    if kind == 6:
        out = 1.0
        for x in numbers:
            out *= x
        return finite(out) if numbers else 0.0
    if kind in (7, 8):
        return math.sqrt(_variance(numbers, kind == 7))
    if kind == 9:
        return fx.total(numbers)
    if kind in (10, 11):
        return _variance(numbers, kind == 10)
    if kind == 12:
        return _median(numbers)
    return _mode(numbers)


def _rect_named(ev, node):
    """The cells an argument names, a reference or a function that hands back cells; None for anything else."""
    while node[0] == "paren":
        node = node[1]
    if node[0] in ("ref", "name", "range"):
        return ev.rect_of(node)
    if node[0] == "func" and node[1] in fx._GIVES_CELLS:
        v = ev.keep(node)
        if isinstance(v, Array) and v.rect is not None:
            return v.rect
    return None


def _cells_counted(ev, rect, skip_hidden, skip_nested):
    """The values of a rectangle's cells, row by row, past the rows and nested totals passed over. skip_hidden is
    True, False, or None for SUBTOTAL's 1 to 11: hidden rows while the sheet is filtered."""
    sheet = rect[0]
    hidden = sheet.hidden_rows if (sheet.in_filter_mode() if skip_hidden is None else skip_hidden) else ()
    nested = ev.book.nested_totals(sheet, *rect[1:]) if skip_nested else ()
    v = ev.cells(rect)
    out = []
    for i, row in enumerate(v.rows[: v.extent[0]]):
        r = rect[1] + i
        if r in hidden:
            continue
        out += [x for j, x in enumerate(row[: v.extent[1]]) if (r, rect[2] + j) not in nested]
    return out


def f_subtotal(ev, args):
    kind = _whole(ev.ev(args[0]))
    skip_hidden = True if kind > 100 else None
    kind = kind - 100 if kind > 100 else kind
    if not 1 <= kind <= 11:
        _fail("#VALUE!")
    xs = []
    for a in args[1:]:
        rect = _rect_named(ev, a)
        if rect is None:
            _fail("#VALUE!")
        xs += _cells_counted(ev, rect, skip_hidden, True)
    return _reduce(kind, xs)


_ARRAY_FORM = {14: lambda xs, k: _kth(xs, k, True), 15: lambda xs, k: _kth(xs, k, False),
               16: lambda xs, k: _percentile(xs, k, False), 17: lambda xs, k: _quartile(xs, k, False),
               18: lambda xs, k: _percentile(xs, k, True), 19: lambda xs, k: _quartile(xs, k, True)}  # fmt: skip


def f_aggregate(ev, args):
    """AGGREGATE(function, options, ...): options 0 to 3 pass over nested totals, 1, 3, 5 and 7 hidden rows (by hand or
    by a filter), 2, 3, 6 and 7 errors. Functions 14 to 19 take an array, which may be worked out (AGGREGATE(15,6,
    A1:A9/(B1:B9="x"),1)), and k; hidden rows and nested totals count only in a reference."""
    kind, options = _whole(ev.ev(args[0])), _whole(ev.ev(args[1]))
    if not 1 <= kind <= 19 or not 0 <= options <= 7 or len(args) < 3:
        _fail("#VALUE!")
    skip_nested, skip_hidden, skip_errors = options <= 3, options % 2 == 1, options in (2, 3, 6, 7)
    if kind >= 14:
        if len(args) < 4:
            _fail("#VALUE!")
        rect = _rect_named(ev, args[2])
        if rect is not None:
            _check_filtered(rect, skip_hidden)
            xs = _cells_counted(ev, rect, skip_hidden, skip_nested)
        else:
            xs = list(ev.array(ev.whole(args[2])).values())
        k = ev.operand(args[3])
        if isinstance(k, Array):
            if ev.arrays and k.extent != (1, 1):
                raise VBANotImplemented("AGGREGATE given several k at once")
            k = k.cell(0, 0)
        numbers = []
        for x in xs:
            if isinstance(x, XLError):
                if not skip_errors:
                    raise _Err(x)
            elif isinstance(x, (int, float)) and not isinstance(x, bool):
                numbers.append(x)
        return _ARRAY_FORM[kind](numbers, num(k))
    xs = []
    for a in args[2:]:
        rect = _rect_named(ev, a)
        if rect is None:
            _fail("#VALUE!")
        _check_filtered(rect, skip_hidden)
        xs += _cells_counted(ev, rect, skip_hidden, skip_nested)
    if skip_errors:
        xs = [x for x in xs if not isinstance(x, XLError)]
    return _reduce(kind, xs)


def _check_filtered(rect, skip_hidden):
    sheet = rect[0]
    if not skip_hidden and sheet.in_filter_mode() and any(rect[1] <= r <= rect[3] for r in sheet.hidden_rows):
        raise VBANotImplemented("AGGREGATE that keeps hidden rows, over rows a filter hides")


# ------------------------------------------------------------------------------------------ the format a formula brings
# As a formula is entered in a General cell, Excel gives the cell a number format from the function its value comes
# from: DATE a date, PMT a currency, IRR a percentage. ROUND, MAX, SUM and a few more keep the format of what they are
# given, and so do + and - (a date and a number make a date, two dates a count of days: General) and the minus sign;
# IF, CHOOSE, ABS, YEAR, *, / and the like give none. A cell with a format of its own keeps it. MacroPress works this
# out for formulas that call a function with a format of its own, and stops where Excel's choice is not known; one
# that calls none stays General, where Excel may give it the format of a cell it reads (not modelled yet).
_CURRENCY, _PERCENT = BUILTIN_FORMATS[8], BUILTIN_FORMATS[9]
_OWN_FORMATS = {"DATE": BUILTIN_FORMATS[14], "TODAY": BUILTIN_FORMATS[14], "NOW": BUILTIN_FORMATS[22],
                "TIME": BUILTIN_FORMATS[18], "IRR": _PERCENT, "RATE": _PERCENT,
                **{name: _CURRENCY for name in ("PMT", "IPMT", "PPMT", "PV", "FV", "NPV", "SLN", "SYD", "DB", "DDB")}}  # fmt: skip
_KEEP_FORMAT = {"ROUND", "ROUNDUP", "ROUNDDOWN", "TRUNC", "INT", "MAX", "MIN", "SUM", "AVERAGE", "MEDIAN"}
# (SECOND, DAYS360 and the .INTL two are counted with HOUR, DAYS, NETWORKDAYS and WORKDAY, which were checked)
_NO_FORMAT = set(
    """IF IFERROR IFNA CHOOSE ABS CEILING TEXT DATEDIF EDATE EOMONTH YEAR MONTH DAY WEEKDAY HOUR MINUTE SECOND NETWORKDAYS
    NETWORKDAYS.INTL WORKDAY WORKDAY.INTL WEEKNUM ISOWEEKNUM YEARFRAC DAYS DAYS360 NPER XNPV XIRR""".split()
)


class _Unknown(Exception):
    pass


def entered_format(text, format_of):
    """The number format a General cell takes as the formula `text` (without its =) is entered in it; None for
    General. format_of(node) gives the format of the first cell of a reference. VBANotImplemented where Excel's choice
    is not known."""
    upper = text.upper()
    if not any(name in upper for name in _OWN_FORMATS):
        return None
    tree = fx.parse(text)
    if not _OWN_FORMATS.keys() & fx.function_names(tree):
        return None
    try:
        return _format_of(tree, format_of)
    except _Unknown:
        raise VBANotImplemented(f"the number format Excel gives ={text} as it is entered") from None


def _format_of(node, format_of):
    kind = node[0]
    if kind in ("ref", "name", "range"):
        return format_of(node)
    if kind in ("paren", "neg"):
        return _format_of(node[1], format_of)
    if kind == "bin":
        op = node[1]
        if op in ("&", "=", "<>", "<", ">", "<=", ">="):
            return None
        a, b = _format_of(node[2], format_of), _format_of(node[3], format_of)
        if op in ("+", "-"):
            return _sum_format(op, a, b)
        if op in ("*", "/") and all(f is None or f in (_CURRENCY, _PERCENT) or is_date_format(f) or is_time_format(f)
                                    for f in (a, b)):  # fmt: skip
            return None
        raise _Unknown
    if kind == "func":
        name, args = node[1], node[2]
        if name in _OWN_FORMATS:
            return _OWN_FORMATS[name]
        if name in _NO_FORMAT:
            return None
        formats = [_format_of(a, format_of) for a in args]
        given = {f for f in formats if f is not None}
        if not given:
            return None
        if name == "MOD" and formats[0] is not None and formats[1:] == [None]:  # MOD(NOW(),1): the time of day
            return formats[0]
        if name in _KEEP_FORMAT and len(given) == 1:
            return given.pop()
        raise _Unknown
    if kind in ("pct", "at"):
        raise _Unknown
    return None  # a number, text, TRUE, an error or an array constant


def _sum_format(op, a, b):
    if b is None:
        return a  # DATE(...)+1, TODAY()-1, PMT(...)+1
    if a is None:
        if is_date_format(b):
            return b  # 1+DATE(...), 1-DATE(...)
        raise _Unknown
    if is_date_format(a) and is_date_format(b):
        return None  # two dates: the days between, or a sum that is no date
    if (is_date_format(a) and is_time_format(b)) or (is_time_format(a) and (is_time_format(b) or op == "-")):
        return a  # DATE(...)+TIME(...), TIME(...)+TIME(...), TIME(...)-DATE(...)
    raise _Unknown


# ------------------------------------------------------------------------------------------------------ registration
FUNCTIONS = {
    "ROUNDUP": f_roundup, "ROUNDDOWN": f_rounddown, "TRUNC": f_trunc, "CEILING": f_ceiling, "FLOOR": f_floor,
    "CEILING.MATH": f_ceiling_math, "FLOOR.MATH": f_floor_math, "CEILING.PRECISE": f_ceiling_precise,
    "ISO.CEILING": f_ceiling_precise, "FLOOR.PRECISE": f_floor_precise, "MROUND": f_mround, "MOD": f_mod,
    "QUOTIENT": f_quotient, "SIGN": f_sign, "POWER": f_power, "SQRT": f_sqrt, "EXP": f_exp, "LN": f_ln, "LOG": f_log,
    "LOG10": f_log10, "PI": f_pi, "EVEN": f_even, "ODD": f_odd, "PRODUCT": f_product,
    "DATE": f_date, "YEAR": f_year, "MONTH": f_month, "DAY": f_day, "WEEKDAY": f_weekday, "EOMONTH": f_eomonth,
    "EDATE": f_edate, "DATEDIF": f_datedif, "DAYS": f_days, "ISOWEEKNUM": f_isoweeknum, "WEEKNUM": f_weeknum,
    "NETWORKDAYS": f_networkdays, "NETWORKDAYS.INTL": f_networkdays_intl, "WORKDAY": f_workday,
    "WORKDAY.INTL": f_workday_intl, "YEARFRAC": f_yearfrac, "DAYS360": f_days360, "TIME": f_time, "HOUR": f_hour,
    "MINUTE": f_minute, "SECOND": f_second,
    "FIND": f_find, "SEARCH": f_search, "SUBSTITUTE": f_substitute, "REPLACE": f_replace, "REPT": f_rept,
    "EXACT": f_exact, "PROPER": f_proper, "TEXTJOIN": f_textjoin, "CHAR": f_char, "CODE": f_code, "UNICHAR": f_unichar,
    "UNICODE": f_unicode, "CLEAN": f_clean, "N": f_n, "T": f_t, "VALUE": f_value,
    "IFS": f_ifs, "SWITCH": f_switch, "IFNA": f_ifna, "ISERROR": f_iserror, "ISERR": f_iserr, "ISNA": f_isna,
    "ISLOGICAL": f_islogical, "ISNONTEXT": f_isnontext, "ISEVEN": f_iseven, "ISODD": f_isodd, "ISREF": f_isref,
    "NA": f_na, "ERROR.TYPE": f_error_type, "TYPE": f_type,
    "MEDIAN": f_median, "MODE": f_mode, "MODE.SNGL": f_mode, "STDEV": f_stdev_s, "STDEV.S": f_stdev_s,
    "STDEVP": f_stdev_p, "STDEV.P": f_stdev_p, "VAR": f_var_s, "VAR.S": f_var_s, "VARP": f_var_p, "VAR.P": f_var_p,
    "LARGE": f_large, "SMALL": f_small, "PERCENTILE": f_percentile_inc, "PERCENTILE.INC": f_percentile_inc,
    "PERCENTILE.EXC": f_percentile_exc, "QUARTILE": f_quartile_inc, "QUARTILE.INC": f_quartile_inc,
    "QUARTILE.EXC": f_quartile_exc, "RANK": f_rank, "RANK.EQ": f_rank, "RANK.AVG": f_rank_avg,
    "COUNTBLANK": f_countblank,
    "PMT": f_pmt, "IPMT": f_ipmt, "PPMT": f_ppmt, "PV": f_pv, "FV": f_fv, "NPER": f_nper, "NPV": f_npv, "XNPV": f_xnpv,
    "IRR": f_irr, "XIRR": f_found_by_trial, "RATE": f_found_by_trial, "SLN": f_sln, "SYD": f_syd, "DB": f_db,
    "DDB": f_ddb,
    "SUBTOTAL": f_subtotal, "AGGREGATE": f_aggregate,
}  # fmt: skip

# How each takes its arguments (see formula.TAKES). w is a value for the functions that came from the Analysis ToolPak:
# a range is taken whole, and several cells are #VALUE! in a formula that is not a dynamic array (EOMONTH(A1:A3,0)).
# TYPE's o is one value of a range but an array kept whole (TYPE({1,2}) is 64).
_VALUES = """ROUNDUP ROUNDDOWN TRUNC CEILING FLOOR CEILING.MATH FLOOR.MATH CEILING.PRECISE ISO.CEILING FLOOR.PRECISE MOD
    SIGN POWER SQRT EXP LN LOG LOG10 EVEN ODD DATE YEAR MONTH DAY WEEKDAY DATEDIF DAYS ISOWEEKNUM DAYS360 TIME HOUR
    MINUTE SECOND FIND SEARCH SUBSTITUTE REPLACE REPT EXACT PROPER CHAR CODE UNICHAR UNICODE CLEAN VALUE IFNA ISERROR
    ISERR ISNA ISLOGICAL ISNONTEXT ERROR.TYPE PMT IPMT PPMT PV FV NPER SLN SYD DB DDB XIRR RATE"""
TAKES = {
    **{name: "v" for name in _VALUES.split()},
    **{name: "w" for name in "MROUND QUOTIENT ISEVEN ISODD EOMONTH EDATE WEEKNUM YEARFRAC".split()},
    "NETWORKDAYS": "wwr", "WORKDAY": "wwr", "NETWORKDAYS.INTL": "wwvr", "WORKDAY.INTL": "wwvr",
    "TEXTJOIN": "rvr", "IFS": "vr" * 127, "SWITCH": "v" + "vr" * 126, "N": "r", "T": "r", "TYPE": "o",
    "LARGE": "rv", "SMALL": "rv", "PERCENTILE": "rv", "PERCENTILE.INC": "rv", "PERCENTILE.EXC": "rv", "QUARTILE": "rv",
    "QUARTILE.INC": "rv", "QUARTILE.EXC": "rv", "RANK": "vrv", "RANK.EQ": "vrv", "RANK.AVG": "vrv",
    "NPV": "vr", "XNPV": "vrr", "IRR": "av", "MODE": "a", "MODE.SNGL": "a", "SUBTOTAL": "vr", "AGGREGATE": "vvar",
}  # fmt: skip

# How many arguments each takes, and which must be references (see formula._ARGUMENTS).
_COUNTS = {
    "ROUNDUP": (2, 2), "ROUNDDOWN": (2, 2), "TRUNC": (1, 2), "CEILING": (2, 2), "FLOOR": (2, 2),
    "CEILING.MATH": (1, 3), "FLOOR.MATH": (1, 3), "CEILING.PRECISE": (1, 2), "ISO.CEILING": (1, 2),
    "FLOOR.PRECISE": (1, 2), "MROUND": (2, 2), "MOD": (2, 2), "QUOTIENT": (2, 2), "SIGN": (1, 1), "POWER": (2, 2),
    "SQRT": (1, 1), "EXP": (1, 1), "LN": (1, 1), "LOG": (1, 2), "LOG10": (1, 1), "PI": (0, 0), "EVEN": (1, 1),
    "ODD": (1, 1), "PRODUCT": (1, 255), "DATE": (3, 3), "YEAR": (1, 1), "MONTH": (1, 1), "DAY": (1, 1),
    "WEEKDAY": (1, 2), "EOMONTH": (2, 2), "EDATE": (2, 2), "DATEDIF": (3, 3), "DAYS": (2, 2), "ISOWEEKNUM": (1, 1),
    "WEEKNUM": (1, 2), "NETWORKDAYS": (2, 3), "NETWORKDAYS.INTL": (2, 4), "WORKDAY": (2, 3), "WORKDAY.INTL": (2, 4),
    "YEARFRAC": (2, 3), "DAYS360": (2, 3), "TIME": (3, 3), "HOUR": (1, 1), "MINUTE": (1, 1), "SECOND": (1, 1),
    "FIND": (2, 3), "SEARCH": (2, 3), "SUBSTITUTE": (3, 4), "REPLACE": (4, 4), "REPT": (2, 2), "EXACT": (2, 2),
    "PROPER": (1, 1), "TEXTJOIN": (3, 254), "CHAR": (1, 1), "CODE": (1, 1), "UNICHAR": (1, 1), "UNICODE": (1, 1),
    "CLEAN": (1, 1), "N": (1, 1), "T": (1, 1), "VALUE": (1, 1), "SWITCH": (3, 254), "IFNA": (2, 2),
    "ISERROR": (1, 1), "ISERR": (1, 1), "ISNA": (1, 1), "ISLOGICAL": (1, 1), "ISNONTEXT": (1, 1), "ISEVEN": (1, 1),
    "ISODD": (1, 1), "ISREF": (1, 1), "NA": (0, 0), "ERROR.TYPE": (1, 1), "TYPE": (1, 1), "MEDIAN": (1, 255),
    "MODE": (1, 255), "MODE.SNGL": (1, 255), "STDEV": (1, 255), "STDEV.S": (1, 255), "STDEVP": (1, 255),
    "STDEV.P": (1, 255), "VAR": (1, 255), "VAR.S": (1, 255), "VARP": (1, 255), "VAR.P": (1, 255), "LARGE": (2, 2),
    "SMALL": (2, 2), "PERCENTILE": (2, 2), "PERCENTILE.INC": (2, 2), "PERCENTILE.EXC": (2, 2), "QUARTILE": (2, 2),
    "QUARTILE.INC": (2, 2), "QUARTILE.EXC": (2, 2), "PMT": (3, 5), "IPMT": (4, 6), "PPMT": (4, 6), "PV": (3, 5),
    "FV": (3, 5), "NPER": (3, 5), "NPV": (2, 255), "XNPV": (3, 3), "IRR": (1, 2), "XIRR": (2, 3), "RATE": (3, 6),
    "SLN": (3, 3), "SYD": (4, 4), "DB": (4, 5), "DDB": (4, 5), "AGGREGATE": (3, 253),
}  # fmt: skip
ARGUMENTS = {
    **{name: (fewest, most, 1, "") for name, (fewest, most) in _COUNTS.items()},
    "IFS": (2, 254, 2, ""),
    "COUNTBLANK": (1, 1, 1, "r"),
    "RANK": (2, 3, 1, "-r"), "RANK.EQ": (2, 3, 1, "-r"), "RANK.AVG": (2, 3, 1, "-r"),
    "SUBTOTAL": (2, 255, 1, "-" + "r" * 254),
}  # fmt: skip

fx.FUNCTIONS.update(FUNCTIONS)
fx.TAKES.update(TAKES)
fx._ARGUMENTS.update(ARGUMENTS)
fx.EXCEL_FUNCTIONS.update(FUNCTIONS, ARGUMENTS)
