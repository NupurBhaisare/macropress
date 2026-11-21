"""Number formats: the ones Excel knows by number, what VBA reads from a cell that has one, and which format a
write leaves in a cell that already has one. Checked against Excel (tests/parity)."""

import re
from decimal import ROUND_HALF_UP, Decimal

from .errors import VBANotImplemented
from .values import fifteen

# The formats Excel knows by number, as VBA's NumberFormat gives them under US regional settings.
BUILTIN_FORMATS = {
    0: "General",
    1: "0",
    2: "0.00",
    3: "#,##0",
    4: "#,##0.00",
    5: "$#,##0_);($#,##0)",
    6: "$#,##0_);[Red]($#,##0)",
    7: "$#,##0.00_);($#,##0.00)",
    8: "$#,##0.00_);[Red]($#,##0.00)",
    9: "0%",
    10: "0.00%",
    11: "0.00E+00",
    12: "# ?/?",
    13: "# ??/??",
    14: "m/d/yyyy",
    15: "d-mmm-yy",
    16: "d-mmm",
    17: "mmm-yy",
    18: "h:mm AM/PM",
    19: "h:mm:ss AM/PM",
    20: "h:mm",
    21: "h:mm:ss",
    22: "m/d/yyyy h:mm",
    37: "#,##0_);(#,##0)",
    38: "#,##0_);[Red](#,##0)",
    39: "#,##0.00_);(#,##0.00)",
    40: "#,##0.00_);[Red](#,##0.00)",
    41: '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
    42: '_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
    43: '_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
    44: '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    45: "mm:ss",
    46: "[h]:mm:ss",
    47: "mm:ss.0",
    48: "##0.0E+0",
    49: "@",
}
TEXT = "@"
DATE, TIME, DATE_AND_TIME = BUILTIN_FORMATS[14], BUILTIN_FORMATS[19], BUILTIN_FORMATS[22]
CURRENCY, CURRENCY_CENTS, CURRENCY_VALUE = BUILTIN_FORMATS[6], BUILTIN_FORMATS[8], BUILTIN_FORMATS[7]
SCIENTIFIC = BUILTIN_FORMATS[11]
REGIONAL = "a number format of the region Excel runs in, number "


def builtin_code(format_id):
    """The code of a number format the file gives no code for: one Excel knows by number. Some of those numbers mean
    a format that depends on the region Excel runs in (East Asian dates); a cell so formatted is not read."""
    return BUILTIN_FORMATS.get(format_id) or f"{REGIONAL}{format_id}"


# What a number format shows as it stands: "text", [colour] or [condition], an escaped \\x, a _x space, a *x fill.
_LITERALS = re.compile(r'"[^"]*"|\[[^\]]*\]|\\.|_.|\*.')
_DATE_PARTS = re.compile(r"y+|d+|m+|h+|s+", re.I)


def vba_code(code):
    """A format code of the file as VBA gives it: a $ that stands bare in one is a dollar sign as it stands (\\$),
    not the currency sign of the regional settings that Excel's own currency formats have."""
    return re.sub(r'("[^"]*"|\[[^\]]*\]|\\.|_.|\*.)|\$', lambda m: m.group(1) or "\\$", code)


def is_date_format(code):
    """Whether VBA reads a number so formatted as a Date: its format shows a day, a month or a year. A format of
    hours, minutes and seconds alone (h:mm, [h]:mm:ss, mm:ss) reads as a Double."""
    if not code or code in ("General", TEXT):
        return False
    if code.startswith(REGIONAL):
        raise VBANotImplemented(f"reading a cell formatted with {code}")
    first = _LITERALS.sub("", code).split(";")[0]
    first = re.sub(r"(?i)am/pm|a/p|(?<=s)\.0+", "", first)  # ss.00 is seconds, not digits
    if re.search(r"[0#?]", first):
        return False
    parts = [p[0].lower() for p in _DATE_PARTS.findall(first)]
    for i, part in enumerate(parts):
        minutes = part == "m" and (parts[i - 1 : i] == ["h"] or parts[i + 1 : i + 2] == ["s"])
        if part in "yd" or (part == "m" and not minutes):
            return True
    return False


_SIGN_WRITTEN_OUT = re.compile(r'\[\$[^\]-]+(-[^\]]*)?\]|"[^"]*[$€£¥₹][^"]*"|\\[$€£¥₹]|[€£¥₹]')


def writes_a_currency_sign(code):
    """Whether a format writes a currency sign out as text, where Excel's own currency formats leave it to the
    regional settings: "$"#,##0.00, [$€-407] #,##0.00, £0.00."""
    return bool(code) and _SIGN_WRITTEN_OUT.search(code) is not None


def is_currency_format(code):
    """Whether VBA reads a number so formatted as a Currency: its format has the currency sign of the regional
    settings, as Excel's own Currency and Accounting formats do ($#,##0.00). A sign written out in the format
    ("$"#,##0.00, [$€-407]) does not count."""
    return bool(code) and "$" in _LITERALS.sub("", code)


_FIXED = re.compile(r"(#,##)?0(\.0+)?")


def shown(x, code=None):
    """The text a cell shows for a number, in a column with room for it, or None where that is not modelled. It is
    for the General format, where the number is short enough to stand as it is (Excel rounds one of more than 11
    characters, and writes a very large or small one with an exponent), and for formats of plain digits: 0, 0.00,
    #,##0 and #,##0.00, with any number of decimals."""
    number = Decimal(repr(fifteen(x)))
    if code in (None, "General", TEXT):
        text = format(number, "f")
        text = text.rstrip("0").rstrip(".") if "." in text else text
        return text if len(text) <= 11 and (x == 0 or abs(x) >= 1e-5) else None
    fixed = _FIXED.fullmatch(code)
    if fixed is None:
        return None
    decimals = len(fixed.group(2) or ".") - 1
    rounded = number.quantize(Decimal(1).scaleb(-decimals), rounding=ROUND_HALF_UP)
    return f"{rounded:,.{decimals}f}" if fixed.group(1) else f"{rounded:.{decimals}f}"


# Excel's own formats by what they are for. A write that brings a format of another kind (a date into a cell of
# percentages) replaces one of these; it leaves a plain number format (0.00, #,##0) or one made by hand in place.
_KINDS = {"currency": (5, 6, 7, 8), "percent": (9, 10), "scientific": (11, 48), "fraction": (12, 13),
          "date": (14, 15, 16, 17, 22), "time": (18, 19, 20, 21, 45, 46, 47), "accounting": (41, 42, 43, 44)}  # fmt: skip
KIND = {BUILTIN_FORMATS[i]: kind for kind, ids in _KINDS.items() for i in ids}


def after_entry(cell, entry):
    """The number format a cell has after a write: `cell` is the one it had, `entry` the one the value brings (a
    Date's m/d/yyyy, the 0% of "5%"; None for a plain number or text). None stands for General."""
    if entry is None or entry == cell:
        return cell
    if cell is None:
        return entry
    if entry not in KIND:
        return cell  # the #,##0 of "1,234": only a General cell takes it
    return entry if cell in KIND and KIND[cell] != KIND[entry] else cell
