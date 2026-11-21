"""What Excel makes of text where it expects a value: written to a cell, as it reads text typed into one, or met in
a formula's arithmetic ("5%"+1).

Checked against Excel with some 300 texts (tests/parity). Excel reads numbers and dates as the machine's regional
settings write them; this follows US settings: "." for decimals, "," between thousands, "$" for currency.
"""

import re

from .errors import VBANotImplemented
from .numfmt import CURRENCY, CURRENCY_CENTS, SCIENTIFIC
from .values import MONTH_NAMES, XLError, typed_number

ERRORS = ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A")

# Digits, with commas only where at least three digits follow each one ("1,234", not "1,23" or "12,5").
_DIGITS = r"(?:\d+(?:,\d{3,})*\.?\d*|\.\d+)"
_NUMBER = re.compile(rf"({_DIGITS})([eE][-+]?\d+)?")
_FRACTION = re.compile(r"(\d+) +(\d+)/(\d{1,2})")


def _amount(text):
    """An unsigned amount: (value, number format), or None. One of "$" before it and "%" before or after it is read.
    The value is None for a number too big to hold: Excel keeps the text."""
    currency = text.startswith("$")
    if currency:
        text = text[1:].lstrip(" ")
    percent = text.endswith("%") or text.startswith("%")
    if percent:
        text = text[:-1].rstrip(" ") if text.endswith("%") else text[1:]
    fraction = _FRACTION.fullmatch(text)
    if fraction and not (currency or percent) and int(fraction.group(3)):
        whole, top, bottom = (int(g) for g in fraction.groups())
        return whole + top / bottom, "# ?/?" if bottom < 10 else "# ??/??"
    number = _NUMBER.fullmatch(text)
    if number is None or (currency and percent):
        return None
    body, exponent = number.groups()
    value = typed_number(body.replace(",", "") + (exponent or ""))
    if value is None:
        return None, SCIENTIFIC if exponent else None  # 1e1000 stays text, in a cell formatted for the number
    if percent:
        value /= 100
    decimals = "." in body
    if exponent:
        return value, SCIENTIFIC
    if currency:
        return value, CURRENCY_CENTS if decimals else CURRENCY
    if percent:
        return value, "0.00%" if decimals else "0%"
    if "," in body:
        return value, "#,##0.00" if decimals else "#,##0"
    return value, None


def number(text):
    """The number a text is, as (value, number format), or None: with its sign in front ("-5", "- 5", "+5"), or as
    parentheses ("(5)"); with "$" in front or "%" at either end; with commas between thousands; with an exponent; or
    as a whole number and a fraction ("1 1/2")."""
    text = text.strip(" ")
    sign = 1
    if text.startswith("$") and text[1:2] in ("-", "(", "+"):  # "$-5", "$(5)"
        text = text[1:2] + "$" + text[2:]
    if text.startswith("(") and text.endswith(")"):
        sign, text = -1, text[1:-1].strip(" ")
    elif text.startswith("-"):
        sign, text = -1, text[1:].lstrip(" ")
    elif text.startswith("+"):
        text = text[1:]
    if text[:1] in ("-", "+", "("):
        return None
    amount = _amount(text)
    if amount is None or amount[0] is None:
        return amount
    return sign * amount[0], amount[1]


_MONTH = "|".join(sorted({m for m in MONTH_NAMES} | {m[:3] for m in MONTH_NAMES}, key=len, reverse=True))
_DATE_LIKE = (
    # 1/2, 12-13, 1-2-3, 2024/01/05, 1 / 2, with a time or not
    re.compile(r"\d{1,4} *[/-] *\d{1,4}(?: *[/-] *\d{1,4})?(?: +\d{1,2}:\d{1,2}(?::\d{1,2})?(?: *[ap]m?)?)?", re.I),
    # 3:30, 1:2:3, 12:00 AM, 3 pm
    re.compile(r"\d{1,2}(?::\d{1,2}){1,2}(?: *[ap]m?)?|\d{1,2} *[ap]m", re.I),
    # Jan 5, 5 Jan 2024, March 5, Jan-24, January 2024: a month's name with a number
    re.compile(rf"(?:\d{{1,4}}[ ,/-]+)?(?:{_MONTH})\.?(?:[ ,/-]+\d{{1,4}}){{0,2}}", re.I),
)


def looks_like_a_date(text):
    text = text.strip(" ")
    return any(p.fullmatch(text) for p in _DATE_LIKE) and any(ch.isdigit() for ch in text)


def value_of(text):
    """The number a formula takes text for in its arithmetic ("5%"+1, SUM("5",1)), or None for #VALUE!. Text with a
    currency sign, a date or a time is not modelled: Excel reads those as the machine's regional settings say."""
    as_number = number(text)
    if as_number is not None:
        if as_number[1] in (CURRENCY, CURRENCY_CENTS):
            raise VBANotImplemented(f"the number in {text!r}: a formula reads a currency sign by regional settings")
        return as_number[0]
    if looks_like_a_date(text):
        raise VBANotImplemented(f"the date or time in {text!r}: a formula reads it by regional settings")
    return None


def read(text):
    """What a General cell holds after `text` is written to it: (value, formula, number format). The value is a
    float, a bool, an error or text; the formula is the text after "=". A date or a time is not modelled: Excel reads
    it as the machine's regional settings say, and fills in the current year."""
    if text.startswith("'"):
        return text[1:], None, None  # the apostrophe says "this is text"
    if text.startswith("=") and len(text) > 1:
        return None, text[1:], None
    upper = text.upper()
    if upper in ("TRUE", "FALSE"):
        return upper == "TRUE", None, None
    if upper in ERRORS:
        return XLError(upper), None, None
    as_number = number(text)
    if as_number is not None:
        return (text if as_number[0] is None else as_number[0]), None, as_number[1]
    if looks_like_a_date(text):
        raise VBANotImplemented(f"writing {text!r} to a cell: Excel turns it into a date or a time")
    return text, None, None
