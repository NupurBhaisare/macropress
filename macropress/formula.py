"""Excel formulas: parse them and evaluate the ones a macro writes or whose inputs it changes.

Covers what macros commonly need: references (A1, $A$1:$B$9, Sheet!A1, 'My sheet'!A:A, 2:5), defined names and table
columns (Tbl[col]), arithmetic, & and comparisons with Excel's (not VBA's) rules, and a set of functions including
the dynamic-array SORT and UNIQUE, whose results spill. Anything else raises VBANotImplemented: a formula is never
given a guessed value.
"""

import datetime as dt
import functools
import math
import re
import unicodedata
from decimal import ROUND_HALF_UP, Decimal

from . import entry
from .errors import VBANotImplemented
from .values import XLError, date_to_serial, fifteen, format_double, typed_number

MAX_ROW, MAX_COL = 1048576, 16384

_TOKEN = re.compile(
    r"""
 (?P<ws>\s+)
|(?P<str>"(?:[^"]|"")*")
|(?P<err>\#(?:N/A|REF!|DIV/0!|VALUE!|NAME\?|NUM!|NULL!|SPILL!|CALC!))
|(?P<ref>(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?(?![\w(\[]))
|(?P<colref>(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}(?![\w(]))
|(?P<rowref>(?:(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!)?\$?\d+:\$?\d+(?![\w(]))
|(?P<num>\d+\.?\d*(?:[eE][-+]?\d+)?|\.\d+(?:[eE][-+]?\d+)?)
|(?P<func>[A-Za-z_][\w.]*(?=\())
|(?P<name>[A-Za-z_][\w.]*(?:\[[^\]]*\])*)
|(?P<op><>|<=|>=|[-+*/^&=<>(),%{};@:])
""",
    re.VERBOSE,
)
REFERENCE_KINDS = ("ref", "colref", "rowref")

# Text no token reads that holds no reference either: a [table column] or [other workbook], and a 'quoted sheet' name
# before a defined name. Taken whole, no part of it is ever read as a reference.
_OPAQUE = re.compile(r"\[(?:[^\[\]]|\[[^\]]*\])*\]|'(?:[^']|'')*'")


def pieces(text):
    """The formula cut into (kind, text) pieces, in order: kind is the token's, or None for text no token reads (an
    operator such as : or @, a {} constant, a [table column]). Joined, the pieces give back the formula. Whatever reads
    or rewrites a formula's text goes through these, so a reference is the same thing everywhere."""
    pos = 0
    while pos < len(text):
        m = _TOKEN.match(text, pos) or _OPAQUE.match(text, pos)
        piece = m.group() if m else text[pos]
        yield (m.lastgroup if m else None), piece
        pos += len(piece)


BIN_PREC = {"=": 1, "<>": 1, "<": 1, ">": 1, "<=": 1, ">=": 1, "&": 2, "+": 3, "-": 3, "*": 4, "/": 4, "^": 5}

# Functions newer than Excel 2007 are stored in the file with a prefix (Excel shows #NAME? without it). These were
# checked against a file Excel saved; ISO.CEILING, NETWORKDAYS.INTL and WORKDAY.INTL have none.
_NEWER_2010_2013 = """CEILING.MATH FLOOR.MATH CEILING.PRECISE FLOOR.PRECISE DAYS ISOWEEKNUM UNICHAR UNICODE RANK.EQ RANK.AVG
    STDEV.S STDEV.P VAR.S VAR.P MODE.SNGL PERCENTILE.INC PERCENTILE.EXC QUARTILE.INC QUARTILE.EXC AGGREGATE"""
FILE_PREFIX = {
    "SORT": "_xlfn._xlws.SORT",
    "FILTER": "_xlfn._xlws.FILTER",
    "UNIQUE": "_xlfn.UNIQUE",
    "XLOOKUP": "_xlfn.XLOOKUP",
    "XMATCH": "_xlfn.XMATCH",
    "SORTBY": "_xlfn.SORTBY",
    "SEQUENCE": "_xlfn.SEQUENCE",
    "LET": "_xlfn.LET",
    "IFS": "_xlfn.IFS",
    "CONCAT": "_xlfn.CONCAT",
    "TEXTJOIN": "_xlfn.TEXTJOIN",
    "SWITCH": "_xlfn.SWITCH",
    "MAXIFS": "_xlfn.MAXIFS",
    "MINIFS": "_xlfn.MINIFS",
    "IFNA": "_xlfn.IFNA",
    "SINGLE": "_xlfn.SINGLE",
    # 2.6, as a file Excel saved has them
    **{
        name: "_xlfn." + name
        for name in """TAKE DROP CHOOSECOLS CHOOSEROWS VSTACK HSTACK TOCOL TOROW EXPAND WRAPROWS WRAPCOLS
       TEXTSPLIT TEXTBEFORE TEXTAFTER LAMBDA ISOMITTED MAP REDUCE SCAN BYROW BYCOL MAKEARRAY""".split()
    },
    **{name: "_xlfn." + name for name in _NEWER_2010_2013.split()},
}


def file_syntax(text):
    """A formula as the file stores it: newer functions get their _xlfn. prefix, @ is _xlfn.SINGLE(), and the names
    LET and LAMBDA bind are _xlpm.x (an optional parameter, [y], _xlop.y where it is bound)."""
    if "@" in text:
        text = _replace_nodes(text, "at", lambda body, node, spans: f"SINGLE({_text_of(body, node[1], spans)})")
    if _BINDS.search(text):
        text = _parameters_marked(text)
    return "".join(FILE_PREFIX.get(p.upper(), p) if kind == "func" else p for kind, p in pieces(text))


_XLFN = re.compile(r"(?i)^_xlfn\.(?:_xlws\.)?")
_XLPM = re.compile(r"(?i)^_xlpm\.")
_XLOP = re.compile(r"(?i)^_xlop\.(.*)$")
_BINDS = re.compile(r"(?i)\b(?:LET|LAMBDA)\(")
_PRIMARY = ("ref", "name", "func", "array", "paren")


def _param_key(node):
    """The name a LET or LAMBDA node binds, as kept to look it up (capitals, without _xlpm.), or None. A name may be
    written like a cell, as LET(A1,1,A1) is (recorded)."""
    if node[0] == "opt":
        return node[1].upper()
    if node[0] == "name":
        return _XLOP.sub(r"\1", _XLPM.sub("", node[1])).upper()
    if node[0] == "ref" and re.fullmatch(r"[A-Za-z]{1,3}\d+", node[1]):
        return node[1].upper()
    return None


def _parameters_marked(text):
    """The formula with every name LET and LAMBDA bind, and every use of one, written as the file keeps them."""
    parser = Parser(text)
    tree = parser.parse()
    body, edits = parser.body, []

    def mark(node):
        start, end = parser.spans[id(node)]
        if node[0] == "opt":
            edits.append((start, end, "_xlop." + node[1]))
        elif not body[start:].lower().startswith(("_xlpm.", "_xlop.")):
            edits.append((start, start, "_xlpm."))

    def walk(node, scope):
        kind, args = node[0], node[2] if node[0] == "func" else ()
        if kind == "func" and node[1] in ("LET", "LAMBDA") and args:
            pairs = (
                list(zip(args[:-1:2], args[1:-1:2], strict=False))
                if node[1] == "LET"
                else [(a, None) for a in args[:-1]]
            )
            inner = set(scope)
            for name, value in pairs:
                if value is not None:
                    walk(value, inner)
                key = _param_key(name)
                if key is not None:
                    mark(name)
                    inner.add(key)
            return walk(args[-1], inner)
        if kind in ("name", "ref") and _param_key(node) in scope:
            return mark(node)
        if kind == "func" and _XLPM.sub("", node[1]).upper() in scope:
            mark(node)
        for child in _children(node):
            walk(child, scope)

    walk(tree, set())
    for start, end, new in sorted(edits, reverse=True):
        body = body[:start] + new + body[end:]
    return ("=" if text.startswith("=") else "") + body


def bound_names(text):
    """The names a formula's LET and LAMBDA bind, as _param_key keeps them; none where it binds none."""
    if not _BINDS.search(text):
        return set()
    try:
        tree = parse(text)
    except VBANotImplemented:
        return set()
    out = set()
    for node in _all(tree):
        if node[0] == "func" and node[1] in ("LET", "LAMBDA") and node[2]:
            out |= {_param_key(a) for a in (node[2][:-1:2] if node[1] == "LET" else node[2][:-1])}
    return out - {None}


def display_syntax(text):
    """A formula as VBA's Formula property shows it: without the file's _xlfn. and _xlpm. prefixes, an optional
    parameter (_xlop.y) as [y], and _xlfn.SINGLE(x) as @x."""
    out = []
    for kind, p in pieces(text):
        if kind == "func":
            p = _XLPM.sub("", _XLFN.sub("", p))
        elif kind == "name":
            p = _XLOP.sub(r"[\1]", _XLPM.sub("", p))
        out.append(p)
    text = "".join(out)
    if "SINGLE(" in text.upper():

        def at(body, node, spans):
            inner = _text_of(body, node[2][0], spans)
            return "@" + (inner if node[2][0][0] in _PRIMARY else f"({inner})")

        text = _replace_nodes(text, "SINGLE", at)
    return text


def _text_of(body, node, spans):
    start, end = spans[id(node)]
    return body[start:end]


def _replace_nodes(text, kind, write):
    """The formula with each node of a kind ("at", or a function's name) written anew by write(body, node, spans);
    nodes inside another of the kind are left to it."""
    parser = Parser(text)
    tree = parser.parse()
    found = []

    def walk(node):
        if node[0] == kind or (node[0] == "func" and node[1] == kind):
            found.append(node)
            return
        for child in _children(node):
            walk(child)

    walk(tree)
    body = parser.body
    for node in sorted(found, key=lambda n: parser.spans[id(n)][0], reverse=True):
        start, end = parser.spans[id(node)]
        body = body[:start] + write(body, node, parser.spans) + body[end:]
    return ("=" if text.startswith("=") else "") + body


def _children(node):
    if node[0] == "func":
        return node[2]
    if node[0] == "call":
        return [node[1], *node[2]]
    if node[0] == "bin":
        return [node[2], node[3]]
    if node[0] == "range":
        return [node[1], node[2]]
    if node[0] in ("neg", "pct", "paren", "at"):
        return [node[1]]
    return []


# ---------------------------------------------------------------------------
# A formula's text as Excel keeps it, and as Formula and Formula2 show it
# ---------------------------------------------------------------------------

# Excel's own functions, whose names it writes in capitals whatever case they are typed in. A name not here may be a
# function of the workbook's VBA, which keeps its own case.
EXCEL_FUNCTIONS = set(
    """ABS ACOS AGGREGATE AND AVERAGE AVERAGEA AVERAGEIF AVERAGEIFS CEILING CHAR CHOOSE CLEAN CODE COLUMN COLUMNS
    CONCAT CONCATENATE COS COUNT COUNTA COUNTBLANK COUNTIF COUNTIFS DATE DATEDIF DATEVALUE DAY DAYS DAYS360 DB DDB
    DOLLAR EDATE EOMONTH EXACT EXP FALSE FILTER FIND FIXED FLOOR FV HLOOKUP HOUR IF IFERROR IFNA IFS INDEX INDIRECT
    INT IPMT IRR ISBLANK ISERR ISERROR ISEVEN ISLOGICAL ISNA ISNONTEXT ISNUMBER ISODD ISREF ISTEXT LARGE LEFT LEN
    LET LN LOG LOG10 LOOKUP LOWER MATCH MAX MAXA MAXIFS MEDIAN MID MIN MINA MINIFS MINUTE MIRR MOD MODE MONTH MROUND
    N NA NETWORKDAYS NOT NOW NPER NPV OFFSET OR PI PMT POWER PPMT PRODUCT PROPER PV QUOTIENT RAND RANDBETWEEN RANK
    RATE REPLACE REPT RIGHT ROUND ROUNDDOWN ROUNDUP ROW ROWS SEARCH SECOND SEQUENCE SIGN SIN SLN SMALL SORT SORTBY
    SQRT STDEV SUBSTITUTE SUBTOTAL SUM SUMIF SUMIFS SUMPRODUCT SWITCH SYD T TAN TEXT TEXTJOIN TIME TIMEVALUE TODAY
    TRANSPOSE TRIM TRUE TRUNC TYPE UNIQUE UPPER VALUE VAR VLOOKUP WEEKDAY WEEKNUM WORKDAY XIRR XLOOKUP XMATCH XNPV
    YEAR YEARFRAC""".split()
)


def normalized(text, find_sheet):
    """A formula (without its =) as Excel keeps what is typed: its own functions' names, references and TRUE and
    FALSE in capitals, a sheet's name as the sheet has it (quoted where it could be read as a cell), a number as
    Excel writes it (-2.5E+3 is -2500), and no space before a comma; other spaces stay where they were typed."""
    out = []
    for kind, piece in pieces(text):
        if kind == "func" and bare_name(piece) in EXCEL_FUNCTIONS:
            piece = piece.upper()
        elif kind in REFERENCE_KINDS:
            name, area = ref_parts(piece)
            sheet = find_sheet(name) if name is not None else None
            prefix = f"{sheet_ref(sheet.name)}!" if sheet is not None else piece[: len(piece) - len(area)]
            piece = prefix + area.upper()
        elif kind == "name" and piece.upper() in ("TRUE", "FALSE"):
            piece = piece.upper()
        elif kind == "num":
            number = typed_number(piece)
            piece = format_double(number) if number is not None else piece
        out.append(piece)
    joined = "".join(out)
    return re.sub(r"\s+(?=,)", "", joined) if " " in joined else joined


# Functions that may hand back several cells or values: where a formula that is not a dynamic array ends on one,
# Formula2 shows it with @ (=@INDEX(A1:A5,0)). For the rest, whose results are one value, it does not.
_MANY_RESULTS = {"INDEX", "IF", "ROW", "COLUMN", "OFFSET", "INDIRECT", "CHOOSE", "IFS", "SWITCH", "IFERROR", "IFNA",
                 "XLOOKUP", "LOOKUP", "TRANSPOSE", "MMULT", "MINVERSE", "FREQUENCY", "TREND", "GROWTH", "LINEST",
                 "LOGEST", "SORT", "SORTBY", "UNIQUE", "FILTER", "SEQUENCE", "RANDARRAY", "TEXTSPLIT", "LET",
                 "SINGLE", "TAKE", "DROP", "CHOOSECOLS", "CHOOSEROWS", "VSTACK", "HSTACK", "TOCOL", "TOROW", "EXPAND",
                 "WRAPROWS", "WRAPCOLS", "MAP", "REDUCE", "SCAN", "BYROW", "BYCOL", "MAKEARRAY", "LAMBDA"}  # fmt: skip
# Of those, the ones a formula that is not a dynamic array is shown to end on with @ before them (=@TAKE(A1:A6,2)),
# and takes the first value of: recorded.
_FIRST_OF = {"FILTER", "TAKE", "SEQUENCE", "TEXTSPLIT"}


class _Unsure(Exception):
    pass


def _implicit(tree, spans, size, top="top"):
    """The @ Excel shows in a formula that is not a dynamic array, as Formula2 shows it: [(offset, node)] of the
    ranges and results it takes one value of. size(node) is a ref's or name's (rows, columns). Raises _Unsure where
    Excel's choice is not known."""
    sites = []

    def many(node):
        """Whether a node may stand for several values once its ranges in value places are taken as one."""
        kind = node[0]
        if kind == "range":
            return True
        if kind in ("ref", "name"):
            return size(node) != (1, 1)
        if kind == "array":
            return len(node[1]) * len(node[1][0]) > 1
        if kind in ("neg", "pct", "paren"):
            return many(node[1])
        if kind == "bin":
            return _arrayish(node[2]) or _arrayish(node[3])
        if kind == "func":
            return _func_many(node)
        return False

    def _arrayish(node):  # in a value place, after its ranges are taken as one
        if node[0] in ("ref", "name"):
            return False
        return many(node)

    def _func_many(node):
        name, args = node[1], node[2]
        if name == "INDEX":  # a row or column of 0, or left empty, takes all of them; an array worked out may be any
            if args and args[0][0] not in ("ref", "name"):
                return True
            numbers = [a for a in args[1:3] if a[0] != "missing"]
            if any(a[0] != "num" for a in numbers):
                raise _Unsure
            return any(a[1] == 0 for a in numbers) or len(numbers) < len(args) - 1
        if name == "IF":
            return any(many(a) for a in args[1:3])
        if name in ("ROW", "COLUMN"):
            if not args or args[0][0] == "missing":
                return False
            if args[0][0] not in ("ref", "name"):
                raise _Unsure
            rows, cols = size(args[0])
            return (rows if name == "ROW" else cols) > 1
        if name in ("XLOOKUP", "LOOKUP"):
            if _arrayish(args[0]):
                raise _Unsure
            return False
        if name == "CHOOSE":  # an index from a range is one value; an array constant picks many
            return any(many(a) for a in args[1:]) or bool(args and args[0][0] == "array" and many(args[0]))
        if name in ("INDIRECT", *_FIRST_OF):
            return True
        if name == "OFFSET":  # as tall and wide as the reference unless told, and a size given is taken as many
            if not args or args[0][0] not in ("ref", "name"):
                raise _Unsure
            given = [a for a in args[3:5] if a[0] != "missing"]
            if any(a[0] != "num" for a in given):
                raise _Unsure
            if any(abs(a[1]) != 1 for a in given):
                return True
            if given:
                raise _Unsure
            return size(args[0]) != (1, 1)
        if name in _MANY_RESULTS:
            raise _Unsure
        return False

    def walk(node, place):
        """place: top (what the formula ends on), value (a function's argument that wants one value), operand (an
        operator's), ref (taken whole), array (worked out as an array)."""
        kind = node[0]
        if kind in ("ref", "name", "range"):
            if place in ("top", "value", "operand") and (kind == "range" or size(node) != (1, 1)):
                sites.append(node)
            return
        if kind == "array":  # an array stays one, but where one value is wanted it is shown with @ (=ROUND(@{1,2},0))
            if place in ("top", "value") and many(node):
                sites.append(node)
            return
        if kind in ("neg", "pct", "paren"):
            return walk(node[1], "operand" if place in ("top", "value") and kind != "paren" else place)
        if kind == "at":
            return
        if kind in ("call", "opt"):
            raise _Unsure
        if kind == "bin":
            inner = "array" if place == "array" else "operand"
            walk(node[2], inner)
            walk(node[3], inner)
            if place in ("top", "value", "operand") and (_arrayish(node[2]) or _arrayish(node[3])):
                raise _Unsure
            return
        if kind == "func":
            name, args = node[1], node[2]
            if place in ("top", "value", "operand") and _func_many(node):
                if place == "value" or name not in (
                    "INDEX",
                    "IF",
                    "ROW",
                    "COLUMN",
                    "CHOOSE",
                    "OFFSET",
                    "INDIRECT",
                    *_FIRST_OF,
                ):
                    raise _Unsure
                sites.append(node)
                place = "ref"  # under the @, what it hands back is taken whole
            if place == "array" or name == "SUMPRODUCT":
                for a in args:
                    walk(a, "array")
                return
            if name == "IF":
                walk(args[0], "value")
                for a in args[1:3]:
                    walk(a, place if place in ("ref", "array") else "ref")
                return
            if name == "CHOOSE":  # an index that is an array constant picks value by value, from values taken whole
                if args and args[0][0] == "array":
                    for a in args[1:]:
                        walk(a, "array")
                    return
                if args:
                    walk(args[0], "value")
                for a in args[1:]:
                    walk(a, place if place in ("ref", "array") else "ref")
                return
            if name not in FUNCTIONS and name not in EXCEL_FUNCTIONS:
                raise _Unsure
            takes = TAKES.get(name)
            for i, a in enumerate(args):
                letter = takes[min(i, len(takes) - 1)] if takes else None
                if letter is None and a[0] in ("ref", "name") and size(a) != (1, 1) and name not in FUNCTIONS:
                    raise _Unsure  # how a function we do not model takes a range is not known
                if letter == "w":  # a range whole, anything else one value: =EOMONTH(A1:A3,0), =EOMONTH(@{1,2},0)
                    letter = "r" if a[0] in ("ref", "name", "range") else "v"
                elif letter == "o":
                    letter = "v"
                walk(a, "value" if letter == "v" else "array" if letter == "a" else "ref")
            return

    walk(tree, top)
    return sites


def ref_size(book, sheet):
    """size(node), for the views below: (rows, columns) of what a ref or name of a formula on `sheet` names."""
    ev = Evaluator(book, sheet, None)

    def size(node):
        try:
            _, r1, c1, r2, c2 = ev.rect_of(node)
        except _Err:
            return (1, 1)
        return (r2 - r1 + 1, c2 - c1 + 1)

    return size


def needs_arrays(text, size):
    """Whether a dynamic-array formula means something else read as a formula that is not one: whether it takes a
    range or an array whole where that one would take one value of it. Saved, it must stay a dynamic array."""
    return formula2_view(text, size) != text


def formula2_view(text, size):
    """How Formula2 shows a formula that is not a dynamic array: with @ where Excel takes one value of a range or of
    what a function hands back. None where where Excel puts them is not known."""
    parser = Parser(text)
    tree = parser.parse()
    try:
        sites = _implicit(tree, parser.spans, size)
    except _Unsure:
        return None
    body = parser.body
    for start in sorted((parser.spans[id(n)][0] for n in sites), reverse=True):
        body = body[:start] + "@" + body[start:]
    return body


def legacy_view(text):
    """How Formula shows a dynamic-array formula: without the @ that a formula that is not one takes as read, before
    a range or result where one value is wanted. None where that is not known: a function it does not model."""
    if "@" not in text:
        return text
    parser = Parser(text)
    tree = parser.parse()
    if any(n[0] == "func" and n[1] not in FUNCTIONS for n in _all(tree)):
        return None
    drop = []

    def walk(node, place):
        kind = node[0]
        if kind == "at" and place in ("top", "value"):
            drop.append(node)
        elif kind in ("neg", "pct", "paren"):
            walk(node[1], place)
        elif kind == "bin":
            walk(node[2], "value")
            walk(node[3], "value")
        elif kind == "func":
            takes = "v" + "r" * len(node[2]) if node[1] == "IF" else TAKES.get(node[1])
            for i, a in enumerate(node[2]):
                letter = takes[min(i, len(takes) - 1)] if takes else "r"
                walk(a, "value" if letter == "v" else "ref")

    walk(tree, "top")
    body = parser.body
    for node in sorted(drop, key=lambda n: parser.spans[id(n)][0], reverse=True):
        start, _ = parser.spans[id(node)]
        body = body[:start] + body[start + 1 :]
    return body


def _all(node):
    yield node
    for child in _children(node):
        yield from _all(child)


def bare_name(name):
    n = name.upper()
    for prefix in ("_XLFN._XLWS.", "_XLFN.", "_XLWS.", "_XLPM."):
        if n.startswith(prefix):
            n = n[len(prefix) :]
    return n


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------


_OPTIONAL = re.compile(r"\[[A-Za-z_][\w.]*\]")


def tokenize(text, offsets=None):
    """The formula's tokens, (kind, text), spaces left out; offsets, a list, is given where each starts. A [name]
    standing alone is LAMBDA's optional parameter ("opt")."""
    out, pos = [], 0
    for kind, piece in pieces(text):
        if kind is None and _OPTIONAL.fullmatch(piece):
            kind = "opt"
        if kind is None:
            raise VBANotImplemented(f"formula syntax at {text[pos : pos + 20]!r} in ={text}")
        if kind != "ws":
            out.append((kind, piece))
            if offsets is not None:
                offsets.append(pos)
        pos += len(piece)
    return out


class Parser:
    """Reads a formula into nodes: ("num", x), ("str", s), ("bool", b), ("err", e), ("ref", text), ("name", text),
    ("func", NAME, [args]), ("array", [[values]]), ("bin", op, a, b), ("neg", a), ("pct", a), ("paren", a),
    ("at", a) for @ (one value of what follows), ("range", a, b) for a:b between references that are not both cells
    (A1:INDEX(A:A,5)), ("missing",) for an argument left empty, ("call", f, [args]) for a call of what a function or
    parentheses give (LAMBDA(x,x*2)(5)), and ("opt", name) for LAMBDA's optional parameter [name]. `spans` keeps where
    each node starts and ends in the text (after the =)."""

    def __init__(self, text):
        self.text = text
        self.offsets = []
        self.body = text[1:] if text.startswith("=") else text
        self.toks = tokenize(self.body, self.offsets)
        self.i = 0
        self.spans = {}

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self):
        t = self.peek()
        self.i += 1
        return t

    def _mark(self, node, first):
        """Remember where the node lies: from token `first` to the last token taken."""
        last = self.i - 1
        self.spans[id(node)] = (self.offsets[first], self.offsets[last] + len(self.toks[last][1]))
        return node

    def parse(self):
        node = self.expr(0)
        if self.i != len(self.toks):
            raise VBANotImplemented(f"formula syntax near {self.peek()[1]!r} in ={self.text}")
        return node

    def expr(self, min_prec):
        first = self.i
        left = self.unary()
        while True:
            kind, val = self.peek()
            if kind != "op" or val not in BIN_PREC or BIN_PREC[val] < min_prec:
                return left
            self.take()
            right = self.expr(BIN_PREC[val] + 1)
            left = self._mark(("bin", val, left, right), first)

    def unary(self):
        first = self.i
        kind, val = self.peek()
        if kind == "op" and val in ("-", "+"):
            self.take()
            operand = self.unary()
            return self._mark(("neg", operand), first) if val == "-" else operand
        if (kind, val) == ("op", "@"):
            self.take()
            return self._mark(("at", self.unary()), first)
        node = self.primary()
        while self.peek() == ("op", "%"):
            self.take()
            node = self._mark(("pct", node), first)
        return node

    def primary(self):
        """A value, and the range from it to another reference after a colon, which binds before any operator."""
        first = self.i
        node = self._mark(self._primary(), first)
        while self.peek() == ("op", "(") and node[0] in ("func", "paren", "call"):
            self.take()
            node = self._mark(("call", node, self._args()), first)
        while self.peek() == ("op", ":"):
            self.take()
            start = self.i
            node = self._mark(("range", node, self._mark(self._primary(), start)), first)
        return node

    def _array(self):
        """An array constant, after its {: rows split by ;, values by , and every row as long as the first. Only
        numbers (with a sign), text, TRUE, FALSE and errors may stand in one; anything else is error 1004 in Excel."""
        rows, row = [], []
        while True:
            kind, val = self.take()
            sign = 1
            if (kind, val) in (("op", "-"), ("op", "+")):
                sign = -1 if val == "-" else 1
                kind, val = self.take()
                if kind != "num":
                    raise _bad_array(self.text)
            if kind == "num":
                number = typed_number(val)
                if number is None:
                    raise _bad_array(self.text)
                row.append(sign * number)
            elif kind == "str":
                row.append(val[1:-1].replace('""', '"'))
            elif kind == "name" and val.upper() in ("TRUE", "FALSE"):
                row.append(val.upper() == "TRUE")
            elif kind == "err":
                row.append(XLError(val))
            else:
                raise _bad_array(self.text)
            kind, val = self.take()
            if (kind, val) == ("op", ","):
                continue
            rows.append(row)
            if rows and len(row) != len(rows[0]):
                raise _bad_array(self.text)
            if (kind, val) == ("op", "}"):
                return ("array", rows)
            if (kind, val) != ("op", ";"):
                raise _bad_array(self.text)
            row = []

    def _primary(self):
        kind, val = self.take()
        if (kind, val) == ("op", "{"):
            return self._array()
        if kind == "num":
            number = typed_number(val)
            if number is None:
                raise VBANotImplemented(f"the number {val} in ={self.text}: too large for Excel")
            return ("num", number)
        if kind == "str":
            return ("str", val[1:-1].replace('""', '"'))
        if kind == "err":
            return ("err", XLError(val))
        if kind in REFERENCE_KINDS:
            return ("ref", val)
        if kind == "func":
            self.take()  # (
            return ("func", bare_name(val), self._args())
        if kind == "opt":
            return ("opt", val[1:-1])
        if kind == "name":
            if val.upper() in ("TRUE", "FALSE"):
                return ("bool", val.upper() == "TRUE")
            return ("name", val)
        if (kind, val) == ("op", "("):
            node = self.expr(0)
            if self.take() != ("op", ")"):
                raise VBANotImplemented(f"formula syntax in ={self.text}")
            return ("paren", node)
        raise VBANotImplemented(f"formula syntax near {val!r} in ={self.text}")


def _args_of(parser):
    """A call's arguments, after its (: one left empty is ("missing",)."""
    args = []
    if parser.peek() == ("op", ")"):
        parser.take()
        return args
    while True:
        if parser.peek() in (("op", ","), ("op", ")")):
            args.append(("missing",))
        else:
            args.append(parser.expr(0))
        k, v = parser.take()
        if (k, v) == ("op", ")"):
            return args
        if (k, v) != ("op", ","):
            raise VBANotImplemented(f"formula syntax in ={parser.text}")


Parser._args = _args_of


def _bad_array(text):
    from .errors import VBARuntimeError  # noqa: PLC0415

    return VBARuntimeError(1004, f"Excel cannot read the array constant in ={text}")


def parse(text):
    return Parser(text).parse()


def function_names(node, out=None):
    """The (bare) names of the functions a parsed formula calls."""
    out = set() if out is None else out
    if node[0] == "func":
        out.add(node[1])
    for child in _children(node):
        function_names(child, out)
    return out


def references(node, out=None):
    """Every ref and name token in a parsed formula."""
    out = [] if out is None else out
    if node[0] in ("ref", "name"):
        out.append(node)
    for child in _children(node):
        references(child, out)
    return out


# ---------------------------------------------------------------------------
# Values with Excel's rules
# ---------------------------------------------------------------------------


class Array:
    """A 2-D block of values: a range read, an array constant, or a result worked out cell by cell. A range read
    holds the cells up to the sheet's last data only; `extent` is the size of the range itself, and its cells past
    the data are blank. A result's rows past the data of what it was worked out from are all alike: `tail` is that
    row (None: blank). `rect` is where a range read lies, (sheet, r1, c1, r2, c2): what a formula that is not a
    dynamic array takes one cell of."""

    LIMIT = 1 << 21  # cells a result is worked out to, one by one: two whole columns

    def __init__(self, rows, extent=None, rect=None, tail=None):
        self.rows = rows
        self.extent = extent or self.shape
        self.rect = rect
        self.tail = tail

    @property
    def shape(self):
        return len(self.rows), (len(self.rows[0]) if self.rows else 0)

    def cell(self, i, j):
        """The value at 0-based (i, j): blank past a range's data, the tail past a result's."""
        if i < len(self.rows):
            row = self.rows[i]
            return row[j] if j < len(row) else None
        return self.tail[j] if self.tail is not None and j < len(self.tail) else None

    def counted(self):
        """Every value that is not blank, with how many cells hold it: each of the data once, and each value of the
        rows past the data, which are alike, once for all of them."""
        rows, cols = self.extent
        held = min(len(self.rows), rows)
        for row in self.rows[:held]:
            for v in row[:cols]:
                yield v, 1
        if held < rows and self.tail is not None:
            for v in self.tail[:cols]:
                if v is not None:
                    yield v, rows - held

    def values(self):
        """Every value that is not blank, row by row, a value of the alike rows past the data as often as it stands."""
        for v, n in self.counted():
            for _ in range(n):
                yield v

    def full(self):
        """The rows of the whole extent, written out."""
        rows, cols = self.extent
        if rows * cols > self.LIMIT:
            raise VBANotImplemented(f"an array of {rows} by {cols} values, more than this works out one by one")
        return [[self.cell(i, j) for j in range(cols)] for i in range(rows)]


class Lambda:
    """What LAMBDA gives: a function of its parameters, [(name, optional)], working out `body` with the names bound
    where it was made (`scope`) and its arguments. A cell that ends on one shows #CALC!."""

    def __init__(self, params, body, scope):
        self.params, self.body, self.scope = params, body, scope


_OMITTED = XLError("omitted")  # what a parameter in [ ] left out is bound to: blank where used (recorded)
_MISSING = XLError("missing")  # an argument left out of a call


def scalar(v):
    if isinstance(v, Array):
        if v.extent == (1, 1):
            return v.cell(0, 0)
        raise VBANotImplemented("an array where a formula expects one value")
    if isinstance(v, Lambda):
        raise VBANotImplemented("a LAMBDA where a formula wants a value")
    return v


def single(v):
    """One value of an array or range: of a 1 by 1 block, its value; else the block itself."""
    return v.cell(0, 0) if isinstance(v, Array) and v.extent == (1, 1) else v


def num(v):
    v = scalar(v)
    if isinstance(v, XLError):
        raise _Err(v)
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, dt.datetime):
        return date_to_serial(v)
    if isinstance(v, str):
        number = entry.value_of(v)  # text is read as it would be typed in a cell: "5%", "1,000", "(5)", "1 1/2"
        if number is not None:
            return number
    raise _Err(XLError("#VALUE!"))


def finite(x):
    """A result of arithmetic: one too large for a number is #NUM!."""
    if not math.isfinite(x):
        raise _Err(XLError("#NUM!"))
    return x


def snap_add(a, b):
    """a + b as Excel gives it where a formula or SUM ends on it: 0 when the two all but cancel, leaving 50 or more
    binary places below the larger of them (0.1+0.2-0.3 is 0, not 5.55E-17)."""
    r = a + b
    if r and math.frexp(max(abs(a), abs(b)))[1] - math.frexp(r)[1] >= 50:
        return 0.0
    return r


def power(x, y):
    """x^y. A negative number has a root when the root is odd ((-8)^(1/3) is -2), and no other fractional power."""
    if x == 0:
        if y > 0:
            return 0.0
        raise _Err(XLError("#NUM!" if y == 0 else "#DIV/0!"))
    try:
        if x < 0 and y != int(y):
            root = fifteen(1 / y)
            if root != int(root) or not int(root) % 2:
                raise _Err(XLError("#NUM!"))
            return finite(-(abs(x) ** y))
        return finite(x**y)
    except OverflowError:
        raise _Err(XLError("#NUM!")) from None


def text(v):
    v = scalar(v)
    if isinstance(v, XLError):
        raise _Err(v)
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return format_double(v)
    return str(v)


def truth(v):
    v = scalar(v)
    if isinstance(v, XLError):
        raise _Err(v)
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    if isinstance(v, str):
        if v.upper() in ("TRUE", "FALSE"):
            return v.upper() == "TRUE"
        raise _Err(XLError("#VALUE!"))
    return bool(v)


class _Err(Exception):
    def __init__(self, err):
        super().__init__(str(err))
        self.err = err


class _NotCells(_Err):
    """A name that stands for a LAMBDA, where cells are wanted: it reads none, and what its LAMBDA reads is known only
    as it is worked out."""


def type_rank(v):
    """Excel's order: numbers, text, logical values, errors; blanks are handled apart."""
    if isinstance(v, XLError):
        return 3
    if isinstance(v, bool):
        return 2
    if isinstance(v, (int, float)):
        return 0
    if isinstance(v, str):
        return 1
    return 3


# Excel's order for symbols (checked against Excel for Mac): punctuation and symbols, then digits, then letters.
# Hyphens and apostrophes are ordinary symbols here, not ignored.
_SYMBOL_ORDER = {ch: i for i, ch in enumerate(" _-,;:!?.'\"()[]{}@*/\\&#%`^+<=>|~$")}


def _char_key(ch):
    if ch in _SYMBOL_ORDER:
        return (0, _SYMBOL_ORDER[ch])
    cat = unicodedata.category(ch)
    if cat == "Nd":
        return (1, unicodedata.digit(ch, 0))
    if cat[0] in "PSZC":
        return (0, 100 + ord(ch))
    return (2, unicodedata.normalize("NFD", ch)[0].casefold())


def text_order_key(s, match_case=False):
    """Text in Excel's sort order: character by character, ignoring case; accents only break ties."""
    key = (tuple(_char_key(ch) for ch in s), unicodedata.normalize("NFD", s.casefold()))
    if match_case:
        key += (tuple(0 if ch.islower() else 1 for ch in s),)  # lower case first
    return key


def sort_value_key(v):
    rank = type_rank(v)
    if rank == 0:
        return (rank, float(v))
    if rank == 1:
        return (rank, text_order_key(v))
    if rank == 2:
        return (rank, bool(v))
    return (rank, str(v))


def compare(a, b, exact=False):
    """Excel's comparison: -1/0/1. Blank meets a number as 0 and text as ''; types never convert. Numbers are
    compared to 15 significant digits (0.1+0.2=0.3 is TRUE), unless exact: a lookup's match is on the full number."""
    a, b = scalar(a), scalar(b)
    for v in (a, b):
        if isinstance(v, XLError):
            raise _Err(v)
    if a is None and b is None:
        return 0
    if a is None:
        a = "" if isinstance(b, str) else (False if isinstance(b, bool) else 0.0)
    if b is None:
        b = "" if isinstance(a, str) else (False if isinstance(a, bool) else 0.0)
    ra, rb = type_rank(a), type_rank(b)
    if ra != rb:
        return -1 if ra < rb else 1
    if ra == 1:
        x, y = text_order_key(a), text_order_key(b)
    elif ra == 0:
        x, y = (float(a), float(b)) if exact else (fifteen(a), fifteen(b))
    else:
        x, y = a, b
    return (x > y) - (x < y)


def round_half_up(x, digits):
    q = Decimal(1).scaleb(-int(digits))
    return float(Decimal(repr(x)).quantize(q, rounding=ROUND_HALF_UP))


# ---------------------------------------------------------------------------
# Evaluation
# ---------------------------------------------------------------------------


def ref_parts(token):
    """'Sheet'!$A$1:$B$2 -> (sheet name or None, area text)."""
    if "!" in token:
        sheet, area = token.rsplit("!", 1)
        if sheet.startswith("'"):
            sheet = sheet[1:-1].replace("''", "'")
        return sheet, area
    return None, token


def sheet_ref(name):
    """A sheet's name as a formula writes it before the !: quoted unless it is a plain name that can't read as a cell."""
    if re.fullmatch(r"[A-Za-z_][\w.]*", name) and not re.fullmatch(r"[A-Za-z]{1,3}\d+|[Rr]\d*[Cc]\d*", name):
        return name
    return "'" + name.replace("'", "''") + "'"


def rename_sheets(text, renamed):
    """The formula with its references to renamed sheets ({old name lower-cased: new name}) naming the new ones, as
    Excel rewrites them; all in one pass, so two sheets can swap names."""
    out = []
    for kind, piece in pieces(text):
        if kind in REFERENCE_KINDS and "!" in piece:
            sheet, area = ref_parts(piece)
            if sheet.lower() in renamed:
                piece = f"{sheet_ref(renamed[sheet.lower()])}!{area}"
        out.append(piece)
    return "".join(out)


def _binary(op, a, b):
    """One operator on two values, as Excel works it out."""
    if op == "&":
        return text(a) + text(b)
    if op in ("=", "<>", "<", ">", "<=", ">="):
        c = compare(a, b)
        return {"=": c == 0, "<>": c != 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0}[op]
    x, y = num(a), num(b)
    if op == "+":
        return finite(x + y)
    if op == "-":
        return finite(x - y)
    if op == "*":
        return finite(x * y)
    if op == "/":
        if y == 0:
            raise _Err(XLError("#DIV/0!"))
        return finite(x / y)
    return power(x, y)


def _sum_end(op, a, b):
    """The + or - a formula ends on: a sum that all but cancels is 0 (see snap_add)."""
    x, y = num(a), num(b)
    return finite(snap_add(x, y if op == "+" else -y))


def _many(v):
    return isinstance(v, Array) and v.extent != (1, 1)


def _each(fn, operands):
    """fn worked out value by value over arrays, as Excel works an operator out over them: a single value goes with
    every value of the other, a single row with every row and a single column with every column; past the end of
    a shorter one the value is #N/A. An error stays in its place. The rows past every array's data are alike, and
    worked out once (the result's tail)."""
    arrays = [o for o in operands if isinstance(o, Array)]
    rows = max(a.extent[0] for a in arrays)
    cols = max(a.extent[1] for a in arrays)

    def at(o, i, j):
        if not isinstance(o, Array):
            return o
        r, c = o.extent
        if (i >= r and r != 1) or (j >= c and c != 1):
            return XLError("#N/A")
        return o.cell(0 if r == 1 else i, 0 if c == 1 else j)

    def one(i, j):
        try:
            v = fn(*[at(o, i, j) for o in operands])
        except _Err as e:
            return e.err
        except ZeroDivisionError:
            return XLError("#DIV/0!")
        if isinstance(v, Array):  # cells handed back for each value: one cell is its value
            if v.extent != (1, 1) and v.rect is not None:
                raise VBANotImplemented("a function handing back a range for each value of an array")
            return v.cell(0, 0)  # an array worked out for each value: its first, as Excel shows (SEQUENCE(A1:A2))
        if isinstance(v, Lambda):
            raise VBANotImplemented("a LAMBDA for each value of an array")
        return v

    start = 0  # from this row on, every row is alike
    for a in arrays:
        if a.extent[0] != 1:
            start = max(start, len(a.rows) if a.extent[0] == rows else a.extent[0])
    start = min(start, rows)
    if start * cols > Array.LIMIT:
        raise VBANotImplemented(f"an array of {rows} by {cols} values, more than this works out one by one")
    body = [[one(i, j) for j in range(cols)] for i in range(start)]
    return Array(body, (rows, cols), None, [one(start, j) for j in range(cols)] if start < rows else None)


class Evaluator:
    """Evaluates one formula of `sheet` against the book (reading cells through `read`). `cell` is where the formula
    is, (row, column): what ROW() gives, and the row or column a range gives a formula that wants one value.
    `arrays` is True for a dynamic-array formula (Formula2) or one entered with Ctrl+Shift+Enter: its ranges are
    taken whole everywhere, as arrays."""

    def __init__(self, book, sheet, read, cell=None, arrays=False, read_rows=None):
        self.book = book
        self.sheet = sheet
        self.read = read  # read(sheet, r, c) -> value (refreshing a stale formula first)
        self.read_rows = read_rows  # read_rows(sheet, r1, c1, r2, c2) -> a block's rows at once; None: one by one
        self.cell = cell
        self.arrays = arrays
        self.areas = {}  # the range arguments of SUMIF and the like, read once: id(node) -> Array
        self.scope = {}  # the names LET and LAMBDA bind where the formula is being worked out (see _param_key)
        self.depth = 0  # LAMBDAs called and not yet worked out

    def run(self, node):
        try:
            if node[0] == "bin" and node[1] in ("+", "-"):  # the formula ends on a sum: see snap_add
                a, b = self.operand(node[2]), self.operand(node[3])
                if _many(a) or _many(b):
                    return _each(lambda x, y, op=node[1]: _sum_end(op, x, y), [a, b])
                return _sum_end(node[1], single(a), single(b))
            v = self.ev(node)
            return XLError("#CALC!") if isinstance(v, Lambda) else v
        except _Err as e:
            return e.err
        except ZeroDivisionError:
            return XLError("#DIV/0!")

    def here(self):
        if self.cell is None:
            raise VBANotImplemented("a formula's own row or column, outside a cell")
        return self.cell

    def rect_of(self, node):
        """(sheet, r1, c1, r2, c2) for a ref or name node, or a range between two references: the box around both,
        on one sheet (#VALUE! across two)."""
        from .excel import parse_area

        key = self.bound_key(node)
        if key is not None:  # a name LET or LAMBDA binds, standing for the cells it was given
            v = self.scope[key]
            if isinstance(v, Array) and v.rect is not None:
                return v.rect
            raise VBANotImplemented(f"{node[1]}, which LET or LAMBDA binds to a value, where cells are wanted")
        if node[0] == "name" and self.named_lambda(node[1]) is not None:
            raise _NotCells(XLError("#VALUE!"))
        if node[0] == "range":
            a, b = self._cells_of(node[1]), self._cells_of(node[2])
            if a[0] is not b[0]:
                raise _Err(XLError("#VALUE!"))
            return (a[0], min(a[1], b[1]), min(a[2], b[2]), max(a[3], b[3]), max(a[4], b[4]))
        if node[0] == "ref":
            sname, area = ref_parts(node[1])
            sheet = self.book.find_sheet(sname) if sname else self.sheet
            if sheet is None:
                raise _Err(XLError("#REF!"))
            rect = parse_area(area.replace("$", ""))
            if rect is None:
                raise VBANotImplemented(f"formula reference {node[1]}")
            return (sheet,) + rect
        name = node[1]
        m = re.match(r"^([A-Za-z_][\w.]*)\[(.*)\]$", name)
        if m and m.group(1).lower() in self.book.tables:
            r = self.sheet.table_range(self.book.tables[m.group(1).lower()], m.group(2))
            return (r.sheet, r.r1, r.c1, r.r2, r.c2)
        ref = self.book.resolve_name(name, self.sheet)
        if ref is None:
            raise _Err(XLError("#NAME?"))
        node2 = ("ref", ref.lstrip("="))
        return self.rect_of(node2)

    def _cells_of(self, node):
        """The cells one side of a range names: a reference, or what a function hands back as cells (#VALUE! for a
        value)."""
        while node[0] == "paren":
            node = node[1]
        if node[0] in ("ref", "name", "range"):
            return self.rect_of(node)
        v = self.ev(node)
        if isinstance(v, Array) and v.rect is not None:
            return v.rect
        raise _Err(XLError("#VALUE!"))

    def range_values(self, rect):
        sheet, r1, c1, r2, c2 = rect
        extent = (r2 - r1 + 1, c2 - c1 + 1)
        b = sheet.used_bounds()
        if b is None:
            return Array([[None]], extent, rect)
        r2 = min(r2, max(b[2], r1))
        c2 = min(c2, max(b[3], c1))
        rows = self.read_rows(sheet, r1, c1, r2, c2) if self.read_rows else None
        if rows is None:
            rows = [[self.read(sheet, r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)]
        return Array(rows, extent, rect)

    def bound_key(self, node):
        """The key of the name a LET or LAMBDA binds that this name (or a name written like a cell) stands for."""
        if not self.scope or node[0] not in ("name", "ref"):
            return None
        key = _param_key(node)
        return key if key in self.scope else None

    def bound(self, key):
        """A bound name where it is used: one cell gives its value; cells, an array, a value or a LAMBDA themselves; a
        parameter left out, blank."""
        v = self.scope[key]
        if v is _OMITTED:
            return None
        if isinstance(v, Array) and v.rect is not None and v.extent == (1, 1):
            return v.cell(0, 0)
        return v

    def named_lambda(self, name):
        """A LAMBDA the workbook keeps as a defined name (Name Manager), worked out once; None for any other name."""
        if name.lower() not in self.book.names:
            return None
        text = self.book.resolve_name(name, self.sheet)
        if text is None or not re.match(r"(?i)^=?(?:_xlfn\.)?LAMBDA\(", text.lstrip()):
            return None
        made = self.book.named_lambdas.get(text)
        if made is None:
            tree = parse("=" + display_syntax(text.lstrip().lstrip("=")))
            made = self.book.named_lambdas[text] = f_lambda(Evaluator(self.book, self.sheet, self.read), tree[2])
        return made

    LAMBDA_DEPTH = 200  # Excel goes on to some thousands (5,000 deep, not 10,000: recorded)

    def invoke(self, lam, args):
        """A LAMBDA called with these argument nodes, each worked out here: a reference stays the cells it names."""
        values = []
        for node in args:
            try:
                values.append(_MISSING if node[0] == "missing" else self.keep(node))
            except _Err as e:
                values.append(e.err)
        return self.apply(lam, values)

    def apply(self, lam, values):
        """A LAMBDA worked out for these values: more of them than it has parameters, or fewer than those not in [ ],
        is #VALUE!; a parameter in [ ] left out or left empty is omitted, any other left empty is blank (recorded)."""
        if not isinstance(lam, Lambda):
            raise VBANotImplemented("calling what is not a LAMBDA")
        params = lam.params
        if len(values) > len(params) or any(not optional for _, optional in params[len(values) :]):
            raise _Err(XLError("#VALUE!"))
        scope = dict(lam.scope)
        for i, (name, optional) in enumerate(params):
            v = values[i] if i < len(values) else _MISSING
            scope[name] = (_OMITTED if optional else None) if v is _MISSING else v
        if self.depth >= self.LAMBDA_DEPTH:
            raise VBANotImplemented(f"a LAMBDA calling itself more than {self.LAMBDA_DEPTH} deep")
        saved, self.scope = self.scope, scope
        self.depth += 1
        try:
            return self.ev(lam.body)
        except RecursionError as err:
            raise VBANotImplemented("a LAMBDA calling itself deeper than MacroPress can follow") from err
        finally:
            self.scope = saved
            self.depth -= 1

    def ev(self, node):
        kind = node[0]
        if kind in ("num", "str", "bool", "const"):
            return node[1]
        if kind == "err":
            raise _Err(node[1])
        if kind == "missing":
            return None
        if kind == "array":
            return Array([list(row) for row in node[1]])
        key = self.bound_key(node)
        if key is not None:
            return self.bound(key)
        if kind == "name":
            lam = self.named_lambda(node[1])
            if lam is not None:
                return lam
        if kind == "call":
            return self.invoke(self.ev(node[1]), node[2])
        if kind == "opt":
            raise VBANotImplemented(f"[{node[1]}] outside the parameters of a LAMBDA")
        if kind in ("ref", "name", "range"):
            rect = self.rect_of(node)
            sheet, r1, c1, r2, c2 = rect
            if (r1, c1) == (r2, c2):
                return self.read(sheet, r1, c1)
            return self.range_values(rect)
        if kind == "paren":
            return self.ev(node[1])
        if kind == "at":
            return self.one(self.keep(node[1]))
        if kind in ("neg", "pct"):
            v = self.operand(node[1])
            fn = (lambda x: -num(x)) if kind == "neg" else (lambda x: num(x) / 100)
            return _each(fn, [v]) if _many(v) else fn(single(v))
        if kind == "bin":
            a, b = self.operand(node[2]), self.operand(node[3])
            if _many(a) or _many(b):
                return _each(lambda x, y, op=node[1]: _binary(op, x, y), [a, b])
            return _binary(node[1], single(a), single(b))
        if kind == "func":
            return self.call(node[1], node[2])
        raise VBANotImplemented(f"formula part {kind}")

    def keep(self, node):
        """An argument as a function that can hand back cells hands it back: a reference stays the cells it names,
        one cell too (IF(c, A1) is the cell A1, which SUM reads as a cell)."""
        while node[0] == "paren":
            node = node[1]
        key = self.bound_key(node)
        if key is not None:
            v = self.scope[key]
            return None if v is _OMITTED else v
        if node[0] == "name" and self.named_lambda(node[1]) is not None:
            return self.named_lambda(node[1])
        if node[0] in ("ref", "name", "range"):
            return self.cells(self.rect_of(node))
        return self.ev(node)

    def cells(self, rect):
        """The cells of a rectangle as a function hands them back: read, and still the cells they are."""
        if rect[1:3] == rect[3:5]:
            return Array([[self.read(rect[0], rect[1], rect[2])]], (1, 1), rect)
        return self.range_values(rect)

    def taken(self, node):
        """An argument a function takes whole whatever the formula (INDEX's array, LOOKUP's and XLOOKUP's): the cells
        a reference names, kept as cells; anything else worked out as an array."""
        while node[0] == "paren":
            node = node[1]
        if node[0] in ("ref", "name", "range"):
            return self.keep(node)
        return self.array(self.whole(node))

    def one(self, v):
        """One value of what was given: a range gives the cell in the formula's own row (of a column) or column (of a
        row), and #VALUE! when it passes the formula by (implicit intersection); an array gives its first value."""
        if not isinstance(v, Array):
            return v
        if v.extent == (1, 1) or v.rect is None:
            return v.cell(0, 0)
        sheet, r1, c1, r2, c2 = v.rect
        r, c = self.here()
        if c1 == c2 and r1 <= r <= r2:
            return self.read(sheet, r, c1)
        if r1 == r2 and c1 <= c <= c2:
            return self.read(sheet, r1, c)
        raise _Err(XLError("#VALUE!"))

    def operand(self, node):
        """What an operator or a function that wants one value works on: in a formula that is not a dynamic array, a
        range gives one cell (see one); in a dynamic array it is an array. An array stays one either way."""
        v = self.ev(node)
        if isinstance(v, Array) and v.rect is not None and not self.arrays:
            return self.one(v)
        return v

    def taken_whole(self, node):
        """A value for a function that came from the Analysis ToolPak (TAKES "w"): in a formula that is not a dynamic
        array a range is taken whole, and several cells are #VALUE!; an array gives its first value."""
        v = self.ev(node)
        if isinstance(v, Array) and not self.arrays:
            if v.rect is not None and v.extent != (1, 1):
                return XLError("#VALUE!")
            return v.cell(0, 0)
        return v

    def whole(self, node):
        """An argument worked out as an array, whatever the formula: SUMPRODUCT's."""
        saved, self.arrays = self.arrays, True
        try:
            return self.ev(node)
        finally:
            self.arrays = saved

    def call(self, name, args):
        """A function, its arguments taken as TAKES says: those that want one value are worked out first, and where
        one is an array the function is worked out for each of its values, an array of results."""
        if name in self.scope:  # a LAMBDA bound by LET or LAMBDA, called: f(1)
            return self.invoke(self.scope[name], args)
        fn = FUNCTIONS.get(name)
        if fn is None:
            lam = self.named_lambda(name)
            if lam is not None:
                return self.invoke(lam, args)
            raise VBANotImplemented(f"formula function {name}")
        takes = TAKES.get(name, "r")
        given, arrays = list(args), {}
        for i, a in enumerate(args):
            letter = takes[min(i, len(takes) - 1)]
            if letter not in "vwo" or a[0] == "missing":
                continue
            try:
                v = self.taken_whole(a) if letter == "w" else self.operand(a)
            except _Err as e:
                v = e.err
            if letter == "o":
                given[i] = ("const", v)
                continue
            if _many(v):
                arrays[i] = v
            else:
                given[i] = ("const", single(v))
        if not arrays:
            return fn(self, given)
        order = sorted(arrays)

        def each(*values):
            for i, v in zip(order, values, strict=True):
                given[i] = ("const", v)
            return fn(self, given)

        return _each(each, [arrays[i] for i in order])

    def arg(self, args, i, default=None, blank=None):
        """Argument i: default when it is left out (IF(c, x)), blank when it is there but empty (IF(c, , x)). The
        older functions read an empty argument as the blank value, None, so 0, FALSE or ""; the newer ones (SORT,
        XLOOKUP) as left out."""
        if i >= len(args):
            return default
        if args[i][0] == "missing":
            return blank
        return self.ev(args[i])

    def array(self, v):
        if isinstance(v, Array):
            return v
        return Array([[v]])


# How each function takes its arguments, a letter each, the last for the rest: v, one value (in a formula that is
# not a dynamic array, a range gives the formula's own row or column of it; an array, and a range in a dynamic array,
# is gone through value by value, the result an array); r, a reference or an array, taken whole; a, the same, and an
# expression there is worked out as an array whatever the formula (INDEX((A1:A9=1)*B1:B9, 0) needs no Ctrl+Shift+Enter,
# MATCH(TRUE, A1:A9>1, 0) does); w, one value for the functions that came from the Analysis ToolPak, which take a
# range whole and are #VALUE! given several cells (library.py has them); o, one value of a range, an array kept
# whole (TYPE's). A function not here takes all as r; IF, CHOOSE and SUMPRODUCT take theirs themselves.
TAKES = {
    "ABS": "v",
    "INT": "v",
    "ROUND": "v",
    "LEN": "v",
    "LEFT": "v",
    "RIGHT": "v",
    "MID": "v",
    "TRIM": "v",
    "UPPER": "v",
    "LOWER": "v",
    "NOT": "v",
    "ISBLANK": "v",
    "ISNUMBER": "v",
    "ISTEXT": "v",
    "IFERROR": "v",
    "CONCATENATE": "v",
    "MATCH": "vrv",
    "VLOOKUP": "vrvv",
    "HLOOKUP": "vrvv",
    "XLOOKUP": "vaarvv",
    "XMATCH": "vavv",
    "LOOKUP": "vaa",
    "INDEX": "avvv",
    "OFFSET": "rvvvv",
    "INDIRECT": "vv",
    "COUNTIF": "rv",
    "SUMIF": "rvr",
    "AVERAGEIF": "rvr",
    "COUNTIFS": "rv" * 127,
    "SUMIFS": "r" + "rv" * 127,
    "AVERAGEIFS": "r" + "rv" * 127,
    "MAXIFS": "r" + "rv" * 127,
    "MINIFS": "r" + "rv" * 127,
}


def _flat(ev, args):
    out = []
    for a in args:
        if a[0] == "missing":
            continue
        v = ev.ev(a)
        if isinstance(v, Array):
            out += list(v.values())
        else:
            out.append(v)
    return out


# A function that may hand back a value picked from an array, which SUM and the like might read either as a value
# written in the formula or as one read from cells. (IF, INDEX and XLOOKUP hand back the cells of a reference as
# cells, an array's value as written; checked for IF and INDEX.)
_PICKS_CELLS = ("XLOOKUP",)


def _operands(ev, args, name):
    """What SUM, COUNT, AND and the like are given, as (value, written, how many): written is True for a value written
    or worked out in the formula itself (SUM("5",TRUE) is 6), False for one read from cells or an array, where these
    functions pass over text and logical values, and None where a function picked it and either could be so. How
    many is 1 but for the alike rows of an array past its data (see Array.counted)."""
    out = []
    for a in args:
        if a[0] == "missing":  # SUM(1,,2): an argument left empty is a 0 written there, FALSE to AND and OR
            out.append((0.0, True, 1))
            continue
        try:
            v = ev.ev(a)
        except _Err as e:
            v = e.err
        if isinstance(v, Array):
            out += [(x, False, n) for x, n in v.counted()]
            continue
        while a[0] == "paren":
            a = a[1]
        cells = True if a[0] in ("ref", "name") else None if a[0] == "func" and a[1] in _PICKS_CELLS else False
        out.append((v, None if cells is None else not cells, 1))
    return out


def _unsure(name, v):
    return VBANotImplemented(f"{name} of {v!r} from a function that may have read it from a cell")


def _numbers(ev, args, name, counting=False):
    """The numbers among a function's operands, each with how many cells hold it. Text written in the formula counts
    if it reads as a number and is #VALUE! if not; logical values written there count as 1 and 0. COUNT passes over
    errors and text it cannot read."""
    out = []
    for v, written, n in _operands(ev, args, name):
        if isinstance(v, XLError):
            if not counting:
                raise _Err(v)
        elif isinstance(v, (bool, str)):
            if written is None:
                raise _unsure(name, v)
            if not written:
                continue
            x = float(v) if isinstance(v, bool) else entry.value_of(v)
            if x is not None:
                out.append((x, n))
            elif not counting:
                raise _Err(XLError("#VALUE!"))
        elif isinstance(v, (int, float)):
            out.append((float(v), n))
    return out


def total(numbers):
    """The numbers added up as SUM adds them: in order, the last addition as snap_add has it. A number may come with
    how many cells hold it, (x, n): added n times over, by one multiplication where that gives the same (0, a whole
    number)."""
    pairs = [x if isinstance(x, tuple) else (x, 1) for x in numbers]
    if not pairs:
        return 0.0
    last, n_last = pairs[-1]
    pairs[-1] = (last, n_last - 1)
    total = 0.0
    for x, n in pairs:
        if n > 1 and float(x).is_integer() and abs(x * n) < 2**53:
            total += x * n
        else:
            for _ in range(n):
                total += x
    return finite(snap_add(total, last))


def f_sum(ev, args):
    return total(_numbers(ev, args, "SUM"))


def f_min(ev, args):
    n = _numbers(ev, args, "MIN")
    return min(x for x, _ in n) if n else 0.0


def f_max(ev, args):
    n = _numbers(ev, args, "MAX")
    return max(x for x, _ in n) if n else 0.0


def f_count(ev, args):
    return float(sum(n for _, n in _numbers(ev, args, "COUNT", counting=True)))


def f_counta(ev, args):
    return float(sum(n for v, _, n in _operands(ev, args, "COUNTA") if v is not None))


def f_average(ev, args):
    numbers = _numbers(ev, args, "AVERAGE")
    if not numbers:
        raise _Err(XLError("#DIV/0!"))
    return total(numbers) / sum(n for _, n in numbers)


def f_if(ev, args):
    """IF: its condition is one value (see Evaluator.operand), or an array gone through value by value. What it hands
    back it hands back as it is: a reference stays the cells it names, so SUM(IF(c, A1:A5)) adds them all up."""
    cond = ev.operand(args[0])
    if _many(cond):
        then = ev.ev(args[1]) if len(args) > 1 and args[1][0] != "missing" else (None if len(args) > 1 else True)
        other = ev.ev(args[2]) if len(args) > 2 and args[2][0] != "missing" else (None if len(args) > 2 else False)
        return _each(lambda c, t, f: t if truth(c) else f, [cond, then, other])
    i = 1 if truth(single(cond)) else 2
    if i >= len(args):
        return i == 1
    return None if args[i][0] == "missing" else ev.keep(args[i])


def f_iferror(ev, args):
    try:
        v = ev.ev(args[0])
    except _Err:
        return ev.arg(args, 1, "")
    if isinstance(v, XLError):
        return ev.arg(args, 1, "")
    return v


def _logicals(ev, args, name):
    """The logical values among AND's or OR's operands: numbers count (0 is FALSE), and so does "TRUE" or "FALSE"
    written in the formula; other text and blank cells are passed over. With none at all the result is #VALUE!."""
    out = []
    for v, written, _ in _operands(ev, args, name):
        if isinstance(v, XLError):
            raise _Err(v)
        if isinstance(v, str):
            if v.upper() not in ("TRUE", "FALSE") or written is False:
                continue
            if written is None:
                raise _unsure(name, v)
            out.append(v.upper() == "TRUE")
        elif v is not None:
            out.append(bool(v))
    if not out:
        raise _Err(XLError("#VALUE!"))
    return out


def f_and(ev, args):
    return all(_logicals(ev, args, "AND"))


def f_or(ev, args):
    return any(_logicals(ev, args, "OR"))


def f_not(ev, args):
    return not truth(ev.ev(args[0]))


def f_concat(ev, args):
    return "".join(text(v) for v in _flat(ev, args))


def f_concatenate(ev, args):
    return "".join(text(ev.ev(a)) for a in args)


def f_len(ev, args):
    return float(len(text(ev.ev(args[0]))))


def f_left(ev, args):
    n = int(num(ev.arg(args, 1, 1.0)))
    if n < 0:
        raise _Err(XLError("#VALUE!"))
    return text(ev.ev(args[0]))[:n]


def f_right(ev, args):
    n = int(num(ev.arg(args, 1, 1.0)))
    if n < 0:
        raise _Err(XLError("#VALUE!"))
    s = text(ev.ev(args[0]))
    return s[max(len(s) - n, 0) :]


def f_mid(ev, args):
    s = text(ev.ev(args[0]))
    start, n = int(num(ev.ev(args[1]))), int(num(ev.ev(args[2])))
    if start < 1 or n < 0:
        raise _Err(XLError("#VALUE!"))
    return s[start - 1 : start - 1 + n]


def f_trim(ev, args):
    return re.sub(" +", " ", text(ev.ev(args[0])).strip(" "))  # Excel's TRIM also collapses inner runs of spaces


def f_upper(ev, args):
    return text(ev.ev(args[0])).upper()


def f_lower(ev, args):
    return text(ev.ev(args[0])).lower()


def f_round(ev, args):
    return round_half_up(fifteen(num(ev.ev(args[0]))), int(num(ev.arg(args, 1, 0.0))))


def f_abs(ev, args):
    return abs(num(ev.ev(args[0])))


def f_int(ev, args):
    return float(math.floor(fifteen(num(ev.ev(args[0])))))  # INT(4.35*100) is 435, though the product falls short


def f_isblank(ev, args):
    return scalar(ev.ev(args[0])) is None


def f_isnumber(ev, args):
    v = scalar(ev.ev(args[0]))
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def f_istext(ev, args):
    return isinstance(scalar(ev.ev(args[0])), str)


def f_now(ev, args):
    now = ev.book.now or dt.datetime.now()
    return date_to_serial(now)


def f_today(ev, args):
    now = ev.book.now or dt.datetime.now()
    return float(date_to_serial(dt.datetime(now.year, now.month, now.day)))


# ---------------------------------------------------------------------------
# Lookups
# ---------------------------------------------------------------------------


def _runs(arr):
    """A row or column's values for a lookup, as runs (first position, last position, value), 1-based: one for each
    value of its data, and one for the alike rows past it; blank cells, which no lookup finds, left out. None for an
    array of more than one row and column."""
    rows, cols = arr.extent
    if rows != 1 and cols != 1:
        return None
    if cols == 1:
        held = [row[0] if row else None for row in arr.rows[:rows]]
        tail = arr.tail[0] if arr.tail else None
    else:
        held, tail = list(arr.rows[0][:cols]) if arr.rows else [], None
    runs = [(i, i, v) for i, v in enumerate(held, 1) if v is not None]
    if tail is not None and len(held) < max(rows, cols):
        runs.append((len(held) + 1, max(rows, cols), tail))
    return runs


def _key(v, limit=False):
    """The value looked up. An error is the result; with limit (VLOOKUP, HLOOKUP, MATCH), text longer than 255
    characters is #VALUE!."""
    v = single(v)
    if isinstance(v, XLError):
        raise _Err(v)
    if isinstance(v, dt.datetime):
        return date_to_serial(v)
    if limit and isinstance(v, str) and len(v) > 255:
        raise _Err(XLError("#VALUE!"))
    return v


def _of_kind(runs, key):
    """The runs whose values are of the key's kind (a number, text, TRUE or FALSE), with their values as compared."""
    kind = _kind(key)[0]
    return [(a, b, x) for a, b, v in runs for k, x in [_kind(v)] if k == kind]


def _cmp(x, key):
    return compare(x, key, exact=True)  # a lookup compares numbers in full, not to 15 digits


def _in_order(runs, descending=False):
    return all(
        (_cmp(a[2], b[2]) >= 0) if descending else (_cmp(a[2], b[2]) <= 0) for a, b in zip(runs, runs[1:], strict=False)
    )


def _unordered(what):
    return VBANotImplemented(f"{what} over values not in order: Excel's answer is wherever its halving search lands")


def _exact(runs, key, wildcards, reverse=False):
    """The position of the first value (the last, reverse) equal to the key: numbers in full, text ignoring case,
    with the wildcards * ? ~ where the lookup takes them, TRUE and FALSE themselves. A number never finds text that
    reads as one, and nothing finds an error."""
    pattern = None
    if wildcards and isinstance(key, str):
        from .excel.filters import wildcard_regex  # noqa: PLC0415 - the excel package imports this module

        pattern = wildcard_regex(key)
    found = _of_kind(runs, key)
    for first, last, x in reversed(found) if reverse else found:
        if pattern is not None:
            if pattern.fullmatch(x) is not None:
                return last if reverse else first
        elif _cmp(x, key) == 0:
            return last if reverse else first
    return None


def _ascending(runs, key):
    """An approximate match (MATCH 1, VLOOKUP TRUE, LOOKUP): the last value at or below the key, last of equal ones,
    among values of its kind, others passed over. Over values not in ascending order, answered only where Excel's
    search cannot land elsewhere: the key at or above them all (the last) or below them all (none)."""
    found = _of_kind(runs, key)
    below = [r for r in found if _cmp(r[2], key) <= 0]
    if len(below) in (0, len(found)) or _in_order(found):
        return below[-1][1] if below else None
    raise _unordered("an approximate match")


def _descending(runs, key):
    """MATCH -1, over values in descending order: the first value equal to the key, else the last above it."""
    found = _of_kind(runs, key)
    above = [r for r in found if _cmp(r[2], key) > 0]
    if len(above) == len(found):
        return found[-1][1] if found else None
    if not above and not any(_cmp(r[2], key) == 0 for r in found):
        return None
    if not _in_order(found, descending=True):
        raise _unordered("a match of type -1")
    equal = [r for r in found if _cmp(r[2], key) == 0]
    return equal[0][0] if equal else above[-1][1]


def _single_place(runs, what):
    """The one position runs of a value hold; refused where it is held in several, and which Excel's search lands on
    is not known."""
    if len(runs) != 1 or runs[0][0] != runs[0][1]:
        raise VBANotImplemented(f"{what} where the value found is there more than once")
    return runs[0][0]


def _nearest(runs, key, mode, reverse):
    """XLOOKUP's and XMATCH's exact match or, by mode, the next smaller (-1) or next larger (1) value, searching
    first to last (last to first, reverse); of equal ones the first the search meets."""
    found = _of_kind(runs, key)
    if len(found) != len(runs):
        raise VBANotImplemented("a next smaller or larger value looked for among values of other kinds")
    hit = _exact(runs, key, False, reverse)
    if hit is not None:
        return hit
    side = [r for r in found if (_cmp(r[2], key) < 0 if mode == -1 else _cmp(r[2], key) > 0)]
    if not side:
        return None
    best = side[0][2]
    for r in side:
        if (_cmp(r[2], best) > 0) if mode == -1 else (_cmp(r[2], best) < 0):
            best = r[2]
    ties = [r for r in side if _cmp(r[2], best) == 0]
    return ties[-1][1] if reverse else ties[0][0]


def _halving(runs, key, mode, descending):
    """XLOOKUP's and XMATCH's binary search (search mode 2, or -2 over values in descending order). Of equal values
    the exact match finds the first (the last, descending); the next smaller ascending is the last of its equals,
    the next larger the first. Where Excel's choice among equal values is not known, refused."""
    found = _of_kind(runs, key)
    if not _in_order(found, descending):
        raise _unordered("a binary search")
    equal = [r for r in found if _cmp(r[2], key) == 0]
    if mode == 0 or (equal and mode in (-1, 1)):
        if not equal:
            return None
        if mode == 0:
            return equal[-1][1] if descending else equal[0][0]
        return _single_place(equal, "a binary search for the exact or the next value")
    side = [r for r in found if (_cmp(r[2], key) < 0 if mode == -1 else _cmp(r[2], key) > 0)]
    if not side:
        return None
    if descending:  # in descending order the next larger value is the last of those above, the next smaller the first
        near = side[-1] if mode == 1 else side[0]
        return _single_place([r for r in side if _cmp(r[2], near[2]) == 0], "a binary search for the next value")
    return side[-1][1] if mode == -1 else side[0][0]


def _looked_up(runs, key, how):
    """how(runs, key), the position found or None. A blank looked up is read as 0 by some accounts and finds nothing
    by others: answered only where the two agree."""
    if key is None:
        if how(runs, 0.0) is not None:
            raise VBANotImplemented("a blank cell looked up, where a 0 would be found")
        return None
    return how(runs, key)


def _column_of(arr, j):
    """Column j of an array, as an array of one column (a range read keeps its alike rows past the data)."""
    return Array([row[j : j + 1] for row in arr.rows], (arr.extent[0], 1), None, arr.tail and arr.tail[j : j + 1])


def _table_lookup(ev, args, across):
    """VLOOKUP and HLOOKUP: the key looked for in the table's first column (first row, across), and the value in the
    column (row) asked for. A column of 0 or less is #VALUE!, past the table #REF!. An approximate match unless the
    last argument says not, or is there but empty."""
    key = _key(ev.ev(args[0]), limit=True)
    table = ev.array(ev.ev(args[1]))
    index = int(num(ev.ev(args[2])))
    approximate = ev.arg(args, 3, True)
    if index < 1:
        raise _Err(XLError("#VALUE!"))
    if index > table.extent[0 if across else 1]:
        raise _Err(XLError("#REF!"))
    first = Array(table.rows[:1], (1, table.extent[1])) if across else _column_of(table, 0)
    exact = approximate is None or not truth(approximate)
    found = _looked_up(_runs(first), key, (lambda runs, k: _exact(runs, k, True)) if exact else _ascending)
    if found is None:
        raise _Err(XLError("#N/A"))
    return table.cell(index - 1, found - 1) if across else table.cell(found - 1, index - 1)


def f_vlookup(ev, args):
    return _table_lookup(ev, args, across=False)


def f_hlookup(ev, args):
    return _table_lookup(ev, args, across=True)


def f_match(ev, args):
    """MATCH(key, array, type): the key's position in one row or column (#N/A in a block). Type 1 (left out) is an
    approximate match over ascending values, 0 exact with wildcards, -1 over descending ones; any other positive or
    negative number counts as 1 or -1."""
    key = _key(ev.ev(args[0]), limit=True)
    runs = _runs(ev.array(ev.ev(args[1])))
    kind = num(ev.arg(args, 2, 1.0))
    if runs is None:
        raise _Err(XLError("#N/A"))
    how = _ascending if kind > 0 else _descending if kind < 0 else (lambda runs, k: _exact(runs, k, True))
    found = _looked_up(runs, key, how)
    if found is None:
        raise _Err(XLError("#N/A"))
    return float(found)


def _modes(ev, args, i):
    """XLOOKUP's and XMATCH's match mode and search mode, from argument i: anything else is #VALUE!."""
    mode = num(ev.arg(args, i, 0.0, blank=0.0))
    search = num(ev.arg(args, i + 1, 1.0, blank=1.0))
    if mode not in (0, -1, 1, 2) or search not in (1, -1, 2, -2):
        raise _Err(XLError("#VALUE!"))
    return int(mode), int(search)


def _searched(runs, key, mode, search):
    if search in (2, -2):
        if mode == 2:
            raise VBANotImplemented("a binary search with wildcards")
        return _halving(runs, key, mode, search == -2)
    if mode in (0, 2):
        return _exact(runs, key, mode == 2, search == -1)
    return _nearest(runs, key, mode, search == -1)


def f_xlookup(ev, args):
    """XLOOKUP(key, look, give, [if not found], [match mode], [search mode]): the key found in a row or column, and
    the row or column of give in its place, as cells. Give of another length than look is #VALUE!."""
    key = _key(ev.ev(args[0]))
    look, give = ev.taken(args[1]), ev.taken(args[2])
    mode, search = _modes(ev, args, 4)
    runs = _runs(look)
    by_row = look.extent[1] == 1 and (look.extent[0] != 1 or give.extent[0] == 1)
    if runs is None or (give.extent[0] if by_row else give.extent[1]) != max(look.extent):
        raise _Err(XLError("#VALUE!"))
    found = _looked_up(runs, key, lambda runs, k: _searched(runs, k, mode, search))
    if found is None:
        if len(args) > 3 and args[3][0] != "missing":
            return ev.ev(args[3])
        raise _Err(XLError("#N/A"))
    return _slice(give, found - 1, by_row)


def _slice(arr, i, by_row):
    """Row i of an array (column i, by_row False), still the cells it was where it is a range; one cell, the cell."""
    rows, cols = arr.extent
    if by_row and cols == 1:
        return _cell_of(arr, i, 0)
    if not by_row and rows == 1:
        return _cell_of(arr, 0, i)
    if by_row:
        rect = arr.rect and (arr.rect[0], arr.rect[1] + i, arr.rect[2], arr.rect[1] + i, arr.rect[4])
        return Array([[arr.cell(i, j) for j in range(cols)]], (1, cols), rect)
    rect = arr.rect and (arr.rect[0], arr.rect[1], arr.rect[2] + i, arr.rect[3], arr.rect[2] + i)
    tail = arr.tail and arr.tail[i : i + 1]
    return Array([[row[i] if i < len(row) else None] for row in arr.rows], (rows, 1), rect, tail)


def f_xmatch(ev, args):
    """XMATCH(key, look, [match mode], [search mode]): the key's position in a row or column (#VALUE! in a block)."""
    key = _key(ev.ev(args[0]))
    runs = _runs(ev.taken(args[1]))
    mode, search = _modes(ev, args, 2)
    if runs is None:
        raise _Err(XLError("#VALUE!"))
    found = _looked_up(runs, key, lambda runs, k: _searched(runs, k, mode, search))
    if found is None:
        raise _Err(XLError("#N/A"))
    return float(found)


def f_lookup(ev, args):
    """LOOKUP(key, look, [give]): an approximate match, always. With give, the value in its place there; without,
    look is a block searched along its first row when wider than tall (else its first column), giving its last row
    (last column)."""
    key = _key(ev.ev(args[0]))
    if isinstance(key, str) and len(key) > 255:
        raise VBANotImplemented("LOOKUP of text longer than 255 characters")
    look = ev.taken(args[1])
    rows, cols = look.extent
    if len(args) > 2 and args[2][0] != "missing":
        give = ev.taken(args[2])
        runs = _runs(look)
        if runs is None or _runs(give) is None:
            raise VBANotImplemented("LOOKUP in a block given a result vector")
        found = _looked_up(runs, key, _ascending)
        if found is not None and found > max(give.extent):
            raise VBANotImplemented("LOOKUP past the end of its result vector")
        if found is None:
            raise _Err(XLError("#N/A"))
        return give.cell(found - 1, 0) if give.extent[1] == 1 else give.cell(0, found - 1)
    across = cols > rows
    first = Array(look.rows[:1], (1, cols)) if across else _column_of(look, 0)
    found = _looked_up(_runs(first), key, _ascending)
    if found is None:
        raise _Err(XLError("#N/A"))
    return look.cell(rows - 1, found - 1) if across else look.cell(found - 1, cols - 1)


def f_choose(ev, args):
    """CHOOSE(index, value, ...): the value the index (a whole number, cut toward 0) picks, the cells themselves
    where it is a reference; outside 1 to the number of values, #VALUE!. An index that is an array picks value by
    value, the values worked out whole: CHOOSE({1,2}, B1:B9, A1:A9) is the two columns side by side."""
    index, values = ev.operand(args[0]), args[1:]

    def pick(i, *given):
        k = num(i)
        if not 1 <= k < len(given) + 1:
            raise _Err(XLError("#VALUE!"))
        return given[int(k) - 1]

    if _many(index):
        return _each(pick, [index, *[ev.whole(a) if a[0] != "missing" else None for a in values]])
    chosen = pick(single(index), *values)
    return None if chosen[0] == "missing" else ev.keep(chosen)


def f_offset(ev, args):
    """OFFSET(reference, rows, columns, [height], [width]): the cells that many rows down and columns across (cut
    toward 0), as tall and wide as asked or as the reference; a negative height or width reaches up or left. Of no
    height or width, or off the sheet, #REF!."""
    base = ev.keep(args[0])
    if not isinstance(base, Array) or base.rect is None:
        raise _Err(XLError("#VALUE!"))
    sheet, r1, c1, r2, c2 = base.rect

    def whole(i, default):
        x = num(ev.arg(args, i, default, blank=default))
        if x < 0 and x != int(x):
            raise VBANotImplemented("OFFSET by a negative number that is not whole")
        return int(x)

    rows, cols = whole(1, 0.0), whole(2, 0.0)
    height, width = whole(3, float(r2 - r1 + 1)), whole(4, float(c2 - c1 + 1))
    if height == 0 or width == 0:
        raise _Err(XLError("#REF!"))
    top, left = r1 + rows, c1 + cols
    t1, t2 = (top, top + height - 1) if height > 0 else (top + height + 1, top)
    u1, u2 = (left, left + width - 1) if width > 0 else (left + width + 1, left)
    if t1 < 1 or u1 < 1 or t2 > MAX_ROW or u2 > MAX_COL:
        raise _Err(XLError("#REF!"))
    return ev.cells((sheet, t1, u1, t2, u2))


def f_indirect(ev, args):
    """INDIRECT(text, [a1]): the cells text names, A1 style or (a1 FALSE) R1C1 style from the formula's own cell, or
    a defined name; text that names none is #REF!. Another workbook's cells stop the run: whether Excel has it open
    is not known."""
    from .excel import parse_area  # noqa: PLC0415 - the excel package imports this module
    from .excel.addresses import R1C1Formula  # noqa: PLC0415

    name = text(ev.arg(args, 0, ""))
    a1 = ev.arg(args, 1, True)  # an empty argument is FALSE, as in the older functions
    if name.lstrip("'").startswith("["):
        raise VBANotImplemented(f"INDIRECT({name!r}): another workbook's cells")
    if a1 is None or not truth(a1):
        try:
            name = R1C1Formula(name).a1(*ev.here())
        except VBANotImplemented:
            raise _Err(XLError("#REF!")) from None
    m = _TOKEN.match(name)
    if not name or m is None or m.end() != len(name) or m.lastgroup not in (*REFERENCE_KINDS, "name"):
        raise _Err(XLError("#REF!"))
    if m.lastgroup == "name":
        try:
            return ev.keep(("name", name))
        except _Err:
            raise _Err(XLError("#REF!")) from None
    area = parse_area(ref_parts(name)[1].replace("$", ""))
    if area is None or area[2] > MAX_ROW or area[3] > MAX_COL:
        raise _Err(XLError("#REF!"))
    return ev.keep(("ref", name))


def _cell_of(arr, i, j):
    """The value at (i, j) of an array; of a range, the cell itself, which a function hands back as a cell."""
    v = arr.cell(i, j)
    if arr.rect is None:
        return v
    sheet, r1, c1 = arr.rect[:3]
    return Array([[v]], (1, 1), (sheet, r1 + i, c1 + j, r1 + i, c1 + j))


def f_index(ev, args):
    """INDEX(array, row, column). One number picks along a single row or column, and is #REF! for a block of both;
    0 for the row or the column takes all of them: a whole column or row of the array, still the cells it was."""
    arr = ev.taken(args[0])
    rows, cols = arr.extent
    r = int(num(ev.arg(args, 1, 0.0, blank=0.0)))
    if len(args) > 2:
        c = int(num(ev.arg(args, 2, 0.0, blank=0.0)))
    elif rows == 1:
        r, c = 1, r
    elif cols == 1:
        c = 1
    else:
        raise _Err(XLError("#REF!"))
    if r < 0 or c < 0:
        raise _Err(XLError("#VALUE!"))
    if r > rows or c > cols:
        raise _Err(XLError("#REF!"))
    if r and c:
        return _cell_of(arr, r - 1, c - 1)
    r1, r2 = (r - 1, r - 1) if r else (0, rows - 1)
    c1, c2 = (c - 1, c - 1) if c else (0, cols - 1)
    if (r1, c1) == (r2, c2):
        return _cell_of(arr, r1, c1)
    rect = arr.rect and (arr.rect[0], arr.rect[1] + r1, arr.rect[2] + c1, arr.rect[1] + r2, arr.rect[2] + c2)
    held = range(r1, min(r2, len(arr.rows) - 1) + 1)  # a range read holds its rows down to the sheet's last data
    return Array([[arr.cell(i, j) for j in range(c1, c2 + 1)] for i in held], (r2 - r1 + 1, c2 - c1 + 1), rect)


def f_unique(ev, args):
    """UNIQUE(array, [by_col], [exactly_once]): its rows (or columns) without the ones seen before, text without
    regard to case; exactly_once keeps those that stand once only, and none is #CALC! (recorded)."""
    arr = ev.array(ev.ev(args[0]))
    by_col = ev.arg(args, 1, False)
    once = ev.arg(args, 2, False)
    by_col, once = by_col is not None and truth(by_col), once is not None and truth(once)
    lines = [list(col) for col in zip(*arr.full(), strict=True)] if by_col else [list(row) for row in arr.rows]
    counts, first = {}, []
    for line in lines:
        key = tuple(  # TRUE and 1 are two values (in Python they are equal)
            ("t", v.casefold()) if isinstance(v, str) else ("b", v) if isinstance(v, bool) else ("v", v) for v in line
        )
        if key not in counts:
            counts[key] = 0
            first.append((key, line))
        counts[key] += 1
    out = [line for key, line in first if not once or counts[key] == 1]
    if not out:
        raise _Err(XLError("#CALC!"))
    return Array([list(row) for row in zip(*out, strict=True)] if by_col else out)


def f_sort(ev, args):
    """SORT(array, [index], [order], [by_col]): its rows by column index (or its columns by row index), blanks last
    whichever the order."""
    arr = ev.array(ev.ev(args[0]))
    idx = int(num(ev.arg(args, 1, 1.0, blank=1.0)))
    order = num(ev.arg(args, 2, 1.0, blank=1.0))
    by_col = ev.arg(args, 3, False)
    by_col = by_col is not None and truth(by_col)
    lines = [list(col) for col in zip(*arr.full(), strict=True)] if by_col else arr.rows
    if order not in (1, -1) or not 1 <= idx <= arr.shape[0 if by_col else 1]:
        raise _Err(XLError("#VALUE!"))
    blanks = [r for r in lines if r[idx - 1] is None]
    rest = sorted(
        (r for r in lines if r[idx - 1] is not None), key=lambda r: sort_value_key(r[idx - 1]), reverse=(order == -1)
    )
    out = rest + blanks
    return Array([list(row) for row in zip(*out, strict=True)] if by_col else out)


def f_let(ev, args):
    """LET(name, value, ..., calculation): each name stands for its value in what follows it, a reference for its
    cells. A value that is an error is that error only where its name is used (LET(x,1/0,5) is 5, recorded)."""
    if len(args) < 3 or not len(args) % 2:
        raise VBANotImplemented("LET with other than names and values, and a calculation")
    saved = ev.scope
    try:
        for name, node in zip(args[:-1:2], args[1:-1:2], strict=True):
            key = _param_key(name)
            if key is None:
                raise VBANotImplemented(f"LET naming {name[0]} {name[1]!s}")
            try:
                v = ev.keep(node)
            except _Err as e:
                v = e.err
            ev.scope = {**ev.scope, key: v}
        return ev.ev(args[-1])
    finally:
        ev.scope = saved


def f_lambda(ev, args):
    """LAMBDA(parameter, ..., calculation): a function, called as LAMBDA(x,x*2)(5), by the name LET gives it, or by a
    defined name; given to MAP and the rest. A parameter in [ ] may be left out (ISOMITTED)."""
    if not args:
        raise VBANotImplemented("LAMBDA()")
    params = []
    for node in args[:-1]:
        key = _param_key(node)
        if key is None or key in [name for name, _ in params]:
            raise VBANotImplemented(f"LAMBDA with a parameter {node[0]} {node[1]!s}")
        params.append((key, node[0] == "opt" or (node[0] == "name" and node[1].lower().startswith("_xlop."))))
    return Lambda(params, args[-1], ev.scope)


def f_isomitted(ev, args):
    """ISOMITTED(parameter): whether a LAMBDA's parameter in [ ] was left out, or left empty (recorded)."""
    key = ev.bound_key(args[0]) if args else None
    if key is None:
        raise VBANotImplemented("ISOMITTED of what is not a parameter of a LAMBDA")
    return ev.scope[key] is _OMITTED


def f_sumproduct(ev, args):
    """SUMPRODUCT: its arguments worked out as arrays, whatever the formula, all of one size (else #VALUE!), and the
    products of their values added up. A value that is not a number (text, TRUE, a blank) counts as 0; an error is
    the result. The sum is the plain one: a sum that all but cancels is not rounded to 0, as SUM's is."""
    arrays = []
    for a in args:
        v = ev.whole(a) if a[0] != "missing" else None
        arrays.append(v if isinstance(v, Array) else Array([[v]]))
    rows, cols = arrays[0].extent
    if any(a.extent != (rows, cols) for a in arrays):
        raise _Err(XLError("#VALUE!"))

    def product(i):
        total = 0.0
        for j in range(cols):
            p = 1.0
            for a in arrays:
                v = a.cell(i, j)
                if isinstance(v, XLError):
                    raise _Err(v)
                p *= float(v) if isinstance(v, (int, float)) and not isinstance(v, bool) else 0.0
            total += p
        return total

    start = min(rows, max(len(a.rows) for a in arrays))  # past every array's data, the rows are alike
    total = sum(product(i) for i in range(start))
    if start < rows:
        total += (rows - start) * product(start)
    return finite(total)


# ---------------------------------------------------------------------------
# SUMIF, COUNTIF and the rest: the cells where criteria hold
# ---------------------------------------------------------------------------

_ERROR_TEXTS = ("#N/A", "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NUM!", "#NULL!", "#SPILL!", "#CALC!")
_CRITERION_OPS = ("<>", "<=", ">=", "<", ">", "=")


@functools.lru_cache(maxsize=4096)
def _reads_as(text):
    """The number text reads as, for a criterion and for a cell it looks for a number in."""
    return entry.value_of(text)  # a date, a time or an amount of money is refused: regional settings read those


def _kind(v):
    """A cell's value as a criterion sees it: ("blank" | "num" | "text" | "bool" | "err", value)."""
    if v is None:
        return "blank", None
    if isinstance(v, XLError):
        return "err", v
    if isinstance(v, bool):
        return "bool", v
    if isinstance(v, (int, float)):
        return "num", float(v)
    if isinstance(v, dt.datetime):
        return "num", date_to_serial(v)
    return "text", v


def _operand(text):
    """What follows a criterion's operator: a number, TRUE or FALSE, an error, or text."""
    number = _reads_as(text)
    if number is not None:
        return number
    if text.upper() in ("TRUE", "FALSE"):
        return text.upper() == "TRUE"
    if text in _ERROR_TEXTS:
        return XLError(text)
    if text.strip(" ").upper() in ("TRUE", "FALSE", *_ERROR_TEXTS):
        raise VBANotImplemented(f"the criterion {text!r}: whether Excel reads it as TRUE, FALSE or an error")
    return text


def _same_number(x, want, want15):
    """x and want equal as Excel compares numbers, to 15 significant digits (want15 is fifteen(want)). Numbers that
    far apart cannot round alike, and are not rounded."""
    return x == want or (abs(x - want) <= 1e-13 * max(abs(x), abs(want)) and fifteen(x) == want15)


def criterion(v):
    """What SUMIF, COUNTIF and the rest look for, as a test of a cell's value. As Excel was found to look:

    - A number finds the numbers equal to it (to 15 digits) and the text that reads as one: 5 finds "05", " 5" and
      "5.0". TRUE and FALSE find themselves, not 1 and 0; an error finds the cells holding it. A blank criterion, or
      one left empty, is the number 0.
    - Text may start with =, <>, <, >, <= or >=. What follows is a number, TRUE or FALSE, or an error if it reads as
      one, and text if not: = and <> match text ignoring case, with the wildcards * ? and ~ ("a~*" is a*).
    - <> finds every cell = would not find, but text that reads as the number: "<>5" finds "5". <, >, <= and >= find
      only cells of the criterion's own kind (">4" passes over "5"), comparing text as Excel orders it.
    - "" finds blank cells and empty text, "=" blank cells only, "<>" every cell that is not blank; any other
      operator alone finds nothing.
    - Text longer than 255 characters is #VALUE!."""
    op = "="
    if isinstance(v, str) and not isinstance(v, XLError):
        if len(v) > 255:
            raise _Err(XLError("#VALUE!"))
        op = next((o for o in _CRITERION_OPS if v.startswith(o)), "")
        rest = v[len(op) :]
        if rest == "":
            if op == "":
                return lambda c: c is None or type(c) is str and c == ""
            if op in ("=", "<>"):
                return (lambda c: c is None) if op == "=" else (lambda c: c is not None)
            return lambda c: False
        v, op = _operand(rest), op or "="
    kind, want = _kind(0.0 if v is None else v)
    if kind == "err" and op not in ("=", "<>"):
        raise VBANotImplemented(f"a criterion that compares with {want} by {op}")
    if kind == "num":
        want15 = fifteen(want)

        def same(c):
            t = type(c)
            if t is float or t is int:
                return _same_number(c, want, want15)
            return t is dt.datetime and _same_number(date_to_serial(c), want, want15)

    elif kind == "text":
        from .excel.filters import wildcard_regex  # noqa: PLC0415 - the excel package imports this module

        pattern = wildcard_regex(want)
        plain = want.lower() if want.isascii() and not any(ch in want for ch in "*?~") else None

        def same(c):  # text only: an error, an XLError, is a str too
            if type(c) is not str:
                return False
            if plain is not None and c.isascii():
                return c.lower() == plain
            return pattern.fullmatch(c) is not None

    else:

        def same(c):  # TRUE or FALSE, or an error
            return type(c) is type(want) and c == want

    if op == "<>":
        return lambda c: not same(c)
    if op == "=":
        if kind != "num":
            return same

        def equal(c):
            if type(c) is str:
                x = _reads_as(c)
                return x is not None and _same_number(x, want, want15)
            return same(c)

        return equal

    def ordered(c):
        k, x = _kind(c)
        if k != kind:
            return False
        if kind == "text" and x == "":
            raise VBANotImplemented(f"whether {op}{want} finds empty text")
        r = compare(x, want)
        return {"<": r < 0, ">": r > 0, "<=": r <= 0, ">=": r >= 0}[op]

    return ordered


def _column(a, j, held):
    """Column j of a range read, over its first `held` rows: blank past its data."""
    line = [row[j] if j < len(row) else None for row in a.rows[:held]]
    return line + [None] * (held - len(line))


def where(areas, criteria, target=None):
    """Where every criterion holds over its range: (how many cells, and the values of target's cells there, row by
    row). The ranges and target are all of one size, else #VALUE!. The rows past every one's data are blank, and
    counted at once."""
    tests = [criterion(c) for c in criteria]
    rows, cols = areas[0].extent
    arrays = areas + ([target] if target is not None else [])
    if any(a.extent != (rows, cols) for a in arrays):
        raise _Err(XLError("#VALUE!"))
    held = min(rows, max(len(a.rows) for a in arrays))
    count, found = 0, []
    for j in range(cols):
        hits = range(held)
        for a, test in zip(areas, tests, strict=True):
            line = _column(a, j, held)
            hits = [i for i in hits if test(line[i])]
        count += len(hits)
        if target is not None:
            line = _column(target, j, held)
            found += [(i, j, line[i]) for i in hits]
    if held < rows and all(test(None) for test in tests):
        count += (rows - held) * cols
    if cols > 1:
        found.sort(key=lambda f: f[:2])
    return count, [v for _, _, v in found]


def _numbers_found(values):
    """The numbers among the cells found, in order: text and TRUE pass over, an error is the result."""
    out = []
    for v in values:
        k, x = _kind(v)
        if k == "err":
            raise _Err(x)
        if k == "num":
            out.append(x)
    return out


def sum_found(values):
    """SUMIF's sum: the plain one, a sum that all but cancels is not 0 (as it is in SUM)."""
    return finite(sum(_numbers_found(values), 0.0))


def average_found(values):
    numbers = _numbers_found(values)
    if not numbers:
        raise _Err(XLError("#DIV/0!"))
    return finite(sum(numbers)) / len(numbers)


def max_found(values):
    return max(_numbers_found(values), default=0.0)


def min_found(values):
    return min(_numbers_found(values), default=0.0)


def _area(ev, node):
    """A range argument of SUMIF and the rest: the cells it names, read once for the formula. What is not cells is
    #VALUE! (COUNTIF(IF(TRUE,{1,2}),1)); a constant or an operator there Excel refuses as the formula is entered."""
    if id(node) not in ev.areas:
        v = ev.keep(node)
        if not isinstance(v, Array) or v.rect is None:
            raise _Err(XLError("#VALUE!"))
        ev.areas[id(node)] = v
    return ev.areas[id(node)]


def _resized(ev, node, extent):
    """SUMIF's and AVERAGEIF's sum range: its first cell and on, over the criteria range's size, whatever its own."""
    while node[0] == "paren":
        node = node[1]
    sheet, r1, c1 = (ev.rect_of(node) if node[0] in ("ref", "name") else _area(ev, node).rect)[:3]
    r2, c2 = r1 + extent[0] - 1, c1 + extent[1] - 1
    if r2 > MAX_ROW or c2 > MAX_COL:
        raise VBANotImplemented("a sum range that runs past the edge of the sheet")
    return ev.range_values((sheet, r1, c1, r2, c2))


def unnamed_reads(tree, rect_of):
    """The cells a formula reads past the references it names, for knowing when it is stale: SUMIF's and AVERAGEIF's
    sum range, read over the criteria range's size (see _resized), and a range between references (A1:INDEX(A:A,5)),
    whose cells lie in the box around every reference inside it. [(sheet, r1, c1, r2, c2)], or None where SUMIF's
    ranges are not references the formula names (INDEX(...)), or where it calls or hands on a LAMBDA kept as a defined
    name, and what it reads is known only as it is worked out."""
    out = []
    for node in _all(tree):  # a LAMBDA kept as a defined name, called or handed on: what its calculation reads
        if node[0] == "name" or (node[0] == "func" and node[1] not in FUNCTIONS):
            try:
                rect_of(node if node[0] == "name" else ("name", node[1]))
            except _NotCells:
                return None
            except (_Err, VBANotImplemented):
                pass
    for node in _all(tree):
        if node[0] == "range":
            boxes = {}
            for ref in references(node):
                try:
                    sheet, r1, c1, r2, c2 = rect_of(ref)
                except _Err:
                    continue  # a #REF! or #NAME? reads nothing
                b = boxes.get(id(sheet), (sheet, r1, c1, r2, c2))
                boxes[id(sheet)] = (sheet, min(b[1], r1), min(b[2], c1), max(b[3], r2), max(b[4], c2))
            out += boxes.values()
        if node[0] != "func" or node[1] not in ("SUMIF", "AVERAGEIF") or len(node[2]) != 3:
            continue
        area, given = node[2][0], node[2][2]
        while area[0] == "paren":
            area = area[1]
        while given[0] == "paren":
            given = given[1]
        if area[0] not in ("ref", "name") or given[0] not in ("ref", "name"):
            return None
        try:
            _, r1, c1, r2, c2 = rect_of(area)
            sheet, s1, t1 = rect_of(given)[:3]
        except _Err:
            continue  # a #REF! or #NAME? reads nothing
        out.append((sheet, s1, t1, min(s1 + r2 - r1, MAX_ROW), min(t1 + c2 - c1, MAX_COL)))
    return out


def _found(ev, args, first):
    """COUNTIFS and the like: the (range, criterion) pairs from argument `first` on."""
    pairs = range(first, len(args) - 1, 2)
    return [_area(ev, args[i]) for i in pairs], [ev.ev(args[i + 1]) for i in pairs]


def f_countif(ev, args):
    return float(where(*_found(ev, args, 0))[0])


def _one_criterion(ev, args):
    """SUMIF's and AVERAGEIF's cells found: in their sum range, or in the criteria range itself."""
    area = _area(ev, args[0])
    target = _resized(ev, args[2], area.extent) if len(args) > 2 else area
    return where([area], [ev.ev(args[1])], target)[1]


def f_sumif(ev, args):
    return sum_found(_one_criterion(ev, args))


def f_averageif(ev, args):
    return average_found(_one_criterion(ev, args))


def _many_criteria(ev, args):
    """SUMIFS's and the like's cells found: in their first argument, all their ranges of one size."""
    return where(*_found(ev, args, 1), _area(ev, args[0]))[1]


def f_sumifs(ev, args):
    return sum_found(_many_criteria(ev, args))


def f_averageifs(ev, args):
    return average_found(_many_criteria(ev, args))


def f_maxifs(ev, args):
    return max_found(_many_criteria(ev, args))


def f_minifs(ev, args):
    return min_found(_many_criteria(ev, args))


# How many arguments a function takes, (fewest, most, step), and which must be references, a letter each ("r"): Excel
# refuses as it is entered (error 1004) a formula that gives another number, or anything but a reference or a function
# that may hand back cells where one is wanted (a constant, an operator, SUM(...)).
_ARGUMENTS = {
    **dict.fromkeys(("AND", "OR", "AVERAGE", "COUNT", "COUNTA", "MAX", "MIN", "SUM", "SUMPRODUCT"), (1, 255, 1, "")),
    "CONCAT": (1, 254, 1, ""),
    "IFERROR": (2, 2, 1, ""),
    "ISNUMBER": (1, 1, 1, ""),
    "ISTEXT": (1, 1, 1, ""),
    "ROUND": (2, 2, 1, ""),
    "TRIM": (1, 1, 1, ""),
    "SORT": (1, 4, 1, ""),
    "UNIQUE": (1, 3, 1, ""),
    "TRANSPOSE": (1, 1, 1, ""),
    "COUNTIF": (2, 2, 1, "r"),
    "SUMIF": (2, 3, 1, "r-r"),
    "AVERAGEIF": (2, 3, 1, "r-r"),
    "COUNTIFS": (2, 254, 2, "r-" * 127),
    **{name: (3, 255, 2, "r" + "r-" * 127) for name in ("SUMIFS", "AVERAGEIFS", "MAXIFS", "MINIFS")},
    "OFFSET": (3, 5, 1, "r"),
    "VLOOKUP": (3, 4, 1, ""),
    "HLOOKUP": (3, 4, 1, ""),
    "MATCH": (2, 3, 1, ""),
    "XLOOKUP": (3, 6, 1, ""),
    "XMATCH": (2, 4, 1, ""),
    "LOOKUP": (2, 3, 1, ""),
    "CHOOSE": (2, 255, 1, ""),
    "INDEX": (2, 4, 1, ""),
    "INDIRECT": (1, 2, 1, ""),
}
_GIVES_CELLS = {"IF", "INDEX", "OFFSET", "INDIRECT", "CHOOSE", "XLOOKUP"}


def check_entry(text):
    """Excel's refusals, as a formula is entered, of what it gives SUMIF, the lookups and the rest (see _ARGUMENTS)."""
    from .errors import VBARuntimeError  # noqa: PLC0415

    upper = text.upper()
    if "#SPILL!" not in upper and "#CALC!" not in upper and not any(name in upper for name in _ARGUMENTS):
        return
    try:
        tree = parse(text)
    except VBANotImplemented:
        return
    for node in _all(tree):
        if node[0] == "err" and str(node[1]) in ("#SPILL!", "#CALC!"):  # results only: Excel will not take them typed
            raise VBARuntimeError(1004, "Application-defined or object-defined error")
        if node[0] != "func" or node[1] not in _ARGUMENTS:
            continue
        name, args = node[1], node[2]
        fewest, most, step, wants = _ARGUMENTS[name]
        refused = not fewest <= len(args) <= most or (len(args) - fewest) % step
        if name in ("LET", "LAMBDA") and not refused:  # names, each bound once (a repeated one is refused: recorded)
            keys = [_param_key(a) for a in (args[:-1:2] if name == "LET" else args[:-1])]
            refused = None in keys or len(set(keys)) != len(keys)
        if name == "SORTBY" and len(args) > 3 and not len(args) % 2:  # by_array and sort_order in pairs (recorded)
            refused = True
        for i, a in enumerate(args if not refused else ()):
            while a[0] == "paren":
                a = a[1]
            if i >= len(wants) or wants[i] != "r" or a[0] in ("ref", "name", "range"):
                continue
            if a[0] == "func" and a[1] in _GIVES_CELLS:
                continue
            unknown = a[0] == "func" and (a[1] not in EXCEL_FUNCTIONS or a[1] in _MANY_RESULTS)
            if unknown or a[0] in ("at", "err"):  # a function that may hand back cells, @, #REF!: not checked
                what = f"{a[1]}(...)" if a[0] == "func" else "@" if a[0] == "at" else a[1]
                raise VBANotImplemented(f"{name} given {what} for a range: whether Excel takes it is not known")
            refused = True
        if refused:
            raise VBARuntimeError(1004, "Application-defined or object-defined error")


def _cells_named(ev, node):
    """The cells an argument names, (sheet, r1, c1, r2, c2): a reference, or what a function hands back as cells."""
    while node[0] == "paren":
        node = node[1]
    if node[0] in ("ref", "name", "range"):
        return ev.rect_of(node)
    v = ev.ev(node)
    if isinstance(v, Array) and v.rect is not None:
        return v.rect
    raise VBANotImplemented("ROW or COLUMN of something that is not a reference")


def f_row(ev, args):
    """ROW(): the formula's own row. ROW(reference): its first row, or, of several rows, an array of them, of which a
    formula that is not a dynamic array shows the first."""
    if not args or args[0][0] == "missing":
        return float(ev.here()[0])
    _, r1, _, r2, _ = _cells_named(ev, args[0])
    return float(r1) if r1 == r2 else Array([[float(r)] for r in range(r1, r2 + 1)])


def f_column(ev, args):
    if not args or args[0][0] == "missing":
        return float(ev.here()[1])
    _, _, c1, _, c2 = _cells_named(ev, args[0])
    return float(c1) if c1 == c2 else Array([[float(c) for c in range(c1, c2 + 1)]])


def f_rows(ev, args):
    v = ev.ev(args[0])
    return float(v.extent[0]) if isinstance(v, Array) else 1.0


def f_columns(ev, args):
    v = ev.ev(args[0])
    return float(v.extent[1]) if isinstance(v, Array) else 1.0


def f_single(ev, args):
    """_xlfn.SINGLE, how the file keeps @: one value of what follows."""
    return ev.one(ev.keep(args[0]))


FUNCTIONS = {
    "SUM": f_sum,
    "MIN": f_min,
    "MAX": f_max,
    "COUNT": f_count,
    "COUNTA": f_counta,
    "AVERAGE": f_average,
    "IF": f_if,
    "IFERROR": f_iferror,
    "AND": f_and,
    "OR": f_or,
    "NOT": f_not,
    "CONCATENATE": f_concatenate,
    "CONCAT": f_concat,
    "LEN": f_len,
    "LEFT": f_left,
    "RIGHT": f_right,
    "MID": f_mid,
    "TRIM": f_trim,
    "UPPER": f_upper,
    "LOWER": f_lower,
    "ROUND": f_round,
    "ABS": f_abs,
    "INT": f_int,
    "ISBLANK": f_isblank,
    "ISNUMBER": f_isnumber,
    "ISTEXT": f_istext,
    "NOW": f_now,
    "TODAY": f_today,
    "MATCH": f_match,
    "VLOOKUP": f_vlookup,
    "XLOOKUP": f_xlookup,
    "HLOOKUP": f_hlookup,
    "XMATCH": f_xmatch,
    "LOOKUP": f_lookup,
    "CHOOSE": f_choose,
    "OFFSET": f_offset,
    "INDIRECT": f_indirect,
    "INDEX": f_index,
    "UNIQUE": f_unique,
    "SORT": f_sort,
    "SUMPRODUCT": f_sumproduct,
    "LET": f_let,
    "LAMBDA": f_lambda,
    "ISOMITTED": f_isomitted,
    "ROW": f_row,
    "COLUMN": f_column,
    "ROWS": f_rows,
    "COLUMNS": f_columns,
    "SINGLE": f_single,
    "COUNTIF": f_countif,
    "COUNTIFS": f_countif,
    "SUMIF": f_sumif,
    "SUMIFS": f_sumifs,
    "AVERAGEIF": f_averageif,
    "AVERAGEIFS": f_averageifs,
    "MAXIFS": f_maxifs,
    "MINIFS": f_minifs,
}

from . import arrays as _newer, library  # noqa: E402, F401 - Excel's other functions, which register themselves here
