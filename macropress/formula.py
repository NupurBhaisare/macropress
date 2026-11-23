"""Excel formulas: parse them and evaluate the ones a macro writes or whose inputs it changes.

Covers what macros commonly need: references (A1, $A$1:$B$9, Sheet!A1, 'My sheet'!A:A, 2:5), defined names and table
columns (Tbl[col]), arithmetic, & and comparisons with Excel's (not VBA's) rules, and a set of functions including
the dynamic-array SORT and UNIQUE, whose results spill. Anything else raises VBANotImplemented: a formula is never
given a guessed value.
"""

import datetime as dt
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
|(?P<op><>|<=|>=|[-+*/^&=<>(),%{};@])
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

# Functions newer than Excel 2007 are stored in the file with a prefix (Excel shows #NAME? without it).
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
}


def file_syntax(text):
    """A formula as the file stores it: newer functions get their _xlfn. prefix, and @ is _xlfn.SINGLE()."""
    if "@" in text:
        text = _replace_nodes(text, "at", lambda body, node, spans: f"SINGLE({_text_of(body, node[1], spans)})")
    return "".join(FILE_PREFIX.get(p.upper(), p) if kind == "func" else p for kind, p in pieces(text))


_XLFN = re.compile(r"(?i)^_xlfn\.(?:_xlws\.)?")
_PRIMARY = ("ref", "name", "func", "array", "paren")


def display_syntax(text):
    """A formula as VBA's Formula property shows it: without the file's _xlfn. prefixes, and _xlfn.SINGLE(x) as @x."""
    text = "".join(_XLFN.sub("", p) if kind == "func" else p for kind, p in pieces(text))
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
    if node[0] == "bin":
        return [node[2], node[3]]
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
                 "SINGLE"}  # fmt: skip


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
        if name == "INDEX":  # a row or column of 0, or left empty, takes all of them
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
        if name == "XLOOKUP":
            if _arrayish(args[0]):
                raise _Unsure
            return False
        if name in _MANY_RESULTS:
            raise _Unsure
        return False

    def walk(node, place):
        """place: top (what the formula ends on), value (a function's argument that wants one value), operand (an
        operator's), ref (taken whole), array (worked out as an array)."""
        kind = node[0]
        if kind in ("ref", "name"):
            if place in ("top", "value", "operand") and size(node) != (1, 1):
                sites.append(node)
            return
        if kind == "array":  # an array stays one, but one a formula ends on is shown with @
            if place == "top" and many(node):
                sites.append(node)
            elif place == "value" and many(node):
                raise _Unsure
            return
        if kind in ("neg", "pct", "paren"):
            return walk(node[1], "operand" if place in ("top", "value") and kind != "paren" else place)
        if kind == "at":
            return
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
                if place == "value" or name not in ("INDEX", "IF", "ROW", "COLUMN"):
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
            if name not in FUNCTIONS and name not in EXCEL_FUNCTIONS:
                raise _Unsure
            takes = TAKES.get(name)
            for i, a in enumerate(args):
                letter = takes[min(i, len(takes) - 1)] if takes else None
                if letter is None and a[0] in ("ref", "name") and size(a) != (1, 1) and name not in FUNCTIONS:
                    raise _Unsure  # how a function we do not model takes a range is not known
                walk(a, "value" if letter == "v" else "ref")
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
    for prefix in ("_XLFN._XLWS.", "_XLFN.", "_XLWS."):
        if n.startswith(prefix):
            n = n[len(prefix) :]
    return n


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------


def tokenize(text, offsets=None):
    """The formula's tokens, (kind, text), spaces left out; offsets, a list, is given where each starts."""
    out, pos = [], 0
    for kind, piece in pieces(text):
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
    ("at", a) for @ (one value of what follows), ("missing",) for an argument left empty. `spans` keeps where each
    node starts and ends in the text (after the =)."""

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
        first = self.i
        return self._mark(self._primary(), first)

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
            args = []
            if self.peek() == ("op", ")"):
                self.take()
                return ("func", bare_name(val), args)
            while True:
                if self.peek() in (("op", ","), ("op", ")")):
                    args.append(("missing",))
                else:
                    args.append(self.expr(0))
                k, v = self.take()
                if (k, v) == ("op", ")"):
                    return ("func", bare_name(val), args)
                if (k, v) != ("op", ","):
                    raise VBANotImplemented(f"formula syntax in ={self.text}")
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
        for a in node[2]:
            function_names(a, out)
    elif node[0] == "bin":
        function_names(node[2], out)
        function_names(node[3], out)
    elif node[0] in ("neg", "pct", "paren", "at"):
        function_names(node[1], out)
    return out


def references(node, out=None):
    """Every ref and name token in a parsed formula."""
    out = [] if out is None else out
    if node[0] in ("ref", "name"):
        out.append(node)
    elif node[0] == "func":
        for a in node[2]:
            references(a, out)
    elif node[0] == "bin":
        references(node[2], out)
        references(node[3], out)
    elif node[0] in ("neg", "pct", "paren", "at"):
        references(node[1], out)
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


def scalar(v):
    if isinstance(v, Array):
        if v.extent == (1, 1):
            return v.cell(0, 0)
        raise VBANotImplemented("an array where a formula expects one value")
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
            return fn(*[at(o, i, j) for o in operands])
        except _Err as e:
            return e.err
        except ZeroDivisionError:
            return XLError("#DIV/0!")

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

    def __init__(self, book, sheet, read, cell=None, arrays=False):
        self.book = book
        self.sheet = sheet
        self.read = read  # read(sheet, r, c) -> value (refreshing a stale formula first)
        self.cell = cell
        self.arrays = arrays

    def run(self, node):
        try:
            if node[0] == "bin" and node[1] in ("+", "-"):  # the formula ends on a sum: see snap_add
                a, b = self.operand(node[2]), self.operand(node[3])
                if _many(a) or _many(b):
                    return _each(lambda x, y, op=node[1]: _sum_end(op, x, y), [a, b])
                return _sum_end(node[1], single(a), single(b))
            return self.ev(node)
        except _Err as e:
            return e.err
        except ZeroDivisionError:
            return XLError("#DIV/0!")

    def here(self):
        if self.cell is None:
            raise VBANotImplemented("a formula's own row or column, outside a cell")
        return self.cell

    def rect_of(self, node):
        """(sheet, r1, c1, r2, c2) for a ref or name node."""
        from .excel import parse_area

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

    def range_values(self, rect):
        sheet, r1, c1, r2, c2 = rect
        extent = (r2 - r1 + 1, c2 - c1 + 1)
        b = sheet.used_bounds()
        if b is None:
            return Array([[None]], extent, rect)
        r2 = min(r2, max(b[2], r1))
        c2 = min(c2, max(b[3], c1))
        return Array([[self.read(sheet, r, c) for c in range(c1, c2 + 1)] for r in range(r1, r2 + 1)], extent, rect)

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
        if kind in ("ref", "name"):
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
        if node[0] in ("ref", "name"):
            rect = self.rect_of(node)
            if rect[1:3] == rect[3:5]:
                return Array([[self.read(rect[0], rect[1], rect[2])]], (1, 1), rect)
            return self.range_values(rect)
        return self.ev(node)

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
        fn = FUNCTIONS.get(name)
        if fn is None:
            raise VBANotImplemented(f"formula function {name}")
        takes = TAKES.get(name, "r")
        given, arrays = list(args), {}
        for i, a in enumerate(args):
            if takes[min(i, len(takes) - 1)] != "v" or a[0] == "missing":
                continue
            try:
                v = self.operand(a)
            except _Err as e:
                v = e.err
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
# is gone through value by value, the result an array); r, a reference or an array, taken whole. A function not here
# takes all as r; IF and SUMPRODUCT take theirs themselves.
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
    "VLOOKUP": "vrv",
    "XLOOKUP": "vrrv",
    "INDEX": "rv",
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


def _line(arr):
    """A row or column's values in order, for a lookup: of the alike rows past an array's data, the first (where
    a match there is first found)."""
    values = [x for row in arr.rows for x in row]
    if arr.tail is not None and len(arr.rows) < arr.extent[0] and arr.extent[1] == 1:
        values.append(arr.tail[0])
    return values


def _match_index(values, key):
    for i, v in enumerate(values):
        if v is not None and not isinstance(v, XLError) and compare(v, key, exact=True) == 0:  # no error matches
            return i
    return None


def f_match(ev, args):
    key = scalar(ev.ev(args[0]))
    arr = ev.array(ev.ev(args[1]))
    mode = num(ev.arg(args, 2, 1.0))
    if mode != 0:
        raise VBANotImplemented("MATCH with an approximate match type")
    values = _line(arr)
    i = _match_index(values, key)
    if i is None:
        raise _Err(XLError("#N/A"))
    return float(i + 1)


def f_vlookup(ev, args):
    key = scalar(ev.ev(args[0]))
    arr = ev.array(ev.ev(args[1]))
    col = int(num(ev.ev(args[2])))
    approx = ev.arg(args, 3, True)  # left out: approximate; empty, VLOOKUP(x, r, 2, ): exact
    if approx is not None and truth(approx):
        raise VBANotImplemented("VLOOKUP with an approximate match")
    if col < 1:
        raise _Err(XLError("#VALUE!"))
    if col > arr.extent[1]:
        raise _Err(XLError("#REF!"))
    i = _match_index(
        _line(Array([row[:1] for row in arr.rows], (arr.extent[0], 1), None, arr.tail and arr.tail[:1])), key
    )
    if i is None:
        raise _Err(XLError("#N/A"))
    return arr.cell(i, col - 1)


def f_xlookup(ev, args):
    key = scalar(ev.ev(args[0]))
    look = ev.array(ev.ev(args[1]))
    ret = ev.array(ev.ev(args[2]))
    if len(args) > 4 and args[4][0] != "missing" and num(ev.ev(args[4])) != 0:
        raise VBANotImplemented("XLOOKUP with a match mode")
    if len(args) > 5 and args[5][0] != "missing" and num(ev.ev(args[5])) != 1:
        raise VBANotImplemented("XLOOKUP with a search mode")
    i = _match_index(_line(look), key)
    if i is None:
        if len(args) > 3 and args[3][0] != "missing":
            return ev.ev(args[3])
        raise _Err(XLError("#N/A"))
    r, c = (i, 0) if look.extent[1] == 1 else (0, i)
    return _cell_of(ret, r, c)


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
    arr = ev.array(ev.ev(args[0]))
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
    arr = ev.array(ev.ev(args[0]))
    by_col = ev.arg(args, 1, False)
    once = ev.arg(args, 2, False)
    if (by_col is not None and truth(by_col)) or (once is not None and truth(once)):
        raise VBANotImplemented("UNIQUE by column or exactly once")
    seen, out = set(), []
    for row in arr.rows:
        key = tuple(  # TRUE and 1 are two values (in Python they are equal)
            ("t", v.casefold()) if isinstance(v, str) else ("b", v) if isinstance(v, bool) else ("v", v) for v in row
        )
        if key not in seen:
            seen.add(key)
            out.append(list(row))
    return Array(out)


def f_sort(ev, args):
    arr = ev.array(ev.ev(args[0]))
    idx = int(num(ev.arg(args, 1, 1.0, blank=1.0)))
    order = num(ev.arg(args, 2, 1.0, blank=1.0))
    by_col = ev.arg(args, 3, False)
    if by_col is not None and truth(by_col):
        raise VBANotImplemented("SORT by column")
    if order not in (1, -1) or not 1 <= idx <= arr.shape[1]:
        raise _Err(XLError("#VALUE!"))
    blanks = [r for r in arr.rows if r[idx - 1] is None]
    rest = sorted(
        (r for r in arr.rows if r[idx - 1] is not None), key=lambda r: sort_value_key(r[idx - 1]), reverse=(order == -1)
    )
    return Array(rest + blanks)


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


def _cells_named(ev, node):
    """The cells an argument names, (sheet, r1, c1, r2, c2): a reference, or what a function hands back as cells."""
    while node[0] == "paren":
        node = node[1]
    if node[0] in ("ref", "name"):
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
    "INDEX": f_index,
    "UNIQUE": f_unique,
    "SORT": f_sort,
    "SUMPRODUCT": f_sumproduct,
    "ROW": f_row,
    "COLUMN": f_column,
    "ROWS": f_rows,
    "COLUMNS": f_columns,
    "SINGLE": f_single,
}
