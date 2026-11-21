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
|(?P<op><>|<=|>=|[-+*/^&=<>(),%])
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
}


def file_syntax(text):
    """A formula as the file stores it: newer functions get their _xlfn. prefix."""
    return "".join(FILE_PREFIX.get(p.upper(), p) if kind == "func" else p for kind, p in pieces(text))


_XLFN = re.compile(r"(?i)^_xlfn\.(?:_xlws\.)?")


def display_syntax(text):
    """A formula as VBA's Formula property shows it: without the file's _xlfn. prefixes."""
    return "".join(_XLFN.sub("", p) if kind == "func" else p for kind, p in pieces(text))


def bare_name(name):
    n = name.upper()
    for prefix in ("_XLFN._XLWS.", "_XLFN.", "_XLWS."):
        if n.startswith(prefix):
            n = n[len(prefix) :]
    return n


# ---------------------------------------------------------------------------
# Parsing
# ---------------------------------------------------------------------------


def tokenize(text):
    out, pos = [], 0
    for kind, piece in pieces(text):
        if kind is None:
            raise VBANotImplemented(f"formula syntax at {text[pos : pos + 20]!r} in ={text}")
        if kind != "ws":
            out.append((kind, piece))
        pos += len(piece)
    return out


class Parser:
    def __init__(self, text):
        self.text = text
        self.toks = tokenize(text[1:] if text.startswith("=") else text)
        self.i = 0

    def peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else (None, None)

    def take(self):
        t = self.peek()
        self.i += 1
        return t

    def parse(self):
        node = self.expr(0)
        if self.i != len(self.toks):
            raise VBANotImplemented(f"formula syntax near {self.peek()[1]!r} in ={self.text}")
        return node

    def expr(self, min_prec):
        left = self.unary()
        while True:
            kind, val = self.peek()
            if kind != "op" or val not in BIN_PREC or BIN_PREC[val] < min_prec:
                return left
            self.take()
            right = self.expr(BIN_PREC[val] + 1)
            left = ("bin", val, left, right)

    def unary(self):
        kind, val = self.peek()
        if kind == "op" and val in ("-", "+"):
            self.take()
            operand = self.unary()
            return ("neg", operand) if val == "-" else operand
        node = self.primary()
        while self.peek() == ("op", "%"):
            self.take()
            node = ("pct", node)
        return node

    def primary(self):
        kind, val = self.take()
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
    elif node[0] in ("neg", "pct", "paren"):
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
    elif node[0] in ("neg", "pct", "paren"):
        references(node[1], out)
    return out


# ---------------------------------------------------------------------------
# Values with Excel's rules
# ---------------------------------------------------------------------------


class Array:
    """A 2-D block of values: a range read, or a dynamic-array result. A range read holds the cells up to the sheet's
    last data only; `extent` is the size of the range itself, and its cells past the data are blank. `rect` is where
    a range read lies, (sheet, r1, c1, r2, c2): what a formula that is not a dynamic array takes one cell of."""

    def __init__(self, rows, extent=None, rect=None):
        self.rows = rows
        self.extent = extent or self.shape
        self.rect = rect

    @property
    def shape(self):
        return len(self.rows), (len(self.rows[0]) if self.rows else 0)

    def cell(self, i, j):
        """The value at 0-based (i, j) of the range: blank past the data."""
        return self.rows[i][j] if i < len(self.rows) and j < len(self.rows[i]) else None


def scalar(v):
    if isinstance(v, Array):
        if v.shape == (1, 1):
            return v.rows[0][0]
        raise VBANotImplemented("an array where a formula expects one value")
    return v


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


class Evaluator:
    """Evaluates one formula of `sheet` against the book (reading cells through `read`)."""

    def __init__(self, book, sheet, read):
        self.book = book
        self.sheet = sheet
        self.read = read  # read(sheet, r, c) -> value (refreshing a stale formula first)

    def run(self, node):
        try:
            if node[0] == "bin" and node[1] in ("+", "-"):  # the formula ends on a sum: see snap_add
                x, y = num(self.ev(node[2])), num(self.ev(node[3]))
                return finite(snap_add(x, y if node[1] == "+" else -y))
            return self.ev(node)
        except _Err as e:
            return e.err
        except ZeroDivisionError:
            return XLError("#DIV/0!")

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
        if kind == "num":
            return node[1]
        if kind in ("str", "bool"):
            return node[1]
        if kind == "err":
            raise _Err(node[1])
        if kind == "missing":
            return None
        if kind in ("ref", "name"):
            rect = self.rect_of(node)
            sheet, r1, c1, r2, c2 = rect
            if (r1, c1) == (r2, c2):
                return self.read(sheet, r1, c1)
            return self.range_values(rect)
        if kind == "paren":
            return self.ev(node[1])
        if kind == "neg":
            return -num(self.ev(node[1]))
        if kind == "pct":
            return num(self.ev(node[1])) / 100
        if kind == "bin":
            op = node[1]
            a, b = self.ev(node[2]), self.ev(node[3])
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
            if op == "^":
                return power(x, y)
        if kind == "func":
            fn = FUNCTIONS.get(node[1])
            if fn is None:
                raise VBANotImplemented(f"formula function {node[1]}")
            return fn(self, node[2])
        raise VBANotImplemented(f"formula part {kind}")

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


def _flat(ev, args):
    out = []
    for a in args:
        if a[0] == "missing":
            continue
        v = ev.ev(a)
        if isinstance(v, Array):
            out += [x for row in v.rows for x in row]
        else:
            out.append(v)
    return out


# Functions that hand back the cells they pick: what they return is read as cells are, not as a value written out.
_PICKS_CELLS = ("INDEX", "XLOOKUP", "IF", "IFERROR")


def _operands(ev, args, name):
    """What SUM, COUNT, AND and the like are given, as (value, written): written is True for a value written or
    worked out in the formula itself (SUM("5",TRUE) is 6), False for one read from cells or an array, where these
    functions pass over text and logical values, and None where a function picked it and either could be so."""
    out = []
    for a in args:
        if a[0] == "missing":  # SUM(1,,2): an argument left empty is a 0 written there, FALSE to AND and OR
            out.append((0.0, True))
            continue
        try:
            v = ev.ev(a)
        except _Err as e:
            v = e.err
        if isinstance(v, Array):
            out += [(x, False) for row in v.rows for x in row]
            continue
        while a[0] == "paren":
            a = a[1]
        cells = True if a[0] in ("ref", "name") else None if a[0] == "func" and a[1] in _PICKS_CELLS else False
        out.append((v, None if cells is None else not cells))
    return out


def _unsure(name, v):
    return VBANotImplemented(f"{name} of {v!r} from a function that may have read it from a cell")


def _numbers(ev, args, name, counting=False):
    """The numbers among a function's operands. Text written in the formula counts if it reads as a number and is
    #VALUE! if not; logical values written there count as 1 and 0. COUNT passes over errors and text it cannot read."""
    out = []
    for v, written in _operands(ev, args, name):
        if isinstance(v, XLError):
            if not counting:
                raise _Err(v)
        elif isinstance(v, (bool, str)):
            if written is None:
                raise _unsure(name, v)
            if not written:
                continue
            n = float(v) if isinstance(v, bool) else entry.value_of(v)
            if n is not None:
                out.append(n)
            elif not counting:
                raise _Err(XLError("#VALUE!"))
        elif isinstance(v, (int, float)):
            out.append(float(v))
    return out


def total(numbers):
    """The numbers added up as SUM adds them: in order, the last addition as snap_add has it."""
    if not numbers:
        return 0.0
    total = 0.0
    for n in numbers[:-1]:
        total += n
    return finite(snap_add(total, numbers[-1]))


def f_sum(ev, args):
    return total(_numbers(ev, args, "SUM"))


def f_min(ev, args):
    n = _numbers(ev, args, "MIN")
    return min(n) if n else 0.0


def f_max(ev, args):
    n = _numbers(ev, args, "MAX")
    return max(n) if n else 0.0


def f_count(ev, args):
    return float(len(_numbers(ev, args, "COUNT", counting=True)))


def f_counta(ev, args):
    return float(sum(1 for v, _ in _operands(ev, args, "COUNTA") if v is not None))


def f_average(ev, args):
    n = _numbers(ev, args, "AVERAGE")
    if not n:
        raise _Err(XLError("#DIV/0!"))
    return total(n) / len(n)


def f_if(ev, args):
    cond = truth(ev.ev(args[0]))
    if cond:
        return ev.arg(args, 1, True)
    return ev.arg(args, 2, False)


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
    for v, written in _operands(ev, args, name):
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
    values = [x for row in arr.rows for x in row]
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
    i = _match_index([row[0] for row in arr.rows], key)
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
    i = _match_index([x for row in look.rows for x in row], key)
    if i is None:
        if len(args) > 3 and args[3][0] != "missing":
            return ev.ev(args[3])
        raise _Err(XLError("#N/A"))
    return ret.cell(i, 0) if look.extent[1] == 1 else ret.cell(0, i)


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
        return arr.cell(r - 1, c - 1)
    r1, r2 = (r - 1, r - 1) if r else (0, rows - 1)
    c1, c2 = (c - 1, c - 1) if c else (0, cols - 1)
    if (r1, c1) == (r2, c2):
        return arr.cell(r1, c1)
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
    "CONCATENATE": f_concat,
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
}
