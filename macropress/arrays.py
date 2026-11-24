"""Excel's functions that make and shape arrays and split text, and LAMBDA's helpers: FILTER, SEQUENCE, SORTBY, TAKE,
DROP, CHOOSECOLS, CHOOSEROWS, VSTACK, HSTACK, TOCOL, TOROW, EXPAND, WRAPROWS, WRAPCOLS, TEXTSPLIT, TEXTBEFORE,
TEXTAFTER, MAP, REDUCE, SCAN, BYROW, BYCOL and MAKEARRAY. Each is written to Excel's recorded answers (tests/parity,
module TNew); where a probe left Excel's answer open, the function stops with VBANotImplemented rather than guess. LET,
LAMBDA and ISOMITTED, which bind names, are the evaluator's own (formula.py).

The functions register themselves in formula.FUNCTIONS, TAKES and the table of arguments Excel checks as a formula is
entered."""

import functools
import math

from . import formula as fx
from .errors import VBANotImplemented
from .formula import Array, Lambda, XLError, _Err, num, scalar, sort_value_key, text, truth

NA = XLError("#N/A")


def _fail(code):
    raise _Err(XLError(code))


def _given(args, i):
    return i < len(args) and args[i][0] != "missing"


def _value(ev, args, i, default):
    """Argument i as one value, or default when it is left out or left empty."""
    return scalar(ev.ev(args[i])) if _given(args, i) else default


def _count(ev, args, i, default):
    """A count (of rows, of columns, an instance): the number with its fraction dropped, or default when left out."""
    return math.trunc(num(_value(ev, args, i, None))) if _given(args, i) else default


def _block(ev, node):
    """An argument taken whole as an array, whatever the formula: a reference keeps the cells it names."""
    if node[0] == "missing":
        raise VBANotImplemented("an array left out")
    return ev.array(ev.taken(node))


def _sized(rows, cols):
    if rows * cols > Array.LIMIT:
        raise VBANotImplemented(f"an array of {rows} by {cols} values, more than this works out one by one")


def _transposed(rows):
    return [list(col) for col in zip(*rows, strict=True)]


def _flagged(v):
    """Whether FILTER keeps a row or column: TRUE or a number other than 0; a blank is FALSE, text #VALUE!."""
    if isinstance(v, XLError):
        raise _Err(v)
    if isinstance(v, str):
        _fail("#VALUE!")
    return v is not None and num(v) != 0


# ------------------------------------------------------------------------------------------------- making and sorting
def f_filter(ev, args):
    """FILTER(array, include, [if_empty]): the rows (include a column as tall as the array) or the columns (a row as
    wide) that include keeps; any other include is #VALUE!. Keeping nothing is if_empty, or #CALC!."""
    arr, include = _block(ev, args[0]), _block(ev, args[1])
    rows, cols = arr.extent
    if include.extent[1] == 1 and include.extent[0] == rows:
        keep = [i for i in range(rows) if _flagged(include.cell(i, 0))]
        out = [[arr.cell(i, j) for j in range(cols)] for i in keep]
    elif include.extent[0] == 1 and include.extent[1] == cols:
        keep = [j for j in range(cols) if _flagged(include.cell(0, j))]
        out = [[arr.cell(i, j) for j in keep] for i in range(rows)] if keep else []
    else:
        _fail("#VALUE!")
    if not keep:
        if _given(args, 2):
            return ev.ev(args[2])
        _fail("#CALC!")
    return Array(out)


def f_sequence(ev, args):
    """SEQUENCE(rows, [columns], [start], [step]): numbers row by row, each the one before it plus step, added (the
    ninth of 0, 0.1, ... is 0.7999999999999999: recorded). No rows or columns is #CALC!, fewer than none #VALUE!."""
    rows, cols = _count(ev, args, 0, 1), _count(ev, args, 1, 1)
    start, step = num(_value(ev, args, 2, 1.0)), num(_value(ev, args, 3, 1.0))
    if rows < 0 or cols < 0:
        _fail("#VALUE!")
    if rows == 0 or cols == 0:
        _fail("#CALC!")
    _sized(rows, cols)
    out, v = [], None
    for _ in range(rows):
        row = []
        for _ in range(cols):
            v = start if v is None else fx.finite(v + step)
            row.append(v)
        out.append(row)
    return Array(out)


def _sort_order(a, b, descending):
    """Two values in SORT's order: blanks last whichever the order, the rest as SORT has them."""
    if (a is None) != (b is None):
        return 1 if a is None else -1
    if a is None:
        return 0
    ka, kb = sort_value_key(a), sort_value_key(b)
    if ka == kb:
        return 0
    return (1 if ka > kb else -1) * (-1 if descending else 1)


def f_sortby(ev, args):
    """SORTBY(array, by_array, [order], ...): the array's rows sorted by columns as tall (or its columns by rows as
    wide), each key in its order (1 or -1, anything else #VALUE!), ties kept in the order they were in. Keys of
    another size, or not all along the same way, are #VALUE!."""
    arr = _block(ev, args[0])
    rows, cols = arr.extent
    keys, along = [], None
    for i in range(1, len(args), 2):
        by = _block(ev, args[i])
        order = num(_value(ev, args, i + 1, 1.0))
        if order not in (1.0, -1.0):
            _fail("#VALUE!")
        if by.extent[1] == 1 and by.extent[0] == rows:
            way, vals = "rows", [by.cell(k, 0) for k in range(rows)]
        elif by.extent[0] == 1 and by.extent[1] == cols:
            way, vals = "columns", [by.cell(0, k) for k in range(cols)]
        else:
            _fail("#VALUE!")
        if along not in (None, way):
            _fail("#VALUE!")
        along = way
        keys.append((vals, order == -1.0))

    def compare(a, b):
        for vals, descending in keys:
            c = _sort_order(vals[a], vals[b], descending)
            if c:
                return c
        return 0

    data = arr.full()
    order = sorted(range(rows if along == "rows" else cols), key=functools.cmp_to_key(compare))
    if along == "rows":
        return Array([data[i] for i in order])
    return Array([[row[j] for j in order] for row in data])


# ---------------------------------------------------------------------------------------------------------- shaping
def f_take(ev, args):
    """TAKE(array, rows, [columns]): the first rows and columns, or the last for a number below 0; more than there
    are is all of them, none #CALC!."""
    arr = _block(ev, args[0])
    rows, cols = arr.extent
    r, c = _count(ev, args, 1, rows), _count(ev, args, 2, cols)
    if r == 0 or c == 0:
        _fail("#CALC!")
    picked_rows = slice(0, r) if r > 0 else slice(max(rows + r, 0), rows)
    picked_cols = slice(0, c) if c > 0 else slice(max(cols + c, 0), cols)
    return Array([row[picked_cols] for row in arr.full()[picked_rows]])


def f_drop(ev, args):
    """DROP(array, rows, [columns]): the array without its first rows and columns, or its last for a number below
    0; dropping all of them is #CALC!."""
    arr = _block(ev, args[0])
    rows, cols = arr.extent
    r, c = _count(ev, args, 1, 0), _count(ev, args, 2, 0)
    kept_rows = slice(r, rows) if r >= 0 else slice(0, max(rows + r, 0))
    kept_cols = slice(c, cols) if c >= 0 else slice(0, max(cols + c, 0))
    out = [row[kept_cols] for row in arr.full()[kept_rows]]
    if not out or not out[0]:
        _fail("#CALC!")
    return Array(out)


def _picked(ev, args, n):
    """CHOOSEROWS' and CHOOSECOLS' numbers, of arrays of them too, as 0-based places; 0 or past the end is
    #VALUE!, below 0 counts from the end."""
    out = []
    for node in args[1:]:
        for row in ev.array(ev.ev(node)).full():
            for v in row:
                k = math.trunc(num(v))
                if k == 0 or abs(k) > n:
                    _fail("#VALUE!")
                out.append(k - 1 if k > 0 else n + k)
    return out


def f_chooserows(ev, args):
    arr = _block(ev, args[0])
    data = arr.full()
    return Array([list(data[i]) for i in _picked(ev, args, arr.extent[0])])


def f_choosecols(ev, args):
    arr = _block(ev, args[0])
    picked = _picked(ev, args, arr.extent[1])
    return Array([[row[j] for j in picked] for row in arr.full()])


def f_vstack(ev, args):
    """VSTACK(array, ...): the arrays one below the other, as wide as the widest; past a narrower one's end, #N/A."""
    blocks = [_block(ev, a) for a in args]
    width = max(b.extent[1] for b in blocks)
    _sized(sum(b.extent[0] for b in blocks), width)
    return Array([row + [NA] * (width - len(row)) for b in blocks for row in b.full()])


def f_hstack(ev, args):
    """HSTACK(array, ...): the arrays side by side, as tall as the tallest; past a shorter one's end, #N/A."""
    blocks = [_block(ev, a) for a in args]
    height = max(b.extent[0] for b in blocks)
    _sized(height, sum(b.extent[1] for b in blocks))
    out = [[] for _ in range(height)]
    for b in blocks:
        data, width = b.full(), b.extent[1]
        for i in range(height):
            out[i] += data[i] if i < len(data) else [NA] * width
    return Array(out)


def _flattened(ev, args):
    """TOCOL's and TOROW's values: row by row, or column by column; ignore 1 passes over blanks, 2 errors, 3 both
    (anything else #VALUE!). None left is #CALC!."""
    arr = _block(ev, args[0])
    ignore = _count(ev, args, 1, 0)
    if ignore not in (0, 1, 2, 3):
        _fail("#VALUE!")
    data = arr.full()
    by_column = _given(args, 2) and truth(ev.ev(args[2]))
    values = [v for col in zip(*data, strict=True) for v in col] if by_column else [v for row in data for v in row]
    out = [v for v in values if not ((ignore & 1 and v is None) or (ignore & 2 and isinstance(v, XLError)))]
    if not out:
        _fail("#CALC!")
    return out


def f_tocol(ev, args):
    return Array([[v] for v in _flattened(ev, args)])


def f_torow(ev, args):
    return Array([_flattened(ev, args)])


def f_expand(ev, args):
    """EXPAND(array, rows, [columns], [pad_with]): the array made larger, padded with #N/A or pad_with; smaller than
    it is #VALUE!."""
    arr = _block(ev, args[0])
    rows, cols = arr.extent
    r, c = _count(ev, args, 1, rows), _count(ev, args, 2, cols)
    if r < rows or c < cols:
        _fail("#VALUE!")
    _sized(r, c)
    pad = _value(ev, args, 3, NA)
    return Array([row + [pad] * (c - cols) for row in arr.full()] + [[pad] * c for _ in range(r - rows)])


def _wrapped(ev, args):
    """The lines WRAPROWS and WRAPCOLS make of a row or column (anything else #VALUE!): `wrap` values each (fewer than
    1, #NUM!), the last padded with #N/A or pad_with."""
    arr = _block(ev, args[0])
    if 1 not in arr.extent:
        _fail("#VALUE!")
    wrap = _count(ev, args, 1, 0)
    if wrap < 1:
        _fail("#NUM!")
    pad = _value(ev, args, 2, NA)
    values = [v for row in arr.full() for v in row]
    lines = [values[k : k + wrap] for k in range(0, len(values), wrap)]
    lines[-1] += [pad] * (wrap - len(lines[-1]))
    return lines


def f_wraprows(ev, args):
    return Array(_wrapped(ev, args))


def f_wrapcols(ev, args):
    return Array(_transposed(_wrapped(ev, args)))


# ------------------------------------------------------------------------------------------------------------- text
def _delimiters(ev, args, i):
    """A delimiter, or an array of them, as text; none when left out."""
    if not _given(args, i):
        return []
    return [text(v) for row in ev.array(ev.ev(args[i])).full() for v in row]


def _found(s, delimiters, fold):
    """Where the delimiters stand in s, left to right, as (start, length): where several match at one place, the
    longest (recorded, in either order). Matches of one that overlap themselves are not modelled."""
    for d in delimiters:
        starts = [i for i in range(len(s)) if d and _at(s, i, d, fold)]
        if any(b - a < len(d) for a, b in zip(starts, starts[1:], strict=False)):
            raise VBANotImplemented(f"the delimiter {d!r} matching where it overlaps itself")
    out, i = [], 0
    while i < len(s):
        length = max((len(d) for d in delimiters if d and _at(s, i, d, fold)), default=0)
        if length:
            out.append((i, length))
            i += length
        else:
            i += 1
    return out


def _at(s, i, d, fold):
    part = s[i : i + len(d)]
    return part.lower() == d.lower() if fold else part == d


def _mode(ev, args, i):
    """match_mode: 0 matches case, 1 ignores it."""
    mode = _count(ev, args, i, 0)
    if mode not in (0, 1):
        _fail("#VALUE!")
    return mode == 1


def _split(s, delimiters, fold):
    parts, last = [], 0
    for start, length in _found(s, delimiters, fold):
        parts.append(s[last:start])
        last = start + length
    return parts + [s[last:]]


def f_textsplit(ev, args):
    """TEXTSPLIT(text, col_delimiter, [row_delimiter], [ignore_empty], [match_mode], [pad_with]): the text's parts,
    split into rows first and each row into columns, as text; rows padded with #N/A or pad_with. No text, or a
    delimiter that is empty, is #VALUE!; ignore_empty passes over empty parts and empty rows (recorded)."""
    s = text(_value(ev, args, 0, ""))
    across, down = _delimiters(ev, args, 1), _delimiters(ev, args, 2)
    if not across and not down:
        raise VBANotImplemented("TEXTSPLIT without a delimiter")
    if s == "" or "" in across + down:
        _fail("#VALUE!")
    if set(across) & set(down):
        raise VBANotImplemented("TEXTSPLIT with a delimiter for rows that is one for columns too")
    fold, pad = _mode(ev, args, 4), _value(ev, args, 5, NA)
    table = [_split(line, across, fold) if across else [line] for line in (_split(s, down, fold) if down else [s])]
    if _given(args, 3) and truth(ev.ev(args[3])):
        table = [row for row in ([part for part in row if part != ""] for row in table) if row]
        if not table:
            raise VBANotImplemented("TEXTSPLIT of text that is all delimiters, ignoring empty parts")
    width = max(len(row) for row in table)
    _sized(len(table), width)
    return Array([row + [pad] * (width - len(row)) for row in table])


def _around(ev, args, after):
    """TEXTBEFORE and TEXTAFTER: the text before or after the instance_num-th delimiter (below 0, counting from the
    end); with match_end the end of the text counts as one after the others (from the end, its start). 0, or more
    than the text's length, is #VALUE!; an instance not found is if_not_found, or #N/A (recorded)."""
    s = text(_value(ev, args, 0, ""))
    delimiters = _delimiters(ev, args, 1)
    n = _count(ev, args, 2, 1)
    fold = _mode(ev, args, 3)
    to_end = _given(args, 4) and truth(ev.ev(args[4]))
    if n == 0 or abs(n) > len(s):
        _fail("#VALUE!")
    if "" in delimiters:  # the empty delimiter stands at the start, or from the end at the end (recorded)
        if len(delimiters) > 1 or abs(n) != 1 or to_end:
            raise VBANotImplemented("TEXTBEFORE or TEXTAFTER with an empty delimiter, past the first or beside others")
        place = (0, 0) if n > 0 else (len(s), 0)
    else:
        found = _found(s, delimiters, fold)
        if to_end:
            found = found + [(len(s), 0)] if n > 0 else [(0, 0)] + found
        if abs(n) > len(found):
            return _value(ev, args, 5, None) if _given(args, 5) else _fail("#N/A")
        place = found[n - 1] if n > 0 else found[n]
    return s[place[0] + place[1] :] if after else s[: place[0]]


def f_textbefore(ev, args):
    return _around(ev, args, False)


def f_textafter(ev, args):
    return _around(ev, args, True)


# --------------------------------------------------------------------------------------------------- LAMBDA's helpers
def _lambda(ev, args):
    """The LAMBDA a helper is given last: anything else is #VALUE! (recorded)."""
    lam = ev.ev(args[-1]) if len(args) > 1 else None
    if not isinstance(lam, Lambda):
        _fail("#VALUE!")
    return lam


def _takes(lam, n):
    """Whether a LAMBDA takes n values: no more than its parameters, no fewer than those not in [ ]."""
    return n <= len(lam.params) and all(optional for _, optional in lam.params[n:])


def _item(block, i, j):
    """Value (i, j) of a block, as the helpers hand it to their LAMBDA: of a reference, the cell (ROW of it works:
    recorded), else the value."""
    if block.rect is None:
        return block.cell(i, j)
    sheet, r1, c1, _, _ = block.rect
    return Array([[block.cell(i, j)]], (1, 1), (sheet, r1 + i, c1 + j, r1 + i, c1 + j))


def _line(ev, block, k, row):
    """Row or column k of a block, the cells themselves where it is a reference."""
    rows, cols = block.extent
    if block.rect is not None:
        sheet, r1, c1, r2, c2 = block.rect
        return ev.cells((sheet, r1 + k, c1, r1 + k, c2) if row else (sheet, r1, c1 + k, r2, c1 + k))
    return Array([[block.cell(k, j) for j in range(cols)]] if row else [[block.cell(i, k)] for i in range(rows)])


def _called(ev, lam, values):
    """The LAMBDA worked out for these values: an error it comes to is its value there."""
    try:
        return ev.apply(lam, values)
    except _Err as e:
        return e.err


def _one(v):
    """One value a LAMBDA gave for one place: an array of more than one, or a LAMBDA, makes the whole #CALC!."""
    if isinstance(v, Array):
        if v.extent != (1, 1):
            _fail("#CALC!")
        return v.cell(0, 0)
    if isinstance(v, Lambda):
        _fail("#CALC!")
    return v


def f_map(ev, args):
    """MAP(array, ..., lambda): the LAMBDA of the arrays' values at each place; past the end of a shorter one (a
    single cell too), #N/A, the LAMBDA not called (recorded)."""
    lam = _lambda(ev, args)
    blocks = [_block(ev, a) for a in args[:-1]]
    if not _takes(lam, len(blocks)):
        _fail("#VALUE!")
    rows, cols = max(b.extent[0] for b in blocks), max(b.extent[1] for b in blocks)
    _sized(rows, cols)
    out = []
    for i in range(rows):
        row = []
        for j in range(cols):
            if any(i >= b.extent[0] or j >= b.extent[1] for b in blocks):
                row.append(NA)
            else:
                row.append(_one(_called(ev, lam, [_item(b, i, j) for b in blocks])))
        out.append(row)
    return Array(out)


def _gone_through(ev, args):
    """What REDUCE and SCAN give at each value, row by row: the LAMBDA of what it gave so far (at first initial, blank
    when left out) and the value. [(row, what it gave)]."""
    lam = _lambda(ev, args)
    if not _takes(lam, 2):
        _fail("#VALUE!")
    block = _block(ev, args[1])
    rows, cols = block.extent
    _sized(rows, cols)
    acc, out = ev.keep(args[0]) if _given(args, 0) else None, []
    for i in range(rows):
        for j in range(cols):
            acc = _called(ev, lam, [acc, _item(block, i, j)])
            out.append((i, acc))
    return out


def f_reduce(ev, args):
    """REDUCE([initial], array, lambda): what the LAMBDA gives at the last value; it may be an array (recorded)."""
    return _gone_through(ev, args)[-1][1]


def f_scan(ev, args):
    """SCAN([initial], array, lambda): what the LAMBDA gives at each value, one each, an array as large as the array;
    an array given at one is #CALC! (recorded)."""
    out = []
    for i, acc in _gone_through(ev, args):
        if i == len(out):
            out.append([])
        out[i].append(_one(acc))
    return Array(out)


def _by(ev, args, row):
    """BYROW and BYCOL: the LAMBDA of each row or column, the cells themselves where the array is a reference
    (recorded); one value each. Without a LAMBDA, #CALC! (recorded)."""
    if len(args) < 2:
        _fail("#CALC!")
    lam = _lambda(ev, args)
    if not _takes(lam, 1):
        _fail("#VALUE!")
    block = _block(ev, args[0])
    values = [_one(_called(ev, lam, [_line(ev, block, k, row)])) for k in range(block.extent[0 if row else 1])]
    return Array([[v] for v in values] if row else [values])


def f_byrow(ev, args):
    return _by(ev, args, True)


def f_bycol(ev, args):
    return _by(ev, args, False)


def f_makearray(ev, args):
    """MAKEARRAY(rows, columns, lambda): the LAMBDA of each row and column number; fewer than 1 is #VALUE!."""
    lam = _lambda(ev, args)
    rows, cols = _count(ev, args, 0, 0), _count(ev, args, 1, 0)
    if rows < 1 or cols < 1:
        _fail("#VALUE!")
    if not _takes(lam, 2):
        _fail("#VALUE!")
    _sized(rows, cols)
    return Array([[_one(_called(ev, lam, [float(i + 1), float(j + 1)])) for j in range(cols)] for i in range(rows)])


FUNCTIONS = {
    "FILTER": f_filter,
    "SEQUENCE": f_sequence,
    "SORTBY": f_sortby,
    "TAKE": f_take,
    "DROP": f_drop,
    "CHOOSEROWS": f_chooserows,
    "CHOOSECOLS": f_choosecols,
    "VSTACK": f_vstack,
    "HSTACK": f_hstack,
    "TOCOL": f_tocol,
    "TOROW": f_torow,
    "EXPAND": f_expand,
    "WRAPROWS": f_wraprows,
    "WRAPCOLS": f_wrapcols,
    "TEXTSPLIT": f_textsplit,
    "TEXTBEFORE": f_textbefore,
    "TEXTAFTER": f_textafter,
    "MAP": f_map,
    "REDUCE": f_reduce,
    "SCAN": f_scan,
    "BYROW": f_byrow,
    "BYCOL": f_bycol,
    "MAKEARRAY": f_makearray,
}
# The text functions are worked out for each value of an array of texts (an array each one gives, its first value:
# recorded), their delimiters taken whole; SEQUENCE's numbers likewise. The rest work out their arguments as arrays,
# whatever the formula (=@FILTER(A1:A6,A1:A6>2): recorded).
TAKES = {
    **dict.fromkeys(set(FUNCTIONS) - {"SEQUENCE", "TEXTSPLIT", "TEXTBEFORE", "TEXTAFTER"}, "a"),
    "SEQUENCE": "vvvv", "TEXTSPLIT": "vrrvvv", "TEXTBEFORE": "vrvvvv", "TEXTAFTER": "vrvvvv",
}  # fmt: skip
_COUNTS = {
    "FILTER": (2, 3), "SEQUENCE": (1, 4), "SORTBY": (2, 255), "TAKE": (2, 3), "DROP": (2, 3), "CHOOSEROWS": (2, 255),
    "CHOOSECOLS": (2, 255), "VSTACK": (1, 254), "HSTACK": (1, 254), "TOCOL": (1, 3), "TOROW": (1, 3), "EXPAND": (2, 4),
    "WRAPROWS": (2, 3), "WRAPCOLS": (2, 3), "TEXTSPLIT": (2, 6), "TEXTBEFORE": (2, 6), "TEXTAFTER": (2, 6),
    "MAP": (2, 254), "REDUCE": (3, 3), "SCAN": (3, 3), "BYROW": (1, 2), "BYCOL": (1, 2), "MAKEARRAY": (3, 3),
    "LAMBDA": (1, 254), "ISOMITTED": (1, 1),
}  # fmt: skip
ARGUMENTS = {**{name: (fewest, most, 1, "") for name, (fewest, most) in _COUNTS.items()}, "LET": (3, 253, 2, "")}

fx.FUNCTIONS.update(FUNCTIONS)
fx.TAKES.update(TAKES)
fx._ARGUMENTS.update(ARGUMENTS)
fx.EXCEL_FUNCTIONS.update(FUNCTIONS, ARGUMENTS)
