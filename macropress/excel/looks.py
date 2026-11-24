"""How a sheet looks: fonts, fills, borders, alignment, widths and heights, tab colours and the window.

None of it is modelled (group 5 of the roadmap). Told to (--skip-appearance), a run passes over the statements that
only change it, where Excel takes the value given, and counts them for its report; otherwise such a statement stops
the run, as any read of how a sheet looks does. What these statements change that a macro can see is kept, as
recorded from Excel (tests/parity, module TAppear): on a protected sheet they are error 1004 unless its protection
allows formatting; a width or height of 0 hides, and any other shows again; the cells and rows they format, and rows
given a height, count in the used range; and where a cell's font, alignment or column width changed, the text a number
shows there (Range.Text) is not known, since Excel shows ### where it no longer fits.
"""

from ..errors import VBANotImplemented
from ..values import VBAObject, default_value, format_double
from .constants import MAX_COL, MAX_ROW


# --- the values Excel takes, each recorded; for any other a run stops rather than guess Excel's error ---
def _number(v):
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _whole(lo, hi):
    return lambda v: _number(v) and v == int(v) and lo <= v <= hi


def _between(lo, hi):
    return lambda v: _number(v) and lo <= v <= hi


def _one_of(*values):
    return lambda v: _number(v) and v in values


def _either(*checks):
    return lambda v: any(check(v) for check in checks)


def _truth(v):
    return isinstance(v, bool) or (_number(v) and v in (0, 1, -1))


def _name(v):
    return isinstance(v, str) and len(v) <= 31  # a font's name of 40 characters is error 1004


COLOR = _whole(-(2**31), 2**31 - 1)  # any Long
THEME = _whole(1, 12)  # xlThemeColorDark1 to xlThemeColorFollowedHyperlink; 13 is error 9
TINT = _between(-1, 1)  # 1.5 is error 5
BORDER = {
    "LineStyle": _one_of(1, -4115, -4118, -4119, 4, 5, 13, -4142),  # xlContinuous to xlLineStyleNone
    "Weight": _one_of(1, 2, -4138, 4),  # xlHairline, xlThin, xlMedium, xlThick
    "Color": COLOR,
    "ColorIndex": _either(_whole(0, 56), _one_of(-4105)),
    "ThemeColor": THEME,
    "TintAndShade": TINT,
}
# member: (what Excel takes, whether it changes the text a number shows, where it no longer fits)
FONT = {
    **dict.fromkeys(("Bold", "Italic", "Superscript", "Subscript", "OutlineFont", "Shadow"), (_truth, True)),
    "Strikethrough": (_truth, False),
    "Underline": (_either(_truth, _one_of(-4142, 2, -4119, 4, 5)), False),
    "Name": (_name, True),
    "FontStyle": (_name, True),
    "Size": (_between(1, 409), True),
    "Color": (COLOR, False),
    "ColorIndex": (_either(_whole(0, 56), _one_of(-4105)), False),  # 57 is error 9
    "ThemeColor": (THEME, False),
    "TintAndShade": (TINT, False),
    "ThemeFont": (_whole(0, 2), True),
}
INTERIOR = {
    "Color": (COLOR, False),
    "ColorIndex": (_either(_whole(1, 56), _one_of(-4142)), False),
    "Pattern": (_one_of(1, -4142, -4105, -4124, -4125, -4126, 17, 18), False),  # xlSolid, xlNone, ..., the greys
    "PatternColor": (COLOR, False),
    "PatternColorIndex": (_either(_whole(1, 56), _one_of(-4105)), False),
    "ThemeColor": (THEME, False),
    "TintAndShade": (TINT, False),
    "PatternThemeColor": (THEME, False),
    "PatternTintAndShade": (TINT, False),
}
ALIGNMENT = {
    "HorizontalAlignment": (_one_of(1, -4131, -4108, -4152, 5, -4130, 7, -4117), True),
    "VerticalAlignment": (_one_of(-4160, -4108, -4107, -4130, -4117), False),
    "Orientation": (_either(_whole(-90, 90), _one_of(-4171, -4170, -4166, -4128)), True),  # 91 is error 1004
    "IndentLevel": (_whole(0, 250), True),  # 251 is error 1004
    "AddIndent": (_truth, True),
    "WrapText": (_truth, True),
    "ShrinkToFit": (_truth, True),
    "ReadingOrder": (_one_of(-5002, -5003, -5004), True),
}
TAB = {"Color": _either(COLOR, lambda v: v is False), "ColorIndex": _either(_whole(1, 56), _one_of(-4142)),
       "ThemeColor": THEME, "TintAndShade": TINT}  # fmt: skip
WINDOW = {
    "Zoom": _whole(10, 400),  # 9 and 401 are error 1004
    **dict.fromkeys(("DisplayGridlines", "DisplayHeadings", "FreezePanes", "Split"), _truth),
    "SplitRow": _whole(0, 30),  # as far as recorded: how far a split may go turns on the window's size
    "SplitColumn": _whole(0, 20),
    "ScrollRow": _whole(1, MAX_ROW),
    "ScrollColumn": _whole(1, MAX_COL),
    "View": _one_of(1),  # xlNormalView
}
WHOLE = _whole(-(2**31), 2**31 - 1)


def _canonical(rules, name):
    return next((member for member in rules if member.lower() == name.lower()), None)


def _shown(v):
    return format_double(v) if _number(v) else repr(v)


def require(app, what):
    """Stop the run at a statement that only changes how a sheet looks, unless it was told to pass over them."""
    if not app.skip_appearance:
        raise VBANotImplemented(
            f"{what}: how a sheet looks is not modelled (--skip-appearance passes over statements that only change it)"
        )


def passing(app, what, value, check):
    """The value a statement that only changes how a sheet looks is given, where the run passes over such statements
    and Excel takes it; otherwise the run stops."""
    require(app, what)
    v = default_value(value)
    if not check(v):
        raise VBANotImplemented(f"{what} = {_shown(v)}: a value Excel may refuse, so it is not passed over")
    return v


def count(app, kind):
    app.passed_over[kind] += 1
    app.note("APPEARANCE-SKIPPED", kind)


def formatted(rng, kind, changes_text):
    """What formatting a range leaves that a macro sees: error 1004 on a protected sheet that does not allow it, the
    cells counted in the used range, and, if the change is one that moves text about, the text numbers show there not
    known."""
    rng._single_sheet_target("Formatting")
    sheet = rng.sheet
    sheet.guard("format")
    sheet.looks.formatted(*rng.rect)
    if changes_text:
        sheet.looks.restyle(*rng.rect)
    count(rng.app, kind)


def setting(member):
    """Range.<member> = value, for the alignment of cells."""
    check, changes_text = ALIGNMENT[member]

    def let(self, value):
        passing(self.app, f"Range.{member}", value, check)
        formatted(self, "alignment", changes_text)

    return let


def size(rng, member, value, top):
    """A width or height given to whole columns or rows: 0 (which hides them) or from 1 to `top`. One between stops the
    run: Excel hides a row given 0.8 points but not 1, and a column given 0.01 of a character but not 0.05."""
    v = passing(rng.app, f"Range.{member}", value, _either(_one_of(0), _between(1, top)))
    rng._single_sheet_target(member)
    return v


def border_around(rng, **given):
    require(rng.app, "Range.BorderAround")
    given = {name: v for name, v in given.items() if v is not None}
    if not given:
        raise VBANotImplemented("Range.BorderAround with no arguments (unverified)")
    for name, v in given.items():
        member = _canonical(BORDER, name)
        passing(rng.app, f"Range.BorderAround {member}:=", v, BORDER[member])
    if "colorindex" in given and "color" in given:
        raise VBANotImplemented("Range.BorderAround given both ColorIndex and Color (unverified)")
    formatted(rng, "Borders", False)


class Look(VBAObject):
    """An object that says how cells look (Font, Interior, Borders, Border): it takes what Excel takes, and is never
    read."""

    rules = {}  # member -> (what Excel takes, whether it changes the text a number shows)
    kind = ""  # what the run's report counts it as

    def __init__(self, rng):
        self.rng = rng

    def vba_member(self, name, args=(), kwargs=None):
        raise VBANotImplemented(f"reading {self.vba_type_name}.{name}: how a sheet looks is not modelled")

    def vba_let_member(self, name, args, kwargs, value):
        member = _canonical(self.rules, name)
        if member is None or args or kwargs:
            raise VBANotImplemented(f"setting {self.vba_type_name}.{name} is not modelled")
        check, changes_text = self.rules[member]
        passing(self.rng.app, f"{self.vba_type_name}.{member}", value, check)
        formatted(self.rng, self.kind, changes_text)


class Font(Look):
    vba_type_name = kind = "Font"
    rules = FONT


class Interior(Look):
    vba_type_name = kind = "Interior"
    rules = INTERIOR


class Border(Look):
    vba_type_name = "Border"
    kind = "Borders"
    rules = {member: (check, False) for member, check in BORDER.items()}


class Borders(Border):
    """Range.Borders: all the edges and inside lines at once, or one of them, Borders(xlEdgeBottom)."""

    vba_type_name = "Borders"

    def vba_index(self, args, kwargs):
        if len(args) != 1 or kwargs:
            raise VBANotImplemented("Borders with other than one index")
        index = default_value(args[0])
        if not _whole(5, 12)(index):  # xlDiagonalDown to xlInsideHorizontal
            raise VBANotImplemented(f"Borders({_shown(index)}) is not modelled")
        return Border(self.rng)

    def vba_member(self, name, args=(), kwargs=None):
        if name.lower() == "item":
            return self.vba_index(list(args), kwargs or {})
        return super().vba_member(name, args, kwargs)


class Tab(VBAObject):
    """Worksheet.Tab: the colour of the sheet's tab. Allowed on a protected sheet (recorded)."""

    vba_type_name = "Tab"

    def __init__(self, sheet):
        self.sheet = sheet

    def vba_member(self, name, args=(), kwargs=None):
        raise VBANotImplemented(f"reading Tab.{name}: how a sheet looks is not modelled")

    def vba_let_member(self, name, args, kwargs, value):
        member = _canonical(TAB, name)
        if member is None or args or kwargs:
            raise VBANotImplemented(f"setting Tab.{name} is not modelled")
        app = self.sheet.book.app
        passing(app, f"Tab.{member}", value, TAB[member])
        protection = self.sheet.book.protection
        if protection is not None and protection.structure:
            raise VBANotImplemented("the tab colour of a sheet in a workbook whose structure is protected (unverified)")
        count(app, "Tab")


class Window(VBAObject):
    """ActiveWindow: how the active sheet is shown (zoom, gridlines, frozen panes, scrolling). Nothing a macro reads
    turns on it, but for the zoom: the text a number shows may."""

    vba_type_name = "Window"

    def __init__(self, app):
        self.app = app

    def vba_member(self, name, args=(), kwargs=None):
        if _canonical(WINDOW, name) is not None:
            raise VBANotImplemented(f"reading Window.{name}: how a sheet looks is not modelled")
        return super().vba_member(name, args, kwargs)

    def vba_let_member(self, name, args, kwargs, value):
        member = _canonical(WINDOW, name)
        if member is None or args or kwargs:
            raise VBANotImplemented(f"setting Window.{name} is not modelled")
        passing(self.app, f"Window.{member}", value, WINDOW[member])
        if member == "Zoom":
            self.app.book.active.looks.restyle(1, 1, MAX_ROW, MAX_COL)
        count(self.app, "ActiveWindow")

    def m_smallscroll(self, down=None, up=None, toright=None, toleft=None):
        require(self.app, "Window.SmallScroll")
        for v in (down, up, toright, toleft):
            if v is not None:
                passing(self.app, "Window.SmallScroll", v, WHOLE)
        count(self.app, "ActiveWindow")
        return True


def _overlap(a, b):
    return a[0] <= b[2] and b[0] <= a[2] and a[1] <= b[3] and b[1] <= a[3]


class SheetLooks:
    """What statements passed over leave on one sheet that a macro can see: the cells and rows they format, and the
    rows they give a height, which count in the used range (recorded: formatting a column, the whole sheet or a width
    does not), and where the text a number shows is no longer known."""

    def __init__(self):
        self.cells = None  # the rectangle around the cells formatted
        self.rows = None  # (first, last) of the rows formatted whole
        self.cols = None  # (first, last) of the columns formatted whole
        self.tall = []  # [(first, last)] of the rows given a height of their own, while they keep it
        self.restyled = []  # rectangles where the text a number shows is not known
        self.unsure = False  # formatting passed over was cleared or moved since: the used range is not known

    def formatted(self, r1, c1, r2, c2):
        rows, cols = c1 == 1 and c2 == MAX_COL, r1 == 1 and r2 == MAX_ROW
        if rows and cols:
            return
        if rows:
            self.rows = (r1, r2) if self.rows is None else (min(self.rows[0], r1), max(self.rows[1], r2))
        elif cols:
            self.cols = (c1, c2) if self.cols is None else (min(self.cols[0], c1), max(self.cols[1], c2))
        elif self.cells is None:
            self.cells = (r1, c1, r2, c2)
        else:
            a = self.cells
            self.cells = (min(a[0], r1), min(a[1], c1), max(a[2], r2), max(a[3], c2))

    def heights(self, r1, r2, own):
        """Rows given a height of their own (own), or fitted to what they hold (AutoFit), which takes them out of the
        used range again. A height given to every row is not known to count."""
        if own and r1 == 1 and r2 == MAX_ROW:
            self.unsure = True
            return
        gap = 1 if own else 0  # a span given a height joins one next to it
        kept = [(a, b) for a, b in self.tall if b < r1 - gap or a > r2 + gap]
        if own:
            joined = [(a, b) for a, b in self.tall if (a, b) not in kept]
            kept.append((min([r1] + [a for a, _ in joined]), max([r2] + [b for _, b in joined])))
        else:
            for a, b in self.tall:
                if a < r1 <= b:
                    kept.append((a, r1 - 1))
                if a <= r2 < b:
                    kept.append((r2 + 1, b))
        self.tall = sorted(kept)

    def used(self):
        """The rows and the columns these add to the used range."""
        rows, cols = [], []
        if self.cells is not None:
            rows += [self.cells[0], self.cells[2]]
            cols += [self.cells[1], self.cells[3]]
        rows += list(self.rows or ()) + ([self.tall[0][0], self.tall[-1][1]] if self.tall else [])
        cols += list(self.cols or ())
        return rows, cols

    def touches(self, r1, c1, r2, c2):
        areas = [self.cells] if self.cells is not None else []
        if self.rows is not None:
            areas.append((self.rows[0], 1, self.rows[1], MAX_COL))
        if self.cols is not None:
            areas.append((1, self.cols[0], MAX_ROW, self.cols[1]))
        return any(_overlap(a, (r1, c1, r2, c2)) for a in areas)

    def restyle(self, r1, c1, r2, c2):
        if len(self.restyled) >= 100:  # a macro formatting cell by cell: the whole sheet, rather than a long list
            self.restyled = [(1, 1, MAX_ROW, MAX_COL)]
        else:
            self.restyled.append((r1, c1, r2, c2))

    def text_known(self, r, c):
        return not any(a[0] <= r <= a[2] and a[1] <= c <= a[3] for a in self.restyled)
