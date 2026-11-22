"""How a sheet stores its cells, and a table (ListObject) over them."""


# ---------------------------------------------------------------------------
# Storage
# ---------------------------------------------------------------------------


class Cell:
    """A cell the sheet holds a record of: what it holds, its number format (None for General), and whether it is
    locked, as a protected sheet asks (cells are, unless their style says not)."""

    __slots__ = ("value", "formula", "numfmt", "locked")

    def __init__(self, value=None, formula=None, numfmt=None, locked=True):
        self.value, self.formula, self.numfmt, self.locked = value, formula, numfmt, locked

    def copy(self):
        return Cell(self.value, self.formula, self.numfmt, self.locked)

    @property
    def style(self):
        return self.numfmt, self.locked


NORMAL = (None, True)  # the style of a cell nothing formats: General, locked


class Table:
    def __init__(self, name, sheet, r1, c1, r2, c2, columns, header_rows=1, totals_rows=0):
        self.name, self.sheet = name, sheet
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2
        self.columns, self.header_rows, self.totals_rows = columns, header_rows, totals_rows
