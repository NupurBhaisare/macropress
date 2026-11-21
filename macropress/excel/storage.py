"""How a sheet stores its cells, and a table (ListObject) over them."""


# ---------------------------------------------------------------------------
# Storage
# ---------------------------------------------------------------------------


class Cell:
    __slots__ = ("value", "formula", "numfmt")

    def __init__(self, value=None, formula=None, numfmt=None):
        self.value, self.formula, self.numfmt = value, formula, numfmt

    def copy(self):
        return Cell(self.value, self.formula, self.numfmt)


class Table:
    def __init__(self, name, sheet, r1, c1, r2, c2, columns, header_rows=1, totals_rows=0):
        self.name, self.sheet = name, sheet
        self.r1, self.c1, self.r2, self.c2 = r1, c1, r2, c2
        self.columns, self.header_rows, self.totals_rows = columns, header_rows, totals_rows
