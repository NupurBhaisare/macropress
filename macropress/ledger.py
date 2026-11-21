"""Behaviours the model relies on, and which of them a run actually used.

status:
  verified    checked against real Excel output (Excel for Mac 16: values and saved view state)
  unverified  believed correct, not yet checked against Excel
  limitation  the model knowingly differs from Excel here
"""

from collections import Counter

ASSUMPTIONS = {
    "GROUP-CLEARCONTENTS": (
        "verified",
        "Selection.ClearContents while several sheets are grouped clears the range on every grouped sheet.",
    ),
    "GROUP-SELECT-ACTIVE": (
        "unverified",
        "Sheets(Array(...)).Select makes the first sheet named in the array the active sheet.",
    ),
    "GROUP-EXTEND-KEEPS-ACTIVE": (
        "unverified",
        "Sheets(...).Select Replace:=False adds sheets to the group and keeps the current active sheet.",
    ),
    "GROUP-SELECT-RANGE": (
        "verified",
        "Range.Select while sheets are grouped selects the same cells on every grouped sheet.",
    ),
    "AF-EXISTING-RANGE": (
        "verified",
        "Range.AutoFilter on a sheet that already has an AutoFilter uses the existing filter range, "
        "not the range named in the code.",
    ),
    "AF-RANGE-NOT-EXTENDED": (
        "unverified",
        "A re-applied AutoFilter keeps its saved range even when data continues below it (Excel may extend it).",
    ),
    "AF-HEADER-ROW": (
        "unverified",
        "Range.AutoFilter on a single row (a header row, or Rows(1)) filters the data below it: the row's current "
        "region, across the row's columns.",
    ),
    "AF-REFILTER-UNHIDES": (
        "unverified",
        "Re-applying an AutoFilter shows every row of its range that passes, including rows hidden by hand.",
    ),
    "AF-MATCH-TEXT": ("verified", "AutoFilter criteria match the cell's text, ignoring case, with * ? ~ wildcards."),
    "END-IGNORES-HIDDEN": ("unverified", "Range.End treats hidden rows like visible ones."),
    "COPY-VISIBLE-IN-FILTER": ("verified", "Copying a range on a filtered sheet copies only the visible rows."),
    "COPY-HIDDEN-NO-FILTER": (
        "unverified",
        "Copying a range with hidden rows on a sheet that is not filtered copies the hidden rows too.",
    ),
    "PASTE-WHOLE-ROWS": (
        "verified",
        "Pasting whole copied rows replaces the whole target rows; blank source cells clear target cells.",
    ),
    "WRITE-TEXT-CONVERSION": (
        "unverified",
        "Text written to a cell is converted the way typed input is: numeric text becomes a number, "
        "'TRUE'/'FALSE' a Boolean, a leading apostrophe forces text.",
    ),
    "WRITE-EMPTY-STRING": ("unverified", 'Writing "" to a cell leaves it empty.'),
    "CMP-NUMBER-TEXT": (
        "unverified",
        "Comparing a number with text that is not numeric gives 'number < text' without an error.",
    ),
    "FORMULAS-RECALCULATED": (
        "unverified",
        "Formulas whose inputs changed are recalculated before their values are read and before saving "
        "(automatic calculation), with Excel's rules for the functions the evaluator covers.",
    ),
    "FORMULA-NOT-EVALUATED": (
        "limitation",
        "A formula using a function or syntax the evaluator does not cover is saved without a value "
        "(Excel calculates it when the file is opened); a macro reading its value stops.",
    ),
    "CALC-MANUAL": (
        "unverified",
        "With Calculation set to manual, a stale formula shows its old value until the workbook is saved.",
    ),
    "SPILL-BLANK-AS-ZERO": ("unverified", "A blank cell in a dynamic-array result (SORT, UNIQUE, ...) spills as 0."),
    "SORT-TEXT-ORDER": (
        "verified",
        "Sorting (Sort, SORT) orders numbers, then text, then logical values, then errors, with blanks last; text "
        "compares character by character ignoring case (accents only break ties), symbols before digits before "
        "letters, hyphens and apostrophes included. Checked in Excel for Mac; Windows Excel may ignore hyphens and "
        "apostrophes.",
    ),
    "SORT-HIDDEN-ROWS": (
        "verified",
        "Sorting a range with hidden rows (filtered or hidden by hand) leaves the hidden rows where they are and "
        "sorts only the visible rows among themselves.",
    ),
    "FORMATS-NOT-MODELLED": (
        "limitation",
        "Column widths, fonts and other formatting are not modelled (AutoFit and similar do nothing).",
    ),
    "ON-ERROR-RESUME-NEXT": (
        "verified",
        "A statement that fails under On Error Resume Next is skipped and its variable keeps its old value.",
    ),
    "MSGBOX-NO-UI": ("limitation", "MsgBox with only an OK button is logged and returns vbOK."),
    "ROUND-DECIMAL": ("unverified", "Round() rounds halves to even on the decimal form of the number."),
    "CDEC-WINDOWS-ONLY": ("limitation", "CDec is evaluated, although Mac VBA reportedly cannot compile it."),
    "NUMBER-TEXT-FORMAT": ("unverified", "Numbers are turned into text with '.' as the decimal separator."),
    "EVENTS-NOT-RUN": (
        "limitation",
        "Event procedures (Worksheet_Change, Workbook_Open, ...) are not run. Excel runs them when the workbook opens "
        "or a macro changes its cells, unless the macro turns events off.",
    ),
}


class Ledger:
    def __init__(self):
        self.counts = Counter()
        self.details = {}

    def note(self, key, detail=None):
        if key not in ASSUMPTIONS:
            raise KeyError(f"unknown assumption {key}")
        self.counts[key] += 1
        if detail and key not in self.details:
            self.details[key] = detail

    def report_lines(self):
        lines = []
        for key, count in sorted(self.counts.items(), key=lambda kv: (ASSUMPTIONS[kv[0]][0], kv[0])):
            status, text = ASSUMPTIONS[key]
            extra = f" (first: {self.details[key]})" if key in self.details else ""
            lines.append(f"- `{key}` [{status}] x{count}: {text}{extra}")
        return lines
