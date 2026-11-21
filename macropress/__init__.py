"""MacroPress: press the buttons of an Excel workbook without Excel.

It reads the workbook's VBA, converts it to Python, and runs it against a model of the workbook (cells, formulas,
sheets, ranges, sorting, filters), then saves the result into a copy of the workbook. What it does not model stops
the run with "not modelled" instead of a guess.

    import macropress
    macropress.list_buttons("invoice.xlsm")
    macropress.press("invoice.xlsm", ["Calculate totals"], out="invoice-done.xlsm")

The command line does the same: macropress buttons invoice.xlsm, macropress press invoice.xlsm --button ... --out ...
"""

from .api import list_buttons, press, succeeded

__version__ = "0.1.0"
__all__ = ["list_buttons", "press", "succeeded", "__version__"]
