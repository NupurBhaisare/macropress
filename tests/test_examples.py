"""The example workbooks: each is exactly what its sources build, and its buttons give what Excel gives.

The values below are the ones Excel for Mac 16 gave when it ran the same macros on the same workbooks.
"""

import datetime as dt
import importlib.util
import zipfile
from pathlib import Path

import pytest

import macropress
from macropress import runtime

EXAMPLES = Path(__file__).resolve().parent.parent / "examples"
NOW = dt.datetime(2026, 4, 1, 9, 30)


def _builder():
    spec = importlib.util.spec_from_file_location("build_examples", EXAMPLES / "build.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@pytest.mark.parametrize("name", ["invoice", "orders-cleanup", "sales-summary"])
def test_each_example_workbook_is_what_its_sources_build(name):
    # so the VBA inside each .xlsm is exactly its .bas files: after changing an example, run python examples/build.py
    build = _builder()
    built = {
        part: data.encode() if isinstance(data, str) else data for part, data in build.build_workbook(name).items()
    }
    with zipfile.ZipFile(EXAMPLES / name / f"{name}.xlsm") as z:
        assert {part: z.read(part) for part in z.namelist()} == built


def _press(tmp_path, name, buttons, answers=()):
    out = tmp_path / f"{name}-done.xlsm"
    report = macropress.press(str(EXAMPLES / name / f"{name}.xlsm"), buttons, str(out), now=NOW, answers=answers)
    if not macropress.succeeded(report):
        return report, None
    snap = runtime.Context(str(out)).snapshot()

    def value(sheet, cell):
        v = snap["sheets"][sheet]["cells"].get(cell)
        return None if v is None else v[1]

    return report, (value, snap["view"]["sheets"])


def test_the_invoice_totals_its_lines_after_a_discount(tmp_path):
    report, (value, _) = _press(tmp_path, "invoice", ["Apply discount", "Calculate totals"], answers=["10", "yes"])
    assert [value("Invoice", f"C{r}") for r in range(5, 10)] == [4.05, 6.52, 10.8, 5.76, 22.49]  # Round: half to even
    assert value("Invoice", "D11") == pytest.approx(87.28)
    assert (value("Invoice", "D12"), value("Invoice", "D13")) == (6.98, pytest.approx(94.26))
    assert report["messages"] == [
        "InputBox: Discount on every unit price, in percent:",
        "MsgBox: Lower every unit price by 10%?",
    ]


def test_the_orders_are_cleaned_onto_their_own_sheet_and_filtered(tmp_path):
    _, (value, view) = _press(tmp_path, "orders-cleanup", ["Clean up", "Show North only"])
    assert [value("Clean", f"A{r}") for r in range(2, 9)] == [1002, 1001, 1005, 1004, 1003, 1006, None]
    assert value("Clean", "C2") == "Maple And Finch"
    assert value("Orders", "C2") == "  blue harbor supply"  # the raw orders stay as they were
    assert view["Clean"]["hidden_rows"] == [[2, 2], [4, 4], [6, 6]]  # South, West and East


def test_the_summary_adds_up_each_region_and_email_is_refused(tmp_path):
    _, (value, _) = _press(tmp_path, "sales-summary", ["Build summary"])
    assert [value("Summary", f"A{r}") for r in range(4, 9)] == ["South", "North", "East", "West", "Total"]
    assert (value("Summary", "B8"), value("Summary", "C8")) == (16, 118)
    assert value("Summary", "D8") == pytest.approx(4891.1)
    assert value("Summary", "A2") == "Built 2026-04-01 09:30:00 from 16 sales"
    report, _ = _press(tmp_path, "sales-summary", ["Email summary"])
    assert "CreateObject is not modelled" in report["error"]
