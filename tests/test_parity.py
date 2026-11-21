"""What Excel answers, MacroPress answers: every probe of tests/parity/probes.py, set beside the answer Excel gave.

Excel's answers were recorded from Excel itself (tests/parity/excel.json; how, and how to add a probe, is in
tests/parity/README.md). Where MacroPress knowingly answers otherwise, or stops rather than answer,
tests/parity/known.json says so with the reason, and the test holds MacroPress to exactly that: a difference can
neither appear nor go away unnoticed.
"""

import importlib.util
from pathlib import Path

import pytest

PARITY = Path(__file__).resolve().parent / "parity"
_spec = importlib.util.spec_from_file_location("parity_run", PARITY / "run.py")
run = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(run)

EXCEL, KNOWN = run.load()
MODULES = sorted({module for module, _name, _lines in run.workbook.load("probes", PARITY / "probes.py").TESTS})


@pytest.fixture(scope="module")
def probes(tmp_path_factory):
    return run.Probes(tmp_path_factory.mktemp("parity"))


@pytest.mark.parametrize("module", MODULES)
def test_macropress_answers_as_excel_does(module, probes):
    known = run.known_tables(KNOWN)
    wrong = []
    for name, outcome in probes.answers(module).items():
        wrong += run.verdict(name, outcome, EXCEL, known)
    assert not wrong, "\n".join(wrong)


def test_every_known_difference_has_been_looked_at_and_names_a_test():
    names = {name for _module, name, _lines in run.workbook.load("probes", PARITY / "probes.py").TESTS}
    listed = [name for group in KNOWN["differences"] for name in group["tests"]] + list(KNOWN["refused"])
    assert not [group["why"] for group in KNOWN["differences"] if group["why"] == run.NOT_LOOKED_AT]
    assert not sorted(set(listed) - names) and not sorted(names - set(EXCEL))
