# Parity with Excel

MacroPress claims to do what Excel does. This folder is where that claim is checked: over a thousand small macros
(the *probes*) were run in Excel itself, each writing down what it found, and the test suite runs the same macros in
MacroPress and compares every answer.

| File | What it is |
| --- | --- |
| `probes.py` | The probes, and the sheets of the workbook they run on. |
| `excel.json` | What Excel answered, one test per line. The first line says which Excel recorded it. |
| `known.json` | The answers MacroPress knowingly gives otherwise, each group with its reason, and the tests it stops at rather than answer. |
| `workbook.py` | Builds the probe workbook from `probes.py`. |
| `run.py` | Runs the probes in MacroPress and compares. The test (`tests/test_parity.py`) uses it. |
| `record.py` | Runs the probes in Excel and writes `excel.json`. Run by hand; the tests never start Excel. |

## A probe

```python
T("TNum", "round1", "N: Round(2.5)", "N: Round(2.675, 2)", "x = 1 / 3", "S: CStr(x)")
```

A test has a module, a name and lines of VBA. A line that starts with `S:`, `N:`, `V:`, `E:` or `X:` is a probe: it
writes one answer (as text, as a number, as type and value, as the error number of a statement) to the test's row of
the `Out` sheet. Each test runs under `On Error Resume Next`, on a fresh copy of the workbook in MacroPress and one
after another in Excel, so a test must not lean on what another left behind.

## When an answer differs

`python tests/parity/run.py` lists the answers that differ from Excel's and are not in `known.json`. Either MacroPress
is wrong, and is fixed, or the difference is one to live with, and is written down:
`python tests/parity/run.py --write` writes `known.json` again from what MacroPress answers now, with new entries
under the reason "not looked at". The test fails until that reason is replaced by a real one.

The answers were recorded on a Mac with day-first regional settings. Some of what is in `known.json` is that Mac and
not MacroPress, which follows US settings and, where Excel for Mac and Excel for Windows differ, Windows: the reasons
say which.

## Adding a probe, and recording again

Add the test to `probes.py`, then record Excel's answers for its module:

```console
$ pip install xlwings
$ python tests/parity/record.py TNum
```

This needs Excel. On a Mac the workbook is opened from Excel's own folder
(`~/Library/Group Containers/UBF8T346G9.Office/macropress-parity`), where Excel opens a file without asking for
access to it; it may still ask whether to enable the macros. Without arguments every test is recorded, which takes
about a minute.
