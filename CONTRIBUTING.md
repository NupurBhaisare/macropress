# Contributing to MacroPress

Thank you for helping. The most useful contributions, roughly in order:

1. **A workbook that doesn't work.** Open an issue with the smallest workbook that shows the problem, the JSON
   report, and what Excel does instead. Use made-up data: this repository is public.
2. **Something that isn't modelled yet**: a VBA function, an Excel member, a worksheet function, or a `Format`
   pattern that stopped your macro with `not modelled`.
3. **Checking MacroPress against Excel.** [tests/parity](tests/parity) holds over a thousand small macros with the
   answers Excel gave. If you have Excel, add probes for something that isn't covered, record what Excel answers,
   and see whether MacroPress agrees. Excel for Windows would be especially welcome: the answers so far come from
   Excel for Mac.
4. Better docs and examples.

Everyone taking part is expected to follow the [Code of Conduct](CODE_OF_CONDUCT.md).

## Setting up

You need Python 3.11 and [uv](https://docs.astral.sh/uv/). On macOS and Linux you also need a C compiler, to build
SimpleParse.

```bash
git clone https://github.com/NupurBhaisare/macropress
cd macropress
uv sync                       # creates .venv with MacroPress and its development tools
uv run pytest                 # the tests: they take about half a minute
uv run ruff check .           # lint
uv run ruff format .          # format
```

[docs/how-it-works.md](docs/how-it-works.md) explains the pipeline, and which file does what.

## How a change should look

- **Refuse rather than guess.** When the model can't do something exactly as Excel does, raise `VBANotImplemented`
  with a message that names what is missing. A half-right result that looks complete is worse than a clear stop.
- **Match Excel, not the documentation.** Excel's documented behaviour and its actual behaviour sometimes differ.
  The way to know is a probe in [tests/parity](tests/parity), with Excel's recorded answer. When a change relies on
  Excel behaving a certain way and you haven't checked it in Excel, add an `unverified` entry to the ledger and call
  `note()` where the model relies on it, so every report that depends on it says so.
- **Never write to the input workbook**, and keep everything the model doesn't understand byte for byte in the saved
  copy.
- **Test what changed, in a way that fails without the change.** Most tests run a few lines of VBA and check what
  `Debug.Print` shows or what the cells hold; see [tests/test_semantics.py](tests/test_semantics.py). Group a new
  check with the test for the same behaviour rather than adding a test per case.
- **Keep the style of the code around it.** Docstrings and comments say what something does and why; `ruff` settles
  formatting and lint.

## Where things go

| To add | Put it in | And |
| --- | --- | --- |
| a VBA function | a `b_<name>` function in [builtins.py](macropress/builtins.py) | add its proper-case name to `CANONICAL` in [codegen.py](macropress/codegen.py), and take it out of `NOT_MODELLED` if it was there |
| a member of an Excel object | an `m_<name>` (method), `get_<name>` or `let_<name>` (property) on its class in [excel/](macropress/excel/) | unknown members already stop the run, so there is no list to update |
| a worksheet function in formulas | an `f_<name>` function in [formula.py](macropress/formula.py), added to its `FUNCTIONS` table | formulas using it are then calculated instead of saved for Excel to calculate |
| a `Format` pattern | `b_format` in [builtins.py](macropress/builtins.py) | |
| a setting of how a sheet looks, passed over by `--skip-appearance` | its member and the values Excel takes, in the tables of [excel/looks.py](macropress/excel/looks.py) | record what Excel takes first; a value outside the table stops the run |
| an assumption about Excel | an entry in [ledger.py](macropress/ledger.py) and a `note()` where it is relied on | |
| a check against Excel | a test in [tests/parity/probes.py](tests/parity/probes.py) | record Excel's answers with `tests/parity/record.py` (see [tests/parity/README.md](tests/parity/README.md)) |

## The examples

The workbooks in [examples/](examples/) are built from the `.bas` files beside them, by
[examples/build.py](examples/build.py), which also describes their sheets and buttons. After changing an example, run
`uv run python examples/build.py`; a test fails if a committed workbook doesn't match its sources. If you can, run the
new buttons in Excel too, and check the cells match.

## Checking behaviour against Excel

To check a behaviour, run the same macro on the same workbook in Excel and in MacroPress, and compare the cells. On
macOS and Windows, [xlwings](https://www.xlwings.org) can drive Excel from Python: open a copy of the workbook, run
the macro with `book.macro("Module1.Name")()`, read `sheet.range("A1:J40").value`, and compare it with what
`macropress.press` saved. Dialogs block Excel, so run the macro a button calls after its `MsgBox` or `InputBox`, as
the invoice example's `LowerPrices` allows.

## Good first issues

- Formula functions that are used everywhere but not calculated yet: `TEXT`, `CELL`, `RANDARRAY`. Excel's functions
  live in [macropress/formula.py](macropress/formula.py) (the core, with `LET` and `LAMBDA`),
  [macropress/library.py](macropress/library.py) (rounding, dates, text, statistics, finance) and
  [macropress/arrays.py](macropress/arrays.py) (the ones that make and shape arrays, and `LAMBDA`'s helpers), each
  written to answers recorded from Excel; the [roadmap](docs/roadmap.md) (group 2) lists their edge cases.
- More `Format` patterns (`"yyyy-mm-dd hh:nn"`, `"0%"`, `"mmm yyyy"`, ...).
- `Range.Delete` and `Range.Insert`, moving every reference to the cells they move.
- User-defined types (`Type ... End Type`): today they fail as a converter error, and should at least stop as
  `not modelled`.
- Listing ActiveX command buttons with the form-control buttons.
- A VBA parser that runs on Python 3.12 and later, which would lift MacroPress's biggest limit.

## Pull requests

Before you open one:

- [ ] `uv run pytest` passes, and the new tests fail without your change.
- [ ] `uv run ruff check .` and `uv run ruff format --check .` pass.
- [ ] New behaviour that rests on an unchecked assumption about Excel has a ledger entry.
- [ ] [docs/supported.md](docs/supported.md) and the [CHANGELOG](CHANGELOG.md) mention what changed for users.

By contributing, you agree that your contributions are licensed under the project's [MIT license](LICENSE).
