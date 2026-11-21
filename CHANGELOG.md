# Changelog

All notable changes to MacroPress are recorded here. The format follows
[Keep a Changelog](https://keepachangelog.com/en/1.1.0/), and versions follow
[Semantic Versioning](https://semver.org/).

## [0.1.0] - not yet released

The first public release.

### Added

- `macropress buttons`, `macropress press` and `macropress convert`, and the same in Python: `macropress.list_buttons`,
  `macropress.press` and `macropress.succeeded`.
- Finding a workbook's buttons: form-control buttons, and shapes and pictures with a macro assigned.
- Converting VBA to readable Python, and running it against a model of the workbook: cells and ranges, formulas that
  recalculate, sheets and named ranges, copy and paste, `Sort`, `AutoFilter`, `WorksheetFunction`, and VBA's own
  rules for values, rounding and errors. See [docs/supported.md](docs/supported.md).
- `--answer` for the questions macros ask with `MsgBox` and `InputBox`, and `--now` for repeatable runs.
- Saving the result to a copy that keeps everything the macros didn't change byte for byte.
- A JSON report of each button's outcome, the cells that changed, and the assumptions about Excel the run relied on.
- Three example workbooks, built from their VBA sources and checked against Excel for Mac.

### Fixed

Found by running some 990 small macros in both Excel and MacroPress and comparing every answer.

- **VBA the converter stopped at or ran wrongly**: `Case Is > 5`; `Enum` blocks (a member counts on from the one
  before it, and `Shade.Red` names one); `Declare` statements and `#If ... #End If`, read as 64-bit Excel for Windows
  reads them; date literals; the `Error` statement; `ParamArray`; the `Mid` statement; `Switch`; an argument left out
  (`f(, 7)`); indexing what a call returns (`Split(s, ",")(0)`); an array element passed `ByRef`, and an array passed
  `ByVal`; `Array()` under `Option Base 1`.
- **Numbers**: `Round` rounds the Double times 10^n, halves to even, as VBA does (`Round(2.675, 2)` is 2.68). A
  `Currency` keeps four decimals and follows VBA's rules for what it gives with other types. A number compared with
  text goes by how each side is declared, as in VBA. `Str`, `CDate`, `DateSerial` with a two-digit year, and text read
  as a number (`"(5)"`, `"5-"`, `"&H10"`) match VBA; `"True"` is no number.
- **Errors**: a member asked of a value that is no object is error 424, and of `Nothing` error 91, where the run used
  to stop with a fault of the converter. A constant MacroPress does not know (`vbSomething`) is refused, where it used
  to be read as an empty variable. A procedure that calls itself without end stops with error 28.
