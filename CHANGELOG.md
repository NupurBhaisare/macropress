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
