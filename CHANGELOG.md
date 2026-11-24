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

- A parity suite (`tests/parity`): over a thousand small macros whose 3,700 answers were recorded from Excel itself
  and are compared with MacroPress's on every test run. The answers MacroPress knowingly gives otherwise are listed
  with their reasons, so a difference can neither appear nor go away unnoticed.
- What recorded macros do in their first lines ([roadmap](docs/roadmap.md), group 1), each recorded from Excel
  first (162 more tests in the parity suite) and the saved copy checked by opening it in Excel:
  - **Setting `NumberFormat`** on cells, whole rows, whole columns and the whole sheet, the code read back as Excel
    reads it; a range down to the last row formats its columns; several formats read as `Null`. Rows and columns
    formatted whole are saved as Excel saves them, and count in `UsedRange`.
  - **Protection**: sheet and workbook protection, read from the file and set by `Protect`, with the passwords Excel
    keeps (either form of hash) and the `Allow...` options; `Locked`; what a protected sheet refuses, and what it lets
    a macro change, as Excel does. `Unprotect` without the password asks for it (`--answer`).
  - **`Visible`**: hiding and showing sheets, very hidden too.
  - **`FormulaR1C1`** and `Formula2R1C1`, read and written, and R1C1 text given to `Formula`.
  - **Filling**: `FillDown`, `FillUp`, `FillRight`, `FillLeft`; a copy onto a larger range; `AutoFill` of formulas,
    copies and number and date series; `Cut`, with the formulas and names that read the cells following them; and
    `Worksheet.Paste`.
- Arrays in formulas ([roadmap](docs/roadmap.md), 2.1), recorded from Excel first (42 more tests in the parity
  suite): a range where one value is wanted gives the formula's own row or column of it, a `Formula2` formula works
  on whole arrays and spills, array constants, `@`, `SUMPRODUCT`, `ROW`, `COLUMN`, `ROWS`, `COLUMNS` and
  `CONCATENATE`; `IF`, `INDEX` and `XLOOKUP` hand back cells. `Formula` and `Formula2` show a formula as Excel shows
  it, `@` and all, and formula text is kept as Excel keeps it. A dynamic-array formula is saved as one, in any
  workbook.
- Conditional sums and counts ([roadmap](docs/roadmap.md), 2.2), recorded from Excel first (19 more tests in the
  parity suite): `SUMIF`, `SUMIFS`, `COUNTIF`, `COUNTIFS`, `AVERAGEIF`, `AVERAGEIFS`, `MAXIFS` and `MINIFS`, in
  formulas and from `WorksheetFunction`. Criteria find what Excel finds: numbers and the text that reads as them,
  operators, wildcards, blanks, `TRUE` and errors. `SUMIF`'s sum range is read over the criteria range's size, and
  the formula is kept current through all of it. A formula giving them a constant where a range is wanted, or the
  wrong number of arguments, is refused as Excel refuses it.
- Lookups ([roadmap](docs/roadmap.md), 2.3), recorded from Excel first (34 more tests in the parity suite):
  approximate matches in `VLOOKUP`, `HLOOKUP`, `MATCH` and `LOOKUP`, wildcards in exact ones, `MATCH` with -1, the
  match and search modes of `XLOOKUP` and `XMATCH`, `CHOOSE`, `OFFSET`, `INDIRECT`, and ranges that end on a
  function's cells (`A1:INDEX(A:A,5)`). Over values not in order a match is answered only where Excel's search cannot
  land elsewhere. The same from `WorksheetFunction`, which takes VBA arrays as tables; `Index` and `XLookup` hand back
  a `Range`.
- Rounding, dates, text and the rest ([roadmap](docs/roadmap.md), 2.4), recorded from Excel first (97 more tests in
  the parity suite, 1,097 answers), in formulas: `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `CEILING`, `FLOOR` and their newer
  forms, `MROUND`, `MOD`, `QUOTIENT`, `POWER`, `SQRT`, `LN`, `LOG`, `EXP`, `EVEN`, `ODD`, `PRODUCT`; `DATE`, `YEAR`,
  `MONTH`, `DAY`, `WEEKDAY`, `EOMONTH`, `EDATE`, `DATEDIF` (its `"md"` as Excel has it), `NETWORKDAYS` and `WORKDAY`
  with their `.INTL` forms, `YEARFRAC` on its five bases, `DAYS360`, `DAYS`, `WEEKNUM`, `ISOWEEKNUM`, `TIME`, `HOUR`,
  `MINUTE`, `SECOND`; `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `REPT`, `EXACT`, `PROPER`, `TEXTJOIN`, `CHAR`, `CODE`,
  `CLEAN`, `VALUE`, `N`, `T`; `IFS`, `SWITCH`, `IFNA`, the `IS` functions, `TYPE`, `ERROR.TYPE`; `MEDIAN`, `MODE`,
  `LARGE`, `SMALL`, `RANK`, `PERCENTILE`, `QUARTILE`, `STDEV`, `VAR`, `COUNTBLANK`; `PMT`, `IPMT`, `PPMT`, `PV`, `FV`,
  `NPER`, `NPV`, `XNPV`, `SLN`, `SYD`, `DB`, `DDB`; and `SUBTOTAL` and `AGGREGATE`, which pass over hidden rows and
  nested totals as Excel does and are worked out again when rows are hidden, shown or filtered. A formula entered in
  a General cell takes the number format Excel gives it (`DATE` a date, `PMT` a currency, `IRR` a percentage), so its
  `Value` is a `Date` or a `Currency` as in Excel. `IRR`, `XIRR` and `RATE`, whose answers stop wherever Excel's
  search took them, are left for Excel.
- Worksheet functions from VBA ([roadmap](docs/roadmap.md), 2.5), recorded from Excel first (63 more tests in the
  parity suite): every function MacroPress calculates is reached as `WorksheetFunction.X` and as `Application.X`
  (`StDev_S` for `STDEV.S`), where Excel lets VBA reach it, and error 438 where it does not. `Application.X` hands
  an error back as its value, `WorksheetFunction.X` raises error 1004. Arrays come back from 1, ranges of cells as a
  `Range`; `Transpose`; `WorksheetFunction.Round` is Excel's. `Evaluate`, `Application.Evaluate`,
  `Worksheet.Evaluate` and `[A1+1]` work a formula's text out against the active sheet.
- Passing over how a sheet looks, when told to ([roadmap](docs/roadmap.md), 1.2), recorded from Excel first (81 more
  tests in the parity suite): `--skip-appearance` (`skip_appearance=True`) lets a run go past statements that only
  set fonts, fills, borders, alignment, widths and heights, tab colours and the window, where Excel takes the value
  given, rather than stop at them, and the report counts them under `appearance_skipped`. What they change that a
  macro sees is kept: a protected sheet refuses them, a width or height of 0 hides and another shows again, formatted
  cells and rows given a height count in `UsedRange`, and a number's `Text` is not guessed once its font or column
  changed. Reading how a sheet looks still stops the run. The constants these statements use (`xlThin`, `xlCenter`,
  `xlThemeColorAccent1`, ...) are known.

### Fixed

Found by running those macros in both Excel and MacroPress and comparing every answer.

- **`AutoFit`** did nothing. Excel shows hidden rows and columns again (not rows a filter hides), refuses a range that
  is neither rows nor columns (`Range("A1").AutoFit`, error 1004) and is refused on a protected sheet; so does
  MacroPress now. `Text` of a number in a column `AutoFit` may have narrowed gave the number, where Excel can show
  `#`. A row hidden by hand now counts in `UsedRange`, and `Range("A1:B3").Columns(2)` is one column (its `Count` is
  1, `For Each` goes through it once), where it counted three cells.

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
- **From VBA**: `WorksheetFunction.Sum`, `Max`, `Min` and `CountA` passed over `Empty`, which Excel counts as 0;
  `[A1+1]` stopped the conversion of its module, where `[A1]` worked; `TypeName` of a `Double` array was `Variant()`.
  `=SUM()`, `=ROUND(1)` and other core functions with the wrong number of arguments are refused as Excel refuses them.
- **Formulas entered by a macro**: `=TODAY()` and `=NOW()` take Excel's date formats, where the cell stayed General;
  a function given the wrong number of arguments, or `#SPILL!` or `#CALC!` typed in, is refused as Excel refuses it;
  the file keeps `_xlfn.` before the names Excel needs it on (`CEILING.MATH`, `STDEV.S`, `AGGREGATE`, ...), where
  Excel would have shown `#NAME?`.
- **Formulas**: numbers are compared to 15 significant digits, as Excel compares them (`=0.1+0.2=0.3` is TRUE), and a
  formula or a `SUM` that ends on a sum which all but cancels gives 0 (`=0.1+0.2-0.3`). `INT` and `ROUND` work on the
  15-digit number (`=INT(4.35*100)` is 435). A number written with more than 15 digits keeps the first 15. An odd root
  of a negative number has a value (`=(-8)^(1/3)`), `0^0` and a result too large are `#NUM!`. Text in a formula's
  arithmetic is read as Excel reads typed text (`"5%"+1`, `"1,000"+1`). `SUM`, `COUNT`, `MAX`, `MIN`, `AVERAGE`, `AND`
  and `OR` count text and logical values written in the formula and pass over those read from cells, and an argument
  left empty is a 0. A formula that wants one value and is given a range takes the cell in its own row or column, and
  `INDEX` with a 0 gives a whole row or column. `WorksheetFunction.Sum`, `Max` and `Min` take values the same way.
- **Lookups**: `INDEX` works an expression it is given out whole in any formula, as Excel does
  (`INDEX((A1:A9=1)*B1:B9, 0)` used to take one row of it); `WorksheetFunction.XLookup` hands back a `Range` and
  takes its match and search modes.
- **Calculation**: under manual calculation a formula keeps its value until something calculates, and one entered then
  is calculated once from what cells show; a workbook saved in manual mode opens in it. A cell filled by an array
  formula entered with Ctrl+Shift+Enter is no longer read with a value its formula has outgrown.
- **Speed**: a write no longer looks at every formula in the workbook to find the ones that read the cell, so a macro
  that writes many cells of a workbook with many formulas takes time in proportion to what it changes.
- **Cells**: text written to a cell is read as Excel reads typed text: `"$5"`, `"5%"`, `"1,234"`, `"(5)"`, `"1e3"`,
  `"1 1/2"` and `"#N/A"` become values with the number format Excel gives them, `"1,23"` and `"5-"` stay text, and
  a date or a time is refused (Excel reads those by regional settings). A cell formatted as Text keeps what is
  written to it. The number formats of empty cells, rows and columns are read, and decide what a write leaves and
  what VBA reads: a `Date` for a format that shows a day, a month or a year, a `Double` for a time alone, a `Currency`
  for Excel's own currency formats. A write replaces one of Excel's own formats of another kind (a date written to a
  cell of percentages) and leaves a plain number format or one made by hand in place. `Value2` reads and writes the
  bare number; a `Currency` written through `Value` arrives rounded to cents. `Clear` resets the number format, and a
  copy brings it, from an empty cell too. `Text` gives what the cell shows for General and plain number formats, and
  is refused where the column's width would decide it.
- **Filters and sorting**: a criterion is matched against what each cell shows, without regard to case or to the
  spaces around it; `<>` a number goes by value. A value, an array or a clear on a filtered range reaches only the
  rows the filter shows. `ShowAllData` with nothing filtered is error 1004, and an `AutoFilter` with a field the range
  does not have leaves no filter behind. `Range.End` passes over hidden rows and columns. A sort puts errors after
  logical values and the empty text first among text.
- **The saved copy**: a changed cell is saved with the number format the model has for it, a new cell takes its
  column's style, an array formula entered with Ctrl+Shift+Enter is saved as one when its value changes, and a number
  too large for Excel is refused rather than written.
- **Running**: `press` takes one button given as a string, paths given as `Path` objects, `environ` for what
  `Environ` returns (`--env NAME=VALUE`) and `timeout` to stop macros that run too long (`--timeout`). A copy that
  could not be saved, over the workbook itself or with another extension, is refused before the macros run instead of
  after. The report lists answers no question took (`answers_unused`), and says when a module's `#If` blocks were
  read as Windows reads them.
- **More from the last recording**: an `AutoFilter` that is given a criterion takes in the rows of data directly
  below its range; `>` and `<` in a filter compare numbers with numbers and text with text, so text that reads as a
  number is not "more than 500"; `INDEX` with one index into a block of several rows and columns is `#REF!`. Every
  behaviour the report used to call unverified has now been checked against Excel.
- **Found while modelling what recorded macros do first**: a write to a protected sheet went through as if it were
  not protected. `NumberFormat` of several cells gave the first cell's format instead of `Null`. A dynamic array was
  kept from spilling by an empty cell that only had a format, and the cells it filled lost theirs. A paste or `Clear`
  of whole rows or columns changed their formats in the model but not in the saved copy. `CutCopyMode` with nothing
  copied gave `False` instead of 0.
- **Found with the arrays in formulas**: `Formula2` read a formula written with `Formula` without the `@` Excel shows
  in it, and formula text was kept as typed (`=sum(a1)`) where Excel keeps `=SUM(A1)`. A dynamic array that spilled a
  range left out the range's blank cells past its data. A `Formula2` formula of one value that takes ranges whole
  was saved as a formula that is not a dynamic array, which Excel then works out another way.

### Changed

- The docs say what MacroPress does not do ([README](README.md#what-it-does-not-do)), how cells, number formats and
  regional settings are treated ([docs/supported.md](docs/supported.md)), and where its one known difference in the
  language lies: a single whole-number type where VBA has `Integer` and `Long`.

