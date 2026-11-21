# What MacroPress models

MacroPress models the parts of VBA and Excel that macros behind buttons use. Anything outside this list stops the
button with `stopped: not modelled` and a message naming what was missing, rather than being skipped or guessed.

What is modelled is checked against Excel itself: [tests/parity](../tests/parity) holds about a thousand small macros
with the answers Excel gave, and every test run compares MacroPress's answers with them. Where MacroPress knowingly
differs from Excel, a run that meets the difference says so in its report, under `unverified_behaviours`;
[macropress/ledger.py](../macropress/ledger.py) lists them all.

## The VBA language

- `Sub` and `Function` procedures, with `ByRef` and `ByVal` arguments, `Optional` arguments and named arguments
  (`Key1:=...`), calls between modules (`Module1.DoThis`) and recursion
- `Dim`, `ReDim` (with `Preserve`), `Const`, `Set`, module-level variables, `Option Explicit`, `Option Compare Text`
- `Enum` blocks, date literals (`#1/5/2024#`), `ParamArray`, and `#If ... #End If`, read as 64-bit Excel for Windows
  reads it (`VBA7`, `Win64` and `Win32` true, `Mac` false)
- `If` / `ElseIf` / `Else` (block and single-line), `Select Case` (values, lists, `x To y` ranges and `Is > x`),
  `For ... Next` (with `Step`), `For Each`, `Do ... Loop` (`While` and `Until`, before or after), `While ... Wend`,
  `With`, `GoTo` and labels, `Exit`, `End`, `Erase`, the `Mid` statement and the `Error` statement
- `On Error GoTo label`, `On Error Resume Next`, `On Error GoTo 0`, `Resume`, `Resume Next` and the `Err` object, with
  VBA's error numbers (for example 9, subscript out of range; 11, division by zero; 13, type mismatch; 1004 from Excel)
- Variants and VBA's conversions between them, `Empty`, `Null`, `Nothing`, error values, arrays with any bounds,
  `Currency` with its four decimals and its own rules of arithmetic, and dates as VBA stores them
- Whole numbers, with one difference from VBA: MacroPress has a single whole-number type. A variable declared
  `Integer` or `Long` overflows where VBA's does, but `TypeName` and `VarType` go by the number's size (a small
  `Long` reads as `Integer`), and two Integers whose sum or product passes 32,767 give the number, where VBA stops
  with error 6
- Sheet and `ThisWorkbook` modules: a sheet module's unqualified `Range` is its own sheet, `Me`, and code names
  (`Sheet1.Range("A1")`)

Not modelled: class modules (`New MyClass`), `Property` procedures, user-defined types (`Type ... End Type`),
UserForms, calling a function that a `Declare` statement brings in from a library (the statement itself is read),
file statements (`Open`, `Print #`, `Close`), and `GoSub`. A constant of VBA or Excel that MacroPress does not know
(`vbSomething`, `xlSomething`) stops the conversion: it is never read as an empty variable.

## VBA functions

Abs, Array, Asc, CBool, CByte, CCur, CDate, CDbl, CDec, CInt, CLng, CSng, CStr, CVErr, CVar, Choose, Chr, Date,
DateAdd, DateDiff, DateSerial, DateValue, Day, DoEvents, Exp, Fix, Format, Hour, IIf, InStr, InStrRev, InputBox, Int,
IsArray, IsDate, IsEmpty, IsError, IsMissing, IsNull, IsNumeric, IsObject, Join, LBound, LCase, LTrim, Left, Len, Log,
Mid, Minute, Month, MsgBox, Now, RGB, RTrim, Replace, Right, Round, Second, Sgn, Space, Split, Sqr, Str, StrComp,
StrReverse, String, Switch, Time, TimeSerial, TimeValue, Timer, Trim, TypeName, UBound, UCase, Val, VarType, Weekday,
Year

Some come with limits, and stop the run outside them:

- `Format` takes these format strings: `yyyy-mm-dd`, `dd/mm/yyyy`, `mm/dd/yyyy`, `yyyymmdd`, `hh:mm`, `hh:mm:ss`,
  `hh:nn:ss`, `yyyy-mm-dd hh:mm:ss`, `yyyy-mm-dd hh:nn:ss`, `0`, `0.0` to `0.0000`, `#,##0`, `#,##0.00`, `Fixed`,
  `Standard` and `@`.
- `MsgBox` with only an OK button is noted in the report and returns `vbOK`. One that asks a question (`vbYesNo`,
  `vbOKCancel`, ...) and `InputBox` take their answers from `--answer`, in order: `yes`, `no`, `ok`, `cancel`, or a
  number for `MsgBox`, and any text for `InputBox`.
- `Now`, `Date`, `Time` and `Timer` start from `--now` (or the real time) and move on by a millisecond each time they
  are read, so a loop that waits for the time to pass ends. `Application.Wait` moves the clock to the time it waits for.
- `Rnd` stops the run (its numbers cannot be repeated). `Environ("NAME")` returns what `--env NAME=VALUE` gives it, and
  stops the run for a name it was not given (its values depend on the machine).
- Text is read as a number the way VBA reads it (`"1,234"`, `"(5)"`, `"5-"`, `"&H1F"`, `"1D3"`), but text with a
  currency sign stops the run: whether `"$5"` is a number depends on the machine's regional settings.
- A run that would not end is stopped by `--timeout`; a procedure that calls itself without end stops with VBA's
  error 28.

## Excel objects

| Object | Members |
| --- | --- |
| Application (and the globals) | ActiveCell, ActiveSheet, ActiveWorkbook, Calculate, Calculation, Cells, Columns, CutCopyMode, DisplayAlerts, DisplayStatusBar, EnableEvents, Intersect, Range, Rows, ScreenUpdating, Selection, Sheets, StatusBar, ThisWorkbook, Wait, Workbooks, WorksheetFunction, Worksheets |
| Workbook | Activate, ActiveSheet, Name, Names, Save (the copy is saved when the run ends), Sheets, Worksheets |
| Sheets, Worksheets | Count, Item, Select (selecting several sheets groups them, as in Excel) |
| Worksheet | Activate, AutoFilterMode, Calculate, Cells, CodeName, Columns, FilterMode, Index, Name (renaming a sheet updates the formulas and names that refer to it), Parent, Range, Rows, Select, ShowAllData, Sort, UsedRange, Visible (read only) |
| Range | Activate, Address, AutoFilter, AutoFit (does nothing: widths are not modelled), Calculate, Cells, Clear, ClearContents, Column, Columns, Copy, Count, CurrentRegion, End, EntireColumn, EntireRow, Formula, Formula2, Hidden, Item, NumberFormat (read only), Offset, Parent, PasteSpecial, Resize, Row, Rows, Select, Sort, Text, Value, Value2, Worksheet |
| Sort (Worksheet.Sort) | SortFields (Add, Add2, Clear, Count), SetRange, Header, MatchCase, Orientation, SortMethod, Apply |
| Name | Name, RefersTo, RefersToRange |
| WorksheetFunction | CountA, Find, Max, Min, Search, Sum, XLookup |

Within those, a few cases stop the run: `PasteSpecial` with `Transpose`, `SkipBlanks` or an `Operation`, or into a
range of another size; `Sort` by colour, with a custom order or left to right; `AutoFilter` with criteria other than
values, comparisons and wildcards; `Address` in R1C1 style; `XLookup` and `VLOOKUP` with approximate matching; and
ranges with several areas (`Range("A1:B2,D1:D5")`, `Union`).

## Cells and their number formats

- **Text written to a cell** is read as Excel reads typed text. `"$5"`, `"5%"`, `"1,234"`, `"(5)"`, `"1e3"`,
  `"1 1/2"`, `"TRUE"` and `"#N/A"` become values, with the number format Excel gives the cell; `"1,23"`, `"5-"` and
  `" TRUE"` stay text; a leading apostrophe keeps anything as text. Text that Excel would turn into a date or a time
  (`"1/2"`, `"Jan 5"`, `"3:30"`) stops the run, because Excel reads it by the machine's regional settings. Write a
  `Date` value instead (`DateSerial(2024, 1, 5)`).
- **Number formats are read, not set.** The formats of cells, rows and columns in the workbook decide what VBA reads
  and what a write leaves. `Value` gives a `Date` for a format that shows a day, a month or a year, a `Double` for a
  time alone, and a `Currency` (four decimals) for Excel's own currency formats; `Value2` gives the bare number. A
  cell formatted as Text keeps what is written to it as text. A write replaces one of Excel's own formats of another
  kind (a date written to a cell of percentages) and leaves a plain number format, or one made by hand, in place.
  `Clear` resets the format, `ClearContents` keeps it, and a copy brings it along.
- **`Text`** gives what the cell shows for text, for the General format and for plain number formats (`0.00`,
  `#,##0.00`). It stops the run for any other format, and wherever the column's width would decide what shows:
  Excel rounds a number that does not fit, or shows `####`.
- **A filtered range**: a value, an array or a clear reaches only the rows the filter shows, as in Excel, and a
  filter matches the text each cell shows (`1,250.00`, not `1250`), without regard to case or to the spaces around it.

Not modelled: `Range.Delete` and `Range.Insert` (moving cells would have to move every reference to them),
`Range.Find`, `Range.SpecialCells`, `Sheets.Add`, `Workbooks.Add`, `Workbooks.Open`, `Workbook.SaveAs`,
`Workbook.Close`, `Application.Run`, setting any format (`NumberFormat`, `Font`, `Interior`, `Borders`), charts,
pivot tables, shapes and comments.

## Formulas

Formulas a macro writes, and formulas already in the workbook, are calculated as Excel would, and recalculated when
the cells they read change: before a macro reads them, and before the copy is saved. Under manual calculation
(`Application.Calculation = xlCalculationManual`, or a workbook saved that way) they keep their values until
`Calculate`, the return to automatic calculation, or the save.

Excel's own habits with numbers are followed: they are compared to 15 significant digits (`=0.1+0.2=0.3` is TRUE), a
formula or a `SUM` that ends on a sum which all but cancels gives 0, text in arithmetic is read as typed text
(`="5%"+1`), and `SUM`, `COUNT` and the like count text and logical values written in the formula but pass over
those read from cells. A formula that wants one value and is given a range takes the cell in its own row or column.

Functions: ABS, AND, AVERAGE, CONCAT, COUNT, COUNTA, IF, IFERROR, INDEX, INT, ISBLANK, ISNUMBER, ISTEXT, LEFT, LEN,
LOWER, MATCH (exact), MAX, MID, MIN, NOT, NOW, OR, RIGHT, ROUND, SORT, SUM, TODAY, TRIM, UNIQUE, UPPER, VLOOKUP
(exact), XLOOKUP (exact). Arithmetic, comparison and `&` operators, references to other sheets, defined names, whole
rows and columns, and dynamic arrays that spill (`UNIQUE`, `SORT`) are all calculated.

A formula using any other function (`SUMIF`, `COUNTIF`, `SUMPRODUCT`, `INDIRECT`, ...), an array constant
(`{1,2,3}`), arithmetic on a whole range (`=A1:A3*2`), or text that Excel would read by regional settings (a date, a
currency sign) does not stop the run. It is saved without a value, for Excel to calculate, and the report names it
(`FORMULA-NOT-EVALUATED`). A macro that then reads that formula's value does stop, since its value is unknown.

## Regional settings

Excel reads and writes numbers and dates the way the machine's regional settings say. MacroPress follows US
settings: `.` for decimals, `,` between thousands, and dates month first. Where the result would turn on the
settings, MacroPress either stops (a date typed as text into a cell, a currency sign in text read as a number) or
goes by US settings and says so in the report: a date turned into text is `1/31/2024`, text such as `"1/5/2024"`
given to `CDate` is the 5th of January, and Excel's own short date format is named `m/d/yyyy`.

Where Excel for Windows and Excel for Mac differ, MacroPress follows Windows: `#If Mac` is false, and `DateDiff`
counts the hour and day boundaries crossed.

## Event procedures

Event procedures (`Worksheet_Change`, `Worksheet_SelectionChange`, `Workbook_Open`, ...) do not run. In Excel they run
by themselves, when the workbook opens or when a macro changes its cells. When a workbook has any, the report says so
(`EVENTS-NOT-RUN`), so you can check whether they would have mattered.

## Buttons

MacroPress lists form-control buttons and shapes (or pictures) with a macro assigned, with the caption they show.
ActiveX command buttons are not listed. Any public macro can be pressed by its name, whether a button runs it or not.
