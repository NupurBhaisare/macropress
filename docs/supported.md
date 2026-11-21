# What MacroPress models

MacroPress models the parts of VBA and Excel that macros behind buttons use. Anything outside this list stops the
button with `stopped: not modelled` and a message naming what was missing, rather than being skipped or guessed.
A few behaviours MacroPress relies on have not yet been checked against Excel; every report lists the ones a run
used, under `unverified_behaviours`, and [macropress/ledger.py](../macropress/ledger.py) lists them all.

## The VBA language

- `Sub` and `Function` procedures, with `ByRef` and `ByVal` arguments, `Optional` arguments and named arguments
  (`Key1:=...`), calls between modules (`Module1.DoThis`) and recursion
- `Dim`, `ReDim` (with `Preserve`), `Const`, `Set`, module-level variables, `Option Explicit`, `Option Compare Text`
- `If` / `ElseIf` / `Else` (block and single-line), `Select Case` (values, lists and `x To y` ranges),
  `For ... Next` (with `Step`), `For Each`, `Do ... Loop` (`While` and `Until`, before or after), `While ... Wend`,
  `With`, `GoTo` and labels, `Exit`, `End`, `Erase`
- `On Error GoTo label`, `On Error Resume Next`, `On Error GoTo 0`, `Resume`, `Resume Next` and the `Err` object, with
  VBA's error numbers (for example 9, subscript out of range; 11, division by zero; 13, type mismatch; 1004 from Excel)
- Variants and VBA's conversions between them, `Empty`, `Null`, `Nothing`, error values, arrays with any bounds,
  `Integer`/`Long` overflow, `Currency`, and dates as VBA stores them
- Sheet and `ThisWorkbook` modules: a sheet module's unqualified `Range` is its own sheet, `Me`, and code names
  (`Sheet1.Range("A1")`)

Not modelled: class modules (`New MyClass`), user-defined types (`Type ... End Type`), UserForms, `Declare` (Windows
API calls), `Case Is`, file statements (`Open`, `Print #`, `Close`), and `GoSub`.

## VBA functions

Abs, Array, Asc, CBool, CByte, CCur, CDate, CDbl, CDec, CInt, CLng, CSng, CStr, CVErr, CVar, Choose, Chr, Date,
DateAdd, DateDiff, DateSerial, DateValue, Day, DoEvents, Exp, Fix, Format, Hour, IIf, InStr, InStrRev, InputBox, Int,
IsArray, IsDate, IsEmpty, IsError, IsMissing, IsNull, IsNumeric, IsObject, Join, LBound, LCase, LTrim, Left, Len, Log,
Mid, Minute, Month, MsgBox, Now, RGB, RTrim, Replace, Right, Round, Second, Sgn, Space, Split, Sqr, Str, StrComp,
StrReverse, String, Time, TimeSerial, TimeValue, Timer, Trim, TypeName, UBound, UCase, Val, VarType, Weekday, Year

Some come with limits, and stop the run outside them:

- `Format` takes these format strings: `yyyy-mm-dd`, `dd/mm/yyyy`, `mm/dd/yyyy`, `yyyymmdd`, `hh:mm`, `hh:mm:ss`,
  `hh:nn:ss`, `yyyy-mm-dd hh:mm:ss`, `yyyy-mm-dd hh:nn:ss`, `0`, `0.0` to `0.0000`, `#,##0`, `#,##0.00`, `Fixed`,
  `Standard` and `@`.
- `MsgBox` with only an OK button is noted in the report and returns `vbOK`. One that asks a question (`vbYesNo`,
  `vbOKCancel`, ...) and `InputBox` take their answers from `--answer`, in order: `yes`, `no`, `ok`, `cancel`, or a
  number for `MsgBox`, and any text for `InputBox`.
- `Now`, `Date`, `Time` and `Timer` start from `--now` (or the real time) and move on by a millisecond each time they
  are read, so a loop that waits for the time to pass ends. `Application.Wait` moves the clock to the time it waits for.
- `Rnd` stops the run (its numbers cannot be repeated), and `Environ` stops it too (its values depend on the machine).

## Excel objects

| Object | Members |
| --- | --- |
| Application (and the globals) | ActiveCell, ActiveSheet, ActiveWorkbook, Calculate, Calculation, Cells, Columns, CutCopyMode, DisplayAlerts, DisplayStatusBar, EnableEvents, Intersect, Range, Rows, ScreenUpdating, Selection, Sheets, StatusBar, ThisWorkbook, Wait, Workbooks, WorksheetFunction, Worksheets |
| Workbook | Activate, ActiveSheet, Name, Names, Save (the copy is saved when the run ends), Sheets, Worksheets |
| Sheets, Worksheets | Count, Item, Select (selecting several sheets groups them, as in Excel) |
| Worksheet | Activate, AutoFilterMode, Calculate, Cells, CodeName, Columns, FilterMode, Index, Name (renaming a sheet updates the formulas and names that refer to it), Parent, Range, Rows, Select, ShowAllData, Sort, UsedRange, Visible |
| Range | Activate, Address, AutoFilter, AutoFit (does nothing: widths are not modelled), Cells, Clear, ClearContents, Column, Columns, Copy, Count, CurrentRegion, End, EntireColumn, EntireRow, Formula, Formula2, Hidden, Item, NumberFormat (read only), Offset, Parent, PasteSpecial, Resize, Row, Rows, Select, Sort, Text, Value, Value2, Worksheet |
| Sort (Worksheet.Sort) | SortFields (Add, Add2, Clear, Count), SetRange, Header, MatchCase, Orientation, SortMethod, Apply |
| Name | Name, RefersTo, RefersToRange |
| WorksheetFunction | CountA, Find, Max, Min, Search, Sum, XLookup |

Within those, a few cases stop the run: `PasteSpecial` with `Transpose`, `SkipBlanks` or an `Operation`, or into a
range of another size; `Sort` by colour, with a custom order or left to right; `AutoFilter` with criteria other than
values, comparisons and wildcards; `Address` in R1C1 style; `XLookup` and `VLOOKUP` with approximate matching; and
ranges with several areas (`Range("A1:B2,D1:D5")`, `Union`).

Not modelled: `Range.Delete` and `Range.Insert` (moving cells would have to move every reference to them),
`Range.Find`, `Range.SpecialCells`, `Sheets.Add`, `Workbooks.Add`, `Workbooks.Open`, `Workbook.SaveAs`,
`Workbook.Close`, `Application.Run`, setting any format (`NumberFormat`, `Font`, `Interior`, `Borders`), charts,
pivot tables, shapes and comments.

## Formulas

Formulas a macro writes, and formulas already in the workbook, are calculated as Excel would, and recalculated when
the cells they read change: before a macro reads them, and before the copy is saved.

Functions: ABS, AND, AVERAGE, CONCAT, COUNT, COUNTA, IF, IFERROR, INDEX, INT, ISBLANK, ISNUMBER, ISTEXT, LEFT, LEN,
LOWER, MATCH (exact), MAX, MID, MIN, NOT, NOW, OR, RIGHT, ROUND, SORT, SUM, TODAY, TRIM, UNIQUE, UPPER, VLOOKUP
(exact), XLOOKUP (exact). Arithmetic, comparison and `&` operators, references to other sheets, defined names, whole
rows and columns, and dynamic arrays that spill (`UNIQUE`, `SORT`) are all calculated.

A formula using any other function (`SUMIF`, `COUNTIF`, `SUMPRODUCT`, `INDIRECT`, ...) does not stop the run. It is
saved without a value, for Excel to calculate, and the report names it (`FORMULA-NOT-EVALUATED`). A macro that then
reads that formula's value does stop, since its value is unknown.

## Event procedures

Event procedures (`Worksheet_Change`, `Worksheet_SelectionChange`, `Workbook_Open`, ...) do not run. In Excel they run
by themselves, when the workbook opens or when a macro changes its cells. When a workbook has any, the report says so
(`EVENTS-NOT-RUN`), so you can check whether they would have mattered.

## Buttons

MacroPress lists form-control buttons and shapes (or pictures) with a macro assigned, with the caption they show.
ActiveX command buttons are not listed. Any public macro can be pressed by its name, whether a button runs it or not.
