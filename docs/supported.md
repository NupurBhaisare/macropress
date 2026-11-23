# What MacroPress models

MacroPress models the parts of VBA and Excel that macros behind buttons use. Anything outside this list stops the
button with `stopped: not modelled` and a message naming what was missing, rather than being skipped or guessed.

What is modelled is checked against Excel itself: [tests/parity](../tests/parity) holds over a thousand small macros
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
| Workbook | Activate, ActiveSheet, Name, Names, Protect, ProtectStructure, ProtectWindows, Save (the copy is saved when the run ends), Sheets, Unprotect, Worksheets |
| Sheets, Worksheets | Count, Item, Select (selecting several sheets groups them, as in Excel) |
| Worksheet | Activate, AutoFilterMode, Calculate, Cells, CodeName, Columns, FilterMode, Index, Name (renaming a sheet updates the formulas and names that refer to it), Parent, Paste, Protect, ProtectContents, ProtectDrawingObjects, Protection (its Allow... options), ProtectionMode, ProtectScenarios, Range, Rows, Select, ShowAllData, Sort, Unprotect, UsedRange, Visible |
| Range | Activate, Address, AutoFill, AutoFilter, AutoFit (does nothing: widths are not modelled), Calculate, Cells, Clear, ClearContents, Column, Columns, Copy, Count, CurrentRegion, Cut, End, EntireColumn, EntireRow, FillDown, FillLeft, FillRight, FillUp, Formula, Formula2, Formula2R1C1, FormulaR1C1, Hidden, Item, Locked, NumberFormat, Offset, Parent, PasteSpecial, Resize, Row, Rows, Select, Sort, Text, Value, Value2, Worksheet |
| Sort (Worksheet.Sort) | SortFields (Add, Add2, Clear, Count), SetRange, Header, MatchCase, Orientation, SortMethod, Apply |
| Name | Name, RefersTo, RefersToRange |
| WorksheetFunction | AverageIf, AverageIfs, Choose, CountA, CountIf, CountIfs, Find, HLookup, Index, Lookup, Match, Max, MaxIfs, Min, MinIfs, Search, Sum, SumIf, SumIfs, VLookup, XLookup, XMatch |

Within those, a few cases stop the run: `PasteSpecial` with `Transpose`, `SkipBlanks` or an `Operation`; `Sort` by
colour, with a custom order or left to right; `AutoFilter` with criteria other than values, comparisons and wildcards;
`Address` in R1C1 style; `XLookup` and `VLOOKUP` with approximate matching; ranges with several areas
(`Range("A1:B2,D1:D5")`, `Union`); and the cases of number formats, protection, fills and cuts listed below.

## Cells and their number formats

- **Text written to a cell** is read as Excel reads typed text. `"$5"`, `"5%"`, `"1,234"`, `"(5)"`, `"1e3"`,
  `"1 1/2"`, `"TRUE"` and `"#N/A"` become values, with the number format Excel gives the cell; `"1,23"`, `"5-"` and
  `" TRUE"` stay text; a leading apostrophe keeps anything as text. Text that Excel would turn into a date or a time
  (`"1/2"`, `"Jan 5"`, `"3:30"`) stops the run, because Excel reads it by the machine's regional settings. Write a
  `Date` value instead (`DateSerial(2024, 1, 5)`).
- **Number formats** decide what VBA reads and what a write leaves, whether the workbook brought them or a macro set
  them. `Value` gives a `Date` for a format that shows a day, a month or a year, a `Double` for a
  time alone, and a `Currency` (four decimals) for Excel's own currency formats; `Value2` gives the bare number. A
  cell formatted as Text keeps what is written to it as text. A write replaces one of Excel's own formats of another
  kind (a date written to a cell of percentages) and leaves a plain number format, or one made by hand, in place.
  `Clear` resets the format, `ClearContents` keeps it, and a copy brings it along.
- **Setting `NumberFormat`** gives the code as Excel reads it: `general` is General, the letters of a date are
  lower-cased (`MM/DD/YYYY` reads back `mm/dd/yyyy`), colours and `AM/PM` are spelled Excel's way, `"$"` becomes a bare
  `$`, and a code Excel cannot read (`[Red`, five sections, `0.00e+00`) is error 1004. A code with letters Excel reads
  its own way (`abc` becomes `\abbc`, `0.00 kg` is refused) stops the run. A whole row, column or sheet takes the
  format itself, so that cells written there later have it; a range that runs to the sheet's last row formats its
  columns, and its cells above keep theirs. Several cells with different formats read as `Null`, and rows and columns
  formatted whole count in `UsedRange`. `Locked` is set and read the same way. `NumberFormatLocal`, `Style` and an
  array of formats stop the run.
- **`Text`** gives what the cell shows for text, for the General format and for plain number formats (`0.00`,
  `#,##0.00`). It stops the run for any other format, and wherever the column's width would decide what shows:
  Excel rounds a number that does not fit, or shows `####`.
- **A filtered range**: a value, an array, a clear, a number format or a fill reaches only the rows the filter shows,
  as in Excel, and so does one cell copied over it; a block of cells pasted there lands whole. A filter matches the
  text each cell shows (`1,250.00`, not `1250`), without regard to case or to the spaces around it.

## Protection and hidden sheets

- **A protected sheet**, protected in the file or by the macro, refuses what Excel refuses, with error 1004: a write
  to a locked cell (a value written over locked and unlocked cells reaches the unlocked ones first), a clear, paste or
  fill touching one, a number format (unless `AllowFormattingCells`), `Locked`, hiding rows or columns (unless
  `AllowFormattingRows` or `AllowFormattingColumns`), sorting (unless `AllowSorting`, and never locked cells) and the
  `AutoFilter` method, whatever `AllowFiltering` says. A paste, fill or `Clear` into unlocked cells changes their
  contents and leaves their formats. `UserInterfaceOnly:=True` lets macros change anything, for the rest of the run,
  as it holds in Excel until the workbook closes.
- **Passwords** are checked against the hash the file keeps, in either of the two forms Excel writes. `Unprotect`
  without the password of a sheet that has one asks for it, as Excel does: pass the answer with `--answer`. A
  password the macro sets is saved as Excel saves one (salted SHA-512), and a protection given back just as it was
  keeps the file's own. `Protect` again on a protected sheet takes the options it is given and keeps the password,
  except with other options and another password, which stops the run.
- **`Workbook.Protect`** locks the structure: renaming, hiding and showing sheets are error 1004.
- **`Visible`**: `False` or `0` hides a sheet, `xlSheetVeryHidden` hides it from the user's own Unhide, `True`, `-1` or
  `1` shows it; anything else is error 1004, as is hiding the last sheet shown. Hiding the active sheet makes the next
  sheet shown to its right active, else the one to its left. A hidden sheet is read and written as usual; activating
  or selecting it is error 1004, as in Excel for Windows.

## Filling, copying and moving cells

- **`FillDown`, `FillUp`, `FillRight`, `FillLeft`** copy the range's first row (last row, first column, last column)
  into the rest of it: values, formulas with their references moved, and formats. A range one row deep fills from the
  row above it, and on the sheet's edge is error 1004.
- **A copy onto a larger range** repeats the copied block where the range holds it a whole number of times down and
  across (a range one cell high or wide counts as holding it); any other range takes the block once, from its first
  cell. `Worksheet.Paste` pastes what was copied into its destination or the selection.
- **`AutoFill`**: a line of formulas and blank cells repeats, as does any line with `xlFillCopy`; a cell filled alone
  is copied, a date counts on by days, a number alone in its line of a larger range counts on by 1, and two numbers
  or more go on along their least-squares trend. `xlFillFormats`, `xlFillValues`, `xlFillSeries` and `xlFillDays`
  behave as in Excel. Text that counts on (`Item 1`, `Q1`, month and weekday names), a destination that runs on both
  down and across, and rows a filter hides stop the run.
- **`Cut`** moves the cells, contents and formats, and leaves them empty and General; every formula and defined name
  that reads them reads them where they are now, one that read the cells they landed on reads `#REF!`, and one that
  read part of a range they belong to is left as it was. A cut on a protected sheet, of part of a filter's range, a
  table or an array formula, of whole rows or columns that are hidden or formatted whole, or that moves the end of a
  range along it (Excel stretches the range) or part of it to another sheet, stops the run. So does saving a copy
  whose conditional formats, validation, hyperlinks or charts cover the cells moved: those parts are not rewritten.

Not modelled: `Range.Delete` and `Range.Insert` (moving cells would have to move every reference to them),
`Range.Find`, `Range.SpecialCells`, `Sheets.Add`, `Workbooks.Add`, `Workbooks.Open`, `Workbook.SaveAs`,
`Workbook.Close`, `Application.Run`, formats other than number formats and `Locked` (`Font`, `Interior`, `Borders`,
`ColumnWidth`), charts, pivot tables, shapes and comments.

## Formulas

Formulas a macro writes, and formulas already in the workbook, are calculated as Excel would, and recalculated when
the cells they read change: before a macro reads them, and before the copy is saved. Under manual calculation
(`Application.Calculation = xlCalculationManual`, or a workbook saved that way) they keep their values until
`Calculate`, the return to automatic calculation, or the save.

Excel's own habits with numbers are followed: they are compared to 15 significant digits (`=0.1+0.2=0.3` is TRUE), a
formula or a `SUM` that ends on a sum which all but cancels gives 0, text in arithmetic is read as typed text
(`="5%"+1`), and `SUM`, `COUNT` and the like count text and logical values written in the formula but pass over
those read from cells.

Functions: ABS, AGGREGATE, AND, AVERAGE, AVERAGEIF, AVERAGEIFS, CEILING, CEILING.MATH, CEILING.PRECISE, CHAR, CHOOSE,
CLEAN, CODE, COLUMN, COLUMNS, CONCAT, CONCATENATE, COUNT, COUNTA, COUNTBLANK, COUNTIF, COUNTIFS, DATE, DATEDIF, DAY,
DAYS, DAYS360, DB, DDB, EDATE, EOMONTH, ERROR.TYPE, EVEN, EXACT, EXP, FIND, FLOOR, FLOOR.MATH, FLOOR.PRECISE, FV,
HLOOKUP, HOUR, IF, IFERROR, IFNA, IFS, INDEX, INDIRECT, INT, IPMT, ISBLANK, ISERR, ISERROR, ISEVEN, ISLOGICAL, ISNA,
ISNONTEXT, ISNUMBER, ISO.CEILING, ISODD, ISOWEEKNUM, ISREF, ISTEXT, LARGE, LEFT, LEN, LN, LOG, LOG10, LOOKUP, LOWER,
MATCH, MAX, MAXIFS, MEDIAN, MID, MIN, MINIFS, MINUTE, MOD, MODE, MODE.SNGL, MONTH, MROUND, N, NA, NETWORKDAYS,
NETWORKDAYS.INTL, NOT, NOW, NPER, NPV, ODD, OFFSET, OR, PERCENTILE, PERCENTILE.EXC, PERCENTILE.INC, PI, PMT, POWER,
PPMT, PRODUCT, PROPER, PV, QUARTILE, QUARTILE.EXC, QUARTILE.INC, QUOTIENT, RANK, RANK.AVG, RANK.EQ, REPLACE, REPT,
RIGHT, ROUND, ROUNDDOWN, ROUNDUP, ROW, ROWS, SEARCH, SECOND, SIGN, SLN, SMALL, SORT, SQRT, STDEV, STDEV.P, STDEV.S,
STDEVP, SUBSTITUTE, SUBTOTAL, SUM, SUMIF, SUMIFS, SUMPRODUCT, SWITCH, SYD, T, TEXTJOIN, TIME, TODAY, TRIM, TRUNC,
TYPE, UNICHAR, UNICODE, UNIQUE, UPPER, VALUE, VAR, VAR.P, VAR.S, VARP, VLOOKUP, WEEKDAY, WEEKNUM, WORKDAY,
WORKDAY.INTL, XLOOKUP, XMATCH, XNPV, YEAR, YEARFRAC. Arithmetic, comparison and `&` operators, references to other
sheets, defined names, whole rows and columns, ranges that end on a function's cells (`A1:INDEX(A:A,5)`), array
constants (`{1,2;3,4}`), and dynamic arrays that spill (`=A1:A5*2`, `UNIQUE`, `SORT`) are all calculated.

**Ranges and arrays** are worked out as Excel works them out:

- In a formula written with `Formula` (or `Value`), a range where one value is wanted (in arithmetic, or given to a
  function such as `ROUND` or `LEN`) gives the cell in the formula's own row, or column, and `#VALUE!` where the range
  passes the formula by: `=SUM(A1:A5*2)` in row 2 is `A2*2`. A range given to a function that takes a reference
  (`SUM`, `MAX`, `COUNT`, `AND`, `CONCAT`) is taken whole; `SUMPRODUCT` works on arrays whatever the formula. An array
  constant stays an array (`=SUM(A1:A5*{1,2})` in row 3 is 9), and a formula that ends on an array shows its first
  value (`=ROW(A2:A4)` is 2).
- In a formula written with `Formula2` every range is taken whole, operators and functions work value by value, and
  the result spills; `@` takes one value, as in Excel. A single row or column goes with every row or column of the
  other side, and past the end of a shorter one the value is `#N/A`.
- `IF`, `INDEX`, `CHOOSE`, `XLOOKUP`, `OFFSET` and `INDIRECT` hand back the cells of a range they pick, not their
  values: `=SUM(IF(c, A1:A5))` adds them all up, and `SUM` passes over text it reads that way as it does text in
  cells.
- `SUMPRODUCT` counts what is not a number as 0, takes arrays of one size only (`#VALUE!` otherwise), and does not
  round to 0 a sum that all but cancels, as `SUM` does.
- **Criteria** (`SUMIF`, `COUNTIF`, `AVERAGEIF`, their `S` forms, `MAXIFS`, `MINIFS`) find what Excel finds. A
  number finds the numbers equal to it to 15 digits and the text that reads as one: `5` and `"5"` find `"05"`, `" 5"`
  and `"5.0"`. Text may start with `=`, `<>`, `<`, `>`, `<=` or `>=`; what follows is a number, `TRUE`, `FALSE` or an
  error (`"#N/A"` finds `#N/A` cells) if it reads as one, else text, matched ignoring case with the wildcards `* ? ~`.
  `<>` finds every cell `=` would not, but text that reads as the number (`"<>5"` finds `"5"`); `<` and `>` find only
  cells of their own kind (`">4"` passes over `"5"`). `""` finds blank cells and empty text, `"="` blank cells only,
  `"<>"` every cell that is not blank. A blank criterion is the number 0; `TRUE` never finds 1. A criterion over 255
  characters is `#VALUE!`. Text that Excel reads by regional settings, a date or an amount of money, in a criterion or
  in a cell a number is looked for in, leaves the formula for Excel to calculate, as below.
- `SUMIF` and `AVERAGEIF` read their sum range from its first cell over the criteria range's size, whatever its own
  (`=SUMIF(A1:A5,"x",B1)` adds up B1:B5), and are recalculated when any of those cells change; the `S` forms want
  ranges of one size (`#VALUE!` otherwise). Their sums are the plain ones, not rounded to 0 as `SUM`'s, and an error
  in a cell found is the result. `AVERAGEIF` that finds nothing is `#DIV/0!`, `MAXIFS` and `MINIFS` 0. A range where
  a criterion is wanted gives the formula's own row of it, as any value does; a constant, an operator or a function
  such as `SUM` where a range is wanted is refused as Excel refuses it (error 1004 as the formula is written), as is
  the wrong number of arguments.
- **Lookups** find what Excel finds. An exact match (`VLOOKUP` and `HLOOKUP` with `FALSE`, `MATCH` with 0) ignores
  case and takes the wildcards `* ? ~` for text; a number never finds text that reads as one, numbers are compared in
  full (not to 15 digits), and a value looked up over 255 characters is `#VALUE!` (but in `XLOOKUP`). An approximate
  match finds the last value at or below the one looked up, the last of equal ones, passing over values of other
  kinds; `MATCH` with -1 the first equal value, else the last above. `XLOOKUP` and `XMATCH` take their match and search
  modes; a binary search finds the first equal value (the last, descending). Over values not in order Excel's answer
  is wherever its halving search lands: MacroPress answers only where it cannot land elsewhere, the value looked up
  above or below them all (`LOOKUP(9.99E+307, A:A)`, `LOOKUP(2, 1/(A:A<>""), B:B)`), and otherwise leaves the formula
  for Excel.
- `INDEX`, `LOOKUP`, `XLOOKUP` and `XMATCH` work an expression they are given out whole in any formula
  (`MATCH(1, INDEX((A1:A9="x")*(B1:B9="y"), 0), 0)`); `MATCH` and `VLOOKUP` take one cell of a range there, as Excel
  does without Ctrl+Shift+Enter, and `CHOOSE` with an array (`CHOOSE({1,2}, B:B, A:A)`) puts its ranges side by side.
- `OFFSET` and `INDIRECT` are worked out again on every change, as in Excel; `INDIRECT` reads A1 or R1C1 text, a
  sheet's name and defined names, and its text naming nothing is `#REF!`. `INDIRECT` to another workbook stops the
  run: whether Excel has it open is not known. A wrong number of arguments, or a constant where `OFFSET` wants a
  reference, is refused as Excel refuses it (error 1004 as the formula is written).
- **The text of a formula** is kept as Excel keeps what is typed: its functions' names and its references in
  capitals, a sheet's name as the sheet has it (quoted where it could be read as a cell, `'AR2'!A1`), numbers as Excel
  writes them, and no space before a comma. `Formula2` shows a formula written with `Formula` with `@` where it takes
  one value (`=SUM(@A1:A5*2)`), and `Formula` shows one written with `Formula2` without them. Where Excel's choice of
  `@` is not known (a function MacroPress does not model, an array where one value is wanted), reading it stops.
- **Saved**, a dynamic-array formula stays one: one that spills, and one of a single value that takes ranges whole
  (`=SUM(A1:A5*2)` from `Formula2`), which read the other way would take one cell of them. A workbook that has no
  dynamic arrays yet is given the part of the file Excel keeps them in. `@` is saved as Excel saves it.

**Rounding, dates, text and the rest** answer as Excel does (group 2.4 of the [roadmap](roadmap.md)):

- **Rounding**: `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `CEILING`, `FLOOR` and their `.MATH` and `.PRECISE` forms, and
  `MROUND` work on the number to 15 digits (`=ROUNDDOWN(4.35*100,0)` is 435, `=FLOOR(0.3,0.1)` 0.3); `QUOTIENT` and
  `MOD` on the number as it is (`=QUOTIENT(4.35*100,1)` is 434). `MOD` is the remainder itself, and `#NUM!` once the
  quotient passes about 1.1E+12. A number and a step of other signs are `#NUM!` in `MROUND`; in `CEILING` a negative
  number with a positive step goes up, toward 0.
- **Dates** are serial numbers in Excel's 1900 system, whose 29th of February 1900 (serial 60) `DATE`, `YEAR`,
  `MONTH`, `DAY` and `WEEKDAY` keep; `EOMONTH` and `EDATE` count months by the calendar. `DATE` rolls months and days
  over and reads a year below 1900 as that many after it; a date's time of day is rounded to the second
  (`=DAY(59.99999999)` is 29). `DATEDIF`'s `"md"` goes negative as Excel's does (31 January to 1 March 2015 is -2), and
  `"yd"` counts as Excel's counts; `YEARFRAC` takes its five bases, `DAYS360` the US and European rules, `NETWORKDAYS`
  and `WORKDAY` holidays and, in their `.INTL` forms, other weekends. `HOUR`, `MINUTE`, `SECOND`, `TIME`, `WEEKNUM`,
  `ISOWEEKNUM` and `DAYS` too. Text that Excel reads as a date by regional settings (`=YEAR("2024-01-05")`) leaves the
  formula for Excel.
- The functions that came from the Analysis ToolPak (`EOMONTH`, `EDATE`, `WORKDAY`, `NETWORKDAYS` and their `.INTL`
  forms, `YEARFRAC`, `WEEKNUM`, `MROUND`, `QUOTIENT`, `ISEVEN`, `ISODD`) take text that reads as a number but not
  `TRUE`, and are `#VALUE!` given a range of several cells in a formula that is not a dynamic array
  (`=EOMONTH(A1:A3,0)`), where other functions take the formula's own row of it.
- **Text**: `FIND` (case counts, no wildcards), `SEARCH`, `SUBSTITUTE` (the nth searched for from the character after
  each found: `=SUBSTITUTE("aaa","aa","b",2)` is "ab"), `REPLACE`, `REPT`, `EXACT`, `PROPER` (a letter after anything
  but a letter in capitals: "It'S"), `TEXTJOIN`, `CLEAN`, `VALUE`, `N`, `T`, `UNICHAR` and `UNICODE`. `CHAR` and `CODE`
  read codes 128 to 255 as Windows does (Windows-1252), where Excel for Mac reads them as Mac Roman. A result of more
  than 32,767 characters is `#VALUE!`.
- **Logic and information**: `IFS` (no condition that holds is `#N/A`), `SWITCH` (compared as `=` compares; then the
  default, or `#N/A`), `IFNA`, the `IS` functions, `ERROR.TYPE`, `TYPE` and `NA`.
- **Statistics**: `MEDIAN`, `MODE` (the first value to reach the count), `LARGE` and `SMALL` (a k with a fraction is
  rounded up), `RANK` in its three forms, `PERCENTILE` and `QUARTILE` in theirs, `STDEV` and `VAR` in theirs, and
  `COUNTBLANK` (which counts `""`). A variance has Excel's last digits: from the sum and the sum of squares, unless
  those cancel to nothing; where the two ways part by more than a last digit, the formula is left for Excel.
- **Finance**: `PMT`, `IPMT`, `PPMT`, `PV`, `FV`, `NPER`, `NPV`, `XNPV`, `SLN`, `SYD`, `DB` and `DDB`, worked out with
  Excel's own arithmetic (a whole power by repeated squaring) to the last digit, all but `PPMT`'s 15th. `IRR`, `XIRR`
  and `RATE` are found by trial, and Excel stops within its tolerance of the rate, where its search took it (two
  guesses give two answers that part in the 10th digit): MacroPress leaves them for Excel, but for cash flows that
  never change sign, which are `#NUM!`.
- **`SUBTOTAL` and `AGGREGATE`** pass over the cells whose formulas call either of them. `SUBTOTAL`'s 1 to 11 pass over
  every hidden row while the sheet is filtered, and none while it is not; 101 to 111 pass over hidden rows always;
  hidden columns count. `AGGREGATE` takes its 19 functions and 8 options, and an array it works out
  (`=AGGREGATE(15,6,A1:A9/(B1:B9="x"),1)`). Hiding or showing a row, or filtering, makes them, and what reads them,
  stale, as a change to a cell they read does.
- **A formula's number format.** Entered in a General cell, a formula takes the format of the function its value comes
  from, as in Excel: `DATE` and `TODAY` a date, `NOW` a date and time, `TIME` a time, `PMT`, `PV`, `FV`, `NPV` and the
  like a currency, `IRR` and `RATE` a percentage, so that `Value` reads a `Date` or a `Currency`. `ROUND`, `MAX`, `SUM`
  and a few more keep the format of what they are given, as do `+` and `-` (a date and a number make a date, two
  dates the days between, General); `IF`, `ABS`, `YEAR`, `*` and `/` give none. A cell with a format of its own keeps
  it. Where Excel's choice is not known (`=DATE(2024,1,1)+PMT(...)`) writing the formula stops the run. A formula that
  calls none of these functions stays General, where Excel may give it the format of a cell it reads (`=A1+1`): not
  modelled yet.

Formulas can be written and read R1C1 style (`FormulaR1C1`, `Formula2R1C1`): one text written to many cells gives
each its own formula, and a relative reference that runs off the sheet comes round from its other edge, as in Excel.
`Formula` reads R1C1 text too where A1 style cannot (`=R1C1`). An A1 reference given to `FormulaR1C1`, which Excel
turns into a name that cannot exist (`='A1'+1`), stops the run.

A formula using any other function (`TEXT`, `IRR`, `FILTER`, ...), or text that Excel would read by regional
settings (a date, a currency sign) does not stop the run. It is saved without a value, for Excel to calculate, and the
report names it (`FORMULA-NOT-EVALUATED`). A macro that then reads that formula's value does stop, since its value is
unknown.

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
