# Roadmap

What MacroPress should model next, in what order, and what to watch for in each piece.

The list comes from what stops buttons in finance workbooks today: formatting, lookup and conditional-sum functions,
rows and sheets that are inserted or deleted, other workbooks and files, and finding or collecting cells. It is
split into eight groups. Each can be built, checked against Excel and released by itself.

## How the order was chosen

1. **How often it stops a macro.** What recorded macros do in their first lines comes first.
2. **What later groups need.** Showing a number as text (group 4) is needed by `Find`, filters on dates, `TEXT` and
   CSV files. Moving references (group 6) is needed by anything that deletes rows.
3. **How surely it can be checked.** Excel's answers are recorded on a Mac ([tests/parity](../tests/parity)). What
   exists only in Excel for Windows (`Scripting.Dictionary`, `FileSystemObject`, Outlook) cannot be recorded there,
   and waits for a Windows recording.

| # | Group | Stops macros today at | Needs | Size |
| --- | --- | --- | --- | --- |
| 1 | [First lines of a recorded macro](#1-the-first-lines-of-a-recorded-macro) | `.Font` (1.1, 1.3, 1.4 and 1.5 are done) | nothing | small, five pieces |
| 2 | [Worksheet functions](#2-worksheet-functions) | `VLOOKUP(…, TRUE)`, `Application.Match`, `INDIRECT`, `TEXT` (2.1 and 2.2 are done) | nothing | large, ships in slices |
| 3 | [Finding and collecting](#3-finding-and-collecting) | `Find`, `SpecialCells`, `Union`, `New Collection`, `RemoveDuplicates` | nothing (a Windows recording for `Dictionary`) | medium |
| 4 | [Numbers as text](#4-numbers-as-text) | `.Text` of a date, `Format(x, "mmm-yy")`, `TEXT()` | 1 | medium |
| 5 | [Appearance](#5-appearance) | `Font`, `Interior`, `Borders`, `ColumnWidth`, `Merge` | 1, 4 | large |
| 6 | [Rows, columns and sheets that come and go](#6-rows-columns-and-sheets-that-come-and-go) | `Rows.Delete`, `Insert`, `Sheets.Add`, `Sheets.Copy` | 3 for filtered deletes, 5 for copies | very large |
| 7 | [Other workbooks and files](#7-other-workbooks-and-files) | `Workbooks.Open`, `SaveAs`, `Open … For Output`, Outlook | 4, 5, 6 | very large |
| 8 | [Left for Excel, or refused on purpose](#8-left-for-excel-or-refused-on-purpose) | `RefreshAll`, `GoalSeek`, PDF, charts | nothing | small |

Formatting is the most common stopper of all, which is why its two cheap halves are in group 1 (number formats, and a
way to carry on past fonts and colours), while writing fonts, fills and borders into the file waits for group 5.

## Rules for every piece

- **Probes first.** Write the probes, record Excel's answers, then model. A case below marked **probe** is one whose
  answer nobody should assume.
- **Refuse what is not settled**, argument by argument if need be: a method can ship with some of its arguments
  refused.
- **Regional settings**: go by US settings and say so in the report, or stop.
- **A dialog is a question.** Where Excel would ask the user, the run takes an answer (`--answer`) or stops. It is
  never assumed away.
- **The saved copy must carry what the model changed, or the save must stop.** Parts of the file MacroPress does not
  read are copied byte for byte, and must not be left pointing at the wrong cells.
- **One push per piece**: probes, model, [supported.md](supported.md), changelog.

## 1. The first lines of a recorded macro

Five small pieces, in this order. None needs new machinery.

**Status**: 1.1, 1.3, 1.4 and 1.5 are done, each recorded from Excel first; what they still refuse is in
[supported.md](supported.md). 1.2 waits for the maintainers' decision. Left open, each with Excel's answer recorded in
the parity suite: format codes with letters Excel reads its own way (`abc` becomes `\abbc`); `AutoFill` of text that
counts on (`Item 1`, `Q1`, months, weekdays) and into two directions at once; a cut that moves the end of a range
along it (Excel stretches the range) or part of a range to another sheet (Excel shrinks it); a second `Protect` with
other options and another password, whose recorded answers disagree.

### 1.1 Setting `NumberFormat`

The model already keeps a number format for every cell, row and column, and the writer already saves a cell's.

- A format does not change what a cell holds. `"@"` on a number leaves a number until it is written again (**probe**).
  A date format makes `Value` a `Date` from then on.
- Excel rewrites what it is given (`"0.00 kg"` may come back as `0.00 "kg"`), and macros compare what they read
  back. **Probe** what `NumberFormat` returns after each kind of code, and which codes are error 1004.
- Whole rows, columns and `Cells`: the format goes to the row or column, not to a million cells. The writer must
  learn to save `<col style>` and `<row s>`, which it does not today.
- A range with several formats reads as `Null`, not as its first cell's.
- A code that equals one of Excel's own formats is Excel's own: [`numfmt.after_entry`](../macropress/numfmt.py)
  decides by it what a later write replaces.
- `NumberFormatLocal`: refuse. `Range.Style = "Comma"`, `"Currency"`, `"Percent"` need named cell styles: **probe**
  what `NumberFormat` reads afterwards.

### 1.2 Carrying on past appearance, when asked to

Until group 5, a switch (`--skip-appearance`) that lets a run pass statements which only change how a sheet looks.
Off by default. *This one is a decision for the maintainers: it trades a faithful-looking copy for a finished run.*

- Appearance only: `Font`, `Interior`, `Borders`, alignment, `WrapText`, `ColumnWidth`, `RowHeight`, `AutoFit`,
  tab colour, `ActiveWindow` zoom, gridlines and frozen panes, `PageSetup`.
- Not appearance: `NumberFormat` (decides what `Value` and `Text` give), `MergeCells` (decides which cells hold
  values), `Hidden` (filters, copies, `End`), `Locked` (which writes fail on a protected sheet).
- A *read* of anything skipped still stops (`If c.Interior.Color = vbYellow`). Colour-coded inputs are common in
  models, and reading them needs group 5.
- `With Selection.Font … End With` and `Set f = c.Font`: the object must take sets and refuse gets.
- The report says what was dropped, by kind and count. The copy must not pass for a formatted one.

### 1.3 `Protect`, `Unprotect` and `Visible`

- The file holds a hash of the password, in one of two forms. `Unprotect` with a wrong password is error 1004; with
  none, on a sheet that has one, Excel asks for it: a question.
- While a sheet is protected, VBA's writes to locked cells fail too, unless it was protected with
  `UserInterfaceOnly:=True`, and that flag is not saved in the file: a workbook just opened never has it.
- `Locked` is part of a cell's style, inherited from rows and columns like a number format. The reader must take it.
- What protection blocks depends on its options (`AllowSorting`, `AllowFiltering`, `AllowInsertingRows`, …): every
  operation in the model has to ask. `Workbook.Protect` blocks adding, deleting, renaming and hiding sheets.
- Saving a new password needs a salt. Excel's is random; derive ours, so that runs stay repeatable.
- `Protect` on a sheet that is already protected: **probe**.
- `Visible` has three states. Hiding the last visible sheet is error 1004. Hiding the active sheet makes another one
  active: **probe** which. A hidden sheet is read and written as usual.

### 1.4 `FormulaR1C1`

- Forms: `RC`, `R[-1]C`, `RC[2]`, `R5C3`, mixed `R5C[1]`, whole rows and columns (`R[-1]`, `C`), ranges, sheet
  names in front, inside functions. Text in quotes is left alone.
- Functions that begin with R or C (`ROUND`, `COUNT`) and defined names such as `RC_Total` are not references.
- Reading gives the A1 formula relative to the cell, with `R[0]C[0]` written `RC`.
- One text written to many cells is the whole point: each cell gets its own A1 formula.
- A reference that leaves the sheet (`R[-1]C` in row 1): **probe** (error 1004, or wrapping round).
- Excel takes some R1C1 text through `Formula` and some A1 text through `FormulaR1C1`: **probe** both.
- `FormulaR1C1Local`, `FormulaArray`: refuse.

### 1.5 Filling

`FillDown`, `FillRight`, a copy onto a range larger than what was copied, `AutoFill`, `Cut`.

- `FillDown` copies the first row into the rest: formulas moved, values copied, number formats too. On one row it
  fills from the row above (**probe**). On a filtered range: **probe** whether hidden rows are filled.
- A copy onto a larger range: one cell fills all of it; a block tiles a range that is a multiple of its size;
  otherwise **probe**. This removes today's refusal ("pasting 1x1 copied cells into …, another size").
- `AutoFill`: the destination must include the source. Formulas always fill as formulas. For values the default
  depends on what is there: one number is copied, two make a series, dates count on, `"Q1"` and `"Item 1"`
  count on, and month and weekday names follow lists kept on the machine. Model copies and formulas; refuse series
  until each kind is probed, and lists for good.
- `Cut` is not a copy and a clear: formulas elsewhere that read the moved cells follow them, and formulas that read
  cells the move overwrote become `#REF!`. Names follow too.
- What `PasteSpecial` still refuses: `Transpose`, `SkipBlanks`, and `Operation` (**probe** what multiplying onto a
  formula leaves).

## 2. Worksheet functions

In formulas, and from VBA. Large, but it ships in slices; each slice is a family of functions and its probes.

### 2.1 Groundwork

**Status**: done, recorded from Excel first. What it settled: a range where one value is wanted gives the formula's
own row or column (in `Formula`), every range is an array (in `Formula2`); array constants stay arrays; `SUMPRODUCT`
works on arrays and does not round a sum that all but cancels; `IF`, `INDEX` and `XLOOKUP` hand back cells, `IFERROR`
values; `Formula2` shows `@` where a `Formula` formula takes one value. Left open: where Excel puts `@` around a
function MacroPress does not model, and arrays of more than two whole columns.

- **Arrays in arithmetic.** `SUMPRODUCT((A1:A9="x")*B1:B9)` needs `=`, `*` and the rest to work cell by cell. Today
  the evaluator refuses a range where one value is wanted.
- **Which functions make their arguments arrays.** Checked already: in a formula that is not a dynamic array,
  `=SUM(F1:F3*2)` and `=MAX(F1:F3*2)` take only the formula's own row of `F1:F3`, while `SUMPRODUCT` takes it all.
  That needs a table, per function and argument, and probes to fill it. `Formula2` evaluates everything as arrays.
- **Array constants** (`{1,2;3,4}`), with text and logical values passed over as they are in a range.
- **The formula's own cell.** `ROW()`, `COLUMN()` and taking one cell of a range need the evaluator to know where
  the formula is. It is told nothing today.
- What is already settled applies everywhere: comparison to 15 digits, the sum that all but cancels, text and
  logical values counted where written and passed over where read, an empty argument as 0.

### 2.2 Conditional sums and counts

`SUMIF(S)`, `COUNTIF(S)`, `AVERAGEIF(S)`, `MAXIFS`, `MINIFS`.

**Status**: done, recorded from Excel first, in formulas and from `WorksheetFunction`. What it settled: `5`, `"5"`
and `"05"` all find the number 5 and the text `"05"`, `" 5"`, `"5.0"`; `"<>5"` finds that text all the same, and
`">4"` passes it over. Text is compared to 15 digits as the number it reads as, so the account numbers
`"1234567890123456"` and `"1234567890123457"` are found by either, as Excel finds them. `"TRUE"` finds `TRUE`, never
the text or 1; `"#N/A"` finds the error; a blank criterion is 0; `"="` finds blank cells, `""` empty text as well.
Sums are not rounded to 0 as `SUM`'s are. A range where a criterion is wanted gives the formula's own row (and
`#VALUE!` outside it, which finds only `#VALUE!` cells). Left open: text that reads as a date or money, in a criterion or in a
cell a number is looked for in, which is read by regional settings.

- A criterion is a value, or an operator and a value: `">5"`, `"<>"` (not blank), `"="` (blank), wildcards
  `* ? ~` for text, no regard to case.
- A number and text that reads as that number count as the same (`5` finds `"5"`). **Probe** the borders: `"05"`,
  `"5.0"`, `" 5"`.
- Numbers longer than 15 digits are compared by their first 15: account numbers held as text can be counted as
  equal when they are not. Reproduce it or refuse it; do not fix it.
- Dates in a criterion's text (`">1/5/2024"`) are read by regional settings: refuse.
- `SUMIF`'s sum range may differ in shape (its top-left cell and the criteria range's size are used). `SUMIFS`
  needs equal shapes, or `#VALUE!`.
- Whether `>` compares numbers with numbers only, as it does in a filter: **probe**. Blank cells, formulas that give
  `""`, errors in the range, criteria over 255 characters, no match in `AVERAGEIF` (`#DIV/0!`).
- Does the total go to 0 where a `SUM` would (the sum that all but cancels)? **Probe**.

### 2.3 Lookups

- **Exact match has wildcards** in `VLOOKUP` and `MATCH` when the value looked up is text (`XLOOKUP` only when
  asked). Today's compare plainly, so `VLOOKUP("a*", …, FALSE)` may already differ from Excel: check that first.
- **Approximate match** bisects, and on data that is not sorted the answer is wherever the bisection lands.
  Recommended: answer when the column is sorted, stop when it is not. Duplicates and mixed types: **probe**.
- A number never matches text that reads as one. A value looked up that is longer than 255 characters is `#VALUE!`.
- `HLOOKUP`, `LOOKUP` in both forms, `MATCH` with -1, the modes of `XLOOKUP` and `XMATCH`, `CHOOSE` and `INDEX`
  handing back cells rather than values (`SUM(INDEX(…):INDEX(…))`).
- `OFFSET` and `INDIRECT` read cells no one can name beforehand: they recalculate on every change, as they do in
  Excel. `INDIRECT` to a closed workbook is `#REF!`.

### 2.4 Rounding, dates, text, the rest

- **Rounding**: `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `MROUND`, `CEILING`, `FLOOR` and their newer forms work on the
  15-digit number (checked for four of them); `MOD` does not. Signs in `CEILING` and `FLOOR`: **probe**.
- **`SUBTOTAL` and `AGGREGATE`** pass over rows a filter hides, and with 101 to 111 rows hidden by hand; they pass
  over other `SUBTOTAL`s. Hidden rows become an input of a formula: hiding or showing a row must make these
  formulas stale, which the index of readers in [book.py](../macropress/excel/book.py) does not know about yet.
- **Dates**: `DATE` (years below 1900, months that overflow), `EOMONTH`, `EDATE`, `DATEDIF` (its `"md"` is known to
  be wrong: reproduce it), `WEEKDAY`, `WEEKNUM`, `NETWORKDAYS`, `WORKDAY`, `YEARFRAC` and its five bases, `DAYS360`.
  Excel's 1900 has a 29th of February. `DATEVALUE` reads text by regional settings: refuse.
- **Text**: `FIND`, `SEARCH`, `SUBSTITUTE`, `REPLACE`, `TEXTJOIN`, `PROPER`, `EXACT`, `REPT`. `CHAR` and `CODE`
  differ between Windows and Mac: follow Windows, and expect the Mac recording to disagree. `TEXT` is group 4.
- **Logic and information**: `IFS` (no match is `#N/A`), `SWITCH`, `IFNA`, `ISERROR`, `ISNA`, `COUNTBLANK` (counts
  `""`), `LARGE`, `SMALL`, `RANK`, `MEDIAN`. `CELL` and `INFO` depend on the machine: refuse. `RAND`: refuse.
- **Finance**: `PMT`, `IPMT`, `PPMT`, `PV`, `FV`, `NPER` are closed forms (a rate of 0 is its own case). `IRR`,
  `XIRR` and `RATE` are found by trial: Excel's first guess, its number of tries and its tolerance decide which
  answer comes back and when `#NUM!` does. Model what probes pin down, and stop where ours does not settle the way
  Excel's does (cash flows that change sign more than once). `NPV` discounts from the first period, `XNPV` from the
  first date. `DB` rounds its rate to three decimals.

### 2.5 From VBA

- `WorksheetFunction.X` raises error 1004 where the function gives an error. `Application.X` returns the error
  value, which is what `IsError(Application.Match(…))` relies on. Both are needed.
- Arguments follow the rules checked for `Sum`, `Max` and `Min`: a value alone counts as one written in a formula,
  a range or an array as cells. A VBA array of one dimension is a row.
- Results come back as `Double`, `String`, `Boolean`, or an array that starts at 1 (`Transpose`, `Index(a, 0, 2)`).
  `Application.Transpose` has limits on size and on the length of text on Windows: **probe** on the Mac.
- `WorksheetFunction.Round` is not VBA's `Round`.
- `Evaluate` and `[…]` beyond a plain reference: a formula in text, read against the active sheet.

## 3. Finding and collecting

### 3.1 `Find`, `FindNext`, `Replace`

- **Arguments left out are remembered** from the last `Find`, by the macro or by the person at the keyboard:
  `LookIn`, `LookAt`, `SearchOrder`. A run has no keyboard before it. Recommended: use the settings of a fresh Excel
  (**probe** them) and say so in the report whenever a macro leans on them.
- The search starts *after* the `After` cell, which is the top-left cell when left out, so that cell is found last.
  Searching backwards from it reaches the last cell first: this is the usual way to find the last used row
  (`Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)`).
- `LookIn:=xlFormulas` reads formula text and constants as typed. `xlValues` reads what cells show, which for dates
  and formatted numbers needs group 4: until then, stop at those cells.
- With `xlValues`, cells in hidden rows are reportedly not found; with `xlFormulas` they are. **Probe**: it matters
  after every filter.
- Whole cell or part of it, wildcards, case, a number or a `Date` as the thing to find (**probe**; refuse dates).
- Nothing found gives `Nothing`, and `.Row` on it is error 91, as members of `Nothing` already are.
- `FindNext` goes on from a cell with the same settings; loops stop by comparing addresses.
- `Replace` writes each cell again as if typed, so it can turn text into a number or break a formula (**probe**).

### 3.2 Ranges of several areas, `Union`, `SpecialCells`

- `Areas`, `Count` (cells in two areas count twice), `Address`, the order of `For Each`.
- `Rows.Count`, `Row` and reading `Value` give the *first area only*. Macros lean on this without knowing
  (`visible.Rows.Count`). Writing `Value` reaches every area.
- `Copy` works when the areas share their rows or their columns, and pastes them closed up: that is how visible rows
  of a filtered list are copied. Any other shape is error 1004.
- `Union` of ranges that touch or overlap: **probe** what `Address` and `Count` give.
- `SpecialCells`: visible cells, blanks, constants and formulas by kind of value, the last cell. None found is error
  1004, which macros catch. It looks no further than the used range. On a single cell it means the whole sheet.
  A formula that gives `""` is not blank.

### 3.3 `Collection`, `Dictionary`, `RemoveDuplicates`

- **`Collection`** is VBA's own and can be recorded on a Mac. It counts from 1. Keys are text, compared without
  regard to case. A key used twice is error 457, a key not there error 5, an index out of range error 9: the usual
  way to build a list without repeats is to add under `On Error Resume Next`, so the numbers matter.
  `Dim c As New Collection` comes back to life when used after `Set c = Nothing`.
- **`Scripting.Dictionary`** does not exist on a Mac, so nothing about it can be recorded there. Reading a key that
  is not there *adds* it. `Add` twice is error 457. Keys compare exactly unless `CompareMode` says otherwise, `1`
  and `"1"` are two keys, and a `Range` used as a key is the cell, not its value. `Keys` and `Items` count from 0,
  in the order added. It stays refused until someone records it on Windows.
- **`RemoveDuplicates`**: which columns, whether there is a header (`xlGuess`: refuse, as `Sort` does), the first of
  each kept. **Probe**: no `Columns` given; `a` against `A`; `1` against `"1"`; the same number in two formats;
  hidden rows; formulas whose rows move up. Cells outside the range do not move.

## 4. Numbers as text

One reader of format codes, used by `Range.Text`, filters, `Find`, the `TEXT` function and VBA's `Format`.
[`numfmt.shown`](../macropress/numfmt.py) covers General and plain digits today.

- **Sections**: positive; negative; zero; text. Conditions (`[>=100]`). Colours, which text ignores. A format with
  one section adds the minus sign; one with a negative section does not.
- **Digits**: `0`, `#`, `?`; a comma between them for thousands; *a comma after them divides by a thousand*
  (`#,##0,` and `0.0,,"M"` are everywhere in finance); `%` multiplies by 100; `E+00`; fractions.
- Rounding is half away from zero, on the 15-digit number. A small negative under `0.00`: **probe** (`-0.00`).
- **Literals**: `"text"`, `\x`, bare characters, `_x` (a space as wide as x). `*x` fills the cell's width, which
  no text can show: stop.
- **Dates and times**: `m` is minutes straight after `h` or straight before `s`, and months otherwise.
  `mmm`, `mmmm`, `ddd`, `dddd` in English under US settings; a format that names another language (`[$-407]`): stop.
  `AM/PM`, `[h]`, `ss.00`. Day 0 is the 0th of January 1900, day 60 the 29th of February 1900, and a negative day
  shows as `####` whatever the width.
- **General**: at most 11 characters, then rounding or an exponent. Finish the rules that `shown` stops at.
- **Width** still decides `Range.Text` alone: a number that does not fit is rounded or shows `####`. Keep stopping.
- **`TEXT()`** uses the same codes. What it does with `_x` and `*x`: **probe**.
- **`Format()`** is its own dialect: `n` for minutes, named formats (`"Currency"`, `"Percent"`, `"Short Date"`,
  `"Yes/No"`), `q`, `w`, `ww`, `y`, and `<`, `>`, `@`, `&`, `!` for text. Named formats and separators go by US
  settings.

## 5. Appearance

- **A style is a combination** of number format, font, fill, border, alignment and protection, kept once in the
  file. Changing one of them over a range must keep each cell's others. Combinations are counted and shared, or the
  file's table of styles grows without end.
- **Rows, columns and the whole sheet** take their style themselves, as number formats do.
- **Reading comes with writing.** `Interior.Color`, `ColorIndex` and `Font.Color` of cells already in the file need
  the workbook's theme and its tints, and its palette of 56 colours, which a workbook may redefine. A colour is a
  `Long` with blue first. A range of mixed values reads as `Null`.
- Conditional formatting changes what is seen, not what `Interior.Color` returns. `DisplayFormat`: refuse.
- **Borders**: the edges and the inside lines of a range are edges of its cells. Line style and weight together
  pick one of the file's border styles. Two cells share an edge: **probe** what the neighbour reads.
- **Widths**: a column's width is counted in digits of the workbook's default font, and the file stores it with the
  cell's padding. Keep a table for the usual fonts, and stop for others. Row heights are in points.
- **`AutoFit`** needs the width of text in a real font, which is not to be had here. It never changes a value. Keep
  the note, and treat the column as wide enough for `Range.Text` afterwards, since that is what it asks for.
- **Merged cells**: only the top-left value is kept, and Excel asks first when others hold data: a question.
  `MergeArea`, `UnMerge`, `Merge Across`. Reading the other cells gives `Empty`; writing them: **probe**. Sorting
  over merged cells of different sizes is an error.
- **The view and the page**: frozen panes, zoom, gridlines, tab colour, `PageSetup` (the print area is a defined
  name). All of it is state in the file. `PrintOut`: refuse.
- `ClearFormats`. And everything that moves cells must now carry their look: copy, fill, sort, insert, delete.
- **Checking the writer**: besides probes that set and read back, have Excel open MacroPress's saved copy and read
  the same properties.

## 6. Rows, columns and sheets that come and go

The largest piece. Both halves change the saved file's structure, which the writer avoids today.

### 6.1 `Insert` and `Delete`

- **What moves**: values, formulas and formats below or to the right; merged areas; hidden rows and row heights;
  the AutoFilter's range; tables; spills; the selection. Whole rows and columns, or only the cells of a range, with
  `Shift`. With `Shift` left out Excel decides by the range's shape: **probe** the rule.
- **What follows**: references to moved cells, in formulas on every sheet and in defined names. A reference into
  deleted cells becomes `#REF!` in the formula's text. A range loses the rows deleted inside it and gains the rows
  inserted inside it; an insert at its first row moves it; an insert just below its last row does not stretch it.
  `A:A` does not change when rows do. A reference held in text (`INDIRECT("A5")`) does not follow.
- [`shift_formula`](../macropress/excel/addresses.py) moves references as a copy does and refuses one pushed off
  the sheet. This needs the other kind of move, by where the insert is, and `#REF!`.
- **What Excel refuses**: pushing cells that are not empty off the sheet, part of an array formula, part of a merged
  area, a protected sheet.
- **With a filter on**: **probe** `Rows("2:100").Delete` and the delete of visible cells' rows. They differ.
- **Loops**: macros delete from the bottom up. One that goes from the top down skips rows in Excel, and must skip
  the same rows here.
- `Insert` puts in the copied cells when something has been copied, not blank ones. `CopyOrigin` picks the format.
- **The saved copy**: rows are renumbered, and every part that names cells is rewritten or the save stops: merged
  cells, conditional formats, validation, hyperlinks, the filter, tables, names, the anchors of buttons and shapes,
  notes, chart series, pivot sources, anything in an extension list. Start with the known ones and stop when a part
  not understood names the sheet, as the check for renamed sheets does today.
- **Speed**: ten thousand rows deleted one at a time must not rewrite every formula each time.

### 6.2 Sheets

- `Sheets.Add`: before the active sheet unless told, and it becomes the active one. The name it gets comes from a
  counter Excel keeps while it is open: **probe** what a fresh workbook gives. Excel must still open the copy when
  the VBA project has no module for the new sheet: check by opening it.
- `Copy` inside the workbook: the name `Name (2)`, names that belong to the sheet, tables renamed, buttons with
  their macros, and the sheet's own code. Copying code means rewriting the VBA project: stop when the sheet has
  any. With no destination it makes a new workbook (group 7).
- `Move`: names that belong to a sheet are stored by its position, and must be renumbered.
- `Delete` asks first unless `DisplayAlerts` is off: a question. The last visible sheet cannot go. References to it
  become `#REF!`. Which sheet is active afterwards: **probe**.
- **Names**: `Names.Add`, `Range.Name =`, for the workbook or for one sheet; names that look like cells are
  refused; a formula that was `#NAME?` comes right when its name appears.
- **Tables** afterwards: `ListRows.Add`, `DataBodyRange` (`Nothing` when empty), columns that fill themselves.

## 7. Other workbooks and files

**First decide what a run may touch.** Recommended: nothing, unless a folder is given (`--files DIR`). Paths a macro
names are looked up inside it, `ThisWorkbook.Path` is that folder, and a path that leads out of it stops the run.
Files written go to their own folder and are listed in the report. No network paths, and no file given to a run is
ever changed.

- **Which workbook is active.** Once a second one is open, every `Range` and `Sheets` with no workbook in front
  goes to the active one. Macros rely on what `Open`, `Add`, `Activate` and `Close` leave active: **probe** each.
- **`Workbooks.Open`**: links may be updated (Excel asks unless told: a question), a password, read-only. A `.csv`
  is read as typed text, by regional settings: numbers by the rules already checked, dates refused, which stops most
  CSVs. That needs a rule of its own. `.xls` and `.xlsb`: refuse. A workbook's `Workbook_Open` runs when it is
  opened.
- **`Workbooks.Add`**: how many sheets depends on a setting of the machine. Say which is assumed.
- **`SaveAs`**: the format by number or by extension; saving macros into an `.xlsx` drops them and asks first; a
  CSV is the active sheet only, as its cells *show* (group 4); writing over a file asks. `Close` with changes
  unsaved asks. All questions.
- **A writer for new workbooks.** Today's writer changes a copy of the original. A workbook made by `Add` has no
  original.
- **Between workbooks**: copies carry formats and names across. Formulas that read another workbook are stored with
  the values last seen, and use those while it is closed. `UpdateLinks`, `BreakLink`: refuse at first.
- **VBA's own file statements**: `Open … For Input/Output/Append`, `Print #`, `Write #` (which quotes text and
  writes dates and logical values its own way), `Line Input #`, `EOF`, `FreeFile`, `Dir` in a loop (the order is
  the file system's), `Kill`, `Name`, `MkDir`. Text is in the machine's code page, with Windows line ends.
  `FileSystemObject`: needs a Windows recording.
- **Asking for a file** (`GetOpenFilename`, `FileDialog`): a question, answered with a path.
- **PDF and printing**: drawing a sheet as Excel does is out of reach. Stay refused, and put the request in the
  report (which sheet, which file) for whoever has Excel.
- **Mail**: never sent. Model enough of Outlook's objects to *keep* a message (to, subject, body, attachments) and
  list it in the report as not sent. Windows only, so it cannot be recorded on the Mac.
- **Queries, databases, the web**: refused.

## 8. Left for Excel, or refused on purpose

- **PivotTables** (`RefreshTable`, `PivotCache.Refresh`, `RefreshAll`). Refreshing one here would mean writing
  Excel's pivot engine. Instead: mark the pivot's cache to refresh when the file is opened, say so in the report,
  and from then on treat the pivot's cells as unknown, so that a macro which reads them, or a `GETPIVOTDATA`
  formula, stops. A new source range is written into the cache.
- **`GoalSeek`** leaves the changing cell where Excel's search stopped, not at the exact answer. Where it stops
  depends on Excel's method, on the workbook's iteration settings and on the starting value, and macros paste the
  result as a value. Keep refusing. If it is wanted: solve it exactly, and say in every report that the last digits
  are not Excel's. Solver, data tables and scenarios are of the same kind.
- **Charts**: making or formatting one stays refused. Charts in the file are kept, and group 6 must keep their
  ranges right or stop.
- **UserForms and range pickers** (`Application.InputBox(Type:=8)`) could take answers as `MsgBox` does.
- **Event procedures** stay as a note in the report. Running `Worksheet_Change` on every write is a decision of its
  own, since it changes results in workbooks that lean on it.
- **Class modules, `Type`, `Property` procedures**: language work that needs none of the above. It can be done at
  any point, and all of it can be recorded on a Mac.
