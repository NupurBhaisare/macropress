# How MacroPress works

```
.xlsm ─► find the buttons ─► read the VBA ─► parse ─► convert to Python ─► run on a model ─► save a copy ─► report
```

Pressing a button in MacroPress takes the same path every time. This page follows it step by step and names the file
that does each part, so you know where to look when you want to change something.

## Three rules

- **Never guess.** When a macro needs something the model doesn't cover, MacroPress raises `VBANotImplemented`. The
  button stops there with `stopped: not modelled`, and the report says what was missing. Unlike VBA's own errors, an
  `On Error` handler in the macro cannot catch it. A run either matches Excel, as far as MacroPress knows, or it stops.
- **Never write to the input.** The workbook is read once. Results go to a copy at another path, and saving over the
  input is refused.
- **Repeatable runs.** `Now`, `Date`, `Time` and `Timer` read a clock that starts at `--now` and moves only as the
  macros read it. Answers to `MsgBox` and `InputBox` come from `--answer`, `Environ` from `--env`, and `Rnd` is
  refused. The same workbook, buttons and answers always give the same result.
- **Excel is the judge.** What MacroPress does is set beside what Excel does: about a thousand small macros were run
  in Excel and their answers recorded, and every test run compares ([tests/parity](../tests/parity)).

## 1. Read the workbook ([xlsx.py](../macropress/xlsx.py))

An `.xlsm` file is a zip of XML parts. MacroPress first checks the sizes the zip declares, and refuses a file that
would unzip to more than 512 MiB, or 256 MiB in one part, before reading anything: zip bombs stop there. It then reads
each sheet's cells (values, formulas including shared ones, and number formats, of empty cells, rows and columns
too), defined names, AutoFilters, tables and the workbook's calculation mode into the workbook model. A workbook on
the 1904 date system is refused, since its dates would be four years off.

## 2. Find the buttons ([xlsx.py](../macropress/xlsx.py), `find_entry_points`)

A button is a drawing that runs a macro when clicked. MacroPress reads two kinds: form-control buttons, stored as VML
shapes with an `<x:FmlaMacro>`, and DrawingML shapes or pictures with a `macro` attribute. For each it records the
caption it shows, the macro it runs and the sheet it sits on.

## 3. Read the VBA ([source.py](../macropress/source.py))

The VBA project lives in `xl/vbaProject.bin`, a compound file holding compressed source for each module.
[oletools](https://github.com/decalage2/oletools) reads it. MacroPress sorts the modules into standard modules,
`ThisWorkbook`, sheet modules and class modules, and splits each into its declarations and procedures. `macropress
convert --vba-dir` reads exported `.bas` and `.cls` files instead.

## 4. Parse ([vbparse.py](../macropress/vbparse.py), [nodes.py](../macropress/nodes.py))

Each procedure is parsed with [vb2py](https://pypi.org/project/vb2py/)'s VBA grammar, and the parse tree is turned
into a small set of nodes: statements (`Assign`, `If`, `For`, `Select`, `With`, `OnError`, ...) and expressions
(`Call`, `Member`, `BinOp`, ...). Each node keeps its line number, so every error can point at a VBA line.

## 5. Convert to Python ([codegen.py](../macropress/codegen.py))

Each procedure becomes a Python function, and each VBA line becomes a comment followed by its translation:

```python
# L12: lastRow = LastItemRow(ws)
v.lastRow = LastItemRow(xl, rt.ref(v, "ws"))
# L13: For r = 5 To lastRow
```

VBA's operators and conversions are not Python's (`"1" + 1` is `2`, `&` joins text, `Empty = 0` is true, `/` and
`\` divide differently, `Round` rounds half to even), so expressions call the runtime (`rt.add`, `rt.cat`, `rt.eq`)
rather than using Python's operators. Variables live in a frame that knows their declared types. Structured code
stays structured (a `GoTo` out of a loop becomes `break`); a procedure with any other `GoTo`, or with
`On Error GoTo`, is written as numbered steps in a loop, so a jump can land on any line.

Text from the workbook (button captions, file and module names) only goes into the generated code as quoted Python
strings or sanitised names: [tests/test_generated_code.py](../tests/test_generated_code.py) tries to break out of
them.

## 6. Run on a model of the workbook

The converted functions run against the model:

| File | What it models |
| --- | --- |
| [runtime.py](../macropress/runtime.py) | A run: the workbook, the clock, the answers, `Debug.Print` and `MsgBox` output, VBA's calls, `ByRef`, `With`, `For Each`, `On Error` and `Err` |
| [values.py](../macropress/values.py) | VBA's values: Variants and their conversions, arithmetic and comparison, `Empty`/`Null`/`Nothing`, arrays, error values, overflow |
| [builtins.py](../macropress/builtins.py) | VBA's functions: `Left`, `InStr`, `Format`, `DateAdd`, `CDbl`, `MsgBox`, ... |
| [excel/](../macropress/excel/) | Excel's objects: `Application`, `Workbook`, `Worksheet`, `Range`, `Sort`, `AutoFilter`, `WorksheetFunction`, copy and paste, the selection and grouped sheets |
| [entry.py](../macropress/entry.py), [numfmt.py](../macropress/numfmt.py) | What Excel makes of text written to a cell (`"$5"`, `"5%"`, `"1,234"`), and number formats: what VBA reads from a cell that has one, and which format a write leaves |
| [formula.py](../macropress/formula.py) | Worksheet formulas: a tokenizer and parser for Excel's formula syntax, and an evaluator for the functions listed in [supported.md](supported.md) |
| [excel/book.py](../macropress/excel/book.py) | Recalculation: which formulas read which cells (an index, so a write finds them without passing the rest), which are stale after a write, and calculating them before their values are read |

Cells hold values the way Excel stores them: numbers as floats, text as strings, `TRUE` and `FALSE` as booleans,
errors as error values, and dates as numbers with a date format. Reading `Range.Value` turns them back into what VBA
would see (a `Date`, a `Currency`). Writing a value that Excel would turn into something else, such as text that looks
like a date, is refused rather than guessed.

## 7. Save a copy ([writer.py](../macropress/writer.py))

MacroPress doesn't write a new workbook. It copies the original package and rewrites only what changed: the cells of
the sheets the macros touched (values, formulas, number formats where needed), hidden rows and AutoFilters, the active
sheet, and defined names. Every other part (the VBA project, styles, drawings, charts, other sheets) is copied byte for
byte, so whatever MacroPress doesn't model is kept, not lost. Changes it cannot save faithfully, such as a sheet added
or deleted, or a renamed sheet that a chart still refers to, stop the save rather than produce a broken file.

## 8. Report ([api.py](../macropress/api.py), [snapshot.py](../macropress/snapshot.py), [ledger.py](../macropress/ledger.py))

Before and after the buttons, MacroPress takes a snapshot of every sheet's cells and view. Their difference becomes
the report's `changes`.

Some of what the model does rests on what Excel does in corners that are easy to get wrong, for example how
`AutoFilter` on a header row picks its range. The **ledger** lists each one with its status: `verified` (checked
against Excel's own output), `unverified` (believed right, not yet checked), or `limitation` (a known difference,
such as following US regional settings). Every time the model relies on one, it notes it, and the report lists the
unverified ones and limitations that applied to the run. Today every entry is verified or a limitation; a new
behaviour starts as unverified, and adding a probe for it to [tests/parity](../tests/parity) is what verifies it.
