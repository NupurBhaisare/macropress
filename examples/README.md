# Examples

Three small workbooks with buttons, each the kind of macro people really write. Every button here was also pressed
in Excel for Mac, and gives the same cells and the same hidden rows there as in MacroPress.

Each workbook is built from the `.bas` file beside it by [build.py](build.py), which also describes its sheets and
buttons, so the VBA inside each `.xlsm` is exactly what its `.bas` file says. Read it before you open the workbook,
as you should with any macro-enabled file. A test checks that each committed workbook matches its sources.

To try them, install MacroPress (see the [README](../README.md#install)) and run the commands below from the
repository's root folder. Each saves its result as `*-done.xlsm`, which you can open in Excel.

## invoice

An invoice with five lines and a tax rate. Its macros are in [invoice/Invoice.bas](invoice/Invoice.bas).

| Button | What it does |
| --- | --- |
| Calculate totals | Writes each line's amount as a formula (`=B5*C5`), then the subtotal (`SUM`), the tax (`ROUND`) and the total |
| Apply discount | Asks for a percentage (`InputBox`), checks with you (`MsgBox ... vbYesNo`), then lowers every unit price, rounding to cents the way VBA does |
| Clear totals | Clears what Calculate totals wrote |

```console
$ macropress press examples/invoice/invoice.xlsm \
    --button "Apply discount" --answer 10 --answer yes \
    --button "Calculate totals" --out invoice-done.xlsm
```

The unit prices become 4.05, 6.52, 10.80, 5.76 and 22.49 (7.25 less 10% is 6.525, and VBA's `Round` rounds a half to
even, so 6.52), and the total comes to 94.26. Without `--answer`, the run stops at the first question and says it
needs an answer.

## orders-cleanup

A list of orders as they often arrive: blank rows, the same order twice, names in every case, stray spaces. Its
macros are in [orders-cleanup/Cleanup.bas](orders-cleanup/Cleanup.bas).

| Button | Sheet | What it does |
| --- | --- | --- |
| Clean up | Orders | Copies the orders to the Clean sheet without blank rows or repeats, with tidy names ("MAPLE AND FINCH " becomes "Maple And Finch"), sorted by date and then order number |
| Show North only | Clean | Filters the clean orders to the North region |
| Show all orders | Clean | Shows every clean order again |

```console
$ macropress press examples/orders-cleanup/orders-cleanup.xlsm \
    --button "Clean up" --button "Show North only" --out orders-done.xlsm
```

Six of the ten rows are kept: two are blank and two repeat an order. The filter hides the rows of the South, West
and East orders, which the report shows as `"hidden_rows_changed": true`. The Orders sheet is left as it was.

## sales-summary

Sixteen sales across four regions, and a summary sheet. Its macros are in
[sales-summary/Summary.bas](sales-summary/Summary.bas).

| Button | What it does |
| --- | --- |
| Build summary | Counts the orders, units and revenue of each region, sorts the regions by revenue, adds a total row and each region's share (formulas), and stamps the time |
| Email summary | Tries to send the workbook with Outlook (`CreateObject("Outlook.Application")`) |

```console
$ macropress press examples/sales-summary/sales-summary.xlsm \
    --button "Build summary" --out summary-done.xlsm --now "2026-04-01 09:30:00"
```

South comes first with 1,372.30 of the 4,891.10 total, and the summary says "Built 2026-04-01 09:30:00 from 16
sales": with `--now`, the time is the same on every run.

```console
$ macropress press examples/sales-summary/sales-summary.xlsm --button "Email summary" --out summary-done.xlsm
```

This one stops: `CreateObject is not modelled (it starts another application or library)`. MacroPress cannot send
mail, and saying so is better than reporting a mail that never went out.

## Changing an example

Edit the `.bas` file, or the sheets and buttons in [build.py](build.py), then rebuild every workbook:

```console
$ python examples/build.py
```

`build.py` writes the parts of an `.xlsm` that are not plain XML by hand: the VBA project (its streams, compressed as
VBA compresses them) and the compound file that holds it. The VBA project has no compiled code, so Excel compiles the
VBA from source when it opens the workbook, exactly as MacroPress reads it.
