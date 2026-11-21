# Security

MacroPress is built to run workbooks from people you don't know, so its security matters. This page describes what
it protects against, what it doesn't, and how to report a problem.

## Reporting a vulnerability

Please report it privately, not in a public issue: use **Report a vulnerability** on the repository's
[Security tab](https://github.com/NupurBhaisare/macropress/security) (GitHub's private vulnerability reporting).
We aim to reply within a week, and to credit you in the fix unless you'd rather we didn't.

Security fixes go into the latest release.

## What MacroPress does with a workbook

- **It never runs Excel or VBA.** It reads the VBA as text, converts it to Python, and runs that Python in its own
  process against a model of the workbook.
- **The model has no way out.** A macro can read and write the workbook's cells and nothing else. Whatever reaches
  outside (`CreateObject`, `Shell`, `Dir`, `Kill`, `Open ... For Output`, `Workbooks.Open`, `SaveAs`, `Environ`)
  stops the run with `not modelled` before it does anything, and a Windows API `Declare` cannot even be converted.
- **Workbook text never becomes code.** Captions, sheet and module names and string literals reach the generated
  Python only as quoted strings or sanitised names. [tests/test_generated_code.py](tests/test_generated_code.py) tries
  to break out of them.
- **The input is never written.** Results go to a copy at another path, and saving over the input is refused.
- **Zip bombs stop before anything is unzipped.** A workbook that declares more than 512 MiB unzipped, or 256 MiB in
  one part, is refused, and Python's zip reader never inflates a part past its declared size.
- **XML entity attacks fail.** Both XML parsers it uses (Python's ElementTree, and lxml 5 or later) refuse external
  entities, so a workbook cannot pull a local file into the saved copy, and refuse entity-expansion bombs.

## What you should still do

MacroPress runs the converted code with your process's rights, and a bug in the converter or the model could let a
crafted workbook do more than it should. For workbooks from people you don't trust:

- Run MacroPress in a separate process (the command line, not the Python API), with a timeout: a macro can loop for a
  long time.
- Limit its memory. A run needs about ten times the unzipped size of the workbook's sheets.
- Run it in a container or sandbox as a user with no other rights, read access to the input, write access to one
  output folder, and no network: MacroPress needs none.
- Treat the saved copy as you would the original: it still holds the original's VBA, which Excel would run if someone
  opened the copy and enabled macros.

## Dependencies

MacroPress reads VBA projects with [oletools](https://github.com/decalage2/oletools), parses VBA with
[vb2py](https://pypi.org/project/vb2py/), and writes XML with [lxml](https://lxml.de). A vulnerability in one of them
may affect MacroPress; keep them updated within the versions `pyproject.toml` allows.
