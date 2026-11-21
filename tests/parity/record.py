"""Record what Excel answers: run every probe in Excel itself and write the answers to excel.json.

    python tests/parity/record.py            every test
    python tests/parity/record.py TFilt ...  only these modules' tests; the other answers are kept

It needs Excel and xlwings (pip install xlwings), and is run by hand when probes are added: the test suite reads
excel.json and never starts Excel. The workbook is built fresh, opened, and closed without saving.

On a Mac it is opened from Excel's own folder, where Excel can open a file without asking for access to it. Excel
may still ask whether to enable the workbook's macros: say yes, they are the probes.
"""

import datetime as dt
import json
import platform
import sys
from pathlib import Path

import workbook
import xlwings

HERE = Path(__file__).resolve().parent
OFFICE = Path.home() / "Library" / "Group Containers" / "UBF8T346G9.Office"


def main(modules):
    folder = (OFFICE if OFFICE.is_dir() else HERE) / "macropress-parity"
    path, _probes, index = workbook.build(folder)
    book = xlwings.Book(str(path))
    try:
        for _row, module, name in index:
            if modules and module not in modules:
                continue
            try:
                book.macro(f"{module}.{name}")()
            except Exception as err:  # noqa: BLE001 - a macro Excel could not run is seen in its answers
                print(f"{module}.{name}: {type(err).__name__}: {str(err)[:200]}")
        grid = book.sheets["Out"].range(f"A1:P{len(index)}").value
        version = f"Excel {book.app.version} on {platform.system()}"
    finally:
        book.close()
    answers = {}
    for (_row, module, name), line in zip(index, grid, strict=True):
        if not modules or module in modules:
            cells = {str(i): v for i, v in enumerate(line[1:], 1) if v is not None}
            answers[name] = {k: f"DATE:{v.isoformat()}" if isinstance(v, dt.datetime) else v for k, v in cells.items()}
    target = HERE / "excel.json"
    if modules:
        answers = {**json.loads(target.read_text(encoding="utf-8"))["answers"], **answers}
    answers = {name: answers[name] for _row, _module, name in index if name in answers}
    lines = ",\n".join(
        f"  {json.dumps(name)}: {json.dumps(cells, ensure_ascii=False)}" for name, cells in answers.items()
    )
    target.write_text(f'{{"recorded_with": {json.dumps(version)},\n "answers": {{\n{lines}\n}}}}\n', encoding="utf-8")
    print(f"wrote {target}: {len(answers)} tests' answers, from {version}")


if __name__ == "__main__":
    main(set(sys.argv[1:]))
