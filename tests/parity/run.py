"""Run the probes in MacroPress and set its answers beside Excel's. Used by the test, and by hand:

    python tests/parity/run.py            list the answers that differ from Excel's and are not in known.json
    python tests/parity/run.py --write    write known.json again, from what MacroPress answers now

known.json holds the answers MacroPress knowingly gives otherwise than Excel, each group with its reason, and the
tests MacroPress stops at rather than answer. A new entry is written with the reason "not looked at": the test fails
until somebody has looked, and either fixed MacroPress or written down why the answers differ.
"""

import json
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path[:0] = [str(HERE), str(HERE.parent.parent)]

import workbook  # noqa: E402

from macropress import codegen, runtime  # noqa: E402
from macropress.errors import VBAError  # noqa: E402
from macropress.source import build_project  # noqa: E402

NOT_LOOKED_AT = "not looked at"


def load():
    """Excel's answers, and what is known to differ: ({test: {probe: answer}}, known.json's content)."""
    excel = json.loads((HERE / "excel.json").read_text(encoding="utf-8"))["answers"]
    known = json.loads((HERE / "known.json").read_text(encoding="utf-8"))
    return excel, known


class Probes:
    """The probe workbook, built once, and its macros converted a module at a time."""

    def __init__(self, folder):
        self.path, self.probes, self.index = workbook.build(folder, with_macros=False)
        self.project = build_project(list(self.probes.module_sources()[0].items()))
        self.modules = sorted({module for _row, module, _name in self.index})

    def answers(self, module):
        """{test: (how its run ended, the message it ended with, {probe: answer})} for each test of the module."""
        tests = [(row, name) for row, m, name in self.index if m == module]
        try:
            macros = self._convert(module, [name for _row, name in tests])
        except VBAError:  # something in the module does not convert: its tests one by one, to find which
            macros = None
        out = {}
        for row, name in tests:
            try:
                macro = (macros or self._convert(module, [name]))[name]
            except VBAError as err:
                out[name] = ("stopped: not modelled", str(err), {})
                continue
            ctx = runtime.Context(str(self.path), skip_appearance=module in self.probes.SKIP_APPEARANCE)
            status, message = runtime.run_macro(macro, ctx)
            ctx.book.recalc_all()
            cells = ctx.book.find_sheet("Out").rows.get(row, {})
            out[name] = (status, message, {str(c - 1): cell.value for c, cell in cells.items() if c > 1})
        return out

    def _convert(self, module, names):
        source = codegen.convert(self.project, "parity.xlsm", [(module, name) for name in names])
        namespace = {}
        exec(compile(source, f"<{module}>", "exec"), namespace)
        return namespace["MACROS"]


def same(a, b):
    """Whether two answers are one: numbers to the 12 digits a cell shows of them."""
    if (
        isinstance(a, (int, float))
        and isinstance(b, (int, float))
        and not isinstance(a, bool)
        and not isinstance(b, bool)
    ):
        return float(f"{a:.12g}") == float(f"{b:.12g}")
    return a == b and type(a) is type(b)


def verdict(name, outcome, excel, known):
    """What is wrong with one test's answers, as lines of text; none when all is as expected. `known` is
    ({test: {probe: MacroPress's answer}}, {test: the start of the message it stops with})."""
    status, message, cells = outcome
    differences, refused = known
    if name in refused:
        if status == "stopped: not modelled" and message.startswith(refused[name]):
            return []
        return [f"{name}: expected to stop with {refused[name]!r}, but: {status} {message}"]
    if status != "completed":
        return [f"{name}: {status}: {message}"]
    wrong = []
    for probe in sorted(set(cells) | set(excel.get(name, {})), key=int):
        got, want = cells.get(probe), excel.get(name, {}).get(probe)
        expected = differences.get(name, {}).get(probe, want)
        if not same(got, expected):
            wrong.append(f"{name} probe {probe}: MacroPress {got!r}, Excel {want!r}")
        elif probe in differences.get(name, {}) and same(got, want):
            wrong.append(f"{name} probe {probe}: listed in known.json as different from Excel, and it is not")
    return wrong


def known_tables(known):
    differences = {}
    for group in known["differences"]:
        for name, probes in group["tests"].items():
            differences.setdefault(name, {}).update(probes)
    return differences, known["refused"]


def main(write):
    excel, known = load()
    tables = known_tables(known)
    why = {name: group["why"] for group in known["differences"] for name in group["tests"]}
    groups, refused, wrong = {}, {}, []
    with tempfile.TemporaryDirectory() as folder:
        probes = Probes(folder)
        for module in probes.modules:
            for name, (status, message, cells) in probes.answers(module).items():
                wrong += verdict(name, (status, message, cells), excel, tables)
                if status == "stopped: not modelled":
                    refused[name] = known["refused"].get(name) or message[:60]
                    continue
                differing = {
                    p: cells.get(p)
                    for p in sorted(set(cells) | set(excel.get(name, {})), key=int)
                    if not same(cells.get(p), excel.get(name, {}).get(p))
                }
                if differing:
                    groups.setdefault(why.get(name, NOT_LOOKED_AT), {})[name] = differing
    print("\n".join(wrong) or "every answer is Excel's, or a difference known.json lists")
    if write:
        order = [g["why"] for g in known["differences"]] + [NOT_LOOKED_AT]
        out = {
            "differences": [{"why": w, "tests": groups[w]} for w in order if w in groups],
            "refused": dict(sorted(refused.items())),
        }
        (HERE / "known.json").write_text(json.dumps(out, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
        print(f"wrote known.json: {sum(len(g['tests']) for g in out['differences'])} tests differ, {len(refused)} stop")


if __name__ == "__main__":
    main("--write" in sys.argv[1:])
