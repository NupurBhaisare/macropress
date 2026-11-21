"""The Python API: list a workbook's buttons, and press them on a copy, without Excel.

Both functions return the report the command line prints as JSON (the command line clips long texts and lists; these
do not). Neither raises for what a workbook holds or its macros do: what stops a run is in the report, as its
"error" or as a step's status.

    >>> import macropress
    >>> macropress.list_buttons("invoice.xlsm")["buttons"]
    [{'caption': 'Calculate totals', 'macro': 'CalculateTotals', 'sheet': 'Invoice'}, ...]
    >>> report = macropress.press("invoice.xlsm", ["Calculate totals"], out="invoice-done.xlsm")
    >>> report["steps"]
    [{'button': 'Calculate totals', 'status': 'completed'}]
"""

import os
import re

from .source import from_workbook
from .xlsx import find_entry_points

# An event procedure in ThisWorkbook or a sheet's module, which Excel runs by itself (Worksheet_Change, Workbook_Open).
EVENT = re.compile(r"^(Workbook|Worksheet|Chart)_[A-Za-z]+$")


def _rows(cells):
    return {int("".join(ch for ch in addr if ch.isdigit())) for addr in cells}


def _changes(before, after, examples=3):
    """What a run changed, per sheet: cells, rows with data, hidden rows, and a few example cells."""
    out = []
    for name, sheet in after["sheets"].items():
        a = sheet["cells"]
        b = before["sheets"].get(name, {}).get("cells", {})
        changed = sorted(
            (k for k in set(a) | set(b) if a.get(k) != b.get(k)),
            key=lambda k: (int("".join(ch for ch in k if ch.isdigit())), k),
        )
        hb = before["view"]["sheets"].get(name, {}).get("hidden_rows")
        ha = after["view"]["sheets"].get(name, {}).get("hidden_rows")
        if not changed and hb == ha:
            continue

        def show(v):
            return "" if v is None else v[1]

        out.append(
            {
                "sheet": name,
                "cells_changed": len(changed),
                "rows_with_data_before": len(_rows(b)),
                "rows_with_data_after": len(_rows(a)),
                "hidden_rows_changed": hb != ha,
                "examples": [
                    {"cell": k, "before": show(b.get(k)), "after": show(a.get(k))} for k in changed[:examples]
                ],
            }
        )
    return out


def list_buttons(workbook):
    """The workbook's buttons (caption, macro, sheet) and its macros. A workbook that cannot be read is reported as
    the report's "error"; macros that cannot be read, as its "macros_error"."""
    out = {"input": workbook, "buttons": []}
    try:
        out["buttons"] = [
            {"caption": label, "macro": macro, "sheet": sheet} for label, macro, sheet in find_entry_points(workbook)
        ]
    except Exception as err:  # noqa: BLE001 - reported, never raised
        out["error"] = f"the workbook could not be read: {type(err).__name__}: {err}"
        return out
    try:
        project = from_workbook(workbook)
        on_buttons = {b["macro"].lower() for b in out["buttons"]}
        out["macros"] = [
            {
                "module": module.name,
                "name": proc.name,
                "kind": proc.kind,
                "scope": proc.scope,
                "lines": [proc.first_line, proc.last_line],
                "on_a_button": proc.name.lower() in on_buttons,
            }
            for module, proc in project.procedures()
        ]
    except Exception as err:  # noqa: BLE001 - the buttons are still worth reporting
        out["macros_error"] = f"{type(err).__name__}: {err}"
    return out


def press(workbook, buttons, out, *, code=None, now=None, answers=()):
    """Convert the workbook's VBA, press `buttons` (captions or macro names) in order on a copy, and save the copy as
    `out`; the workbook itself is never written. A button that stops ends the run: the steps that ran, and the copy
    as they left it, are still reported and saved.

    code     also save the converted Python to this path
    now      a datetime: what Now, Date and Time return when the run starts, so a run can be repeated exactly
    answers  the answers to the macros' MsgBox and InputBox questions, in the order they are asked
    """
    from . import codegen, runtime
    from .errors import VBAEnd, VBAError, VBARuntimeError

    result = {"input": workbook, "steps": []}
    stage = "the workbook could not be read"
    try:
        entry = [(label, macro) for label, macro, _ in find_entry_points(workbook)]
        result["buttons_available"] = [{"caption": label, "macro": m} for label, m in entry]
        project = from_workbook(workbook)
        wanted = [b.strip() for b in buttons]
        known = {m.lower() for _, m in entry}
        extra = [
            (None, w)
            for w in wanted
            if w.lower() not in known and not any((label or "").lower() == w.lower() for label, _ in entry)
        ]
        pressed = [
            (label, m) for label, m in entry if any(w.lower() in ((label or "").lower(), m.lower()) for w in wanted)
        ]
        stage = "the VBA could not be converted"
        try:
            source = codegen.convert(project, os.path.basename(workbook), entry + extra)
        except Exception:  # noqa: BLE001 - a button that is not pressed must not stop the ones that are
            source = codegen.convert(project, os.path.basename(workbook), pressed + extra)
        if code:
            os.makedirs(os.path.dirname(os.path.abspath(code)), exist_ok=True)
            with open(code, "w", encoding="utf-8") as f:
                f.write(source)
            result["code"] = code
        namespace = {}
        exec(compile(source, code or "<converted>", "exec"), namespace)
        by_caption = {k.lower(): (k, f) for k, f in namespace["BUTTONS"].items()}
        macros = {k.lower(): (k, f) for k, f in namespace["MACROS"].items()}
        by_macro = {m.lower(): label for label, m in entry}
        shared = {}  # a caption two buttons show, with different macros: pressing it by caption would guess
        for label, m in entry:
            if label:
                shared.setdefault(label.lower(), set()).add(m)
        stage = "the workbook could not be read"
        ctx = runtime.Context(workbook, now=now, answers=list(answers))
        events = [f"{m.name}.{p.name}" for m, p in project.procedures() if m.kind != "standard" and EVENT.match(p.name)]
        if events:
            ctx.ledger.note("EVENTS-NOT-RUN", ", ".join(events))
        before = ctx.snapshot()
        stage = "the buttons could not be pressed"
        for want in wanted:
            if len(shared.get(want.lower(), ())) > 1 and want.lower() not in macros:
                message = f"{len(shared[want.lower()])} buttons show this caption, for macros " + ", ".join(
                    sorted(shared[want.lower()])
                )
                result["steps"].append({"button": want, "status": "not found", "message": message + ": name the macro"})
                break
            hit = by_caption.get(want.lower()) or macros.get(want.lower())
            if hit is None:
                stale = next((m for label, m in pressed if want.lower() in ((label or "").lower(), m.lower())), None)
                message = "no button or macro by that name"
                if stale:
                    message = f"its macro {stale!r} is not in this workbook (deleted, or in another workbook)"
                result["steps"].append({"button": want, "status": "not found", "message": message})
                break
            name, fn = hit
            step = {"button": name if want.lower() in by_caption else by_macro.get(want.lower(), name)}
            try:
                fn(ctx)
                step["status"] = "completed"
            except VBAEnd:
                step["status"] = "completed"
            except VBARuntimeError as err:
                step.update(status="stopped by a VBA error", message=str(err))
            except VBAError as err:
                step.update(status="stopped: not modelled", message=str(err))
            except Exception as err:  # noqa: BLE001 - a fault in the converter stops this button, not the run
                step.update(status="stopped: converter error", message=f"{type(err).__name__}: {err}")
            result["steps"].append(step)
            if step["status"] != "completed":
                break
        stage = "the updated workbook could not be saved"
        os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        ctx.save(out)
        result["output"] = out
        stage = "the changes could not be listed"
        result["changes"] = _changes(before, ctx.snapshot())
        result["unverified_behaviours"] = [line for line in ctx.ledger.report_lines() if "[verified]" not in line]
        result["messages"] = [f"{k}: {t}" for k, t in ctx.messages]
    except Exception as err:  # noqa: BLE001 - whatever stops the run is reported, with the steps that ran
        result["error"] = f"{stage}: {type(err).__name__}: {err}"
    return result


def succeeded(report):
    """Whether a press report ran every button it was asked to: no error, and every step completed."""
    return "error" not in report and all(s["status"] == "completed" for s in report["steps"])
