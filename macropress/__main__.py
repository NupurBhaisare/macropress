"""Command line: macropress <command> ... (or python -m macropress)

buttons   WORKBOOK                         the workbook's buttons and macros, as JSON
press     WORKBOOK --button B... --out X   convert, press buttons in order on a copy, save X, print a JSON summary
convert   WORKBOOK --out FILE.py           convert the workbook's VBA (all buttons) into a Python file
"""

import argparse
import datetime as dt
import json
import os
import sys

from . import __version__
from .api import list_buttons, press, succeeded
from .runtime import SKIP_APPEARANCE_HELP
from .source import from_directory, from_workbook
from .xlsx import find_entry_points

# The JSON report is meant for scripts and AI agents, so it stays small whatever the workbook holds or its macros do:
# each text is clipped, and each list keeps its first items with its full count beside it as <name>_total (a macro
# can Debug.Print once per row, and a cell holds 32,767 characters). Texts are clipped here and only here, where their
# full length is known, so a clipped one says how much was cut.
REPORT_TEXT = 500
REPORT_ITEMS = 1000


def _clip(text):
    return text if len(text) <= REPORT_TEXT else text[:REPORT_TEXT] + f"... [{len(text) - REPORT_TEXT} more characters]"


def _bounded(value):
    if isinstance(value, dict):
        out = {}
        for key, item in value.items():
            if isinstance(item, (list, tuple)) and len(item) > REPORT_ITEMS:
                out[f"{key}_total"] = len(item)
                item = item[:REPORT_ITEMS]
            out[key] = _bounded(item)
        return out
    if isinstance(value, (list, tuple)):
        return [_bounded(item) for item in value]
    if value is None or isinstance(value, (bool, int, float)):
        return value
    return _clip(str(value))


def _report(result, code):
    """Print the report, bounded, as the last line of output; return the exit code."""
    print(json.dumps(_bounded(result)))
    return code


def cmd_buttons(args):
    """The workbook's buttons (caption, sheet, macro) and its macros, as JSON on the last line."""
    out = list_buttons(args.workbook)
    return _report(out, 1 if "error" in out else 0)


def cmd_press(args):
    """Convert the workbook's VBA, press buttons in order on a copy, save it; print a JSON summary as the last line.
    Whatever stops the run is reported there as its error, with the steps that ran, never as a traceback."""
    try:
        now = dt.datetime.strptime(args.now, "%Y-%m-%d %H:%M:%S") if args.now else None
    except ValueError:
        return _report({"input": args.workbook, "steps": [], "error": "--now must be 'YYYY-MM-DD HH:MM:SS'"}, 1)
    environ = dict(item.split("=", 1) for item in args.env or [] if "=" in item)
    if len(environ) != len(args.env or []):
        return _report({"input": args.workbook, "steps": [], "error": "--env must be NAME=VALUE"}, 1)
    result = press(
        args.workbook,
        args.button,
        args.out,
        code=args.code,
        now=now,
        answers=args.answer or [],
        environ=environ,
        timeout=args.timeout,
        skip_appearance=args.skip_appearance,
    )
    return _report(result, 0 if succeeded(result) else 1)


def cmd_convert(args):
    from . import codegen

    project = from_directory(args.vba_dir) if args.vba_dir else from_workbook(args.workbook)
    entry = [(label, macro) for label, macro, _ in find_entry_points(args.workbook)]
    entry += [(None, m) for m in args.macro or [] if all(m.lower() != e[1].lower() for e in entry)]
    if not entry:
        print("no buttons found; name the macros to convert with --macro", file=sys.stderr)
        return 1
    conv = codegen.Conversion(project, os.path.basename(args.workbook), entry, force_flat=args.steps)
    code = conv.source(args.out)  # imports macropress from wherever it runs, as an installed package
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(code)
    print(f"wrote {args.out}: {len(conv.procs)} procedures, {len(code.splitlines())} lines")
    for label, macro in entry:
        print(f"  {('button ' + repr(label)) if label else 'macro'} -> {macro}")
    return 0


def main(argv=None):
    ap = argparse.ArgumentParser(
        prog="macropress", description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    ap.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("convert", help="convert the workbook's VBA into a Python file")
    p.add_argument("workbook")
    p.add_argument("--out", required=True, help="the Python file to write")
    p.add_argument("--macro", action="append", help="also convert this macro (buttons are always included)")
    p.add_argument("--vba-dir", help="folder of exported .bas/.cls modules (default: read the workbook's VBA)")
    p.add_argument("--steps", action="store_true", help="write every procedure as numbered steps (for testing)")
    p = sub.add_parser("buttons", help="list the workbook's buttons and macros, as JSON")
    p.add_argument("workbook")
    p = sub.add_parser("press", help="press buttons on a copy of the workbook, without Excel")
    p.add_argument("workbook")
    p.add_argument("--button", action="append", required=True, help="button caption or macro name, in order")
    p.add_argument("--out", required=True, help="where to save the updated workbook")
    p.add_argument("--code", help="where to save the converted Python")
    p.add_argument("--now", help="value of Now(), 'YYYY-MM-DD HH:MM:SS'")
    p.add_argument("--answer", action="append", help="answer to the next MsgBox/InputBox question (repeatable)")
    p.add_argument("--env", action="append", metavar="NAME=VALUE", help='what Environ("NAME") gives (repeatable)')
    p.add_argument("--timeout", type=float, metavar="SECONDS", help="stop the macros when they have run this long")
    p.add_argument("--skip-appearance", action="store_true", help=SKIP_APPEARANCE_HELP)
    args = ap.parse_args(argv)
    return {"buttons": cmd_buttons, "press": cmd_press, "convert": cmd_convert}[args.cmd](args)


if __name__ == "__main__":
    sys.exit(main())
