# Using MacroPress from an AI agent

People send agents spreadsheets, and many of those spreadsheets are run with buttons: "press Prepare, then Run
Split". An agent with MacroPress can do that itself, on a Linux sandbox with no Excel. It lists the buttons, presses
the right ones in order, and hands back the updated workbook with a summary of what changed.

MacroPress suits agents for four reasons:

- **The report is made for a model to read**: each button's status, which cells changed with examples, and what
  could not be done. It is JSON, it never holds a traceback, and its texts and lists are clipped so it stays small.
- **It refuses rather than guesses**: when a macro needs something MacroPress doesn't model, the step says so, and
  the agent can tell the user instead of handing back a workbook that is quietly wrong.
- **The original is never touched**: results go to a copy, so the agent can try, look at the result, and try again.
- **Runs repeat exactly**: with a fixed `--now` and the same answers, the same buttons always give the same workbook.

## Two tools

Give the agent one tool to look, and one to act. In JSON Schema, which most agent frameworks and model APIs accept:

```json
{
  "name": "list_workbook_buttons",
  "description": "List the buttons in an Excel workbook (.xlsm), with the caption each shows and the macro it runs. Call this before pressing any.",
  "input_schema": {
    "type": "object",
    "properties": {"path": {"type": "string", "description": "Path to the .xlsm workbook"}},
    "required": ["path"]
  }
}
```

```json
{
  "name": "press_workbook_buttons",
  "description": "Press buttons in an Excel workbook, in order, on a copy of it, without Excel. Returns a JSON report: each button's status, the cells that changed, the path of the saved copy, and what could not be done. The original workbook is never changed.",
  "input_schema": {
    "type": "object",
    "properties": {
      "path": {"type": "string", "description": "Path to the .xlsm workbook"},
      "buttons": {
        "type": "array", "items": {"type": "string"},
        "description": "Button captions or macro names, in the order to press them"
      },
      "answers": {
        "type": "array", "items": {"type": "string"},
        "description": "Answers to the questions the macros ask with MsgBox or InputBox, in the order asked: yes, no, ok, or text"
      }
    },
    "required": ["path", "buttons"]
  }
}
```

## A small implementation

Call the command-line tool in a subprocess rather than importing MacroPress: that keeps a run's memory and time
apart from the agent's, and lets the agent run on any Python version while MacroPress runs on 3.11 (install it as a
command-line tool, as in the README's [Install](../README.md#install) section).

```python
import json
import os
import subprocess


def _macropress(*args, timeout):
    done = subprocess.run(["macropress", *args], capture_output=True, text=True, timeout=timeout)
    return json.loads(done.stdout.strip().splitlines()[-1])  # the report is always the last line


def list_workbook_buttons(path):
    return _macropress("buttons", path, timeout=120)


def press_workbook_buttons(path, buttons, answers=(), out_dir="outputs"):
    os.makedirs(out_dir, exist_ok=True)
    out = os.path.join(out_dir, os.path.splitext(os.path.basename(path))[0] + "-done.xlsm")
    # --button=... keeps a caption that starts with "-" from being read as an option
    args = ["press", path, f"--out={out}", *(f"--button={b}" for b in buttons), *(f"--answer={a}" for a in answers)]
    return _macropress(*args, timeout=900)
```

### As an MCP server

The same two functions work as tools in a [Model Context Protocol](https://modelcontextprotocol.io) server, here with
the official Python SDK (`pip install mcp`):

```python
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("macropress")


@mcp.tool()
def list_buttons(path: str) -> dict:
    """List the buttons in an Excel workbook (.xlsm), with the caption each shows and the macro it runs."""
    return list_workbook_buttons(path)


@mcp.tool()
def press_buttons(path: str, buttons: list[str], answers: list[str] | None = None) -> dict:
    """Press buttons in an Excel workbook, in order, on a copy, without Excel; return the report."""
    return press_workbook_buttons(path, buttons, answers or [])


if __name__ == "__main__":
    mcp.run()
```

## What to tell the agent

The tools work best with a few rules in the agent's instructions:

- List the buttons before pressing any, and press them in the order a person would. If the user's request doesn't
  make the order clear, ask.
- A step that is `stopped: not modelled` means MacroPress cannot do that part. Tell the user what was missing (the
  step's `message`), and don't present the copy as finished.
- When a step stops because a macro asks a question (`pass an answer`), ask the user, then press again with their
  answers in `answers`.
- Say the run worked only when every step is `completed`. Mention `unverified_behaviours` when the user is relying on
  exact results.
- Hand back the workbook at `output`. The original is unchanged.

## Running it safely

MacroPress has no way to reach files, the network or other programs from a macro (see [SECURITY.md](../SECURITY.md)),
but a workbook from a stranger still deserves a sandbox:

- **Time**: a macro can loop for a long time. Give each run a timeout, as above.
- **Memory**: MacroPress holds the workbook in memory, roughly ten times the size of its unzipped sheets. A workbook
  that unzips to 200 MB needs about 2 GB. It refuses anything that unzips to over 512 MiB.
- **Files**: give it read access to the input and write access to one output folder, as a user with no other rights.
- **Network**: it needs none, so run it without.
