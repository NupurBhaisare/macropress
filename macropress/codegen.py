"""Convert a workbook's VBA into a Python module that runs on the workbook.

Each procedure becomes a Python function. Where the VBA's jumps allow it, the code is
structured (if / while / for), with a GoTo out of a loop turned into `break` and
On Error Resume Next turned into a try/except around each statement. Otherwise the
procedure is written as a list of numbered steps, which reproduces any GoTo or
On Error GoTo exactly.
"""

import datetime as dt
import keyword
import os
import re
from decimal import Decimal

from . import nodes as n
from .builtins import BUILTINS, FRAME_AWARE, VB_CONSTANTS
from .compiler import Compiler
from .errors import VBACompileError, VBANotImplemented, VBAParseError
from .excel import XL_CONSTANTS, Globals
from .values import (
    EMPTY,
    NOTHING,
    NULL,
    NUMERIC_TYPES,
    normalise_type,
    op_add,
    op_concat,
    op_div,
    op_intdiv,
    op_mod,
    op_mul,
    op_neg,
    op_sub,
)
from .vbparse import parse_procedure

CANONICAL = {
    name.lower(): name
    for name in [
        "Len",
        "Left",
        "Right",
        "Mid",
        "Trim",
        "LTrim",
        "RTrim",
        "UCase",
        "LCase",
        "InStr",
        "InStrRev",
        "Replace",
        "Split",
        "Join",
        "StrComp",
        "StrReverse",
        "Space",
        "String",
        "Chr",
        "Asc",
        "Val",
        "Str",
        "CStr",
        "CInt",
        "CLng",
        "CDbl",
        "CSng",
        "CDec",
        "CCur",
        "CBool",
        "CDate",
        "CVar",
        "CByte",
        "Int",
        "Fix",
        "Abs",
        "Sgn",
        "Sqr",
        "Exp",
        "Log",
        "Round",
        "Now",
        "Date",
        "Time",
        "Timer",
        "Year",
        "Month",
        "Day",
        "Hour",
        "Minute",
        "Second",
        "Weekday",
        "DateSerial",
        "TimeSerial",
        "DateAdd",
        "DateDiff",
        "DateValue",
        "TimeValue",
        "Format",
        "IsEmpty",
        "IsNull",
        "IsNumeric",
        "IsDate",
        "IsArray",
        "IsObject",
        "IsError",
        "IsMissing",
        "TypeName",
        "VarType",
        "Array",
        "UBound",
        "LBound",
        "IIf",
        "Choose",
        "CVErr",
        "RGB",
        "MsgBox",
        "InputBox",
        "Environ",
        "DoEvents",
        "Randomize",
        "Rnd",
    ]
}

# VBA's own functions and statements the model does not cover. A call to one stops the conversion as "not modelled":
# "not defined" would read as a mistake in the macro.
NOT_MODELLED = {
    name.lower(): (name, reason)
    for reason, names in {
        "it starts another application or library": "CreateObject GetObject CallByName",
        "it uses files and folders": (
            "Dir CurDir FileCopy FileDateTime FileLen GetAttr SetAttr Kill MkDir RmDir ChDir ChDrive FreeFile EOF LOF "
            "Loc Seek Reset"
        ),
        "it reaches outside the workbook": (
            "Shell SendKeys AppActivate Beep Command GetSetting GetAllSettings SaveSetting DeleteSetting LoadPicture"
        ),
        "it shows a UserForm": "Load Unload",
        "it runs code held in text": "Run ExecuteExcel4Macro",
        "it is not covered yet": (
            "StrConv Filter FormatCurrency FormatNumber FormatPercent FormatDateTime MonthName WeekdayName DatePart Hex "
            "Oct ChrW AscW Atn Cos Sin Tan Pmt PPmt IPmt PV FV NPV IRR MIRR NPer Rate DDB SLN SYD Partition QBColor GoSub"
        ),
    }.items()
    for name in names.split()
}


# Excel's own objects that are not modelled, by the names VBA code uses without a qualifier. Read as an undeclared
# variable, Windows(1).Zoom = 80 would be "object required", an error On Error Resume Next skips.
NOT_MODELLED_GLOBALS = {
    name.lower(): name
    for name in (
        "ActiveChart ActivePrinter Windows Charts Names CommandBars Dialogs AddIns UserForms Assistant "
        "ThisCell Excel4MacroSheets DialogSheets"
    ).split()
}
# The libraries VBA's own functions can be named with: VBA.Left, Strings.Trim, DateTime.Now.
VBA_LIBRARIES = {"vba", "strings", "datetime", "conversion", "information", "interaction", "math", "financial"}

# How a comparison reads a number against a string goes by how each side is declared (values.compare): a numeric
# type "n", a String "s", or a Variant "v". What a VBA function returns, where it is not a Variant; with a type
# character (Left$) every one of them returns a String.
FUNCTION_KINDS = {
    **dict.fromkeys("cstr replace join strreverse typename inputbox".split(), "s"),
    **dict.fromkeys(
        "cint clng cdbl csng cbyte ccur cbool cdate len instrrev asc val sqr exp log rgb timer rnd ubound lbound "
        "vartype isnumeric isdate isempty isnull isarray iserror ismissing isobject msgbox".split(),
        "n",
    ),
}
MEMBER_KINDS = {
    **dict.fromkeys(("name", "address", "codename"), "s"),
    **dict.fromkeys(("row", "column", "count", "index"), "n"),
}


def type_kind(type_name):
    t = normalise_type(type_name)
    if t == "string":
        return "s"
    return "n" if t in NUMERIC_TYPES or t in ("boolean", "date") else "v"


def _undefined(proc, name):
    """The error for a call to `name` that no procedure, VBA function or Excel global answers."""
    known = NOT_MODELLED.get(name.lower())
    if known:
        return VBANotImplemented(f"{proc.name}: {known[0]} is not modelled ({known[1]})")
    return VBACompileError(f"{proc.name}: Sub or Function not defined: {name}")


BINOPS = {
    "+": "add",
    "-": "sub",
    "*": "mul",
    "/": "div",
    "\\": "idiv",
    "mod": "mod",
    "^": "pow_",
    "&": "cat",
    "and": "and_",
    "or": "or_",
    "xor": "xor",
    "eqv": "eqv",
    "imp": "imp",
    "=": "eq",
    "<>": "ne",
    "<": "lt",
    ">": "gt",
    "<=": "le",
    ">=": "ge",
    "like": "like",
    "is": "is_",
}
COMPARISONS = {"eq", "ne", "lt", "gt", "le", "ge", "like"}

GLOBAL_MEMBERS = {name.split("_", 1)[1] for name in dir(Globals) if name.startswith(("get_", "m_"))}
SHEET_RELATIVE = {"range", "cells", "rows", "columns", "evaluate"}  # in a sheet's module, these are that sheet's
RESERVED = {"rt", "xl", "v", "w", "st", "pc", "os", "sys", "BUTTONS", "MACROS"}


def ident_ok(name):
    return name.isidentifier() and not keyword.iskeyword(name)


# Names that Python's own numbers, text and dates answer to (s.count, x.real). A VBA value that is no object must say
# so when asked for one (error 424), so these members are asked through rt.member, never as Python attributes.
PYTHON_ATTRIBUTES = {a for t in (int, float, str, bool, dt.datetime, Decimal, type(None)) for a in dir(t)}


def member_ok(name):
    return ident_ok(name) and name not in PYTHON_ATTRIBUTES


def py_ident(name):
    """`name` as a Python identifier. A workbook's VBA project names its modules and procedures; nothing checks
    those names before they reach the generated code."""
    ident = "".join(ch if ("a" + ch).isidentifier() else "_" for ch in name)
    if not ident.isidentifier():
        ident = "_" + ident
    return ident + "_" if keyword.iskeyword(ident) else ident


class Conversion:
    """Generates one Python module for a VBA project and a set of entry points."""

    def __init__(self, project, workbook_name, entry_points, force_flat=False):
        self.project = project
        self.workbook_name = workbook_name
        self.entry_points = entry_points  # [(button caption or None, macro name)]
        self.force_flat = force_flat
        self.procs = {}  # (module lower, proc lower) -> (module, src, proc)
        self.py_names = {}
        self.module_ids = {}  # module name -> its M_ variable in the generated code
        self.structured = {}
        self.missing = []  # entry points whose macro is not in the project: [(caption, macro)]
        self.object_modules = {  # a sheet's module or the workbook's, by name: in code, its name stands for the object
            m.name.lower(): m for m in project.modules.values() if m.kind in ("document", "workbook")
        }
        modules = list(project.modules.values())
        self.externals = {k: v for m in modules for k, v in m.externals.items()}  # Declare: name -> (name, library)
        self.enums = {k: v for m in modules for k, v in m.enums.items()}  # Enum name -> its members
        self.types = {k: v for m in modules for k, v in m.types.items()}  # Type ... End Type names

    def type_name(self, name, what):
        """A declared type as the runtime knows it: an Enum is a Long. A Type ... End Type is not modelled."""
        key = normalise_type(name)
        if key in self.types:
            raise VBANotImplemented(f"{what}: the user-defined type {self.types[key]} (Type ... End Type)")
        return "long" if key in self.enums else name

    # --- what to convert ---
    def find(self, name, module=None):
        key = name.lower()
        if "." in key:  # Module1.Macro, as a button can name it
            mod_name, _, key = key.rpartition(".")
            mod = next((m for m in self.project.modules.values() if m.name.lower() == mod_name), None)
            src = mod.procedures.get(key) if mod is not None else None
            return (mod, src) if src is not None else (None, None)
        if module is not None and key in module.procedures:
            return module, module.procedures[key]
        for mod in self.project.modules.values():
            src = mod.procedures.get(key)
            if src is not None and (src.scope != "Private" or mod is module):
                return mod, src
        return None, None

    def collect(self):
        todo = []
        for label, macro in self.entry_points:
            module, src = self.find(macro)
            if src is None:  # a button left pointing at a deleted macro, or at another workbook's
                self.missing.append((label, macro))
                continue
            todo.append((module, src))
        while todo:
            module, src = todo.pop()
            key = (module.name.lower(), src.name.lower())
            if key in self.procs:
                continue
            if module.decl_error:
                raise VBAParseError(module.decl_error)
            proc = parse_procedure(src)
            proc.module = module
            self.procs[key] = (module, src, proc)
            for name in sorted(_called_names(proc.body)):
                m2, s2 = self.find(name, module)
                if s2 is not None:
                    todo.append((m2, s2))
        used = set()
        for (mkey, pkey), (module, _src, proc) in sorted(self.procs.items()):
            name = proc.name if ident_ok(proc.name) and proc.name not in RESERVED else f"{py_ident(proc.name)}_"
            if name.lower() in used:
                name = f"{py_ident(module.name)}_{py_ident(proc.name)}"
            while name.lower() in used:
                name += "_"
            used.add(name.lower())
            self.py_names[(mkey, pkey)] = name

    def py_name(self, module, src):
        return self.py_names[(module.name.lower(), src.name.lower())]

    def module_id(self, module):
        """The module's M_ variable: a valid identifier, and one per module, whatever the module is called."""
        if module.name not in self.module_ids:
            ident = f"M_{py_ident(module.name)}"
            while ident in self.module_ids.values():
                ident += "_"
            self.module_ids[module.name] = ident
        return self.module_ids[module.name]

    # --- output ---
    def source(self, out_path=None, runtime_dir=None):
        self.collect()
        lines = []
        buttons = [(label, macro) for label, macro in self.entry_points if label and (label, macro) not in self.missing]
        doc = [f"Python converted by macropress from the VBA of {self.workbook_name}.", ""]
        if buttons:
            doc.append("Buttons:")
            doc += [f"  {label!r:45} -> {macro}" for label, macro in buttons]
            doc.append("")
        script = os.path.basename(out_path) if out_path else "this_file.py"
        example = " ".join(f'--button "{label}"' for label, _ in buttons[:2]) or f"--macro {self.entry_points[0][1]}"
        doc += [
            "Run it on a workbook (the input file is never modified):",
            f"  python3 {script} INPUT.xlsm OUTPUT.xlsm {example}",
            "",
            "Each step shows the VBA line it comes from. The code calls macropress.runtime (rt) to keep",
            "VBA's and Excel's behaviour: comparisons and empty cells, the active sheet and selection,",
            "grouped sheets, filters, copy and paste. It needs the macropress package on the Python path.",
        ]
        lines.append(doc_literal("\n".join(doc) + "\n"))
        lines += ["import os", "import sys", ""]
        if out_path and runtime_dir:
            rel = os.path.relpath(os.path.abspath(runtime_dir), os.path.dirname(os.path.abspath(out_path)))
            lines.append(f"sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), {rel!r}))")
        lines += ["from macropress import runtime as rt  # noqa: E402", ""]
        modules = sorted({module.name: module for module, _, _ in self.procs.values()}.values(), key=lambda m: m.name)
        for module in modules:
            lines += self.module_vars(module)
        for key in sorted(self.procs, key=lambda k: (k[0], self.procs[k][1].first_line)):
            module, src, proc = self.procs[key]
            lines += ["", ""] + ProcGen(self, module, src, proc).generate()
        lines += ["", ""]
        lines.append("BUTTONS = {")
        for label, macro in buttons:
            module, src = self.find(macro)
            lines.append(f"    {label!r}: {self.py_name(module, src)},")
        lines.append("}")
        lines.append("MACROS = {")
        for _key, (module, src, proc) in sorted(self.procs.items()):
            if src.kind == "sub" and not proc.params:
                lines.append(f"    {proc.name!r}: {self.py_name(module, src)},")
        lines.append("}")
        lines += ["", "", 'if __name__ == "__main__":', "    sys.exit(rt.main(BUTTONS, MACROS, __doc__))", ""]
        return "\n".join(lines)

    def module_vars(self, module):
        decls = []
        for d in module.declarations:
            if isinstance(d, n.Dim):
                for decl in d.decls:
                    decls.append(self.decl_tuple(decl, module, {}))
        opts = [o.lower() for o in module.options]
        kind = f", kind={module.kind!r}" if module.kind != "standard" else ""
        base = ", option_base=1" if "option base 1" in opts else ""
        return [
            f"{self.module_id(module)} = rt.Module({module.name!r}, decls={decls!r}, "
            f"option_compare_text={'option compare text' in opts}, "
            f"option_explicit={'option explicit' in opts}{kind}{base})"
        ]

    def decl_tuple(self, decl, module, consts, static=False):
        base = 1 if any(o.lower() == "option base 1" for o in module.options) else 0
        bounds = None
        if decl.bounds is not None:
            bounds = []
            for lo, hi in decl.bounds:
                lo_v = const_value(lo, consts, self, module) if lo is not None else base
                bounds.append((int(lo_v), int(const_value(hi, consts, self, module))))
        if decl.is_new:
            raise VBANotImplemented(f"'Dim {decl.name} As New {decl.type_name}'")
        t = (decl.name, bounds, decl.dynamic, self.type_name(decl.type_name, decl.name))
        return t + (True,) if static else t

    def module_consts(self, module):
        out = {}
        for d in module.declarations:
            if isinstance(d, n.Const):
                for name, expr, _ in d.defs:
                    out[name.lower()] = const_value(expr, out, self, module)
        return out


def _called_names(stmts):
    names = set()

    def expr(e):
        if isinstance(e, n.Name):
            names.add(e.name)
        elif isinstance(e, n.Call):
            expr(e.func)
            for a in e.args:
                if a.expr is not None:
                    expr(a.expr)
        elif isinstance(e, n.Member):
            if isinstance(e.obj, n.Name):
                names.add(f"{e.obj.name}.{e.name}")  # Module1.Helper: a procedure named with its module
            expr(e.obj)
        elif isinstance(e, n.BinOp):
            expr(e.left)
            expr(e.right)
        elif isinstance(e, (n.Unary, n.Paren)):
            expr(e.operand if isinstance(e, n.Unary) else e.expr)

    for s in walk_statements(stmts):
        for e in statement_exprs(s):
            expr(e)
    return names


def walk_statements(stmts):
    for s in stmts:
        yield s
        for body in child_blocks(s):
            yield from walk_statements(body)


def child_blocks(s):
    if isinstance(s, n.If):
        return [b for _, b in s.branches] + ([s.else_body] if s.else_body is not None else [])
    if isinstance(s, (n.For, n.ForEach, n.While, n.DoLoop, n.With)):
        return [s.body]
    if isinstance(s, n.Select):
        return [b for _, b in s.cases] + ([s.else_body] if s.else_body is not None else [])
    return []


def statement_exprs(s):
    if isinstance(s, n.Assign):
        return [s.target, s.value]
    if isinstance(s, n.CallStmt):
        return [s.target] + [a.expr for a in (s.args or []) if a.expr is not None]
    if isinstance(s, n.If):
        return [c for c, _ in s.branches]
    if isinstance(s, n.For):
        return [x for x in (s.var, s.start, s.end, s.step) if x is not None]
    if isinstance(s, n.ForEach):
        return [s.var, s.collection]
    if isinstance(s, n.While):
        return [s.cond]
    if isinstance(s, n.DoLoop):
        return [c[1] for c in (s.pre, s.post) if c]
    if isinstance(s, n.With):
        return [s.obj]
    if isinstance(s, n.Select):
        out = [s.subject]
        for conds, _ in s.cases:
            for c in conds:
                out += [x for x in (c.expr, c.expr2) if x is not None]
        return out
    return []


def const_value(expr, consts, conv, module):
    if isinstance(expr, n.Literal):
        return expr.value
    if isinstance(expr, n.Paren):
        return const_value(expr.expr, consts, conv, module)
    if isinstance(expr, n.Name):
        key = expr.name.lower()
        if key in consts:
            return consts[key]
        if key in VB_CONSTANTS:
            return VB_CONSTANTS[key]
        if key in XL_CONSTANTS:
            return XL_CONSTANTS[key]
        mc = conv.module_consts(module) if conv is not None else {}
        if key in mc:
            return mc[key]
    if isinstance(expr, n.Unary) and expr.op == "-":
        return op_neg(const_value(expr.operand, consts, conv, module))
    if isinstance(expr, n.BinOp):
        ops = {"+": op_add, "-": op_sub, "*": op_mul, "/": op_div, "\\": op_intdiv, "mod": op_mod, "&": op_concat}
        if expr.op in ops:
            return ops[expr.op](
                const_value(expr.left, consts, conv, module), const_value(expr.right, consts, conv, module)
            )
    raise VBANotImplemented(f"constant expression {expr}")


def doc_literal(text):
    """A docstring holding `text`, which comes partly from the workbook (button captions, its file name): a quote or
    backslash in it stays text, never ends the string or starts an escape."""
    if any(ord(ch) < 32 and ch not in "\n\t" for ch in text):
        return repr(text)
    body = text.replace("\\", "\\\\")
    if '"""' in body or body.endswith('"'):  # quotes that could end the docstring: escape them all
        body = body.replace('"', '\\"')
    return f'"""{body}"""'


def py_literal(value):
    if value is EMPTY:
        return "rt.EMPTY"
    if value is NULL:
        return "rt.NULL"
    if value is NOTHING:
        return "rt.NOTHING"
    if isinstance(value, dt.datetime):
        return f"rt.datetime({value.year}, {value.month}, {value.day}, {value.hour}, {value.minute}, {value.second})"
    return repr(value)


# ---------------------------------------------------------------------------
# One procedure
# ---------------------------------------------------------------------------


class ProcGen:
    def __init__(self, conv, module, src, proc):
        self.conv, self.module, self.src, self.proc = conv, module, src, proc
        self.lines = []
        self.counter = 0
        self.with_stack = []
        self.rn = False
        self.text_cmp = any(o.lower() == "option compare text" for o in module.options)
        self.text_arg = ", text=True" if self.text_cmp else ""  # for comparisons: the module's Option Compare
        self.explicit = any(o.lower() == "option explicit" for o in module.options)
        first = src.first_line
        self.source_lines = {first + i: line.strip() for i, line in enumerate(src.text.split("\n"))}
        self.locals, self.arrays, self.consts, self.decls = set(), set(), {}, []
        self.types = {}  # the declared type of each of the procedure's own names
        self.select_kinds = {}
        for p in proc.params:
            self.locals.add(p.name.lower())
            self.types[p.name.lower()] = conv.type_name(p.type_name, p.name)
            if p.is_array:
                self.arrays.add(p.name.lower())
        if proc.kind == "function":
            self.locals.add(proc.name.lower())
            self.types[proc.name.lower()] = conv.type_name(proc.return_type, proc.name)
        module_consts = conv.module_consts(module)
        for s in walk_statements(proc.body):
            if isinstance(s, n.Const):
                for name, expr, _ in s.defs:
                    self.consts[name.lower()] = const_value(expr, {**module_consts, **self.consts}, conv, module)
            elif isinstance(s, n.Dim):
                for decl in s.decls:
                    self.locals.add(decl.name.lower())
                    self.types[decl.name.lower()] = conv.type_name(decl.type_name, decl.name)
                    if decl.bounds is not None or decl.dynamic:
                        self.arrays.add(decl.name.lower())
                    self.decls.append(
                        conv.decl_tuple(
                            decl, module, {**module_consts, **self.consts}, static=s.is_static or proc.is_static
                        )
                    )
            elif isinstance(s, n.ReDim):
                for decl in s.decls:
                    self.arrays.add(decl.name.lower())

    # --- helpers ---
    def emit(self, ind, text):
        self.lines.append("    " * ind + text)

    def new(self, prefix):
        self.counter += 1
        return f"{prefix}{self.counter}"

    def vba_comment(self, ind, line):
        text = self.source_lines.get(line)
        if text and line != getattr(self, "last_commented", None):
            self.emit(ind, f"# L{line}: {text}")
            self.last_commented = line

    def var_ref(self, name):
        """How generated code names a VBA variable."""
        key = name.lower()
        scope = "v"
        if key not in self.locals and key not in self.conv_module_vars(self.module):
            for mod in self.conv.project.modules.values():
                if mod is not self.module and key in self.conv_module_vars(mod):
                    scope = self.conv.module_id(mod)
                    break
        return f"{scope}.{name}" if ident_ok(name) else f"{scope}[{name!r}]", scope

    def conv_module_vars(self, module):
        out = set()
        for d in module.declarations:
            if isinstance(d, n.Dim):
                for decl in d.decls:
                    if module is self.module or (d.scope or "").lower() in ("public", "global"):
                        out.add(decl.name.lower())
        return out

    def is_var(self, name):
        key = name.lower()
        if key in self.locals or key in self.conv_module_vars(self.module):
            return True
        return any(key in self.conv_module_vars(m) for m in self.conv.project.modules.values() if m is not self.module)

    def is_array(self, name):
        key = name.lower()
        if key in self.arrays:
            return True
        for d in self.module.declarations:
            if isinstance(d, n.Dim):
                for decl in d.decls:
                    if decl.name.lower() == key and (decl.bounds is not None or decl.dynamic):
                        return True
        return False

    def const(self, key):
        if key in self.consts:
            return True, self.consts[key]
        mc = self.conv.module_consts(self.module)
        if key in mc:
            return True, mc[key]
        for mod in self.conv.project.modules.values():
            mc = self.conv.module_consts(mod)
            if key in mc:
                return True, mc[key]
        return False, None

    # --- how an expression is declared: a numeric type "n", a String "s" or a Variant "v" ---
    def declared_type(self, name):
        key = name.lower()
        if key in self.types:
            return self.types[key]
        modules = [self.module] + [m for m in self.conv.project.modules.values() if m is not self.module]
        for mod in modules:
            for d in mod.declarations:
                if isinstance(d, n.Dim) and (mod is self.module or (d.scope or "").lower() in ("public", "global")):
                    for decl in d.decls:
                        if decl.name.lower() == key:
                            return self.conv.type_name(decl.type_name, decl.name)
        return "variant"

    def kind(self, e):
        if isinstance(e, n.Literal):
            return "s" if isinstance(e.value, str) else "n"
        if isinstance(e, n.Paren):
            return self.kind(e.expr)
        if isinstance(e, n.Unary):
            return "n" if self.kind(e.operand) != "v" else "v"
        if isinstance(e, n.BinOp):
            left, right = self.kind(e.left), self.kind(e.right)
            if "v" in (left, right):
                return "v"
            if e.op == "&" or (e.op == "+" and left == right == "s"):
                return "s"
            return "n"
        if isinstance(e, n.Name):
            key = e.name.lower()
            what = self.resolve(e.name)
            if what == "keyword":
                return "n" if key in ("true", "false") else "v"
            if what == "var":
                return "v" if self.is_array(e.name) else type_kind(self.declared_type(e.name))
            if what in ("const", "vbconst"):
                value = self.const(key)[1] if what == "const" else VB_CONSTANTS[key]
                return "s" if isinstance(value, str) else "n"
            if what == "xlconst":
                return "n"
            if what == "proc":
                return self.proc_kind(*self.conv.find(e.name, self.module))
            if what == "builtin":
                return "s" if e.suffix == "$" else FUNCTION_KINDS.get(key, "v")
            return "v"
        if isinstance(e, n.Member):
            module, src = self.qualified_proc(e)
            if src is not None:
                return self.proc_kind(module, src)
            if isinstance(e.obj, n.Name) and e.obj.name.lower() in self.conv.enums and not self.is_var(e.obj.name):
                return "n"
            return MEMBER_KINDS.get(e.name.lower(), "v")
        if isinstance(e, n.Call):
            func = e.func
            if isinstance(func, n.Name):
                if self.is_var(func.name):  # an element of an array: as the array is declared
                    return type_kind(self.declared_type(func.name)) if self.is_array(func.name) else "v"
                return self.kind(func)
            if isinstance(func, n.Member):
                module, src = self.qualified_proc(func)
                return self.proc_kind(module, src) if src is not None else "v"
        return "v"

    def proc_kind(self, module, src):
        """How a Function's result is declared, read from its first line: ... ) As String, or Function Name$( ..."""
        header = src.text.split("\n")[0].split("'")[0]
        if re.search(rf"(?i)\bFunction\s+{re.escape(src.name)}\$", header):
            return "s"
        m = re.search(r"(?i)\bAs\s+([\w.]+)\s*$", header.rpartition(")")[2] if ")" in header else header)
        return type_kind(self.conv.type_name(m.group(1), src.name)) if m else "v"

    def kinds(self, *exprs):
        """The k argument of a comparison: left out when every side is a Variant."""
        k = "".join(self.kind(e) for e in exprs)
        return f", k={k!r}" if k.strip("v") else ""

    # --- expressions ---
    def given_a_value(self, name):
        """Whether the module's code ever assigns to the name, or loops with it."""
        word = re.escape(name)
        assigned = rf"(?im)(?:^|:|\bthen\b|\belse\b)[ \t]*(?:set[ \t]+|let[ \t]+)?{word}[ \t]*(?:\([^=\n]*\))?[ \t]*="
        return re.search(rf"{assigned}|(?i:\bfor[ \t]+(?:each[ \t]+)?{word}\b)", self.module.source) is not None

    def resolve(self, name):
        """What a bare name means, in the order VBA looks it up."""
        key = name.lower()
        if key in ("true", "false", "empty", "null", "nothing"):
            return "keyword"
        if key == "me":
            return "me"
        if self.is_var(name):
            return "var"
        if self.const(key)[0]:
            return "const"
        if self.conv.find(name, self.module)[1] is not None:
            return "proc"
        if key in self.conv.object_modules:
            return "object"  # a sheet by its code name, or the workbook by its module's
        if key in BUILTINS:
            return "builtin"
        if key in VB_CONSTANTS:
            return "vbconst"
        if key in ("err", "debug"):
            return "errdebug"
        if key in XL_CONSTANTS:
            return "xlconst"
        if key in GLOBAL_MEMBERS:
            return "global"
        if key in NOT_MODELLED_GLOBALS:
            raise VBANotImplemented(f"{self.proc.name}: {NOT_MODELLED_GLOBALS[key]} is not modelled")
        if re.match(r"(?:vb|xl|mso)[a-z]", key) and not self.given_a_value(name):
            # a name the code only reads: a constant of VBA or Excel this does not know, not a variable left Empty
            raise VBANotImplemented(f"{self.proc.name}: the constant {name} is not modelled")
        if self.explicit:
            raise VBACompileError(f"{self.proc.name}: variable not defined: {name}")
        self.locals.add(key)  # an undeclared name is a new Variant variable
        return "var"

    def name_expr(self, name):
        key = name.lower()
        kind = self.resolve(name)
        if kind == "keyword":
            return {"true": "True", "false": "False"}.get(key, f"rt.{key.upper()}")
        if kind == "var":
            return self.var_ref(name)[0]
        if kind == "const":
            return py_literal(self.const(key)[1])
        if kind == "proc":
            module, src = self.conv.find(name, self.module)
            return f"{self.conv.py_name(module, src)}(xl)"
        if kind == "me":
            return self.object_ref(self.module)
        if kind == "object":
            return self.object_ref(self.conv.object_modules[key])
        if kind == "builtin":
            return f"rt.{CANONICAL.get(key, name)}()"
        if kind == "vbconst":
            return py_literal(VB_CONSTANTS[key])
        if kind == "errdebug":
            return f"xl.{key.capitalize()}"
        if kind == "xlconst":
            return f"rt.{name}"
        return self.global_ref(name)

    def object_ref(self, module):
        """What a sheet's or the workbook's module stands for: Sheet1 in Sheet1.Range("A1"), or Me inside it."""
        if module.kind == "document":
            return f"rt.codename(xl, {module.name!r})"
        if module.kind == "workbook":
            return "xl.ThisWorkbook"
        if module.kind == "class":
            raise VBANotImplemented(f"{self.proc.name}: Me in a class module")
        raise VBACompileError(f"{self.proc.name}: invalid use of Me keyword")

    def global_ref(self, name):
        """A name usable without a qualifier. In a sheet's module, Range, Cells, Rows, Columns and Evaluate (so [A1])
        are that sheet's."""
        if self.module.kind == "document" and name.lower() in SHEET_RELATIVE:
            return f"rt.scope(xl, {self.conv.module_id(self.module)}).{name}"
        return f"xl.{name}"

    def qualified_proc(self, e):
        """Module1.Helper or Sheet1.Refresh: a procedure named with its module, as (module, src), or (None, None)."""
        if isinstance(e, n.Member) and isinstance(e.obj, n.Name) and not self.is_var(e.obj.name):
            return self.conv.find(f"{e.obj.name}.{e.name}")
        return None, None

    def args(self, args, byref=False):
        """Argument list. For calls to VBA procedures a plain variable is passed by reference."""
        out = []
        for a in args:
            e = a.expr
            if e is None:
                code = "None"
            elif byref and isinstance(e, n.Name) and self.resolve(e.name) == "var":
                code = f"rt.ref({self.var_ref(e.name)[1]}, {e.name!r})"
            elif byref and isinstance(e, n.Call) and isinstance(e.func, n.Name) and self.is_var(e.func.name):
                code = f"rt.ref_item({self.var_ref(e.func.name)[0]}, {self.args(e.args)})"  # an array's element
            else:
                code = self.ex(e)
            if a.name and not ident_ok(a.name):  # From:=1: a name Python keeps for itself
                out.append(f"**{{{a.name!r}: {code}}}")
            else:
                out.append(f"{a.name}={code}" if a.name else code)
        return ", ".join(out)

    def library_function(self, e):
        """VBA.Left or Strings.Trim: the VBA function named with its library, as a plain name; else None."""
        if isinstance(e, n.Member) and isinstance(e.obj, n.Name) and e.obj.name.lower() in VBA_LIBRARIES:
            if not self.is_var(e.obj.name) and e.name.lower() in BUILTINS:
                return n.Name(e.name)
        return None

    def ex(self, e):
        if isinstance(e, n.Literal):
            return py_literal(e.value)
        if isinstance(e, n.Name):
            return self.name_expr(e.name)
        if isinstance(e, n.WithRef):
            if not self.with_stack:
                raise VBACompileError(f"{self.proc.name}: '.' used outside a With block")
            return self.with_stack[-1]
        if isinstance(e, n.Member):
            module, src = self.qualified_proc(e)
            if src is not None:
                return f"{self.conv.py_name(module, src)}(xl)"
            if self.library_function(e) is not None:
                return self.ex(self.library_function(e))
            if isinstance(e.obj, n.Name) and e.obj.name.lower() in self.conv.enums and not self.is_var(e.obj.name):
                found, value = self.const(e.name.lower())  # Colour.Red: a member of an Enum, named with it
                if found:
                    return py_literal(value)
            obj = self.ex(e.obj)
            return f"{obj}.{e.name}" if member_ok(e.name) else f"rt.member({obj}, {e.name!r})"
        if isinstance(e, n.Call):
            return self.call_expr(e)
        if isinstance(e, n.BinOp):
            fn = BINOPS[e.op]
            extra = self.text_arg if fn in COMPARISONS else ""
            if fn in COMPARISONS and fn != "like":
                extra += self.kinds(e.left, e.right)
            return f"rt.{fn}({self.ex(e.left)}, {self.ex(e.right)}{extra})"
        if isinstance(e, n.Unary):
            return f"rt.{'not_' if e.op == 'not' else 'neg'}({self.ex(e.operand)})"
        if isinstance(e, n.Paren):
            return f"rt.val({self.ex(e.expr)})"
        if isinstance(e, n.New):
            raise VBANotImplemented(f"'New {e.type_name}'")
        raise VBANotImplemented(f"expression {type(e).__name__}")

    def call_expr(self, e):
        func = e.func
        if isinstance(func, n.Name):
            key = func.name.lower()
            own_function = self.proc.kind == "function" and key == self.proc.name.lower()
            if self.is_var(func.name) and not own_function:
                if self.is_array(func.name):
                    return f"{self.var_ref(func.name)[0]}[{self.args(e.args)}]"
                return f"rt.index({self.var_ref(func.name)[0]}, {self.args(e.args)})"
            module, src = self.conv.find(func.name, self.module)
            if src is not None:
                a = self.args(e.args, byref=True)
                return f"{self.conv.py_name(module, src)}(xl{', ' + a if a else ''})"
            self.check_external(func.name)
            if key in BUILTINS:
                args = [a for a in (self.args(e.args), "_frame=v" if key in FRAME_AWARE else "") if a]
                return f"rt.{CANONICAL.get(key, func.name)}({', '.join(args)})"
            if key in GLOBAL_MEMBERS:
                return f"{self.global_ref(func.name)}({self.args(e.args)})"
            raise _undefined(self.proc, func.name)
        if isinstance(func, n.Member):
            module, src = self.qualified_proc(func)
            if src is not None:
                a = self.args(e.args, byref=True)
                return f"{self.conv.py_name(module, src)}(xl{', ' + a if a else ''})"
            if self.library_function(func) is not None:
                return self.call_expr(n.Call(self.library_function(func), e.args))
            obj, a = self.ex(func.obj), self.args(e.args)
            if member_ok(func.name):
                return f"{obj}.{func.name}({a})"
            return f"rt.member({obj}, {func.name!r}{', ' + a if a else ''})"
        # what a call gave, indexed: Split(s, ",")(0)
        return f"rt.index({self.ex(func)}, {self.args(e.args)})"

    def check_external(self, name):
        """A procedure of another library, brought in with Declare: calling it is not modelled."""
        if name.lower() in self.conv.externals:
            declared, library = self.conv.externals[name.lower()]
            raise VBANotImplemented(
                f"{self.proc.name}: {declared} is a procedure of {library} (a Declare statement); "
                "calls outside VBA and Excel are not modelled"
            )

    # --- simple statements ---
    def simple(self, s):
        """Python lines for a statement with no nested block."""
        if isinstance(s, n.Assign):
            return [self.assign(s)]
        if isinstance(s, n.CallStmt):
            return [self.call_statement(s)]
        if isinstance(s, n.ReDim):
            out = []
            for decl in s.decls:
                bounds = ", ".join(
                    f"({self.ex(lo) if lo is not None else 0}, {self.ex(hi)})" for lo, hi in decl.bounds or []
                )
                self.arrays.add(decl.name.lower())
                self.locals.add(decl.name.lower()) if not self.is_var(decl.name) else None
                scope = self.var_ref(decl.name)[1]
                out.append(
                    f"rt.redim({scope}, {decl.name!r}, [{bounds}], preserve={s.preserve}, type_name={decl.type_name!r})"
                )
            return out
        if isinstance(s, n.Erase):
            return [f"rt.erase({self.var_ref(name)[1]}, {name!r})" for name in s.names]
        if isinstance(s, n.End):
            return ["raise rt.VBAEnd()"]
        if isinstance(s, n.Unsupported):
            what = s.kind.replace("_", " ")
            message = f"line {s.line}: {what} {s.text[:60]!r}"
            return [f"rt.not_modelled({message!r})"]
        if isinstance(s, (n.Option, n.Dim, n.Const, n.Label)):
            return []
        raise VBANotImplemented(f"line {getattr(s, 'line', '?')}: statement {type(s).__name__}")

    def assign(self, s):
        t = s.target
        value = self.ex(s.value)
        if isinstance(t, n.Call) and isinstance(t.func, n.Name) and t.func.name.lower() == "mid" and not s.is_set:
            if not self.is_var("mid") and self.conv.find("mid", self.module)[1] is None and 2 <= len(t.args) <= 3:
                # The Mid statement: Mid(s, 2, 1) = "X" puts the text over part of the variable's.
                text, start, *length = [self.ex(a.expr) for a in t.args]
                changed = f"rt.mid_statement({text}, {start}, {length[0] if length else 'None'}, {value})"
                return self.assign_to(t.args[0].expr, changed, False, s.line)
        return self.assign_to(t, value, s.is_set, s.line)

    def assign_to(self, t, value, is_set, line):
        """The Python line that gives the target `t` the value of the Python expression `value`."""
        s = n.Assign(t, None, is_set, line=line)
        if s.is_set:
            if isinstance(t, n.Name):
                if not self.is_var(t.name):
                    self.locals.add(t.name.lower())
                return f"rt.set_({self.var_ref(t.name)[1]}, {t.name!r}, {value})"
            raise VBANotImplemented(f"line {s.line}: Set on this kind of target")
        if isinstance(t, n.Name):
            key = t.name.lower()
            if self.const(key)[0] and not self.is_var(t.name):
                raise VBACompileError(f"line {s.line}: assignment to constant {t.name}")
            if not self.is_var(t.name) and key in GLOBAL_MEMBERS:
                return f"{self.global_ref(t.name)} = {value}"
            if not self.is_var(t.name):
                if self.explicit:
                    raise VBACompileError(f"{self.proc.name}: variable not defined: {t.name}")
                self.locals.add(key)
            return f"{self.var_ref(t.name)[0]} = {value}"
        if isinstance(t, n.Call):
            func = t.func
            if isinstance(func, n.Name) and self.is_var(func.name):
                if self.is_array(func.name):
                    return f"{self.var_ref(func.name)[0]}[{self.args(t.args)}] = {value}"
                return f"rt.let_index({self.var_ref(func.name)[0]}, ({self.args(t.args)},), {value})"
            return f"rt.let({self.ex(t)}, {value})"
        if isinstance(t, n.Member):
            obj = self.ex(t.obj)
            if member_ok(t.name):
                return f"{obj}.{t.name} = {value}"
            return f"rt._let_member({obj}, {t.name!r}, {value})"
        raise VBANotImplemented(f"line {s.line}: assignment target {type(t).__name__}")

    def call_statement(self, s):
        t = s.target
        if s.args is None:
            if isinstance(t, n.Name):
                key = t.name.lower()
                module, src = self.conv.find(t.name, self.module)
                if src is not None:
                    return f"{self.conv.py_name(module, src)}(xl)"
                self.check_external(t.name)
                if key in BUILTINS:
                    return f"rt.{CANONICAL.get(key, t.name)}()"
                if key in GLOBAL_MEMBERS:
                    return f"{self.global_ref(t.name)}()"
                raise _undefined(self.proc, t.name)
            if isinstance(t, n.Member):
                module, src = self.qualified_proc(t)
                if src is not None:
                    return f"{self.conv.py_name(module, src)}(xl)"
                return f"{self.ex(t)}()"
            return self.ex(t)
        return self.ex(n.Call(t, s.args))

    # --- driving ---
    def generate(self):
        p = self.proc
        params = [
            (
                x.name,
                x.byval,
                x.optional,
                (const_value(x.default, self.consts, self.conv, self.module) if x.default is not None else None),
                self.types[x.name.lower()],
                x.paramarray,
            )
            for x in p.params
        ]
        problems, breaks = analyse(p) if not self.conv.force_flat else (["flat form requested"], set())
        self.breaks = breaks
        py = self.conv.py_name(self.module, self.src)
        button = [label for label, macro in self.conv.entry_points if label and macro.lower() == p.name.lower()]
        where = f"VBA: {self.module.name}.{p.name} (lines {self.src.first_line}-{self.src.last_line})"
        doc = (f'Button "{button[0]}". ' if button else "") + where
        body_lines = []
        saved, self.lines = self.lines, body_lines
        if problems:
            self.flat(problems)
        else:
            self.block(p.body, 1, top=True)
            if p.kind == "function":
                self.emit(1, f"return rt.result(v, {p.name!r})")
        self.lines = saved
        header = [f"def {py}(xl, *args, **kwargs):", f"    {doc_literal(doc)}"]
        frame = (
            f"    v = rt.Frame(xl, {self.conv.module_id(self.module)}, {p.name!r}, decls={self.decls!r}, "
            f"params={params!r}, args=args, kwargs=kwargs"
            + (f", returns=({p.name!r}, {self.types[p.name.lower()]!r})" if p.kind == "function" else "")
            + ")"
        )
        return header + [frame] + body_lines

    def block(self, stmts, ind, top=False):
        start = len(self.lines)
        for s in stmts:
            self.statement(s, ind, top)
        if not any(line.strip() and not line.strip().startswith("#") for line in self.lines[start:]):
            self.emit(ind, "pass")

    def wrap_rn(self, ind, line, code_lines):
        if not code_lines:
            return
        if not self.rn:
            for c in code_lines:
                self.emit(ind, c)
            return
        self.emit(ind, "try:")
        for c in code_lines:
            self.emit(ind + 1, c)
        self.emit(ind, "except rt.CAUGHT as err:")
        self.emit(ind + 1, f"rt.skipped(err, {line})")

    def cond(self, expr, line, on_error=True):
        if self.rn:
            return f"rt.cond_rn(lambda: {self.ex(expr)}, {on_error}, {line})"
        return f"rt.cond({self.ex(expr)})"

    def statement(self, s, ind, top):
        line = getattr(s, "line", 0)
        if isinstance(s, n.Label):
            self.emit(ind, f"# {s.name}:")
            return
        if isinstance(s, (n.Dim, n.Const, n.Option)):
            return
        self.vba_comment(ind, line)
        if isinstance(s, n.OnError):
            self.rn = s.mode == "resume_next"
            self.emit(ind, "rt.on_error()")
            return
        if isinstance(s, n.GoTo):
            self.emit(ind, "break")
            return
        if isinstance(s, n.Exit):
            if s.kind in ("for", "do"):
                self.emit(ind, "break")
            elif self.proc.kind == "function":
                self.emit(ind, f"return rt.result(v, {self.proc.name!r})")
            else:
                self.emit(ind, "return")
            return
        if isinstance(s, n.If):
            for i, (c, body) in enumerate(s.branches):
                self.emit(ind, f"{'if' if i == 0 else 'elif'} {self.cond(c, line)}:")
                self.block(body, ind + 1)
            if s.else_body is not None:
                self.emit(ind, "else:")
                self.block(s.else_body, ind + 1)
            return
        if isinstance(s, n.For):
            if not isinstance(s.var, n.Name):
                raise VBANotImplemented(f"line {line}: For with a counter that is not a plain variable")
            k = self.new("")
            if not self.is_var(s.var.name):
                self.locals.add(s.var.name.lower())
            scope = self.var_ref(s.var.name)[1]
            counter = self.var_ref(s.var.name)[0]
            step = self.ex(s.step) if s.step is not None else "1"
            init = f"rt.for_init({scope}, {s.var.name!r}, {self.ex(s.start)}, {self.ex(s.end)}, {step})"
            test = f"rt.for_test({counter}, end{k}, step{k})"
            if self.rn:
                init, test = f"rt.header(lambda: {init}, {line})", f"rt.header(lambda: {test}, {line})"
            self.emit(ind, f"end{k}, step{k} = {init}")
            self.emit(ind, f"while {test}:")
            self.block(s.body, ind + 1)
            nxt = f"{counter} = rt.add({counter}, step{k})"
            if self.rn:
                nxt = f"rt.header(lambda: rt.step({scope}, {s.var.name!r}, step{k}), {line})"
            self.emit(ind + 1, nxt)
            return
        if isinstance(s, n.ForEach):
            if not isinstance(s.var, n.Name):
                raise VBANotImplemented(f"line {line}: For Each with a target that is not a plain variable")
            k = self.new("item")
            if not self.is_var(s.var.name):
                self.locals.add(s.var.name.lower())
            coll = f"rt.each({self.ex(s.collection)})"
            if self.rn:
                coll = f"rt.header(lambda: {coll}, {line})"
            self.emit(ind, f"for {k} in {coll}:")
            self.emit(ind + 1, f"rt.assign_each({self.var_ref(s.var.name)[1]}, {s.var.name!r}, {k})")
            self.block(s.body, ind + 1)
            return
        if isinstance(s, n.While):
            self.emit(ind, f"while {self.cond(s.cond, line)}:")
            self.block(s.body, ind + 1)
            return
        if isinstance(s, n.DoLoop):
            if s.pre:
                kind, c = s.pre
                test = self.cond(c, line, True) if kind == "while" else f"not {self.cond(c, line, False)}"
                self.emit(ind, f"while {test}:")
            else:
                self.emit(ind, "while True:")
                if not s.post:
                    self.emit(ind + 1, "rt.tick()")  # no condition to count its turns: see runtime.tick
            self.block(s.body, ind + 1)
            if s.post:
                kind, c = s.post
                exit_test = f"not {self.cond(c, line, False)}" if kind == "while" else self.cond(c, line, True)
                self.emit(ind + 1, f"if {exit_test}:")
                self.emit(ind + 2, "break")
            return
        if isinstance(s, n.With):
            name = self.new("w")
            value = self.ex(s.obj)
            if self.rn:
                value = f"rt.try_value(lambda: {value}, rt.NOTHING, {line})"
            self.emit(ind, f"{name} = {value}")
            self.with_stack.append(name)
            self.block(s.body, ind)
            self.with_stack.pop()
            return
        if isinstance(s, n.Select):
            name = self.new("case")
            subject = f"rt.val({self.ex(s.subject)})"
            if self.rn:
                subject = f"rt.header(lambda: {subject}, {line})"
            self.emit(ind, f"{name} = {subject}")
            first = True
            for conds, body in s.cases:
                test = " or ".join(self.case_test(name, s.subject, c) for c in conds)
                if self.rn:
                    test = f"rt.cond_rn(lambda: {test}, True, {line})"
                self.emit(ind, f"{'if' if first else 'elif'} {test}:")
                first = False
                self.block(body, ind + 1)
            if s.else_body is not None:
                if first:
                    self.block(s.else_body, ind)
                else:
                    self.emit(ind, "else:")
                    self.block(s.else_body, ind + 1)
            return
        if isinstance(s, (n.End, n.Unsupported)):
            for c in self.simple(s):
                self.emit(ind, c)
            return
        self.wrap_rn(ind, line, self.simple(s))

    def case_test(self, value, subject, c):
        """The test of one Case condition against the Select Case value held in the Python name `value`."""
        if c.kind == "eq":
            return f"rt.case_eq({value}, {self.ex(c.expr)}{self.text_arg}{self.kinds(subject, c.expr)})"
        if c.kind == "range":
            k = self.kinds(subject, c.expr, c.expr2)
            return f"rt.case_range({value}, {self.ex(c.expr)}, {self.ex(c.expr2)}{self.text_arg}{k})"
        compare = BINOPS[c.op]  # Case Is > 5
        return f"rt.cond(rt.{compare}({value}, {self.ex(c.expr)}{self.text_arg}{self.kinds(subject, c.expr)}))"

    # --- step-by-step form (any GoTo / On Error GoTo) ---
    def flat(self, problems):
        compiled = Compiler(self.proc).compile()
        code = compiled.code
        self.emit(1, f"# Written as numbered steps because of: {'; '.join(sorted(set(problems)))}")
        self.emit(1, "w = []")
        self.emit(1, "st = {}")
        self.emit(1, f"DEPTH = {tuple(ins.depth for ins in code)!r}")
        self.emit(1, f"KIND = {tuple(ins.op for ins in code)!r}")
        labels = {name: pc for name, pc in compiled.labels.items()}
        self.emit(1, f"LABELS = {labels!r}")
        self.emit(1, "pc = 0")
        self.emit(1, f"while pc < {len(code)}:")
        self.emit(2, "try:")
        for i, ins in enumerate(code):
            self.emit(3, f"{'if' if i == 0 else 'elif'} pc == {i}:")
            if ins.line:
                self.vba_comment(4, ins.line)
            for c in self.flat_instruction(i, ins, labels):
                self.emit(4, c)
        if not code:
            self.emit(3, "pass")
        self.emit(2, "except rt.CAUGHT as err:")
        self.emit(3, "pc = rt.handle_error(v, err, pc, w, DEPTH, KIND[pc])")
        if self.proc.kind == "function":
            self.emit(1, f"return rt.result(v, {self.proc.name!r})")

    def flat_instruction(self, i, ins, labels):
        op = ins.op
        self.with_stack = ["w[-1]"] if ins.depth > 0 or op == "with_pop" else []
        nxt = f"pc = {i + 1}"
        ret = f"return rt.result(v, {self.proc.name!r})" if self.proc.kind == "function" else "return"
        if op == "stmt":
            s = ins.a
            if isinstance(s, n.OnError):
                mode = {"resume_next": '"resume_next"', "goto": '"goto"', "goto0": "None"}[s.mode]
                target = labels.get(s.label) if s.label else None
                return [f"rt.set_on_error(v, {mode}, {target!r})", nxt]
            if isinstance(s, n.Resume):
                if s.target is None:
                    return ['pc = rt.resume(v, "resume", None, w, DEPTH)']
                if s.target == "next":
                    return ['pc = rt.resume(v, "next", None, w, DEPTH)']
                return [f'pc = rt.resume(v, "label", {labels[s.target]}, w, DEPTH)']
            return self.simple(s) + [nxt]
        if op in ("jif", "jit"):
            neg = "" if op == "jif" else "not "
            return [f"pc = {i + 1} if {neg}rt.cond({self.ex(ins.a)}) else rt.jump(w, DEPTH, {ins.b})"]
        if op == "jmp":
            return [f"pc = rt.jump(w, DEPTH, {ins.a})"]
        if op == "goto":
            return [f"pc = rt.jump(w, DEPTH, {ins.b})"]
        if op == "return":
            return [ret]
        if op == "for_init":
            s = ins.b
            if not isinstance(s.var, n.Name):
                raise VBANotImplemented(f"line {s.line}: For with a counter that is not a plain variable")
            if not self.is_var(s.var.name):
                self.locals.add(s.var.name.lower())
            step = self.ex(s.step) if s.step is not None else "1"
            scope = self.var_ref(s.var.name)[1]
            return [
                f"st[{ins.a}] = rt.for_init({scope}, {s.var.name!r}, {self.ex(s.start)}, {self.ex(s.end)}, {step})",
                nxt,
            ]
        if op == "for_test":
            scope = self.var_ref(ins.b.var.name)[1]
            return [
                f"pc = {i + 1} if rt.for_ok({scope}, {ins.b.var.name!r}, st, {ins.a}) else rt.jump(w, DEPTH, {ins.c})"
            ]
        if op == "for_next":
            scope = self.var_ref(ins.b.var.name)[1]
            return [f"rt.for_next({scope}, {ins.b.var.name!r}, st, {ins.a})", f"pc = {ins.c}"]
        if op == "each_init":
            if not self.is_var(ins.b.var.name):
                self.locals.add(ins.b.var.name.lower())
            return [f"st[{ins.a}] = [rt.each({self.ex(ins.b.collection)}), 0]", nxt]
        if op == "each_next":
            scope = self.var_ref(ins.b.var.name)[1]
            return [
                f"pc = {i + 1} if rt.each_next({scope}, {ins.b.var.name!r}, st, {ins.a}) else rt.jump(w, DEPTH, {ins.c})"
            ]
        if op == "with_push":
            self.with_stack = ["w[-1]"] if ins.depth > 0 else []
            return [
                "try:",
                f"    w.append({self.ex(ins.a)})",
                "except rt.CAUGHT:",
                "    w.append(rt.NOTHING)",
                "    raise",
                nxt,
            ]
        if op == "with_pop":
            return ["w.pop()", nxt]
        if op == "select":
            self.select_kinds[ins.a] = ins.b
            return [f"st[{ins.a}] = rt.val({self.ex(ins.b)})", nxt]
        if op == "case":
            tests = [self.case_test(f"st[{ins.a}]", self.select_kinds[ins.a], c) for c in ins.b]
            return [f"pc = {i + 1} if ({' or '.join(tests)}) else rt.jump(w, DEPTH, {ins.c})"]
        raise VBANotImplemented(f"instruction {op}")


def analyse(proc):
    """Can this procedure be written as structured Python? Returns (problems, gotos that are breaks)."""
    problems, breaks = [], set()

    def walk(stmts, loops, top):
        for i, s in enumerate(stmts):
            if isinstance(s, n.OnError):
                if not top:
                    problems.append("On Error inside a block")
                if s.mode == "goto":
                    problems.append("On Error GoTo")
            elif isinstance(s, n.Resume):
                problems.append("Resume")
            elif isinstance(s, n.GoTo):
                if loops and s.label in loops[-1][1]:
                    breaks.add(id(s))
                else:
                    problems.append(f"GoTo {s.label}")
            elif isinstance(s, n.Exit) and s.kind in ("for", "do"):
                want = (n.For, n.ForEach) if s.kind == "for" else (n.DoLoop,)
                if not loops or not isinstance(loops[-1][0], want):
                    problems.append(f"Exit {s.kind} across another loop")
            elif isinstance(s, (n.For, n.ForEach, n.While, n.DoLoop)):
                after = set()
                for nxt in stmts[i + 1 :]:
                    if isinstance(nxt, n.Label):
                        after.add(nxt.name)
                    else:
                        break
                walk(s.body, loops + [(s, after)], False)
            else:
                for body in child_blocks(s):
                    walk(body, loops, False)

    walk(proc.body, [], True)
    return problems, breaks


def convert(project, workbook_name, entry_points, out_path=None, runtime_dir=None, force_flat=False):
    return Conversion(project, workbook_name, entry_points, force_flat).source(out_path, runtime_dir)
