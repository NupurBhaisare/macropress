"""Parse VBA with vb2py's grammar and build the converter's syntax tree.

vb2py's own Python output is not used. Only its grammar (SimpleParse) is: it gives a
complete parse tree, which is converted here into the nodes in nodes.py.
"""

import bisect
import re

from simpleparse.parser import Parser
from vb2py import utils as vb_utils, vbparser

from . import nodes as n
from .errors import VBANotImplemented, VBAParseError
from .values import hex_value, parse_date_text

_PARSER = None


def _parser():
    global _PARSER
    if _PARSER is None:
        grammar = vb_utils.loadGrammarFrom(vbparser.GRAMMAR_FILE, data={"dialect": None})
        _PARSER = Parser(grammar, "line")
    return _PARSER


class CST:
    """A node of the SimpleParse tree, with absolute offsets."""

    __slots__ = ("name", "start", "end", "kids", "src")

    def __init__(self, tup, src):
        self.name, self.start, self.end, kids = tup
        self.src = src
        self.kids = [CST(k, src) for k in (kids or [])]

    @property
    def text(self):
        return self.src[self.start : self.end]

    def child(self, name):
        for k in self.kids:
            if k.name == name:
                return k
        return None

    def children(self, name):
        return [k for k in self.kids if k.name == name]

    def find(self, name):
        if self.name == name:
            return self
        for k in self.kids:
            hit = k.find(name)
            if hit is not None:
                return hit
        return None


# ---------------------------------------------------------------------------
# Splitting a module into procedures
# ---------------------------------------------------------------------------

PROC_START = re.compile(
    r"^\s*(?:(Public|Private|Friend|Global)\s+)?(?:(Static)\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)",
    re.I,
)
PROC_END = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.I)


# What vb2py's grammar does not read is rewritten first, line for line, so every line keeps its number.

# Conditional compilation is decided as 64-bit Excel for Windows decides it.
COMPILER_CONSTANTS = {"vba7": -1, "vba6": -1, "win64": -1, "win32": -1, "win16": 0, "mac": 0}
_CC_TOKEN = re.compile(r"\s*(?:(\d+)|([A-Za-z_]\w*)|(<>|<=|>=|[=<>()+\-*]))")
_CC_BINARY = {"or": 1, "xor": 1, "and": 2, "=": 4, "<>": 4, "<": 4, ">": 4, "<=": 4, ">=": 4, "+": 5, "-": 5, "*": 6}


def _cc_value(text, consts):
    """A conditional-compilation expression's value, with VBA's True (-1) and False (0); a name not defined is 0."""
    tokens, pos = [], 0
    while text[pos:].strip():
        m = _CC_TOKEN.match(text, pos)
        if m is None:
            raise VBAParseError(f"cannot read the conditional-compilation expression {text.strip()!r}")
        number, name, op = m.groups()
        tokens.append(int(number) if number is not None else (name or op).lower())
        pos = m.end()
    at = 0

    def prefix():
        nonlocal at
        if at >= len(tokens):
            raise VBAParseError(f"cannot read the conditional-compilation expression {text.strip()!r}")
        tok = tokens[at]
        at += 1
        if isinstance(tok, int):
            return tok
        if tok == "(":
            value = expr(0)
            at += 1  # the closing parenthesis
            return value
        if tok == "not":
            return ~expr(3)
        if tok == "-":
            return -expr(7)
        if tok in ("true", "false"):
            return -1 if tok == "true" else 0
        return consts.get(tok, COMPILER_CONSTANTS.get(tok, 0))

    def expr(min_power):
        nonlocal at
        left = prefix()
        while at < len(tokens) and _CC_BINARY.get(tokens[at], -1) >= min_power:
            op = tokens[at]
            at += 1
            right = expr(_CC_BINARY[op] + 1)
            left = {
                "or": lambda a, b: a | b, "xor": lambda a, b: a ^ b, "and": lambda a, b: a & b,
                "=": lambda a, b: -(a == b), "<>": lambda a, b: -(a != b), "<": lambda a, b: -(a < b),
                ">": lambda a, b: -(a > b), "<=": lambda a, b: -(a <= b), ">=": lambda a, b: -(a >= b),
                "+": lambda a, b: a + b, "-": lambda a, b: a - b, "*": lambda a, b: a * b,
            }[op](left, right)  # fmt: skip
        return left

    return expr(0)


def conditional_compilation(lines, where):
    """Keep the lines of the #If branches that are compiled, and blank the others and the # lines themselves.
    Returns the lines and whether there were any."""
    out, stack, consts, active, used = [], [], {}, True, False
    for number, line in enumerate(lines, 1):
        text = line.strip()
        if not text.startswith("#"):
            out.append(line if active else "")
            continue
        used = True
        out.append("")
        body = text.split("'")[0].strip()
        const = re.match(r"#\s*const\s+(\w+)\s*=\s*(.+)$", body, re.I)
        branch = re.match(r"#\s*(if|elseif)\b(.*?)(?:\bthen)?\s*$", body, re.I)
        if const:
            if active:
                consts[const.group(1).lower()] = _cc_value(const.group(2), consts)
        elif branch and branch.group(1).lower() == "if":
            taken = active and _cc_value(branch.group(2), consts) != 0
            stack.append([active, taken])
            active = taken
        elif not stack:
            raise VBAParseError(f"{where}: line {number}: {text!r} without #If")
        elif branch:
            parent, taken = stack[-1]
            active = parent and not taken and _cc_value(branch.group(2), consts) != 0
            stack[-1][1] = taken or active
        elif re.match(r"#\s*else\b", body, re.I):
            parent, taken = stack[-1]
            active = parent and not taken
            stack[-1][1] = True
        elif re.match(r"#\s*end\s+if\b", body, re.I):
            active = stack.pop()[0]
        else:
            raise VBAParseError(f"{where}: line {number}: cannot read {text!r}")
    if stack:
        raise VBAParseError(f"{where}: #If without #End If")
    return out, used


_DECLARE = re.compile(
    r'^\s*(?:(?:Public|Private)\s+)?Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+(\w+)\s+Lib\s+"([^"]*)"', re.I
)
_ENUM = re.compile(r"^\s*(?:(Public|Private)\s+)?Enum\s+(\w+)\s*(?:'.*)?$", re.I)
_ENUM_MEMBER = re.compile(r"^\s*\[?(\w+)\]?\s*(?:=\s*([^']+?))?\s*(?:'.*)?$")
_TYPE = re.compile(r"^\s*(?:(?:Public|Private)\s+)?Type\s+(\w+)\s*(?:'.*)?$", re.I)


def module_declarations(module, lines):
    """Read what the grammar leaves out of a module's declarations, and return their lines without it:
    - Declare: the name of each procedure of another library (calling one is not modelled), its lines blanked
    - Enum: each member becomes a constant, on its own line
    - Type: its name, so a variable of that type can be refused by name"""
    out = [text for _, text in lines]
    i = 0
    while i < len(out):
        text = out[i]
        end = i
        while out[end].rstrip().endswith(" _") and end + 1 < len(out):
            end += 1  # a statement carried over several lines
        logical = " ".join(t.rstrip().removesuffix("_") for t in out[i : end + 1])
        declare, enum, user_type = _DECLARE.match(logical), _ENUM.match(text), _TYPE.match(text)
        if declare:
            module.externals[declare.group(1).lower()] = (declare.group(1), declare.group(2))
            out[i : end + 1] = [""] * (end + 1 - i)
        elif enum:
            scope, members, previous = enum.group(1) or "Public", [], None
            out[i] = ""
            i += 1
            while i < len(out) and not re.match(r"^\s*End\s+Enum\b", out[i], re.I):
                member = _ENUM_MEMBER.match(out[i])
                if out[i].strip() and not out[i].lstrip().startswith(("'", "Rem ")) and member:
                    name, value = member.group(1), member.group(2)
                    value = value or (f"{previous} + 1" if previous else "0")
                    out[i] = f"{scope} Const {name} = {value}"
                    members.append(name)
                    previous = name
                else:
                    out[i] = ""
                i += 1
            if i < len(out):
                out[i] = ""
            module.enums[enum.group(2).lower()] = members
            end = i
        elif user_type:
            module.types[user_type.group(1).lower()] = user_type.group(1)
        i = end + 1
    return out


def split_module(name, code, kind="standard"):
    """Cut a module into its declarations section and one text per procedure."""
    lines = code.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    module = n.Module(name=name, kind=kind, source=code)
    lines, module.conditional = conditional_compilation(lines, name)
    decl_lines = []
    i = 0
    first_proc_seen = False
    while i < len(lines):
        m = PROC_START.match(lines[i])
        if m and not lines[i].lstrip().startswith("'"):
            first_proc_seen = True
            start = i
            j = i
            while j < len(lines) and not PROC_END.match(lines[j]):
                j += 1
            if j >= len(lines):
                raise VBAParseError(f"{name}: procedure {m.group(4)} at line {i + 1} has no End statement")
            kind_word = m.group(3).split()
            proc_kind = " ".join(w.lower() for w in kind_word)
            module.procedures[m.group(4).lower()] = n.ProcSource(
                name=m.group(4),
                kind=proc_kind,
                scope=(m.group(1) or "Public").capitalize(),
                first_line=start + 1,
                last_line=j + 1,
                text="\n".join(lines[start : j + 1]),
            )
            i = j + 1
            continue
        if not first_proc_seen:
            decl_lines.append((i + 1, lines[i]))
        i += 1
    if decl_lines:
        first = decl_lines[0][0]
        text = "\n".join(module_declarations(module, decl_lines))
        try:
            stmts = Builder(text, first).parse_lines()
        except VBAParseError as err:  # told when one of the module's procedures is used: the others still run
            module.decl_error = f"{name}: {err}"
            stmts = []
        for s in stmts:
            if isinstance(s, n.Option):
                module.options.append(s.text)
            elif isinstance(s, (n.Dim, n.Const)):
                module.declarations.append(s)
            elif isinstance(s, n.Unsupported) and s.kind in (
                "external_declaration",
                "user_type_definition",
                "enumeration_definition",
                "event_definition",
            ):
                module.declarations.append(s)
    return module


def parse_procedure(src: n.ProcSource) -> n.Procedure:
    if src.kind.startswith("property"):
        raise VBANotImplemented(f"{src.name}: Property procedures are not modelled")
    b = Builder(src.text, src.first_line)
    stmts = b.parse_lines()
    procs = [s for s in stmts if isinstance(s, n.Procedure)]
    if len(procs) != 1:
        raise VBAParseError(f"{src.name}: expected one procedure, parsed {len(procs)}")
    return procs[0]


# ---------------------------------------------------------------------------
# Builder: parse tree -> nodes
# ---------------------------------------------------------------------------

BIN_PREC = {
    "imp": 1,
    "eqv": 2,
    "xor": 3,
    "or": 4,
    "and": 5,
    "=": 7,
    "<>": 7,
    "<": 7,
    ">": 7,
    "<=": 7,
    ">=": 7,
    "like": 7,
    "is": 7,
    "&": 8,
    "+": 9,
    "-": 9,
    "mod": 10,
    "\\": 11,
    "*": 12,
    "/": 12,
    "^": 14,
}
UN_PREC = {"not": 6, "-": 13, "+": 13}


# A date literal is handed to the grammar as a call: on its own, the grammar stops at the colon of #12:30:00 PM#.
DATE_MARK = "MacroPressDateLiteral"
_DATE_BODY = re.compile(
    r"^\s*(?:\d{1,4}[/-]\d{1,2}[/-]\d{1,4})?\s*(?:\d{1,2}:\d{1,2}(?::\d{1,2})?\s*(?:[AaPp][Mm])?)?\s*$"
)


def mark_date_literals(line):
    if "#" not in line or line.lstrip().lower().startswith("rem "):
        return line
    out, i = [], 0
    while i < len(line):
        ch = line[i]
        if ch == '"':  # a string: to its closing quote ("" is a quote inside it)
            j = i + 1
            while j < len(line) and not (line[j] == '"' and line[j + 1 : j + 2] != '"'):
                j += 2 if line[j] == '"' else 1
            out.append(line[i : j + 1])
            i = j + 1
        elif ch == "'":  # a comment: the rest of the line
            out.append(line[i:])
            break
        else:
            j = line.find("#", i + 1) if ch == "#" else -1
            if j > 0 and line[i + 1 : j].strip() and _DATE_BODY.match(line[i + 1 : j]):
                out.append(f'{DATE_MARK}("{line[i + 1 : j].strip()}")')
                i = j + 1
            else:
                out.append(ch)
                i += 1
    return "".join(out)


class Builder:
    def __init__(self, text, first_line=1):
        self.first_line = first_line
        prepared = []
        for line in text.split("\n"):
            stripped = line.strip()
            if stripped.startswith("Attribute ") and "=" in stripped:
                stripped = ""  # module/procedure attributes (VB_Name, VB_ProcData, ...)
            prepared.append(mark_date_literals(stripped))
        self.src = vbparser.makeSafeFromUnicode("\n".join(prepared) + "\n\n")
        self.newlines = [i for i, ch in enumerate(self.src) if ch == "\n"]

    # --- driving the parser ---
    def parse_lines(self):
        parser = _parser()
        pos = 0
        out = []
        while pos < len(self.src):
            ok, tree, nxt = parser.parse(self.src, start=pos)
            if not ok or nxt <= pos:
                rest = self.src[pos:]
                if rest.strip():
                    raise VBAParseError(f"line {self.line_of(pos)}: cannot parse {rest.strip().splitlines()[0]!r}")
                break
            for t in tree or []:
                node = CST(t, self.src)
                out.extend(self.top_level(node))
            pos = nxt
        return out

    def line_of(self, offset):
        return bisect.bisect_left(self.newlines, offset) + self.first_line

    def text(self, node):
        return vbparser.makeUnicodeFromSafe(node.text)

    def top_level(self, node):
        if node.name in ("line_body", "line"):
            return self.line(node) if node.name == "line" else self.line_body(node)
        if node.name == "label_statement":
            return [n.Label(self.label_name(node), line=self.line_of(node.start))]
        if node.name in ("line_end", "NEWLINE"):
            return []
        return self.statement(node)

    # --- lines and blocks ---
    def lines_in(self, node):
        for k in node.kids:
            if k.name == "line":
                yield k
            elif k.name in (
                "block",
                "block_content",
                "sub_block",
                "fn_block",
                "if_block",
                "else_block",
                "else_if_block",
                "while_block",
                "do_block",
                "for_each_body",
                "case_body",
                "case_comment_block",
                "property_block",
            ):
                yield from self.lines_in(k)

    def block(self, node):
        if node is None:
            return []
        out = []
        for line in self.lines_in(node):
            out.extend(self.line(line))
        return out

    def line(self, node):
        out = []
        for k in node.kids:
            if k.name == "label_statement":
                out.append(n.Label(self.label_name(k), line=self.line_of(k.start)))
            elif k.name == "line_body":
                out.extend(self.line_body(k))
        return out

    def line_body(self, node):
        out = []
        for k in node.kids:
            if k.name in ("line_end", "colon", "NEWLINE"):
                continue
            out.extend(self.statement(k))
        return out

    def label_name(self, node):
        lab = node.find("label")
        return (lab.text if lab is not None else node.text.rstrip(":")).strip().lower()

    # --- statements ---
    def statement(self, node):
        name = node.name
        line = self.line_of(node.start)
        if name in ("single_statement", "valid_statement", "compound_statement"):
            out = []
            for k in node.kids:
                out.extend(self.statement(k))
            return out
        handler = getattr(self, "st_" + name, None)
        if handler is not None:
            label = node.child("label_definition")
            result = handler(node, line)
            stmts = result if isinstance(result, list) else [result]
            if label is not None:
                stmts = [n.Label(self.label_name(label), line=line)] + stmts
            return stmts
        if name in ("comment_statement", "line_end", "NEWLINE", "colon"):
            return []
        if name == "label_statement":
            return [n.Label(self.label_name(node), line=line)]
        return [n.Unsupported(name, self.text(node).strip(), line=line)]

    def st_sub_definition(self, node, line):
        return self._procedure(node, "sub", "sub_start_definition", "sub_block", line)

    def st_fn_definition(self, node, line):
        return self._procedure(node, "function", "fn_start_definition", "fn_block", line)

    def _procedure(self, node, kind, start_name, block_name, line):
        start = node.child(start_name)
        ident = start.child("identifier")
        scope = start.child("scope")
        params = []
        fpl = start.child("formal_param_list")
        if fpl is not None:
            for fp in fpl.children("formal_param"):
                params.append(self.formal_param(fp))
        ret = start.child("type_definition")
        return n.Procedure(
            name=strip_marker(ident.text),
            kind=kind,
            params=params,
            body=self.block(node.child(block_name)),
            scope=(scope.text.capitalize() if scope is not None else "Public"),
            is_static=start.child("static") is not None,
            return_type=self.type_of(ret),
            line=line,
        )

    def formal_param(self, fp):
        obj = fp.child("object")  # the grammar reads "ParamArray xs()" as an object with an empty argument list
        ident = obj.find("identifier") if obj is not None else fp.child("identifier")
        name = strip_marker(ident.text)
        sem = fp.child("passing_semantics")
        default = fp.child("default_value")
        return n.Param(
            name=name,
            byval=sem is not None and sem.text.strip().lower() == "byval",
            optional=fp.child("optional") is not None,
            default=self.expression(default.child("expression")) if default is not None else None,
            is_array=fp.child("array_indicator") is not None
            or (obj is not None and obj.child("parameter_list") is not None),
            type_name=self.type_of(fp.child("type_definition")),
            paramarray=fp.child("param_array") is not None,
        )

    def type_of(self, type_def):
        if type_def is None:
            return "variant"
        t = type_def.child("type")
        return t.text.strip().lower() if t is not None else "variant"

    def st_implicit_call_statement(self, node, line):
        target = self.object_expr(node.child("callable_object"))
        bare = node.child("bare_list")
        args = self.arg_list(bare) if bare is not None else []
        # The grammar does not reserve Call, GoTo, Erase and Stop, so a line holding one of them
        # comes back as a call to a procedure with that name.
        keyword = target.name.lower() if isinstance(target, n.Name) else None
        if keyword == "call" and len(args) == 1 and args[0].name is None:
            return n.CallStmt(args[0].expr, None, line=line)
        if keyword == "goto" and len(args) == 1 and isinstance(args[0].expr, (n.Name, n.Literal)):
            label = args[0].expr.name if isinstance(args[0].expr, n.Name) else str(args[0].expr.value)
            return n.GoTo(label.lower(), line=line)
        if keyword == "erase" and args and all(isinstance(a.expr, n.Name) for a in args):
            return n.Erase([a.expr.name for a in args], line=line)
        if keyword == "stop" and not args:
            return n.Unsupported("stop_statement", "Stop", line=line)
        if keyword == "error" and len(args) == 1 and args[0].name is None:  # Error 13 raises that error
            return n.CallStmt(n.Member(n.Name("Err"), "Raise"), args, line=line)
        return n.CallStmt(target, args if args else None, line=line)

    def st_inline_implicit_call(self, node, line):
        simple = node.child("simple_expr")
        target = self.simple_operand(simple)
        args = self.arg_list(node.child("bare_list"))
        return n.CallStmt(target, args if args else None, line=line)

    def st_explicit_call_statement(self, node, line):
        q = node.child("qualified_object")
        target = self.object_expr(q.child("bare_object"))
        return n.CallStmt(
            n.Call(target, self.arg_list(q.child("parameter_list").child("list").child("bare_list"))), None, line=line
        )

    def st_call_statement(self, node, line):
        target = self.object_expr(node.child("object"))
        lst = node.child("list")
        if lst is not None:
            target = n.Call(target, self.arg_list(lst.child("bare_list")))
        return n.CallStmt(target, None, line=line)

    def st_assignment_statement(self, node, line):
        body = node.child("assignment_body")
        return n.Assign(
            self.object_expr(body.child("object")), self.expression(body.child("expression")), False, line=line
        )

    def st_set_statement(self, node, line):
        target = self.object_expr(node.child("object"))
        value = self.expression(node.child("expression"))
        if node.child("new_keyword") is not None:
            value = n.New(self.text(node.child("expression")).strip())
        return n.Assign(target, value, True, line=line)

    def st_dim_statement(self, node, line):
        scoped = node.child("scoped_dim")
        scope = None
        if scoped is not None:
            sc = scoped.child("scope")
            scope = sc.text.strip().capitalize() if sc is not None else None
            basic = scoped.child("basic_dim")
        else:
            basic = node.child("unscoped_dim").child("basic_dim")
        return n.Dim(self.decls(basic), scope=scope, is_static=(scope == "Static"), line=line)

    def st_redim_statement(self, node, line):
        return n.ReDim(
            self.decls(node.child("basic_dim")), preserve=node.child("preserve_keyword") is not None, line=line
        )

    def decls(self, basic):
        out = []
        for od in basic.children("object_definition"):
            last = od.find("dim_expression_object_last_part")
            ident = last.child("identifier") if last is not None else od.find("identifier")
            size = od.find("size_definition")
            unsized = od.find("unsized_definition")
            type_def = od.find("type_definition")
            bounds = None
            if size is not None:
                bounds = []
                for part in size.kids:
                    if part.name == "size_range":
                        lo, hi = part.children("size")
                        bounds.append(
                            (self.expression(lo.child("expression")), self.expression(hi.child("expression")))
                        )
                    elif part.name == "size":
                        bounds.append((None, self.expression(part.child("expression"))))
            name = strip_marker(ident.text)
            marker = ident.text.strip()[len(name) :]
            out.append(
                n.VarDecl(
                    name=name,
                    bounds=bounds,
                    dynamic=unsized is not None,
                    type_name=self.type_of(type_def)
                    if type_def is not None
                    else TYPE_CHARACTERS.get(marker, "variant"),
                    is_new=type_def is not None and type_def.child("new_keyword") is not None,
                )
            )
        return out

    def st_const_statement(self, node, line):
        defs = []
        for cd in node.children("const_definition"):
            defs.append(
                (
                    strip_marker(cd.child("identifier").text),
                    self.expression(cd.child("expression")),
                    self.type_of(cd.child("type_definition")),
                )
            )
        scope = node.child("scope")
        return n.Const(defs, scope=scope.text.strip().capitalize() if scope is not None else None, line=line)

    def st_option_statement(self, node, line):
        return n.Option(" ".join(self.text(node).split()), line=line)

    def st_inline_if_statement(self, node, line):
        cond = self.expression(node.child("condition").child("expression"))
        body = self.inline_block(node.child("inline_if_block").child("inline_block"))
        else_node = node.child("inline_else_block")
        else_body = self.inline_block(else_node.child("inline_block")) if else_node is not None else None
        return n.If([(cond, body)], else_body, line=line)

    def inline_block(self, node):
        out = []
        for k in node.kids:
            if k.name == "valid_statement":
                out.extend(self.statement(k))
            elif k.name == "inline_implicit_call":
                out.append(self.st_inline_implicit_call(k, self.line_of(k.start)))
        return out

    def st_if_statement(self, node, line):
        start = node.child("if_start_statement")
        branches = [(self.expression(start.child("condition").child("expression")), self.block(node.child("if_block")))]
        for ei in node.children("else_if_statement"):
            cond = self.expression(ei.child("condition").child("expression"))
            inline = ei.child("else_if_inline")
            body = (
                self.inline_block(inline.child("inline_block"))
                if inline is not None
                else self.block(ei.child("else_if_block"))
            )
            branches.append((cond, body))
        else_stmt = node.child("else_statement")
        else_body = self.block(else_stmt.child("else_block")) if else_stmt is not None else None
        return n.If(branches, else_body, line=line)

    def st_for_statement(self, node, line):
        start = node.child("for_start_line")
        exprs = start.children("expression")
        step = start.child("for_stepping")
        return n.For(
            self.object_expr(start.child("object")),
            self.expression(exprs[0]),
            self.expression(exprs[1]),
            self.expression(step.child("expression")) if step is not None else None,
            self.block(node.child("block")),
            line=line,
        )

    def st_inline_for_statement(self, node, line):
        exprs = node.children("expression")
        step = node.child("for_stepping")
        return n.For(
            self.object_expr(node.child("object")),
            self.expression(exprs[0]),
            self.expression(exprs[1]),
            self.expression(step.child("expression")) if step is not None else None,
            self.block(node.child("block")),
            line=line,
        )

    def st_for_each_statement(self, node, line):
        start = node.child("for_each_start_line")
        return n.ForEach(
            self.object_expr(start.child("object")),
            self.expression(start.child("expression")),
            self.block(node.child("for_each_body")),
            line=line,
        )

    def st_while_statement(self, node, line):
        start = node.child("while_start_statement")
        return n.While(self.expression(start.child("expression")), self.block(node.child("while_block")), line=line)

    def st_do_statement(self, node, line):
        start = node.child("do_start_statement")
        end = node.child("do_end_statement")
        pre = post = None
        for clause_name in ("while_clause", "until_clause"):
            c = start.child(clause_name)
            if c is not None:
                pre = (clause_name.split("_")[0], self.expression(c.child("expression")))
        for clause_name, kind in (("post_while_clause", "while"), ("post_until_clause", "until")):
            c = end.child(clause_name)
            if c is not None:
                inner = c.kids[0]
                post = (kind, self.expression(inner.child("expression")))
        return n.DoLoop(pre, post, self.block(node.child("do_block")), line=line)

    def st_with_statement(self, node, line):
        start = node.child("with_start_statement")
        return n.With(self.expression(start.child("expression")), self.block(node.child("block")), line=line)

    def st_select_statement(self, node, line):
        subject = self.expression(node.child("select_start_statement").child("expression"))
        cases = []
        for item in node.children("case_item_block"):
            conds = []
            case_list = item.child("case_item_start").child("case_list")
            for ce in case_list.children("case_expression"):
                exprs = ce.children("expression")
                if ce.child("to_keyword") is not None:
                    conds.append(n.CaseCond("range", self.expression(exprs[0]), self.expression(exprs[1])))
                    continue
                e = self.expression(exprs[0])
                # The grammar reads "Is > 5" as a comparison of a variable called Is with 5.
                is_test = (
                    isinstance(e, n.BinOp)
                    and e.op in ("=", "<>", "<", ">", "<=", ">=")
                    and isinstance(e.left, n.Name)
                    and e.left.name.lower() == "is"
                )
                conds.append(n.CaseCond("is", e.right, op=e.op) if is_test else n.CaseCond("eq", e))
            cases.append((conds, self.block(item.child("case_body"))))
        else_block = node.child("case_else_block")
        else_body = self.block(else_block.child("case_body")) if else_block is not None else None
        return n.Select(subject, cases, else_body, line=line)

    def st_goto_statement(self, node, line):
        return n.GoTo(self.label_name(node), line=line)

    def st_on_statement(self, node, line):
        if node.child("on_error_resume") is not None:
            return n.OnError("resume_next", line=line)
        goto = node.child("on_error_goto")
        if goto is not None:
            label = self.label_name(goto)
            if label == "0":
                return n.OnError("goto0", line=line)
            return n.OnError("goto", label, line=line)
        return n.Unsupported("on_goto", self.text(node).strip(), line=line)

    def st_resume_statement(self, node, line):
        loc = node.child("resume_location")
        if loc is None:
            return n.Resume(None, line=line)
        if loc.text.strip().lower() == "next":
            return n.Resume("next", line=line)
        return n.Resume(self.label_name(loc), line=line)

    def st_exit_statement(self, node, line):
        return n.Exit(node.text.split()[-1].strip().lower(), line=line)

    def st_end_statement(self, node, line):
        return n.End(line=line)

    # --- expressions ---
    def expression(self, node):
        if node is None:
            return None
        if node.name == "condition":
            node = node.child("expression")
        toks = []
        self.flatten(node, toks)
        if not toks:
            raise VBAParseError(f"line {self.line_of(node.start)}: empty expression {self.text(node)!r}")
        expr, pos = self.pratt(toks, 0, 0)
        if pos != len(toks):
            raise VBAParseError(f"line {self.line_of(node.start)}: cannot read expression {self.text(node)!r}")
        return expr

    def flatten(self, node, toks):
        for k in node.kids:
            name = k.name
            if name in ("pre_named_argument", "passing_semantics", "addressof"):
                continue
            if name == "new_keyword":
                raise VBAParseError(f"line {self.line_of(k.start)}: 'New' inside an expression is not supported")
            if name == "pre_operator":
                toks.append(("un", self.pre_op(k)))
            elif name == "sign":
                toks.append(("un", k.text.strip()))
            elif name == "par_expression":
                self.flatten_par(k, toks)
            elif name == "operation":
                self.add_operation(k, toks)
            elif name in ("wsp",):
                continue
            elif name == "line_expression":
                raise VBAParseError(f"line {self.line_of(k.start)}: unsupported expression {self.text(k)!r}")

    def pre_op(self, k):
        if k.child("pre_not") is not None or k.text.strip().lower() == "not":
            return "not"
        raise VBAParseError(f"line {self.line_of(k.start)}: 'TypeOf' is not supported")

    def flatten_par(self, node, toks):
        inner = node.child("expression")
        if inner is not None:
            toks.append(("opnd", n.Paren(self.expression(inner))))
            return
        base = node.child("base_expression")
        if base is None:
            raise VBAParseError(f"line {self.line_of(node.start)}: unsupported expression {self.text(node)!r}")
        for k in base.kids:
            if k.name == "simple_expr":
                self.flatten_simple(k, toks)
            elif k.name == "operation":
                self.add_operation(k, toks)

    def flatten_simple(self, node, toks):
        for k in node.kids:
            if k.name == "pre_operator":
                toks.append(("un", self.pre_op(k)))
            elif k.name == "sign":
                toks.append(("un", k.text.strip()))
            elif k.name in ("call", "atom"):
                toks.append(("opnd", self.operand(k)))
            elif k.name == "channelid":
                raise VBAParseError(f"line {self.line_of(k.start)}: file channels are not supported")

    def simple_operand(self, node):
        toks = []
        self.flatten_simple(node, toks)
        expr, pos = self.pratt(toks, 0, 0)
        return expr

    def add_operation(self, k, toks):
        op = " ".join(k.text.split()).lower()
        if op in ("and not", "or not"):
            toks.append(("bin", op.split()[0]))
            toks.append(("un", "not"))
        elif op == "not":
            toks.append(("un", "not"))
        elif op in BIN_PREC:
            toks.append(("bin", op))
        else:
            raise VBAParseError(f"line {self.line_of(k.start)}: unsupported operator {op!r}")

    def pratt(self, toks, pos, min_prec):
        left, pos = self.prefix(toks, pos)
        while pos < len(toks) and toks[pos][0] == "bin" and BIN_PREC[toks[pos][1]] >= min_prec:
            op = toks[pos][1]
            right, pos = self.pratt(toks, pos + 1, BIN_PREC[op] + 1)
            left = n.BinOp(op, left, right)
        return left, pos

    def prefix(self, toks, pos):
        if pos >= len(toks):
            raise VBAParseError("incomplete expression")
        kind, val = toks[pos]
        if kind == "un":
            operand, pos = self.pratt(toks, pos + 1, UN_PREC[val])
            if val == "+":
                return operand, pos
            return n.Unary(val, operand), pos
        if kind == "opnd":
            return val, pos + 1
        raise VBAParseError(f"unexpected operator {val!r}")

    def operand(self, node):
        if node.name == "call":
            expr = self.object_expr(node.child("object"))
            pl = node.child("parameter_list")
            if pl is not None:
                expr = n.Call(expr, self.arg_list(pl.child("list").child("bare_list")))
            return expr
        if node.name == "atom":
            obj = node.child("object")
            if obj is not None:
                return self.object_expr(obj)
            return n.Literal(self.literal(node.child("literal")))
        raise VBAParseError(f"line {self.line_of(node.start)}: unsupported operand {self.text(node)!r}")

    def object_expr(self, node):
        expr = None
        with_ref = False
        for k in node.kids:
            if k.name == "implicit_object":
                with_ref = True
            elif k.name == "primary":
                rng = k.child("range_definition")
                if rng is not None:
                    expr = n.Call(n.Name("Range"), [n.Arg(n.Literal(self.text(rng).strip()[1:-1]))])
                    continue
                name = strip_marker(k.text.strip())
                expr = n.Member(n.WithRef(), name) if with_ref else n.Name(name, k.text.strip()[len(name) :])
            elif k.name == "attribute":
                expr = n.Member(expr, strip_marker(k.text.strip()))
            elif k.name == "parameter_list":
                expr = n.Call(expr, self.arg_list(k.child("list").child("bare_list")))
            elif k.name == "range_definition":
                raise VBAParseError(f"line {self.line_of(k.start)}: unsupported [..] reference")
        if expr is None:
            raise VBAParseError(f"line {self.line_of(node.start)}: cannot read {self.text(node)!r}")
        if isinstance(expr, n.Call) and isinstance(expr.func, n.Name) and expr.func.name == DATE_MARK:
            return n.Literal(parse_date_literal(expr.args[0].expr.value))
        return expr

    def arg_list(self, bare):
        if bare is None:
            return []
        args = []
        for k in bare.kids:
            if k.name == "positional_item":
                item = k.child("bare_list_item")
                args.append(self.arg(item) if item is not None else n.Arg(None))
            elif k.name == "bare_list_item":
                args.append(self.arg(k))
        return args

    def arg(self, item):
        expr_node = item.child("expression")
        named = expr_node.child("pre_named_argument")
        name = named.child("named_argument").text.strip().lower() if named is not None else None
        return n.Arg(self.expression(expr_node), name)

    def literal(self, node):
        k = node.kids[0]
        t = k.text
        if k.name == "stringliteral":
            return vbparser.makeUnicodeFromSafe(t[1:-1].replace('""', '"'))
        if k.name in ("integer", "longinteger"):
            return int(t.rstrip("%&@!^"))
        if k.name == "floatnumber":
            return float(t.rstrip("#!@"))
        if k.name == "hexinteger":
            return hex_value(int(t[2:].rstrip("&"), 16), long=t.endswith("&"))
        if k.name == "octinteger":
            return hex_value(int(t.lstrip("&oO").rstrip("&"), 8), long=t.endswith("&"))
        if k.name == "binaryinteger":
            return int(t[2:], 2)
        if k.name == "dateliteral":
            return parse_date_literal(t)
        raise VBAParseError(f"unsupported literal {t!r}")


TYPE_CHARACTERS = {"$": "string", "%": "integer", "&": "long", "!": "single", "#": "double", "@": "currency"}


def strip_marker(name):
    return name.rstrip("$%#&!@")


def parse_date_literal(t):
    """A date literal's date: month first, as the VBA editor writes them (#1/31/2024 12:30:00 PM#)."""
    d = parse_date_text(t.strip("#"))
    if d is None:
        raise VBAParseError(f"unsupported date literal {t!r}")
    return d
