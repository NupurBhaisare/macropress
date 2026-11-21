"""Compiles a VBA procedure to a flat list of instructions.

GoTo, Exit For and On Error Resume Next need a procedure as numbered steps; codegen writes those steps
as Python when a procedure cannot be written as structured code.
"""

from . import nodes as n
from .errors import VBACompileError
from .values import coerce


class Var:
    __slots__ = ("value", "type_name")

    def __init__(self, value, type_name="variant"):
        self.value = value
        self.type_name = type_name

    def let(self, value):
        self.value = coerce(value, self.type_name)


class Instr:
    __slots__ = ("op", "a", "b", "c", "line", "depth")

    def __init__(self, op, a=None, b=None, c=None, line=0, depth=0):
        self.op, self.a, self.b, self.c, self.line, self.depth = op, a, b, c, line, depth


class Compiled:
    def __init__(self, proc, code, labels, decls, lines):
        self.proc, self.code, self.labels, self.decls, self.lines = proc, code, labels, decls, lines


class Compiler:
    def __init__(self, proc):
        self.proc = proc
        self.code = []
        self.labels = {}
        self.depth = 0
        self.loop_exits = []  # stack of (kind, [instr indexes to patch])
        self.decls = []
        self.lines = set()
        self.next_id = 0

    def emit(self, op, a=None, b=None, c=None, line=0):
        self.code.append(Instr(op, a, b, c, line, self.depth))
        if line:
            self.lines.add(line)
        return len(self.code) - 1

    def new_id(self):
        self.next_id += 1
        return self.next_id

    def compile(self):
        self.block(self.proc.body)
        for ins in self.code:
            if ins.op == "goto":
                if ins.a not in self.labels:
                    raise VBACompileError(f"{self.proc.name}: label {ins.a!r} not defined")
                ins.b = self.labels[ins.a]
        return Compiled(self.proc, self.code, self.labels, self.decls, self.lines)

    def block(self, stmts):
        for s in stmts:
            self.stmt(s)

    def stmt(self, s):
        if isinstance(s, n.Label):
            self.labels[s.name] = len(self.code)
        elif isinstance(s, (n.Dim, n.Const)):
            self.decls.append(s)  # declarations are set up when the procedure starts
        elif isinstance(s, n.If):
            end_jumps = []
            for cond, body in s.branches:
                jif = self.emit("jif", cond, None, line=s.line)
                self.block(body)
                end_jumps.append(self.emit("jmp", None))
                self.code[jif].b = len(self.code)
            if s.else_body is not None:
                self.block(s.else_body)
            for j in end_jumps:
                self.code[j].a = len(self.code)
        elif isinstance(s, n.For):
            lid = self.new_id()
            self.emit("for_init", lid, s, line=s.line)
            test = self.emit("for_test", lid, s, None, line=s.line)
            self.loop_exits.append(("for", []))
            self.block(s.body)
            self.emit("for_next", lid, s, test, line=s.line)
            exit_at = len(self.code)
            self.code[test].c = exit_at
            for j in self.loop_exits.pop()[1]:
                self.code[j].a = exit_at
        elif isinstance(s, n.ForEach):
            lid = self.new_id()
            self.emit("each_init", lid, s, line=s.line)
            test = self.emit("each_next", lid, s, None, line=s.line)
            self.loop_exits.append(("for", []))
            self.block(s.body)
            self.emit("jmp", test)
            exit_at = len(self.code)
            self.code[test].c = exit_at
            for j in self.loop_exits.pop()[1]:
                self.code[j].a = exit_at
        elif isinstance(s, n.While):
            test = self.emit("jif", s.cond, None, line=s.line)
            self.loop_exits.append(("while", []))
            self.block(s.body)
            self.emit("jmp", test)
            self.code[test].b = len(self.code)
            self.loop_exits.pop()
        elif isinstance(s, n.DoLoop):
            top = len(self.code)
            pre = None
            if s.pre:
                pre = self.emit("jif" if s.pre[0] == "while" else "jit", s.pre[1], None, line=s.line)
            self.loop_exits.append(("do", []))
            self.block(s.body)
            if s.post:
                self.emit("jit" if s.post[0] == "while" else "jif", s.post[1], top, line=s.line)
            else:
                self.emit("jmp", top)
            exit_at = len(self.code)
            if pre is not None:
                self.code[pre].b = exit_at
            for j in self.loop_exits.pop()[1]:
                self.code[j].a = exit_at
        elif isinstance(s, n.With):
            self.emit("with_push", s.obj, line=s.line)
            self.depth += 1
            self.block(s.body)
            self.emit("with_pop")
            self.depth -= 1
        elif isinstance(s, n.Select):
            sid = self.new_id()
            self.emit("select", sid, s.subject, line=s.line)
            end_jumps = []
            for conds, body in s.cases:
                test = self.emit("case", sid, conds, None, line=s.line)
                self.block(body)
                end_jumps.append(self.emit("jmp", None))
                self.code[test].c = len(self.code)
            if s.else_body is not None:
                self.block(s.else_body)
            for j in end_jumps:
                self.code[j].a = len(self.code)
        elif isinstance(s, n.GoTo):
            self.emit("goto", s.label, None, line=s.line)
        elif isinstance(s, n.Exit):
            if s.kind in ("for", "do"):
                for kind, patches in reversed(self.loop_exits):
                    if (kind == "for") == (s.kind == "for") and kind != "while":
                        patches.append(self.emit("jmp", None, line=s.line))
                        break
                else:
                    raise VBACompileError(f"{self.proc.name} line {s.line}: Exit {s.kind} not within {s.kind}")
            else:
                self.emit("return", line=s.line)
        else:
            self.emit("stmt", s, line=getattr(s, "line", 0))
