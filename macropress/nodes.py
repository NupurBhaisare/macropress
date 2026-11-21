"""Syntax tree for the VBA the converter translates (built from vb2py's parse tree)."""

from dataclasses import dataclass, field
from typing import Any

# --- expressions -----------------------------------------------------------


@dataclass
class Literal:
    value: Any


@dataclass
class Name:
    name: str


@dataclass
class WithRef:
    """The object of the enclosing With block (a leading '.')."""


@dataclass
class Member:
    obj: Any
    name: str


@dataclass
class Arg:
    expr: Any  # None: a missing positional argument
    name: str | None = None  # named argument (Field:=1)


@dataclass
class Call:
    func: Any
    args: list[Arg]


@dataclass
class BinOp:
    op: str
    left: Any
    right: Any


@dataclass
class Unary:
    op: str
    operand: Any


@dataclass
class Paren:
    expr: Any


@dataclass
class New:
    type_name: str


# --- statements (every statement carries its source line) -----------------


@dataclass
class Assign:
    target: Any
    value: Any
    is_set: bool = False
    line: int = 0


@dataclass
class CallStmt:
    target: Any
    args: list[Arg] | None = None  # None: no arguments after the name
    line: int = 0


@dataclass
class VarDecl:
    name: str
    bounds: list | None = None  # [(lo_expr or None, hi_expr)] for fixed-size arrays
    dynamic: bool = False  # Dim a() As ...
    type_name: str = "variant"
    is_new: bool = False


@dataclass
class Dim:
    decls: list[VarDecl]
    scope: str | None = None  # Public / Private / Global / Static at module level
    is_static: bool = False
    line: int = 0


@dataclass
class ReDim:
    decls: list[VarDecl]
    preserve: bool = False
    line: int = 0


@dataclass
class Const:
    defs: list  # [(name, expr, type_name)]
    scope: str | None = None
    line: int = 0


@dataclass
class If:
    branches: list  # [(condition, [statements])]
    else_body: list | None = None
    line: int = 0


@dataclass
class For:
    var: Any
    start: Any
    end: Any
    step: Any
    body: list
    line: int = 0


@dataclass
class ForEach:
    var: Any
    collection: Any
    body: list
    line: int = 0


@dataclass
class While:
    cond: Any
    body: list
    line: int = 0


@dataclass
class DoLoop:
    pre: tuple | None  # ("while"|"until", expr)
    post: tuple | None
    body: list
    line: int = 0


@dataclass
class With:
    obj: Any
    body: list
    line: int = 0


@dataclass
class CaseCond:
    kind: str  # "eq" | "range" | "is"
    expr: Any
    expr2: Any = None
    op: str | None = None


@dataclass
class Select:
    subject: Any
    cases: list  # [([CaseCond], [statements])]
    else_body: list | None = None
    line: int = 0


@dataclass
class GoTo:
    label: str
    line: int = 0


@dataclass
class Label:
    name: str
    line: int = 0


@dataclass
class OnError:
    mode: str  # "resume_next" | "goto" | "goto0"
    label: str | None = None
    line: int = 0


@dataclass
class Resume:
    target: str | None = None  # None: Resume, "next": Resume Next, else a label
    line: int = 0


@dataclass
class Erase:
    names: list[str]
    line: int = 0


@dataclass
class Exit:
    kind: str  # sub | function | property | for | do
    line: int = 0


@dataclass
class End:
    line: int = 0


@dataclass
class Option:
    text: str
    line: int = 0


@dataclass
class Unsupported:
    kind: str
    text: str
    line: int = 0


# --- procedures and modules ------------------------------------------------


@dataclass
class Param:
    name: str
    byval: bool = False
    optional: bool = False
    default: Any = None
    is_array: bool = False
    type_name: str = "variant"
    paramarray: bool = False


@dataclass
class Procedure:
    name: str
    kind: str  # sub | function | property get/let/set
    params: list[Param]
    body: list
    scope: str = "Public"
    is_static: bool = False
    return_type: str = "variant"
    line: int = 0
    module: Any = None


@dataclass
class ProcSource:
    """A procedure's text, parsed only when it is first called."""

    name: str
    kind: str
    scope: str
    first_line: int
    last_line: int
    text: str
    parsed: Procedure | None = None
    error: Exception | None = None


@dataclass
class Module:
    name: str
    kind: str  # standard | document | class
    declarations: list = field(default_factory=list)
    procedures: dict = field(default_factory=dict)  # lower name -> ProcSource
    options: list = field(default_factory=list)
    source: str = ""
