"""Load the VBA project of a workbook (oletools) or from a folder of exported modules."""

import os
import re

from . import vbparse
from .errors import VBANotImplemented


class Project:
    def __init__(self, modules):
        self.modules = {m.name.lower(): m for m in modules}

    def procedures(self):
        for m in self.modules.values():
            for p in m.procedures.values():
                yield m, p


def module_kind(code):
    if re.search(r"^\s*VERSION\s+1\.0\s+CLASS", code, re.M | re.I):
        return "class"
    if re.search(r'^Attribute VB_Base\s*=\s*"0\{00020819', code, re.M):
        return "workbook"  # ThisWorkbook, whatever this Excel's language calls it
    if re.search(r'^Attribute VB_Base\s*=\s*"0\{000208', code, re.M):
        return "document"  # a sheet's module
    return "standard"


def _module_name(code, fallback):
    m = re.search(r'^Attribute VB_Name\s*=\s*"([^"]+)"', code, re.M)
    return m.group(1) if m else fallback


def _strip_class_header(code):
    """Remove the VERSION/BEGIN...END header of exported class and form modules (keeps line count)."""
    lines = code.split("\n")
    out, in_header = [], False
    for line in lines:
        s = line.strip()
        if re.match(r"^VERSION\s+\d", s, re.I):
            out.append("")
            continue
        if s.upper() == "BEGIN":
            in_header = True
            out.append("")
            continue
        if in_header:
            if s.upper() == "END":
                in_header = False
            out.append("")
            continue
        out.append(line)
    return "\n".join(out)


def build_project(named_sources):
    modules = []
    for name, code in named_sources:
        code = code.replace("\r\n", "\n").replace("\r", "\n")
        kind = module_kind(code)
        code = _strip_class_header(code)
        modules.append(vbparse.split_module(_module_name(code, name), code, kind))
    return Project(modules)


def from_workbook(path):
    try:
        from oletools.olevba import VBA_Parser
    except ImportError as err:
        raise VBANotImplemented(
            "reading VBA from a workbook needs oletools (pip install oletools), or pass --vba-dir with exported modules"
        ) from err
    from .xlsx import open_package

    open_package(path).close()  # oletools reads its OLE parts (the VBA project) whole: one too big is refused first
    parser = VBA_Parser(path)
    sources = []
    try:
        if not parser.detect_vba_macros():
            return Project([])
        for _, stream, _, code in parser.extract_macros():
            if isinstance(code, bytes):
                code = code.decode("latin-1")
            sources.append((stream.split("/")[-1], code))
    finally:
        parser.close()
    return build_project(sources)


def from_directory(path):
    sources = []
    for name in sorted(os.listdir(path)):
        if name.lower().endswith((".bas", ".cls", ".vba", ".frm")):
            with open(os.path.join(path, name), encoding="latin-1") as f:
                sources.append((os.path.splitext(name)[0], f.read()))
    return build_project(sources)
