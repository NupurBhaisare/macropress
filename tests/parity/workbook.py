"""Build the probe workbook: the sheets and macros of probes.py as an .xlsm, with the builder the examples use."""

import importlib.util
from pathlib import Path

HERE = Path(__file__).resolve().parent


def load(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def build(folder, with_macros=True):
    """Write parity.xlsm into `folder`. Returns its path, the probes module, and [(row, module, test)]: the row of
    the Out sheet each test writes its answers to. Excel needs the macros in the workbook; MacroPress's own test
    reads them from probes.py, and saves the seconds it takes to pack them, with_macros=False."""
    probes = load("parity_probes", HERE / "probes.py")
    builder = load("build_examples", HERE.parent.parent / "examples" / "build.py")
    sources, index = probes.module_sources()
    vba = Path(folder) / "vba"
    vba.mkdir(parents=True, exist_ok=True)
    if with_macros:
        for name, source in sources.items():
            (vba / f"{name}.bas").write_text(source, encoding="utf-8")
    else:
        builder.vba_project = lambda modules, project_id: b""
    spec = {"sheets": probes.sheets(index), "styles": probes.STYLES, "names": probes.NAMES}
    path = Path(folder) / "parity.xlsm"
    builder.write_workbook(path, builder.build_workbook("parity", spec, folder=vba))
    return path, probes, index


def probes_of(lines):
    """A test's probes, in the order of the columns they write to."""
    return [line for line in lines if line[:3] in ("S: ", "N: ", "V: ", "E: ", "X: ")]
