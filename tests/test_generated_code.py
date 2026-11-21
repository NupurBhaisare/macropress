"""Text from the workbook (button captions, its file name) goes into the generated module's docstrings: it must stay
text there, whatever quotes or backslashes it holds."""

from macropress import codegen
from macropress.source import build_project


def test_a_caption_can_neither_break_the_module_nor_run_as_code():
    injected = 'Run """\nINJECTED = True\n"""'
    path = "Save to C:\\Users\\reports"
    project = build_project([("Module1", "Sub A()\nEnd Sub\nSub B()\nEnd Sub\n")])
    source = codegen.convert(project, 'C:\\files\\"""book.xlsm', [(injected, "A"), (path, "B")])
    namespace = {}
    exec(compile(source, "<generated>", "exec"), namespace)
    assert "INJECTED" not in namespace
    assert set(namespace["BUTTONS"]) == {injected, path}


def test_a_button_whose_macro_is_gone_does_not_stop_the_others():
    project = build_project([("Module1", "Sub Go()\nEnd Sub\n")])
    source = codegen.convert(project, "book.xlsm", [("Old", "Deleted"), ("Go", "Module1.Go")])
    namespace = {}
    exec(compile(source, "<generated>", "exec"), namespace)
    assert set(namespace["BUTTONS"]) == {"Go"}


def test_a_module_name_cannot_run_as_code():
    project = build_project([("Mod\nINJECTED = True\n#", "Dim n As Long\nSub Go()\nn = 1\nEnd Sub\n")])
    namespace = {}
    exec(compile(codegen.convert(project, "book.xlsm", [(None, "Go")]), "<generated>", "exec"), namespace)
    assert "INJECTED" not in namespace
