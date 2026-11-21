"""A workbook file through the whole path a button press takes: read it, run converted VBA on it, save the copy.

The workbook is built here, a minimal package Excel can open, so no real workbook is needed.
"""

import json
import re
import zipfile

import pytest

from macropress import api, codegen, runtime, xlsx
from macropress.__main__ import REPORT_ITEMS, REPORT_TEXT, _bounded, main
from macropress.errors import VBANotImplemented
from macropress.excel import Book, Cell
from macropress.source import build_project, from_workbook
from macropress.xlsx import find_entry_points, read_workbook

PARTS = {
    "[Content_Types].xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        "</Types>"
    ),
    "_rels/.rels": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
        'officeDocument" Target="xl/workbook.xml"/></Relationships>'
    ),
    "xl/workbook.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>'
    ),
    "xl/_rels/workbook.xml.rels": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/'
        'worksheet" Target="worksheets/sheet1.xml"/></Relationships>'
    ),
    "xl/worksheets/sheet1.xml": (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1*2</f><v>2</v></c></row>'
        '<row r="2"><c r="A2" t="inlineStr"><is><t>keep</t></is></c></row>'
        "</sheetData></worksheet>"
    ),
}


def _workbook(path, parts=PARTS):
    with zipfile.ZipFile(path, "w") as z:
        for name, xml in parts.items():
            z.writestr(name, xml)
    return path


def _run(ctx, vba):
    project = build_project([("Module1", "Sub Test()\n" + vba + "\nEnd Sub\n")])
    namespace = {}
    exec(compile(codegen.convert(project, "test", [(None, "Test")]), "<generated>", "exec"), namespace)
    namespace["Test"](ctx)
    return [text for kind, text in ctx.messages if kind == "Debug.Print"]


def _press(path, out, vba):
    ctx = runtime.Context(str(path))
    _run(ctx, vba)
    ctx.save(str(out))
    return read_workbook(str(out)).find_sheet("Data")


def _value(sheet, r, c):
    cell = sheet.rows.get(r, {}).get(c)
    return None if cell is None else cell.value


def test_written_cells_are_saved_and_the_rest_is_kept(tmp_path):
    sheet = _press(_workbook(tmp_path / "in.xlsx"), tmp_path / "out.xlsx", 'Range("C1").Value = "new"')
    assert _value(sheet, 1, 3) == "new"
    assert _value(sheet, 2, 1) == "keep"
    assert _value(sheet, 1, 1) == 1.0


def test_a_formula_reading_a_changed_cell_is_saved_recalculated(tmp_path):
    sheet = _press(_workbook(tmp_path / "in.xlsx"), tmp_path / "out.xlsx", 'Range("A1").Value = 5')
    assert _value(sheet, 1, 1) == 5.0
    assert sheet.rows[1][2].formula == "A1*2"
    assert _value(sheet, 1, 2) == 10.0


def test_the_original_is_never_changed(tmp_path):
    original = _workbook(tmp_path / "in.xlsx")
    before = original.read_bytes()
    _press(original, tmp_path / "out.xlsx", 'Range("A1").ClearContents')
    assert original.read_bytes() == before


def test_the_copy_is_never_written_over_the_input(tmp_path):
    original = _workbook(tmp_path / "in.xlsx")
    before = original.read_bytes()
    with pytest.raises(VBANotImplemented, match="over the input"):
        _press(original, original, 'Range("A1").Value = 5')
    assert original.read_bytes() == before


def test_pasting_values_keeps_the_destinations_date_format():
    book = Book.blank(("Data",))
    book.find_sheet("Data").rows[1] = {1: Cell(45296.0), 2: Cell(None, None, "yyyy-mm-dd")}
    vba = 'Range("A1").Copy\nRange("B1").PasteSpecial xlPasteValues\nDebug.Print TypeName(Range("B1").Value)'
    assert _run(runtime.Context(book=book), vba) == ["Date"]


@pytest.mark.parametrize(
    ("paste", "type_name"),
    [
        ("xlPasteValues", "Double"),
        ("xlPasteFormulas", "Double"),
        ("xlPasteValuesAndNumberFormats", "Date"),
        ("xlPasteFormulasAndNumberFormats", "Date"),
        ("xlPasteAll", "Date"),
    ],
)
def test_a_paste_brings_number_formats_only_when_its_type_says_so(paste, type_name):
    book = Book.blank(("Data",))
    book.find_sheet("Data").rows[1] = {1: Cell(45296.0, None, "yyyy-mm-dd"), 2: Cell(1.0, None, "0.00")}
    vba = f'Range("A1").Copy\nRange("B1").PasteSpecial {paste}\nDebug.Print TypeName(Range("B1").Value)'
    assert _run(runtime.Context(book=book), vba) == [type_name]


def test_a_name_read_from_the_cell_that_uses_it_is_not_guessed():
    book = Book.blank(("Data",))
    book.names = {"fixed": [(None, "Data!$A$1")], "moving": [(None, "Data!A1")]}
    ctx = runtime.Context(book=book)
    assert _run(ctx, 'Range("A1").Value = 3\nDebug.Print Range("fixed").Value') == ["3"]
    with pytest.raises(VBANotImplemented, match="relative reference"):
        _run(ctx, 'Debug.Print Range("moving").Value')


def test_a_workbook_on_the_1904_date_system_is_refused(tmp_path):
    workbook = PARTS["xl/workbook.xml"].replace("<sheets>", '<workbookPr date1904="1"/><sheets>')
    path = _workbook(tmp_path / "in.xlsx", {**PARTS, "xl/workbook.xml": workbook})
    with pytest.raises(VBANotImplemented, match="1904"):
        read_workbook(str(path))


def test_a_shared_formula_moves_its_references_into_each_cell(tmp_path):
    sheet = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="shared" ref="B1:B2" si="0">Data!A1*2</f><v>2</v></c></row>'
        '<row r="2"><c r="A2"><v>5</v></c><c r="B2"><f t="shared" si="0"/><v>10</v></c></row>'
        "</sheetData></worksheet>"
    )
    path = _workbook(tmp_path / "in.xlsx", {**PARTS, "xl/worksheets/sheet1.xml": sheet})
    saved = _press(path, tmp_path / "out.xlsx", 'Range("A2").Value = 7')
    assert saved.rows[2][2].formula == "Data!A2*2"
    assert _value(saved, 2, 2) == 14.0


def _sheet(rows):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>'
        f"{rows}</sheetData></worksheet>"
    )


def test_a_formulas_saved_value_stands_until_a_cell_it_reads_changes(tmp_path):
    cells = '<c r="A1"><v>2</v></c><c r="B1"><f>SUM({1,2})+A1</f><v>5</v></c><c r="C1"><f>A1*3</f></c>'
    path = _workbook(tmp_path / "in.xlsx", {**PARTS, "xl/worksheets/sheet1.xml": _sheet(f'<row r="1">{cells}</row>')})
    # C1 was saved with no value: it is calculated. B1, which the evaluator can't read, keeps its saved value while
    # A1 does, and once A1 changes its value is unknown, not guessed.
    saved = _press(path, tmp_path / "out.xlsx", 'Range("Z99").Value = 1')
    assert (_value(saved, 1, 2), _value(saved, 1, 3)) == (5.0, 6.0)
    with pytest.raises(VBANotImplemented, match="could not be calculated"):
        _press(path, tmp_path / "out2.xlsx", 'Range("A1").Value = 5\nx = Range("B1").Value')


def test_calculation_follows_what_the_workbook_file_says(tmp_path):
    manual = PARTS["xl/workbook.xml"].replace("</workbook>", '<calcPr calcMode="manual"/></workbook>')
    path = _workbook(tmp_path / "in.xlsx", {**PARTS, "xl/workbook.xml": manual})
    ctx = runtime.Context(str(path))
    # saved with manual calculation: B1 keeps its value when A1 changes, and is calculated as the copy is saved
    assert _run(ctx, 'Range("A1").Value = 5\nDebug.Print Range("B1").Value') == ["2"]
    ctx.save(str(tmp_path / "out.xlsx"))
    assert _value(read_workbook(str(tmp_path / "out.xlsx")).find_sheet("Data"), 1, 2) == 10.0
    never = manual.replace('calcMode="manual"', 'calcMode="manual" calcOnSave="0"')
    path = _workbook(tmp_path / "in2.xlsx", {**PARTS, "xl/workbook.xml": never})
    assert _value(_press(path, tmp_path / "out2.xlsx", 'Range("A1").Value = 5'), 1, 2) == 2.0
    # an array formula entered over B1:B2: what B2 shows is the formula's doing, and unknown once A2 changes
    cells = (
        '<row r="1"><c r="A1"><v>1</v></c><c r="B1"><f t="array" ref="B1:B2">A1:A2</f><v>1</v></c></row>'
        '<row r="2"><c r="A2"><v>2</v></c><c r="B2"><v>2</v></c></row>'
    )
    path = _workbook(tmp_path / "in3.xlsx", {**PARTS, "xl/worksheets/sheet1.xml": _sheet(cells)})
    assert _run(runtime.Context(str(path)), 'Debug.Print Range("B2").Value') == ["2"]
    with pytest.raises(VBANotImplemented, match="could not be calculated"):
        _run(runtime.Context(str(path)), 'Range("A2").Value = 5\nx = Range("B2").Value')


STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>'
    '<fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>'
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs></styleSheet>'
)


def test_saving_formats_written_dates_and_refuses_a_rename_a_chart_would_miss(tmp_path):
    sheet = PARTS["xl/worksheets/sheet1.xml"].replace('<c r="A1"><v>1</v></c>', '<c r="A1" s="1"><v>1</v></c>')
    chart = '<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:f>Data!$A$1:$A$2</c:f></c:chartSpace>'  # fmt: skip
    parts = {**PARTS, "xl/styles.xml": STYLES, "xl/worksheets/sheet1.xml": sheet, "xl/charts/chart1.xml": chart}
    path = _workbook(tmp_path / "in.xlsx", parts)
    saved = _press(path, tmp_path / "out.xlsx", 'Range("A1").Value = DateSerial(2024, 1, 6)\nRange("E1").Value = Now')
    assert (saved.rows[1][1].numfmt, saved.rows[1][5].numfmt) == ("yyyy-mm-dd", "m/d/yy h:mm")
    with pytest.raises(VBANotImplemented, match="xl/charts/chart1.xml still refers to"):
        _press(path, tmp_path / "out2.xlsx", 'Sheets("Data").Name = "Renamed"')


@pytest.mark.parametrize("limit", ["MAX_PART_BYTES", "MAX_TOTAL_BYTES"])
def test_a_workbook_that_unzips_too_big_is_refused_before_anything_is_read(tmp_path, monkeypatch, limit):
    path = str(_workbook(tmp_path / "in.xlsm"))
    monkeypatch.setattr(xlsx, limit, 100)
    for read in (read_workbook, find_entry_points, from_workbook):
        with pytest.raises(VBANotImplemented, match="unzips to"):
            read(path)


def test_what_stops_a_command_is_reported_in_its_json_not_as_a_traceback(tmp_path, capsys):
    broken = tmp_path / "broken.xlsm"
    broken.write_bytes(b"not a workbook")
    workbook = PARTS["xl/workbook.xml"].replace("<sheets>", '<workbookPr date1904="1"/><sheets>')
    dated_1904 = _workbook(tmp_path / "1904.xlsm", {**PARTS, "xl/workbook.xml": workbook})
    out = str(tmp_path / "out.xlsm")
    for argv, error in (
        (["buttons", str(broken)], "the workbook could not be read: BadZipFile"),
        (["press", str(broken), "--button", "Go", "--out", out], "the workbook could not be read: BadZipFile"),
        (["press", str(dated_1904), "--button", "Go", "--out", out], "the workbook could not be read: .*1904"),
    ):
        assert main(argv) == 1
        assert re.match(error, json.loads(capsys.readouterr().out.splitlines()[-1])["error"])


def test_the_json_report_stays_small_whatever_the_macros_print():
    report = _bounded({"messages": ["x" * 5000] * (REPORT_ITEMS + 1), "steps": []})
    assert report["messages_total"] == REPORT_ITEMS + 1 and len(report["messages"]) == REPORT_ITEMS
    assert report["messages"][0] == "x" * REPORT_TEXT + f"... [{5000 - REPORT_TEXT} more characters]"
    assert "steps_total" not in report


def test_event_procedures_excel_would_have_run_are_named_in_the_report(tmp_path, monkeypatch):
    sheet_module = (
        'Attribute VB_Name = "Sheet1"\nAttribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"\n'
        "Private Sub Worksheet_Change(ByVal Target As Range)\nEnd Sub\n"
    )
    project = build_project([("Sheet1", sheet_module), ("Module1", 'Sub Go()\nRange("A1").Value = 1\nEnd Sub\n')])
    monkeypatch.setattr(api, "from_workbook", lambda path: project)
    report = api.press(str(_workbook(tmp_path / "in.xlsm")), ["Go"], str(tmp_path / "out.xlsm"))
    assert api.succeeded(report)
    assert any(
        "EVENTS-NOT-RUN" in line and "Sheet1.Worksheet_Change" in line for line in report["unverified_behaviours"]
    )


def test_buttons_are_listed_by_the_text_they_show(tmp_path):
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    parts = {
        **PARTS,
        "xl/worksheets/_rels/sheet1.xml.rels": (
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'<Relationship Id="rId1" Type="{rel}vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>'
            f'<Relationship Id="rId2" Type="{rel}drawing" Target="../drawings/drawing1.xml"/></Relationships>'
        ),
        "xl/drawings/vmlDrawing1.vml": (
            '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel">'
            '<v:shape id="_x0000_s1025"><v:textbox><div><font>Clear &amp; Reset</font></div></v:textbox>'
            '<x:ClientData ObjectType="Button"><x:FmlaMacro>[0]!Reset_Sheet</x:FmlaMacro></x:ClientData></v:shape></xml>'
        ),
        "xl/drawings/drawing1.xml": (
            '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
            'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><xdr:twoCellAnchor>'
            '<xdr:sp macro="[0]!Build_Report"><xdr:nvSpPr><xdr:cNvPr id="2" name="Rectangle 3"/></xdr:nvSpPr>'
            "<xdr:txBody><a:p><a:r><a:t>Run &amp; Save</a:t></a:r></a:p></xdr:txBody></xdr:sp>"
            "</xdr:twoCellAnchor></xdr:wsDr>"
        ),
    }
    path = _workbook(tmp_path / "in.xlsm", parts)
    assert set(find_entry_points(str(path))) == {
        ("Clear & Reset", "Reset_Sheet", "Data"),
        ("Run & Save", "Build_Report", "Data"),
    }
