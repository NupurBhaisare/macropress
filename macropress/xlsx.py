"""Read an .xlsx/.xlsm package into the workbook model (standard library only)."""

import datetime as dt
import html
import os
import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile

from .errors import VBANotImplemented
from .excel import (
    PENDING,
    AutoFilterState,
    Book,
    BoolCriterion,
    Cell,
    OpCriterion,
    Sheet,
    Table,
    UnsupportedCriterion,
    ValuesCriterion,
    col_num,
    parse_area,
    shift_formula,
)
from .numfmt import builtin_code, vba_code
from .values import XLError, date_to_serial

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
PKG_REL = "{http://schemas.openxmlformats.org/package/2006/relationships}"

CUSTOM_OPS = {
    "equal": "=",
    "lessThan": "<",
    "lessThanOrEqual": "<=",
    "notEqual": "<>",
    "greaterThanOrEqual": ">=",
    "greaterThan": ">",
}


def _ln(tag):
    return tag.rsplit("}", 1)[-1]


# A workbook is unzipped in memory, so one that unzips to far more than real workbooks do is refused before anything
# is read: a compressed-file bomb. A large real workbook unzips to 195 MiB, three sheets of 65 MiB, and takes about
# 2 GB of memory to press; the limits leave room above that. zipfile never inflates a part past the size its entry
# declares, so the declared sizes bound what is read.
MAX_PART_BYTES = 256 << 20
MAX_TOTAL_BYTES = 512 << 20


def open_package(path):
    """The workbook's zip, refused when it unzips to more than this reads."""
    z = zipfile.ZipFile(path)
    sizes = [info.file_size for info in z.infolist()]
    total, largest = sum(sizes), max(sizes, default=0)
    if total > MAX_TOTAL_BYTES or largest > MAX_PART_BYTES:
        z.close()
        raise VBANotImplemented(
            f"a workbook this big: it unzips to {total >> 20} MiB with a {largest >> 20} MiB part, over the "
            f"{MAX_TOTAL_BYTES >> 20} MiB ({MAX_PART_BYTES >> 20} MiB a part) this reads"
        )
    return z


class Package:
    def __init__(self, path):
        self.zip = open_package(path)
        self.names = set(self.zip.namelist())

    def xml(self, part):
        return ET.fromstring(self.zip.read(part))

    def rels(self, part):
        rp = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        if rp not in self.names:
            return []
        out = []
        for rel in self.xml(rp):
            target = rel.get("Target")
            if rel.get("TargetMode") != "External":
                target = (
                    target.lstrip("/")
                    if target.startswith("/")
                    else posixpath.normpath(posixpath.join(posixpath.dirname(part), target))
                )
            out.append((rel.get("Id"), rel.get("Type", "").rsplit("/", 1)[-1], target))
        return out


def read_workbook(path):
    pkg = Package(path)
    book = Book(os.path.basename(path))
    sst = _shared_strings(pkg)
    formats = _cell_formats(pkg)
    wb = pkg.xml("xl/workbook.xml")
    rels = {rid: target for rid, _, target in pkg.rels("xl/workbook.xml")}
    pr = wb.find(NS + "workbookPr")
    book.date1904 = pr is not None and pr.get("date1904") in ("1", "true")
    if book.date1904:
        raise VBANotImplemented("a workbook on the 1904 date system (its dates would be off by four years)")
    calc = wb.find(NS + "calcPr")
    if calc is not None:
        if calc.get("fullPrecision") in ("0", "false"):
            raise VBANotImplemented("a workbook set to 'precision as displayed' (numbers are rounded as they show)")
        book.manual = calc.get("calcMode") == "manual"
        book.calc_on_save = calc.get("calcOnSave") not in ("0", "false")
    sheet_parts = []
    for s in wb.iter(NS + "sheet"):
        part = rels.get(s.get(REL_NS + "id"))
        sheet = Sheet(book, s.get("name"), state=s.get("state", "visible"))
        book.sheets.append(sheet)
        sheet_parts.append((sheet, part))
    uncalculated = False
    for sheet, part in sheet_parts:
        if part and part in pkg.names:
            uncalculated |= _read_sheet(pkg, part, sheet, sst, formats, book)
    for dn in wb.iter(NS + "definedName"):
        local = dn.get("localSheetId")
        scope = book.sheets[int(local)].name if local is not None else None
        book.names.setdefault(dn.get("name").lower(), []).append((scope, (dn.text or "").strip()))
    view = wb.find(f"{NS}bookViews/{NS}workbookView")
    tab = int(view.get("activeTab", 0)) if view is not None else 0
    book.active = book.sheets[tab] if tab < len(book.sheets) else book.sheets[0]
    book.selected = [s for s in book.sheets if getattr(s, "_tab_selected", False)]
    if book.active not in book.selected:
        book.selected = [book.active] + book.selected
    if uncalculated:
        book._formulas()  # formulas saved without their values are calculated when first read
    return book


def _shared_strings(pkg):
    if "xl/sharedStrings.xml" not in pkg.names:
        return []
    out = []
    for si in pkg.xml("xl/sharedStrings.xml").iter(NS + "si"):
        parts = []
        for child in si:
            tag = _ln(child.tag)
            if tag == "t":
                parts.append(child.text or "")
            elif tag == "r":
                t = child.find(NS + "t")
                parts.append((t.text or "") if t is not None else "")
        out.append("".join(parts))
    return out


def _cell_formats(pkg):
    """The number format of each cell style, as VBA names it; General is None."""
    if "xl/styles.xml" not in pkg.names:
        return []
    root = pkg.xml("xl/styles.xml")
    codes = custom_formats(root)
    xfs = root.find(NS + "cellXfs")
    ids = [int(xf.get("numFmtId", 0)) for xf in (xfs if xfs is not None else [])]
    return [codes.get(i) or builtin_code(i) for i in ids]


def custom_formats(styles_root):
    """{number: code} of the number formats the file defines itself."""
    return {int(nf.get("numFmtId")): vba_code(nf.get("formatCode")) for nf in styles_root.iter(NS + "numFmt")}


def _read_sheet(pkg, part, sheet, sst, formats, book):
    """Read one sheet's cells and view into `sheet`; True when a formula in it was saved without its value."""
    masters = {}
    uncalculated = False
    root = pkg.xml(part)
    pr = root.find(NS + "sheetPr")
    if pr is not None:
        sheet.code_name = pr.get("codeName")
    dim = root.find(NS + "dimension")
    if dim is not None:
        area = parse_area(dim.get("ref", "A1").split(":")[-1])
        if area:
            sheet.dimension = (area[2], area[3])
    sv = root.find(f"{NS}sheetViews/{NS}sheetView")
    if sv is not None:
        sheet._tab_selected = sv.get("tabSelected") == "1"
        pane = sv.find(NS + "pane")
        active_pane = pane.get("activePane") if pane is not None else None
        sels = sv.findall(NS + "selection")
        chosen = None
        for sel in sels:
            if sel.get("pane") == active_pane:
                chosen = sel
        if chosen is None and sels:
            chosen = sels[-1]
        if chosen is not None:
            active = parse_area(chosen.get("activeCell", "A1")) or (1, 1, 1, 1)
            sq = (chosen.get("sqref", chosen.get("activeCell", "A1")) or "A1").split()[0]
            area = parse_area(sq) or active
            sheet.sel = area
            sheet.active_cell = (active[0], active[1])

    def number_format(style):
        style = int(style or 0)
        code = formats[style] if style < len(formats) else "General"
        return None if code == "General" else code

    default = root.find(NS + "sheetFormatPr")
    if default is not None and default.get("defaultColWidth"):
        sheet.default_width = float(default.get("defaultColWidth"))
    cols = root.find(NS + "cols")
    for col in cols if cols is not None else []:
        first, last = int(col.get("min")), int(col.get("max"))
        if col.get("width"):
            sheet.col_widths.append((first, last, float(col.get("width"))))
        if col.get("hidden") in ("1", "true"):
            sheet.hidden_cols.update(range(first, last + 1))
        if number_format(col.get("style")) is not None:  # cells of the column with no record of their own
            sheet.col_formats.update(dict.fromkeys(range(first, last + 1), number_format(col.get("style"))))
    data = root.find(NS + "sheetData")
    for row in data if data is not None else []:
        r = int(row.get("r"))
        if row.get("hidden") in ("1", "true"):
            sheet.hidden_rows.add(r)
        if row.get("customFormat") in ("1", "true"):
            sheet.row_formats[r] = number_format(row.get("s"))
        cells = {}
        for c in row.findall(NS + "c"):
            m = re.match(r"([A-Z]+)(\d+)$", c.get("r"))
            col = col_num(m.group(1))
            t = c.get("t", "n")
            v = c.find(NS + "v")
            f = c.find(NS + "f")
            numfmt = number_format(c.get("s"))
            value = None
            if t == "inlineStr":
                is_ = c.find(NS + "is")
                value = "".join(x.text or "" for x in is_.iter(NS + "t")) if is_ is not None else ""
            elif v is not None and v.text is not None:
                if t == "s":
                    value = sst[int(v.text)]
                elif t == "str":
                    value = v.text
                elif t == "b":
                    value = v.text == "1"
                elif t == "e":
                    value = XLError(v.text)
                elif t == "d":
                    value = date_to_serial(dt.datetime.fromisoformat(v.text.replace("Z", "")))
                else:
                    value = float(v.text)
            elif v is not None and t in ("str", "s"):
                value = ""
            formula = None
            if f is not None:
                text = f.text
                if f.get("t") == "shared":
                    si = f.get("si")
                    if text:
                        masters[si] = (r, col, text)
                    elif si in masters:
                        mr, mc, mt = masters[si]
                        text = shift_formula(mt, r - mr, col - mc)
                formula = text
                if f.get("t") == "array":
                    area = parse_area(f.get("ref") or c.get("r"))
                    if area and c.get("cm") is not None:  # a dynamic-array formula and its spill
                        sheet.spills[(r, col)] = area
                        sheet.dynamic_cm[(r, col)] = c.get("cm")
                    elif area:  # an array formula entered with Ctrl+Shift+Enter, over one cell or several
                        sheet.array_cells[(r, col)] = area
                if formula is not None and value is None:  # saved without its value: calculated when it is read
                    value, uncalculated = PENDING, True
            if value is None and formula is None and numfmt == sheet.inherited_format(r, col):
                continue  # an empty cell formatted as its row or column is anyway: no record needed
            cells[col] = Cell(value, formula, numfmt)
        if cells:
            sheet.rows[r] = cells
    for (ar, ac), (r1, c1, r2, c2) in [*sheet.spills.items(), *sheet.array_cells.items()]:
        for rr in range(r1, r2 + 1):
            for cc in range(c1, c2 + 1):
                if (rr, cc) != (ar, ac):  # its values are the formula's: unknown once the formula's are
                    sheet.spill_owner[(rr, cc)] = (ar, ac)
    merges = root.find(NS + "mergeCells")
    for mc in merges if merges is not None else []:
        area = parse_area(mc.get("ref"))
        if area:
            sheet.merged.append(area)
    af = root.find(NS + "autoFilter")
    if af is not None:
        sheet.autofilter = _autofilter(af)
    in_filter = sheet.autofilter is not None and sheet.autofilter.criteria
    for r in sheet.hidden_rows:
        inside = in_filter and sheet.autofilter.r1 < r <= sheet.autofilter.r2
        if not inside:
            sheet.manual_hidden.add(r)
    for _, rtype, target in pkg.rels(part):
        if rtype == "table" and target in pkg.names:
            tbl = _table(pkg.xml(target), sheet.name)
            if tbl is not None:
                book.tables[tbl.name.lower()] = tbl
    return uncalculated


def _autofilter(el):
    area = parse_area(el.get("ref"))
    state = AutoFilterState(*area)
    for fc in el.findall(NS + "filterColumn"):
        field = int(fc.get("colId")) + 1
        crit = None
        for child in fc:
            kind = _ln(child.tag)
            if kind == "filters":
                if child.findall(NS + "dateGroupItem"):
                    crit = UnsupportedCriterion("dateGroupItem")
                else:
                    crit = ValuesCriterion(
                        [x.get("val", "") for x in child.findall(NS + "filter")],
                        blank=child.get("blank") in ("1", "true"),
                    )
            elif kind == "customFilters":
                parts = [
                    OpCriterion(CUSTOM_OPS.get(cf.get("operator", "equal"), "="), cf.get("val", ""))
                    for cf in child.findall(NS + "customFilter")
                ]
                crit = (
                    parts[0]
                    if len(parts) == 1
                    else BoolCriterion("and" if child.get("and") in ("1", "true") else "or", parts[0], parts[1])
                )
            else:
                crit = UnsupportedCriterion(kind)
        if crit is not None:
            state.criteria[field] = crit
    return state


def _table(root, sheet_name):
    area = parse_area(root.get("ref", ""))
    if area is None:
        return None
    columns = [tc.get("name") for tc in root.iter(NS + "tableColumn")]
    return Table(
        root.get("displayName") or root.get("name"),
        sheet_name,
        *area,
        columns,
        int(root.get("headerRowCount", 1)),
        int(root.get("totalsRowCount", 0)),
    )


def find_entry_points(path):
    """Buttons and shapes that run a macro: [(label, macro, sheet)]."""
    pkg = Package(path)
    wb = pkg.xml("xl/workbook.xml")
    rels = {rid: target for rid, _, target in pkg.rels("xl/workbook.xml")}
    out = []
    for s in wb.iter(NS + "sheet"):
        part = rels.get(s.get(REL_NS + "id"))
        for _, rtype, target in pkg.rels(part or ""):
            if target not in pkg.names:
                continue
            if rtype == "vmlDrawing":
                text = pkg.zip.read(target).decode("utf-8", "replace")
                for shape in re.findall(r"<v:shape\b.*?</v:shape>", text, re.S):
                    macro = re.search(r"<x:FmlaMacro>(.*?)</x:FmlaMacro>", shape, re.S)
                    if macro:
                        caption = " ".join(
                            re.sub(
                                r"<[^>]+>", " ", " ".join(re.findall(r"<v:textbox.*?</v:textbox>", shape, re.S))
                            ).split()
                        )
                        macro_name = html.unescape(macro.group(1)).split("!")[-1]
                        out.append((html.unescape(caption), macro_name, s.get("name")))
            elif rtype == "drawing":
                root = pkg.xml(target)
                for el in root.iter():
                    if el.get("macro"):
                        text = " ".join(" ".join(x.text or "" for x in el.iter() if _ln(x.tag) == "t").split())
                        name = next((x.get("name") for x in el.iter() if _ln(x.tag) == "cNvPr"), "")
                        out.append((text or name, el.get("macro").split("!")[-1], s.get("name")))
    return out
