"""Save the workbook model back into an .xlsm/.xlsx: a copy of the original with the changes applied.

Every part of the original file is copied unchanged (VBA project, buttons, styles, drawings...)
except the parts that changed:
  - each changed sheet: cell values and formulas, hidden rows, AutoFilter, selection, grouping
  - workbook.xml: active sheet, sheet names, filter ranges, recalculation flag
  - calcChain.xml is dropped when formulas changed (Excel rebuilds it)
Cell formatting is kept: a cleared cell keeps its style, as with ClearContents.
"""

import copy
import os
import re
import zipfile

from lxml import etree

from .errors import VBANotImplemented
from .excel import PENDING, area_text, col_letters, parse_area
from .formula import file_syntax, rename_sheets
from .values import XLError
from .xlsx import BUILTIN_FORMATS, NS, REL_NS, Package, read_workbook

M = NS[1:-1]
R = REL_NS[1:-1]
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

# Child order of <worksheet>, used to insert an element at a valid position.
SHEET_ORDER = [
    "sheetPr",
    "dimension",
    "sheetViews",
    "sheetFormatPr",
    "cols",
    "sheetData",
    "sheetCalcPr",
    "sheetProtection",
    "protectedRanges",
    "scenarios",
    "autoFilter",
    "sortState",
    "dataConsolidate",
    "customSheetViews",
    "mergeCells",
    "phoneticPr",
    "conditionalFormatting",
    "dataValidations",
    "hyperlinks",
    "printOptions",
    "pageMargins",
    "pageSetup",
    "headerFooter",
    "rowBreaks",
    "colBreaks",
    "customProperties",
    "cellWatches",
    "ignoredErrors",
    "smartTags",
    "drawing",
    "legacyDrawing",
    "legacyDrawingHF",
    "picture",
    "oleObjects",
    "controls",
    "webPublishItems",
    "tableParts",
    "extLst",
]


def q(tag):
    return f"{{{M}}}{tag}"


def _stored(sheet, r, c):
    """A cell as the sheet holds it. The writer compares cells as they are stored: reading one with get_cell would
    first calculate the formulas a copy read from the file has left uncalculated, and hide that the model's differ."""
    return sheet.rows.get(r, {}).get(c)


def _cell_key(cell):
    if cell is None or (cell.value is None and cell.formula is None):
        return None
    v = cell.value
    tag = (
        "p"
        if v is PENDING
        else "b"
        if isinstance(v, bool)
        else "e"
        if isinstance(v, XLError)
        else "n"
        if isinstance(v, float)
        else "s"
        if isinstance(v, str)
        else "x"
    )
    return (tag, None if v is PENDING else v, cell.formula, cell.numfmt)


def _number_text(x):
    return str(int(x)) if float(x).is_integer() and abs(x) < 1e15 else repr(float(x))


def save_workbook(book, original_path, out_path):
    """Write `book` (as changed by macros) to out_path, starting from original_path."""
    if os.path.realpath(out_path) == os.path.realpath(original_path):
        raise VBANotImplemented("saving over the input workbook: write the updated copy to another path")
    book.recalc_all()  # Excel recalculates before saving
    base = read_workbook(original_path)
    pkg = Package(original_path)
    if len(base.sheets) != len(book.sheets):
        raise VBANotImplemented("saving after sheets were added or deleted")
    sst = _shared_string_index(pkg)
    styles = _Styles(pkg.zip.read("xl/styles.xml")) if "xl/styles.xml" in pkg.names else None
    renamed = {
        old.name.lower(): new.name for old, new in zip(base.sheets, book.sheets, strict=False) if old.name != new.name
    }
    rels = {rid: target for rid, _, target in pkg.rels("xl/workbook.xml")}
    wb_root = etree.fromstring(pkg.zip.read("xl/workbook.xml"))
    sheet_parts = [rels[s.get(f"{{{R}}}id")] for s in wb_root.iter(q("sheet"))]

    patched = {}
    formulas_changed = False
    pending_formulas = False
    filter_refs = {}
    for i, (old, new) in enumerate(zip(base.sheets, book.sheets, strict=False)):
        part = sheet_parts[i]
        cells = _changed_cells(old, new)
        listed = {(r, c) for r, c, _ in cells}
        for k in set(old.spills) | set(new.spills):  # a spill that changed size: its formula's ref changes
            if old.spills.get(k) != new.spills.get(k) and k not in listed and _stored(new, *k) is not None:
                cells.append((k[0], k[1], _stored(new, *k)))
        cells.sort(key=lambda x: (x[0], x[1]))
        if any(c is not None and c.value is PENDING for _, _, c in cells):
            pending_formulas = True
        if any(
            (_stored(old, r, c) is not None and _stored(old, r, c).formula is not None)
            or (cell is not None and cell.formula is not None)
            for r, c, cell in cells
        ):
            formulas_changed = True
        old_af = _af_desc(old.autofilter)
        new_af = _af_desc(new.autofilter)
        selected = new in book.selected
        view_changed = (old.sel, old.active_cell, old in base.selected) != (new.sel, new.active_cell, selected)
        if not (cells or old.hidden_rows != new.hidden_rows or old_af != new_af or view_changed):
            continue
        patched[part] = _patch_sheet(
            pkg.zip.read(part), old, new, cells, sst, selected, old_af != new_af, _default_cm(book), styles
        )
        if old_af != new_af:
            filter_refs[i] = new.autofilter
    patched["xl/workbook.xml"] = _patch_workbook(wb_root, base, book, pending_formulas, filter_refs, renamed)
    if styles is not None and styles.changed:
        patched["xl/styles.xml"] = etree.tostring(styles.root, xml_declaration=True, encoding="UTF-8", standalone=True)
    if renamed:
        _check_renamed(pkg, patched, renamed)
    dropped = set()
    if formulas_changed and "xl/calcChain.xml" in pkg.names:
        dropped.add("xl/calcChain.xml")
        patched["xl/_rels/workbook.xml.rels"] = _drop_rel(pkg.zip.read("xl/_rels/workbook.xml.rels"), "calcChain")
        patched["[Content_Types].xml"] = _drop_override(pkg.zip.read("[Content_Types].xml"), "/xl/calcChain.xml")

    with zipfile.ZipFile(out_path, "w") as out:
        for info in pkg.zip.infolist():
            if info.filename in dropped:
                continue
            data = patched.get(info.filename)
            out.writestr(info, data if data is not None else pkg.zip.read(info.filename))
    return sorted(p for p in patched if p.startswith("xl/worksheets/"))


class _Styles:
    """The workbook's cell styles. A written cell that is General in the file and has a number format in the model (a
    Date written, a format pasted) gets a copy of its style with that format (its font, fill and borders kept), made
    once for each style and format. A cell with a number format in the file keeps it: the model may not know it,
    since a formatted cell with no value is not read."""

    # The formats VBA's own writes give a General cell (a Date, with or without its time), as Excel stores them.
    WRITTEN = {"m/d/yyyy": 14, "m/d/yyyy h:mm": 22}

    def __init__(self, xml):
        self.root = etree.fromstring(xml)
        self.xfs = self.root.find(q("cellXfs"))
        self.codes = dict(BUILTIN_FORMATS)
        for nf in self.root.iter(q("numFmt")):
            self.codes[int(nf.get("numFmtId"))] = nf.get("formatCode")
        self.ids = {**{code: i for i, code in self.codes.items()}, **self.WRITTEN}
        self.made = {}
        self.changed = False

    def _code(self, s):
        xf = self.xfs[s] if self.xfs is not None and 0 <= s < len(self.xfs) else None
        return self.codes.get(int(xf.get("numFmtId", 0)), "General") if xf is not None else "General"

    def apply(self, el, numfmt):
        s, want = int(el.get("s", 0)), numfmt
        if want is None or self.xfs is None or self._code(s) != "General":
            return
        if (s, want) not in self.made:
            xf = copy.deepcopy(self.xfs[s]) if 0 <= s < len(self.xfs) else etree.Element(q("xf"))
            xf.set("numFmtId", str(self._format_id(want)))
            xf.set("applyNumberFormat", "1")
            self.xfs.append(xf)
            self.xfs.set("count", str(len(self.xfs)))
            self.made[(s, want)] = len(self.xfs) - 1
            self.changed = True
        el.set("s", str(self.made[(s, want)]))

    def _format_id(self, code):
        if code not in self.ids:
            fmts = self.root.find(q("numFmts"))
            if fmts is None:
                fmts = etree.Element(q("numFmts"))
                self.root.insert(0, fmts)  # the first child of styleSheet
            new_id = max([163, *self.codes]) + 1
            etree.SubElement(fmts, q("numFmt"), numFmtId=str(new_id), formatCode=code)
            fmts.set("count", str(len(fmts)))
            self.codes[new_id] = code
            self.ids[code] = new_id
        return self.ids[code]


def _check_renamed(pkg, patched, renamed):
    """A renamed sheet that a part of the file still names (a chart, a pivot table, a validation rule: parts this does
    not rewrite) would leave that part pointing at no sheet in Excel: refused, as not modelled."""
    names = "|".join(re.escape(n) for n in renamed)
    quoted = "|".join(re.escape(n.replace("'", "''")) for n in renamed)
    ref = re.compile(rf"'(?:{quoted})'!|(?<![\w.'])(?:{names})!|sheet=\"(?:{names})\"", re.I)
    for part in sorted(pkg.names):
        if not part.startswith("xl/") or not part.endswith(".xml") or part.endswith("sharedStrings.xml"):
            continue  # docProps lists names Excel rewrites on saving; cell text may say anything
        data = patched.get(part) or pkg.zip.read(part)
        text = data.decode("utf-8", "replace")
        if part.startswith("xl/worksheets/"):
            text = re.sub(r"<sheetData.*?</sheetData>", "", text, flags=re.S)  # its formulas are rewritten
        m = ref.search(text)
        if m:
            raise VBANotImplemented(f"renaming a sheet that {part} still refers to ({m.group(0)})")


def _default_cm(book):
    """The cell-metadata index the file uses for dynamic-array formulas, if it has any."""
    for sheet in book.sheets:
        for cm in sheet.dynamic_cm.values():
            return cm
    return None


def _shared_string_index(pkg):
    from .xlsx import _shared_strings

    index = {}
    for i, text in enumerate(_shared_strings(pkg)):
        index.setdefault(text, i)
    return index


def _changed_cells(old, new):
    keys = set()
    for sheet in (old, new):
        for r, row in sheet.rows.items():
            for c in row:
                keys.add((r, c))
    out = []
    for r, c in sorted(keys):
        a, b = _stored(old, r, c), _stored(new, r, c)
        if _cell_key(a) != _cell_key(b):
            out.append((r, c, b if _cell_key(b) is not None else None))
    return out


def _af_desc(af):
    if af is None:
        return None
    return (af.rect, tuple(sorted((k, repr(v.describe())) for k, v in af.criteria.items())))


# ---------------------------------------------------------------------------
# Sheets
# ---------------------------------------------------------------------------


def _patch_sheet(xml, old, new, cells, sst, selected, filter_changed, default_cm=None, styles=None):
    root = etree.fromstring(xml)
    data = root.find(q("sheetData"))
    rows = {int(r.get("r")): r for r in data.findall(q("row"))}
    cmap = {}
    for row in rows.values():
        for c_el in row.findall(q("c")):
            area = parse_area(c_el.get("r"))
            cmap[(area[0], area[1])] = c_el

    touched = {(r, c) for r, c, _ in cells}
    if any(cmap.get(k) is not None and _has_shared_formula(cmap[k]) for k in touched):
        _expand_shared_formulas(cmap, old)
    for k in touched:
        el = cmap.get(k)
        f = el.find(q("f")) if el is not None else None
        if f is not None and f.get("t") == "array" and k not in old.spills:
            raise VBANotImplemented(f"changing part of an array formula ({old.name}!{col_letters(k[1])}{k[0]})")

    touched_rows = set()
    for r, c, cell in cells:
        el = cmap.get((r, c))
        if cell is None:
            if el is not None:
                _clear(el)
                if not (set(el.attrib) - {"r"}):
                    el.getparent().remove(el)
                    del cmap[(r, c)]
            continue
        if el is None:
            row_el = _row(data, rows, r)
            el = _new_cell(row_el, r, c)
            cmap[(r, c)] = el
        spill = new.spills.get((r, c)) if cell.formula is not None else None
        if spill is not None:
            cm = new.dynamic_cm.get((r, c)) or default_cm
            if cm is None:
                raise VBANotImplemented("saving a new dynamic-array formula in a workbook that has none")
            spill = (spill, cm)
        _set(el, cell, sst, spill, (r, c) in new.spill_owner)
        if styles is not None:
            styles.apply(el, cell.numfmt)
        touched_rows.add(r)

    for r in old.hidden_rows ^ new.hidden_rows:
        row_el = _row(data, rows, r)
        if r in new.hidden_rows:
            row_el.set("hidden", "1")
        elif "hidden" in row_el.attrib:
            del row_el.attrib["hidden"]
        touched_rows.add(r)
    for r in touched_rows:
        if r in rows and "spans" in rows[r].attrib:
            del rows[r].attrib["spans"]

    if filter_changed:
        _write_autofilter(root, new.autofilter)
    _set_filter_mode(root, new.in_filter_mode())
    _write_view(root, new, selected)
    _update_dimension(root, new)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _has_shared_formula(el):
    f = el.find(q("f"))
    return f is not None and f.get("t") == "shared"


def _expand_shared_formulas(cmap, old):
    """Give every shared-formula cell its own formula text, so cells can change independently."""
    for (r, c), el in cmap.items():
        f = el.find(q("f"))
        if f is not None and f.get("t") == "shared":
            cell = _stored(old, r, c)
            text = cell.formula if cell is not None else f.text
            for attr in ("t", "si", "ref"):
                f.attrib.pop(attr, None)
            f.text = text


def _clear(el):
    for child in list(el):
        if etree.QName(child).localname in ("v", "f", "is"):
            el.remove(child)
    for attr in ("t", "cm", "vm"):
        el.attrib.pop(attr, None)


def _row(data, rows, r):
    if r in rows:
        return rows[r]
    row_el = etree.Element(q("row"), r=str(r))
    later = [x for x in rows if x > r]
    if later:
        rows[min(later)].addprevious(row_el)
    else:
        data.append(row_el)
    rows[r] = row_el
    return row_el


def _new_cell(row_el, r, c):
    el = etree.Element(q("c"), r=f"{col_letters(c)}{r}")
    if row_el.get("customFormat") == "1" and row_el.get("s"):
        el.set("s", row_el.get("s"))
    for existing in row_el.findall(q("c")):
        if parse_area(existing.get("r"))[1] > c:
            existing.addprevious(el)
            return el
    row_el.append(el)
    return el


def _set(el, cell, sst, dynamic=None, spilled=False):
    """dynamic: (spill rectangle, cm) for a dynamic-array formula; spilled: a value filled in by one."""
    _clear(el)
    value = cell.value
    if cell.formula is not None:
        f = etree.SubElement(el, q("f"))
        f.text = file_syntax(cell.formula)
        if dynamic is not None:
            f.set("t", "array")
            f.set("ref", area_text(*dynamic[0], absolute=False))
            el.set("cm", dynamic[1])
        if value is PENDING or value is None:
            return
    if isinstance(value, bool):
        el.set("t", "b")
        etree.SubElement(el, q("v")).text = "1" if value else "0"
    elif isinstance(value, XLError):
        el.set("t", "e")
        etree.SubElement(el, q("v")).text = str(value)
    elif isinstance(value, float):
        etree.SubElement(el, q("v")).text = _number_text(value)
    elif isinstance(value, str):
        if cell.formula is not None or spilled:
            el.set("t", "str")
            etree.SubElement(el, q("v")).text = value
        elif value in sst:
            el.set("t", "s")
            etree.SubElement(el, q("v")).text = str(sst[value])
        else:
            el.set("t", "inlineStr")
            is_ = etree.SubElement(el, q("is"))
            t = etree.SubElement(is_, q("t"))
            t.text = value
            if value != value.strip():
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _insert_in_order(root, el):
    name = etree.QName(el).localname
    rank = SHEET_ORDER.index(name)
    for child in root:
        local = etree.QName(child).localname
        if local in SHEET_ORDER and SHEET_ORDER.index(local) > rank:
            child.addprevious(el)
            return
    root.append(el)


def _write_autofilter(root, af):
    old = root.find(q("autoFilter"))
    if af is None:
        if old is not None:
            root.remove(old)
        return
    el = etree.Element(q("autoFilter"), ref=area_text(*af.rect, absolute=False))
    for field, crit in sorted(af.criteria.items()):
        fc = etree.SubElement(el, q("filterColumn"), colId=str(field - 1))
        _write_criterion(fc, crit)
    if old is not None:
        old.addprevious(el)
        root.remove(old)
    else:
        _insert_in_order(root, el)


OPS_XML = {
    "=": "equal",
    "<": "lessThan",
    "<=": "lessThanOrEqual",
    "<>": "notEqual",
    ">=": "greaterThanOrEqual",
    ">": "greaterThan",
}


def _write_criterion(fc, crit):
    from .excel import BoolCriterion, OpCriterion, ValuesCriterion

    if isinstance(crit, ValuesCriterion):
        f = etree.SubElement(fc, q("filters"))
        if crit.blank:
            f.set("blank", "1")
        for value in sorted(crit.values):
            etree.SubElement(f, q("filter"), val=value)
        return
    parts = [crit] if isinstance(crit, OpCriterion) else [crit.a, crit.b] if isinstance(crit, BoolCriterion) else None
    if parts is None or not all(isinstance(p, OpCriterion) for p in parts):
        raise VBANotImplemented("saving this kind of AutoFilter criterion")
    cf = etree.SubElement(fc, q("customFilters"))
    if isinstance(crit, BoolCriterion) and crit.kind == "and":
        cf.set("and", "1")
    for p in parts:
        attrs = {"val": p.operand}
        if p.op != "=":
            attrs["operator"] = OPS_XML[p.op]
        etree.SubElement(cf, q("customFilter"), **attrs)


def _set_filter_mode(root, on):
    pr = root.find(q("sheetPr"))
    if on:
        if pr is None:
            pr = etree.Element(q("sheetPr"))
            root.insert(0, pr)
        pr.set("filterMode", "1")
    elif pr is not None and "filterMode" in pr.attrib:
        del pr.attrib["filterMode"]


def _write_view(root, sheet, selected):
    views = root.find(q("sheetViews"))
    if views is None:
        views = etree.Element(q("sheetViews"))
        _insert_in_order(root, views)
    view = views.find(q("sheetView"))
    if view is None:
        view = etree.SubElement(views, q("sheetView"), workbookViewId="0")
    if selected:
        view.set("tabSelected", "1")
    elif "tabSelected" in view.attrib:
        del view.attrib["tabSelected"]
    pane = view.find(q("pane"))
    active_pane = pane.get("activePane") if pane is not None else None
    sel = None
    for s in view.findall(q("selection")):
        if s.get("pane") == active_pane:
            sel = s
    if sel is None:
        sels = view.findall(q("selection"))
        sel = sels[-1] if sels else etree.SubElement(view, q("selection"))
        if active_pane and sel.get("pane") is None and not sels:
            sel.set("pane", active_pane)
    r, c = sheet.active_cell
    sel.set("activeCell", f"{col_letters(c)}{r}")
    sel.set(
        "sqref",
        area_text(*sheet.sel, absolute=False)
        if sheet.sel[0:2] != sheet.sel[2:4]
        else f"{col_letters(sheet.sel[1])}{sheet.sel[0]}",
    )


def _update_dimension(root, sheet):
    dim = root.find(q("dimension"))
    bounds = sheet.used_bounds()
    if dim is None or bounds is None:
        return
    r2, c2 = bounds[2], bounds[3]
    old = parse_area(dim.get("ref", "A1").split(":")[-1])
    if old is not None:
        r2, c2 = max(r2, old[2]), max(c2, old[3])
    dim.set("ref", f"A1:{col_letters(c2)}{r2}")


# ---------------------------------------------------------------------------
# Workbook-level parts
# ---------------------------------------------------------------------------


def _patch_workbook(root, base, book, pending_formulas, filter_refs, renamed=None):
    for el, old, new in zip(root.iter(q("sheet")), base.sheets, book.sheets, strict=False):
        if old.name != new.name:
            el.set("name", new.name)
    if renamed:  # a defined name follows its sheet's new name, as in Excel
        for d in root.iter(q("definedName")):
            d.text = rename_sheets(d.text or "", renamed)
    view = root.find(f"{q('bookViews')}/{q('workbookView')}")
    if view is not None and book.active is not None:
        tab = book.sheets.index(book.active)
        view.set("activeTab", str(tab))
        if int(view.get("firstSheet", 0)) > tab:
            view.set("firstSheet", str(tab))
    if pending_formulas:
        calc = root.find(q("calcPr"))
        if calc is None:
            calc = etree.SubElement(root, q("calcPr"))
        calc.set("fullCalcOnLoad", "1")
    if filter_refs:
        names = root.find(q("definedNames"))
        if names is None:
            names = etree.Element(q("definedNames"))
            calc = root.find(q("calcPr"))
            (calc.addprevious(names) if calc is not None else root.append(names))
        for index, af in filter_refs.items():
            existing = [
                d
                for d in names.findall(q("definedName"))
                if d.get("name") == "_xlnm._FilterDatabase" and d.get("localSheetId") == str(index)
            ]
            if af is None:
                for d in existing:
                    names.remove(d)
                continue
            sheet_name = book.sheets[index].name.replace("'", "''")
            text = f"'{sheet_name}'!{area_text(*af.rect)}"
            if existing:
                existing[0].text = text
            else:
                d = etree.SubElement(
                    names, q("definedName"), name="_xlnm._FilterDatabase", localSheetId=str(index), hidden="1"
                )
                d.text = text
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _drop_rel(xml, rel_type_suffix):
    root = etree.fromstring(xml)
    for rel in list(root):
        if rel.get("Type", "").endswith("/" + rel_type_suffix):
            root.remove(rel)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _drop_override(xml, part_name):
    root = etree.fromstring(xml)
    for el in list(root):
        if el.get("PartName") == part_name:
            root.remove(el)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
