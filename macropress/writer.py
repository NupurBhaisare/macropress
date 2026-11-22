"""Save the workbook model back into an .xlsm/.xlsx: a copy of the original with the changes applied.

Every part of the original file is copied unchanged (VBA project, buttons, styles, drawings...)
except the parts that changed:
  - each changed sheet: cell values and formulas, hidden rows, AutoFilter, selection, grouping
  - workbook.xml: active sheet, sheet names, filter ranges, recalculation flag
  - styles.xml, when a written cell's number format changed (a Date written to a General cell)
  - calcChain.xml is dropped when formulas changed (Excel rebuilds it)
A cell's font, fill and borders are kept, whatever the macro did: a cleared cell keeps its look.
"""

import copy
import hashlib
import math
import os
import re
import zipfile

from lxml import etree

from .errors import VBANotImplemented
from .excel import PENDING, area_text, col_letters, parse_area
from .excel.addresses import moved_references
from .formula import file_syntax, rename_sheets
from .numfmt import BUILTIN_FORMATS, builtin_code
from .values import XLError
from .xlsx import NS, REL_NS, Package, custom_formats, read_workbook

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
    if not math.isfinite(x):
        raise VBANotImplemented(f"saving {x!r} in a cell: a number too large for Excel")
    return str(int(x)) if float(x).is_integer() and abs(x) < 1e15 else repr(float(x))


def save_workbook(book, original_path, out_path):
    """Write `book` (as changed by macros) to out_path, starting from original_path."""
    if os.path.realpath(out_path) == os.path.realpath(original_path):
        raise VBANotImplemented("saving over the input workbook: write the updated copy to another path")
    if book.calc_on_save:
        book.recalc_all()  # Excel recalculates before saving, unless the workbook says not to
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
        lines = _changed_lines(old, new)
        protection = _protection_change(book, old, new)
        changed = cells or any(lines) or protection is not False or old_af != new_af or view_changed
        if not (changed or old.hidden_rows != new.hidden_rows):
            continue
        patched[part] = _patch_sheet(
            pkg.zip.read(part), old, new, cells, sst, selected, old_af != new_af, _default_cm(book), styles, lines,
            protection,
        )  # fmt: skip
        if old_af != new_af:
            filter_refs[i] = new.autofilter
    if book.moves:
        if renamed:
            raise VBANotImplemented("saving after cells were cut and sheets renamed in one run")
        _check_moves(pkg, book, sheet_parts)
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


def locked_style(xf):
    """Whether a cell style (an <xf>) locks its cells: it does unless its <protection> says locked="0"."""
    protection = xf.find(q("protection")) if xf is not None else None
    return protection is None or protection.get("locked") not in ("0", "false")


def _set_locked(xf, locked):
    protection = xf.find(q("protection"))
    if locked:
        if protection is not None:
            protection.attrib.pop("locked", None)
            if not protection.attrib:
                xf.remove(protection)
        return
    if protection is None:
        protection = etree.Element(q("protection"))
        ext = xf.find(q("extLst"))
        (ext.addprevious(protection) if ext is not None else xf.append(protection))
    protection.set("locked", "0")
    xf.set("applyProtection", "1")


class _Styles:
    """The workbook's cell styles. A cell, row or column whose number format or lock in the model is not its style's
    in the file (a Date written to a General cell, a format set, pasted or cleared) gets a copy of its style with
    that format and lock (its font, fill and borders kept), made once for each style, format and lock."""

    def __init__(self, xml):
        self.root = etree.fromstring(xml)
        self.xfs = self.root.find(q("cellXfs"))
        self.codes = {**BUILTIN_FORMATS, **custom_formats(self.root)}
        self.ids = {code: i for i, code in sorted(self.codes.items(), reverse=True)}  # the lowest number of each
        self.made = {}
        self.changed = False

    def apply(self, el, numfmt, locked=True, attr="s"):
        """Give the element (a <c> or <row>; a <col> with attr="style") a style with this number format and lock."""
        if self.xfs is None:
            return
        s, want = int(el.get(attr, 0)), numfmt or "General"
        xf = self.xfs[s] if 0 <= s < len(self.xfs) else None
        have = int(xf.get("numFmtId", 0)) if xf is not None else 0
        same = want in (self.codes.get(have), builtin_code(have)) or self.ids.get(want) == have
        if same and locked_style(xf) == locked:
            if el.get(attr) is None and s != 0:
                el.set(attr, str(s))
            return
        if (s, want, locked) not in self.made:
            new = copy.deepcopy(xf) if xf is not None else etree.Element(q("xf"))
            if not same:
                new.set("numFmtId", str(self._format_id(want)))
                new.set("applyNumberFormat", "1")
            _set_locked(new, locked)
            self.xfs.append(new)
            self.xfs.set("count", str(len(self.xfs)))
            self.made[(s, want, locked)] = len(self.xfs) - 1
            self.changed = True
        el.set(attr, str(self.made[(s, want, locked)]))

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


_SQREF = re.compile(r'<(conditionalFormatting|dataValidation|hyperlink)\b[^>]*?\b(?:sqref|ref)="([^"]+)"')
_SHEET_REF = re.compile(r"((?:'(?:[^']|'')+'|[A-Za-z_][\w.]*))!(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)")


def _check_moves(pkg, book, sheet_parts):
    """Cells cut and pasted take their formulas and names with them in the model. Parts of the file this does not
    rewrite (conditional formats, validation, hyperlinks, charts, pivot tables) would be left on the cells as they
    were: a save where one covers or reads the cells moved, or where they landed, is refused, as not modelled."""
    index = {s.name.lower(): i for i, s in enumerate(book.sheets)}
    areas = {}  # sheet index -> the rectangles moves took cells from or put them on
    for src, rect, dst, landed in book.moves:
        areas.setdefault(index[src.lower()], []).append(rect)
        areas.setdefault(index[dst.lower()], []).append(landed)

    def hit(i, text):
        found = parse_area(text.replace("$", "")) if text else None
        return found is not None and any(not (found[2] < a[0] or found[0] > a[2] or found[3] < a[1] or found[1] > a[3])
                                         for a in areas.get(i, ()))  # fmt: skip

    for i in areas:
        xml = pkg.zip.read(sheet_parts[i]).decode("utf-8", "replace")
        xml = re.sub(r"<sheetData.*?</sheetData>", "", xml, flags=re.S)
        for kind, sqref in _SQREF.findall(xml):
            if any(hit(i, part) for part in sqref.split()):
                raise VBANotImplemented(f"saving cells cut and pasted that the sheet's {kind} covers")
    for part in sorted(pkg.names):
        if not part.startswith("xl/") or part.startswith("xl/worksheets/") or not part.endswith(".xml"):
            continue
        if part in ("xl/workbook.xml", "xl/styles.xml", "xl/sharedStrings.xml"):
            continue
        for name, ref in _SHEET_REF.findall(pkg.zip.read(part).decode("utf-8", "replace")):
            i = index.get(name.strip("'").replace("''", "'").lower())
            if i is not None and hit(i, ref):
                raise VBANotImplemented(f"saving cells cut and pasted that {part} reads ({name}!{ref})")


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
    """[(row, column, cell)] for each cell whose content, number format or lock the macros changed; the cell is None
    for one that is empty now."""
    keys = set()
    for sheet in (old, new):
        for r, row in sheet.rows.items():
            for c in row:
                keys.add((r, c))
    out = []
    for r, c in sorted(keys):
        a, b = _stored(old, r, c), _stored(new, r, c)
        if _cell_key(a) != _cell_key(b) or old.style_at(r, c) != new.style_at(r, c):
            out.append((r, c, b if _cell_key(b) is not None else None))
    return out


def _changed_lines(old, new):
    """The rows and the columns formatted whole whose style the macros changed: ({row: style or None}, {col: style})."""
    rows = {r: new.row_style(r) for r in {*old.row_formats, *new.row_formats} if old.row_style(r) != new.row_style(r)}
    cols = {*old.col_formats, *old.col_unlocked, *new.col_formats, *new.col_unlocked}
    return rows, {c: new.col_style(c) for c in sorted(cols) if old.col_style(c) != new.col_style(c)}


def _af_desc(af):
    if af is None:
        return None
    return (af.rect, tuple(sorted((k, repr(v.describe())) for k, v in af.criteria.items())))


# ---------------------------------------------------------------------------
# Sheets
# ---------------------------------------------------------------------------


def _patch_sheet(
    xml, old, new, cells, sst, selected, filter_changed, default_cm=None, styles=None, lines=({}, {}), protection=False
):
    root = etree.fromstring(xml)
    data = root.find(q("sheetData"))
    rows = {int(r.get("r")): r for r in data.findall(q("row"))}
    column_styles = [
        (int(col.get("min")), int(col.get("max")), col.get("style"))
        for col in root.iterfind(f"{q('cols')}/{q('col')}")
        if col.get("style")
    ]
    cmap = {}
    for row in rows.values():
        for c_el in row.findall(q("c")):
            area = parse_area(c_el.get("r"))
            cmap[(area[0], area[1])] = c_el

    touched = {(r, c) for r, c, _ in cells}
    if any(cmap.get(k) is not None and _has_shared_formula(cmap[k]) for k in touched):
        _expand_shared_formulas(cmap, old)
    keep_formula = set()  # array formulas entered with Ctrl+Shift+Enter whose values alone changed
    for k in touched:
        el = cmap.get(k)
        f = el.find(q("f")) if el is not None else None
        if f is not None and f.get("t") == "array" and k not in old.spills:
            before, now = _stored(old, *k), _stored(new, *k)
            if now is None or now.formula != before.formula or now.value is PENDING:
                raise VBANotImplemented(f"changing part of an array formula ({old.name}!{col_letters(k[1])}{k[0]})")
            keep_formula.add(k)

    touched_rows = set()
    for r, c, cell in cells:
        el = cmap.get((r, c))
        if cell is None:  # emptied, or empty with another style than it had (Clear, a format set or pasted)
            if el is None and new.style_at(r, c) != old.inherited(r, c):
                el = cmap[(r, c)] = _new_cell(_row(data, rows, r), r, c, column_styles)
            if el is not None:
                _clear(el)
                if styles is not None:
                    styles.apply(el, *new.style_at(r, c))
                if not (set(el.attrib) - {"r"}):
                    el.getparent().remove(el)
                    del cmap[(r, c)]
            continue
        if el is None:
            row_el = _row(data, rows, r)
            el = _new_cell(row_el, r, c, column_styles)
            cmap[(r, c)] = el
        spill = new.spills.get((r, c)) if cell.formula is not None else None
        if spill is not None:
            cm = new.dynamic_cm.get((r, c)) or default_cm
            if cm is None:
                raise VBANotImplemented("saving a new dynamic-array formula in a workbook that has none")
            spill = (spill, cm)
        _set(el, cell, sst, spill, (r, c) in new.spill_owner, keep=el.find(q("f")) if (r, c) in keep_formula else None)
        if styles is not None:
            styles.apply(el, *cell.style)
        touched_rows.add(r)

    row_styles, col_styles = lines
    if (row_styles or col_styles) and styles is None:
        raise VBANotImplemented("saving formats of whole rows or columns in a workbook with no styles part")
    for r, style in row_styles.items():
        row_el = _row(data, rows, r)
        if style is None:
            for attr in ("customFormat", "s"):
                row_el.attrib.pop(attr, None)
        else:
            row_el.set("customFormat", "1")
            styles.apply(row_el, *style)
    if col_styles:
        _write_col_styles(root, col_styles, styles, new)
    if row_styles or col_styles:
        # a cell the file holds but the model needs no record of takes its row's or column's new style; one the file
        # does not hold, whose row or column has another style now, is given an element to keep its own
        for (r, c), el in cmap.items():
            if (r, c) not in touched and old.style_at(r, c) != new.style_at(r, c):
                styles.apply(el, *new.style_at(r, c))
        for r, row in new.rows.items():
            for c, cell in row.items():
                if (r, c) not in cmap and cell.style != new.inherited(r, c):
                    el = cmap[(r, c)] = _new_cell(_row(data, rows, r), r, c, column_styles)
                    styles.apply(el, *cell.style)

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
    if protection is not False:
        _write_protection(root, protection)
    _set_filter_mode(root, new.in_filter_mode())
    _write_view(root, new, selected)
    _update_dimension(root, new)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _write_col_styles(root, col_styles, styles, sheet):
    """Give the columns their new styles: a <col> that covers more columns than change is cut in parts, and a column
    with none is given one, at the sheet's default width (a <col> with no width hides its column)."""
    cols = root.find(q("cols"))
    if cols is None:
        cols = etree.Element(q("cols"))
        _insert_in_order(root, cols)
    runs = []  # (first, last, style) of consecutive columns given the same style
    for c, style in col_styles.items():
        if runs and runs[-1][1] == c - 1 and runs[-1][2] == style:
            runs[-1][1] = c
        else:
            runs.append([c, c, style])
    for lo, hi, style in runs:
        for el in list(cols):
            first, last = int(el.get("min")), int(el.get("max"))
            if first <= hi and last >= lo and (first < lo or last > hi):  # cut it where the run starts and ends
                for a, b in ((first, lo - 1), (max(first, lo), min(last, hi)), (hi + 1, last)):
                    if a <= b:
                        part = copy.deepcopy(el)
                        part.set("min", str(a))
                        part.set("max", str(b))
                        el.addprevious(part)
                cols.remove(el)
        covered = []
        for el in cols:
            first, last = int(el.get("min")), int(el.get("max"))
            if lo <= first and last <= hi:
                styles.apply(el, *style, attr="style")
                covered.append((first, last))
        start = lo
        for first, last in sorted(covered) + [(hi + 1, hi + 1)]:
            if start < first and tuple(style) != (None, True):
                el = etree.Element(q("col"), min=str(start), max=str(first - 1), width=_default_width(root, sheet))
                styles.apply(el, *style, attr="style")
                later = [x for x in cols if int(x.get("min")) > start]
                (later[0].addprevious(el) if later else cols.append(el))
            start = last + 1


def _default_width(root, sheet):
    """The width, as the file writes widths, of a column the sheet gives no width of its own: the sheet's default if
    the file says it, else the one Excel for Windows works out from the base width and the widest digit of Calibri 11
    (seven pixels), in whole steps of eight pixels. A workbook with another font may show it a little otherwise."""
    pr = root.find(q("sheetFormatPr"))
    if pr is not None and pr.get("defaultColWidth"):
        return pr.get("defaultColWidth")
    if sheet.book.app is not None:
        sheet.book.app.note("COLUMN-WIDTH-DEFAULT")
    base = int(pr.get("baseColWidth", 8)) if pr is not None else 8
    pixels = -(-(base * 7 + 5) // 8) * 8
    return repr(math.trunc(pixels / 7 * 256) / 256)


def _protection_change(book, old, new):
    """The attributes of the sheet's <sheetProtection> now, None for no protection, or False when it has not changed.
    A password a macro gave is hashed with a salt made from the file and the sheet's name: the same press of the same
    workbook saves the same copy."""
    salt = hashlib.sha256(book.seed + new.name.encode("utf-8")).digest()[:16]
    before, after = (p.attrs(salt) if p is not None else None for p in (old.protection, new.protection))
    return False if before == after else after


def _write_protection(root, attrs):
    old = root.find(q("sheetProtection"))
    if old is not None:
        root.remove(old)
    if attrs is not None:
        _insert_in_order(root, etree.Element(q("sheetProtection"), attrs))


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


def _new_cell(row_el, r, c, column_styles=()):
    """A cell the file has no record of. It looks as its row does, or else its column."""
    el = etree.Element(q("c"), r=f"{col_letters(c)}{r}")
    if row_el.get("customFormat") in ("1", "true"):
        style = row_el.get("s")
    else:
        style = next((s for first, last, s in column_styles if first <= c <= last), None)
    if style:
        el.set("s", style)
    for existing in row_el.findall(q("c")):
        if parse_area(existing.get("r"))[1] > c:
            existing.addprevious(el)
            return el
    row_el.append(el)
    return el


def _set(el, cell, sst, dynamic=None, spilled=False, keep=None):
    """dynamic: (spill rectangle, cm) for a dynamic-array formula; spilled: a value filled in by one; keep: the
    cell's <f> element, to stay as the file has it."""
    _clear(el)
    value = cell.value
    if keep is not None:
        el.append(keep)
    elif cell.formula is not None:
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


WORKBOOK_ORDER = ["fileVersion", "fileSharing", "workbookPr", "workbookProtection", "bookViews", "sheets"]


def _patch_workbook(root, base, book, pending_formulas, filter_refs, renamed=None):
    for el, old, new in zip(root.iter(q("sheet")), base.sheets, book.sheets, strict=False):
        if old.name != new.name:
            el.set("name", new.name)
        if old.state != new.state:
            if new.state == "visible":
                el.attrib.pop("state", None)
            else:
                el.set("state", new.state)
    salt = hashlib.sha256(book.seed + b"workbook").digest()[:16]
    protection = [p.attrs(salt) if p is not None else None for p in (base.protection, book.protection)]
    if protection[0] != protection[1]:
        old_el = root.find(q("workbookProtection"))
        if old_el is not None:
            root.remove(old_el)
        if protection[1] is not None:
            el = etree.Element(q("workbookProtection"), protection[1])
            later = [c for c in root if etree.QName(c).localname in WORKBOOK_ORDER[4:]]
            (later[0].addprevious(el) if later else root.append(el))
    if renamed:  # a defined name follows its sheet's new name, as in Excel
        for d in root.iter(q("definedName")):
            d.text = rename_sheets(d.text or "", renamed)
    if book.moves:  # and the cells it names, where they were cut and pasted
        sheets = {s.name.lower(): s for s in book.sheets}
        for d in root.iter(q("definedName")):
            text = d.text or ""
            for src, rect, dst, landed in book.moves:
                cut, dest = (sheets[src.lower()], *rect), (sheets[dst.lower()], landed[0], landed[1])
                text = moved_references(text, None, None, cut, dest, lambda n: sheets.get(n.lower()))
            d.text = text
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
