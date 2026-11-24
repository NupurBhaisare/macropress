"""Build the example workbooks from the sources beside this script.

    python examples/build.py

Each example is a folder holding its VBA (one .bas file per module) and the workbook built from it. The sheets and
buttons are described in EXAMPLES below; the VBA project inside each .xlsm is written from the .bas files alone, with no
compiled code, so what a workbook runs is exactly what its .bas files say. Excel compiles the VBA when it opens the
workbook, and MacroPress reads the same source.

The builder writes the parts of an .xlsm that are not plain XML by hand:
- the VBA project (MS-OVBA: its dir, PROJECT and module streams, compressed as VBA compresses them)
- the compound file that holds it (MS-CFB, version 3)
It is deterministic: the same sources give the same workbook.
"""

import datetime as dt
import math
import struct
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent

# ---------------------------------------------------------------------------------------------------------------------
# The examples
# ---------------------------------------------------------------------------------------------------------------------


def d(text):
    return dt.date.fromisoformat(text)


# A sheet is (name, columns, rows, buttons): columns are widths, or (width, style) for a column formatted as a whole; a
# row is a list of cells, where a cell is a value or a (value, style) pair (styles: "bold", "money", "date",
# "percent"), and (None, style) is an empty cell formatted ahead for what the macros write there; a button is
# (caption, macro, from cell, to cell).
MONEY, DATE = (None, "money"), (None, "date")
ORDER_HEADER = [("Order", "bold"), ("Date", "bold"), ("Customer", "bold"), ("Region", "bold"), ("Amount", "bold")]

EXAMPLES = {
    "invoice": {
        "sheets": [
            (
                "Invoice",
                [22, 8, 12, 12],
                [
                    [("Invoice INV-0042", "bold")],
                    ["Tax rate", (0.08, "percent")],
                    [],
                    [("Item", "bold"), ("Qty", "bold"), ("Unit price", "bold"), ("Amount", "bold")],
                    ["Notebook", 3, (4.5, "money"), MONEY],
                    ["Pens (box of 10)", 2, (7.25, "money"), MONEY],
                    ["Stapler", 1, (12, "money"), MONEY],
                    ["Printer paper (ream)", 5, (6.4, "money"), MONEY],
                    ["Desk lamp", 1, (24.99, "money"), MONEY],
                    [],
                    [None, None, None, MONEY],
                    [None, None, None, MONEY],
                    [None, None, None, MONEY],
                ],
                [
                    ("Calculate totals", "CalculateTotals", "F2", "H4"),
                    ("Apply discount", "ApplyDiscount", "F5", "H7"),
                    ("Clear totals", "ClearTotals", "F8", "H10"),
                ],
            )
        ],
    },
    "orders-cleanup": {
        "sheets": [
            (
                "Orders",
                [8, 12, 24, 10, 12],
                [
                    ORDER_HEADER,
                    [1001, (d("2026-03-02"), "date"), "  blue harbor supply", "North", 1250],
                    [],
                    [1002, (d("2026-03-01"), "date"), "MAPLE AND FINCH ", "South", 830.5],
                    [1003, (d("2026-03-04"), "date"), "orbit cycles", "East", 412],
                    [1002, (d("2026-03-01"), "date"), "Maple and Finch", "South", 830.5],
                    [],
                    [1004, (d("2026-03-03"), "date"), " granite peak tools", "North", 2100],
                    [1005, (d("2026-03-02"), "date"), "sunny day bakery", "West", 96.75],
                    [1003, (d("2026-03-04"), "date"), "Orbit Cycles", "East", 412],
                    [1006, (d("2026-03-05"), "date"), "lighthouse BOOKS", "North", 540],
                ],
                [("Clean up", "CleanUpOrders", "G2", "I4")],
            ),
            (
                "Clean",
                [8, 12, 24, 10, 12],
                [ORDER_HEADER] + [[None, DATE, None, None, MONEY] for _ in range(20)],
                [
                    ("Show North only", "ShowNorthOnly", "G2", "I4"),
                    ("Show all orders", "ShowAllOrders", "G5", "I7"),
                ],
            ),
        ],
    },
    "sales-summary": {
        "sheets": [
            (
                "Sales",
                [12, 10, 12, 8, 12],
                [
                    [("Date", "bold"), ("Region", "bold"), ("Product", "bold"), ("Units", "bold"), ("Revenue", "bold")],
                    *[
                        [(d(day), "date"), region, product, units, (revenue, "money")]
                        for day, region, product, units, revenue in [
                            ("2026-01-05", "North", "Kettle", 12, 359.4),
                            ("2026-01-06", "South", "Toaster", 7, 314.3),
                            ("2026-01-09", "East", "Kettle", 4, 119.8),
                            ("2026-01-12", "West", "Blender", 3, 194.7),
                            ("2026-01-15", "North", "Toaster", 9, 404.1),
                            ("2026-01-20", "South", "Blender", 5, 324.5),
                            ("2026-01-27", "East", "Toaster", 6, 269.4),
                            ("2026-02-02", "North", "Blender", 2, 129.8),
                            ("2026-02-04", "West", "Kettle", 10, 299.5),
                            ("2026-02-10", "South", "Kettle", 8, 239.6),
                            ("2026-02-16", "East", "Blender", 7, 454.3),
                            ("2026-02-23", "West", "Toaster", 4, 179.6),
                            ("2026-03-02", "North", "Kettle", 15, 449.25),
                            ("2026-03-05", "South", "Toaster", 11, 493.9),
                            ("2026-03-11", "East", "Kettle", 9, 269.55),
                            ("2026-03-18", "West", "Blender", 6, 389.4),
                        ]
                    ],
                ],
                [],
            ),
            (
                "Summary",
                [14, 10, 10, 12, 10],
                [[("Sales by region", "bold")], ["Press Build summary to fill this in."], []]
                + [[None, None, None, MONEY, (None, "percent")] for _ in range(12)],
                [
                    ("Build summary", "BuildSummary", "G2", "I4"),
                    ("Email summary", "EmailSummary", "G5", "I7"),
                ],
            ),
        ],
    },
}

# ---------------------------------------------------------------------------------------------------------------------
# VBA compression (MS-OVBA 2.4.1)
# ---------------------------------------------------------------------------------------------------------------------


def _copy_token_help(difference):
    bit_count = max((difference - 1).bit_length(), 4)
    return bit_count, (0xFFFF >> bit_count) + 3


def compress(data):
    """Compress `data` as VBA does: chunks of 4096 bytes, each a sequence of literal bytes and back-references."""
    out = bytearray(b"\x01")
    for chunk_start in range(0, len(data), 4096):
        chunk_end = min(chunk_start + 4096, len(data))
        chunk = bytearray()
        pos = chunk_start
        while pos < chunk_end:
            flags, tokens = 0, bytearray()
            for bit in range(8):
                if pos >= chunk_end:
                    break
                best_len, best_at = 0, 0
                for candidate in range(pos - 1, chunk_start - 1, -1):
                    n = 0
                    while pos + n < chunk_end and data[candidate + n] == data[pos + n]:
                        n += 1
                    if n > best_len:
                        best_len, best_at = n, candidate
                if best_len >= 3:
                    bit_count, max_len = _copy_token_help(pos - chunk_start)
                    length = min(best_len, max_len)
                    tokens += struct.pack("<H", ((pos - best_at - 1) << (16 - bit_count)) | (length - 3))
                    flags |= 1 << bit
                    pos += length
                else:
                    tokens.append(data[pos])
                    pos += 1
            chunk.append(flags)
            chunk += tokens
        assert len(chunk) + 2 <= 4098, "a chunk did not compress; text this size always does"
        out += struct.pack("<H", (len(chunk) + 2 - 3) | 0xB000) + chunk
    return bytes(out)


# ---------------------------------------------------------------------------------------------------------------------
# The compound file that holds the VBA project (MS-CFB, version 3: 512-byte sectors)
# ---------------------------------------------------------------------------------------------------------------------

FREE, END, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF


class _Entry:
    def __init__(self, name, kind, data=b""):
        self.name, self.kind, self.data = name, kind, data  # kind: 1 storage, 2 stream, 5 root
        self.children, self.left, self.right, self.child, self.black = {}, NOSTREAM, NOSTREAM, NOSTREAM, True
        self.start, self.size = 0, 0


def _sibling_tree(entries):
    """Arrange a storage's children as the red-black tree the format asks for: a balanced tree in the format's name
    order (shorter names first, then by upper case), with the deepest level red when the tree is not perfect."""
    entries = sorted(entries, key=lambda e: (len(e.name), e.name.upper()))
    depths = {}

    def build(items, depth):
        if not items:
            return NOSTREAM
        mid = len(items) // 2
        node = items[mid]
        depths[node] = depth
        node.left, node.right = build(items[:mid], depth + 1), build(items[mid + 1 :], depth + 1)
        return node.sid

    root = build(entries, 0)
    deepest = max(depths.values(), default=0)
    for node, depth in depths.items():
        node.black = not (deepest > 0 and depth == deepest)
    return root


def compound_file(streams):
    """A compound file holding `streams`, a {"Storage/Stream": bytes} mapping."""
    root = _Entry("Root Entry", 5)
    for path, data in streams.items():
        node = root
        *storages, name = path.split("/")
        for storage in storages:
            node = node.children.setdefault(storage, _Entry(storage, 1))
        node.children[name] = _Entry(name, 2, data)
    entries = []

    def walk(node):
        node.sid = len(entries)
        entries.append(node)
        for child in node.children.values():
            walk(child)

    walk(root)
    for node in entries:
        if node.children:
            node.child = _sibling_tree(list(node.children.values()))

    sectors, fat = [], []

    def allocate(data):
        count = math.ceil(len(data) / 512)
        start = len(sectors)
        for i in range(count):
            sectors.append(data[i * 512 : (i + 1) * 512].ljust(512, b"\0"))
            fat.append(start + i + 1 if i < count - 1 else END)
        return start if count else END

    mini, minifat = bytearray(), []
    for node in entries:
        if node.kind != 2:
            continue
        node.size = len(node.data)
        if node.size >= 4096:
            node.start = allocate(node.data)
        elif node.size:
            count = math.ceil(node.size / 64)
            node.start = len(mini) // 64
            minifat += [node.start + i + 1 if i < count - 1 else END for i in range(count)]
            mini += node.data.ljust(count * 64, b"\0")
        else:
            node.start = END
    root.start, root.size = (allocate(bytes(mini)), len(mini)) if mini else (END, 0)
    minifat_bytes = struct.pack(f"<{len(minifat)}I", *minifat)
    minifat_bytes += b"\xff" * (-len(minifat_bytes) % 512)
    minifat_start = allocate(minifat_bytes) if minifat else END

    directory = bytearray()
    for node in entries:
        name = node.name.encode("utf-16-le")
        directory += struct.pack(
            "<64sHBBIII16sIQQIQ",
            name,
            len(name) + 2,
            node.kind,
            1 if node.black else 0,
            node.left,
            node.right,
            node.child,
            b"",
            0,
            0,
            0,
            node.start,
            node.size,
        )
    while len(directory) % 512:
        directory += struct.pack("<64sHBBIII16sIQQIQ", b"", 0, 0, 0, NOSTREAM, NOSTREAM, NOSTREAM, b"", 0, 0, 0, 0, 0)
    directory_start = allocate(bytes(directory))

    fat_count = 0
    while fat_count * 128 < len(fat) + fat_count:
        fat_count += 1
    assert fat_count <= 109, "more than the header's 109 FAT sectors"
    fat_start = len(sectors)
    fat += [FATSECT] * fat_count
    fat += [FREE] * (fat_count * 128 - len(fat))
    for i in range(fat_count):
        sectors.append(struct.pack("<128I", *fat[i * 128 : (i + 1) * 128]))

    header = struct.pack(
        "<8s16sHHHHH6s9I",
        bytes.fromhex("D0CF11E0A1B11AE1"),
        b"",
        0x003E,
        0x0003,
        0xFFFE,
        9,
        6,
        b"",
        0,
        fat_count,
        directory_start,
        0,
        4096,
        minifat_start,
        math.ceil(len(minifat) / 128),
        END,
        0,
    )
    header += struct.pack("<109I", *([fat_start + i for i in range(fat_count)] + [FREE] * (109 - fat_count)))
    return header + b"".join(sectors)


# ---------------------------------------------------------------------------------------------------------------------
# The VBA project (MS-OVBA 2.3)
# ---------------------------------------------------------------------------------------------------------------------

WORKBOOK_BASE = "0{00020819-0000-0000-C000-000000000046}"
SHEET_BASE = "0{00020820-0000-0000-C000-000000000046}"


def document_module(name, base):
    """The source of a document module (ThisWorkbook or a sheet's) with no code of its own."""
    return (
        f'Attribute VB_Name = "{name}"\r\nAttribute VB_Base = "{base}"\r\nAttribute VB_GlobalNameSpace = False\r\n'
        "Attribute VB_Creatable = False\r\nAttribute VB_PredeclaredId = True\r\nAttribute VB_Exposed = True\r\n"
        "Attribute VB_TemplateDerived = False\r\nAttribute VB_Customizable = True\r\n"
    )


def _encrypt(project_id, data, seed):
    """The encryption the PROJECT stream's CMG, DPB and GC values use (MS-OVBA 2.4.3.2), as hex."""
    key = sum(project_id.encode("ascii")) & 0xFF
    out = [seed, seed ^ 2, seed ^ key]
    previous = key
    plain = [seed] * ((seed & 6) // 2) + list(struct.pack("<I", len(data))) + list(data)
    for byte in plain:
        out.append(byte ^ ((out[-2] + previous) & 0xFF))
        previous = byte
    return bytes(out).hex().upper()


def vba_project(modules, project_id):
    """vbaProject.bin for `modules`, a list of (name, kind, source) with kind "document" or "standard"."""

    def record(record_id, data=b""):
        return struct.pack("<HI", record_id, len(data)) + data

    cp = "cp1252"
    info = (
        record(0x0001, struct.pack("<I", 1))  # PROJECTSYSKIND: 32-bit Windows
        + record(0x0002, struct.pack("<I", 0x409))  # PROJECTLCID
        + record(0x0014, struct.pack("<I", 0x409))  # PROJECTLCIDINVOKE
        + record(0x0003, struct.pack("<H", 1252))  # PROJECTCODEPAGE
        + record(0x0004, b"VBAProject")  # PROJECTNAME
        + record(0x0005)  # PROJECTDOCSTRING
        + record(0x0040)
        + record(0x0006)  # PROJECTHELPFILEPATH
        + record(0x003D)
        + record(0x0007, struct.pack("<I", 0))  # PROJECTHELPCONTEXT
        + record(0x0008, struct.pack("<I", 0))  # PROJECTLIBFLAGS
        + struct.pack("<HIIH", 0x0009, 4, 1, 0)  # PROJECTVERSION
        + record(0x000C)  # PROJECTCONSTANTS
        + record(0x003C)
    )
    # No references: the VBA and Excel libraries are always there, and the examples use nothing else.
    listing = record(0x000F, struct.pack("<H", len(modules))) + record(0x0013, struct.pack("<H", 0xFFFF))
    for name, kind, _ in modules:
        listing += (
            record(0x0019, name.encode(cp))  # MODULENAME
            + record(0x0047, name.encode("utf-16-le"))  # MODULENAMEUNICODE
            + record(0x001A, name.encode(cp))  # MODULESTREAMNAME
            + record(0x0032, name.encode("utf-16-le"))
            + record(0x001C)  # MODULEDOCSTRING
            + record(0x0048)
            + record(0x0031, struct.pack("<I", 0))  # MODULEOFFSET: the stream holds only the source
            + record(0x001E, struct.pack("<I", 0))  # MODULEHELPCONTEXT
            + record(0x002C, struct.pack("<H", 0xFFFF))  # MODULECOOKIE
            + record(0x0022 if kind == "document" else 0x0021)  # MODULETYPE
            + record(0x002B)  # terminator
        )
    dir_stream = info + listing + record(0x0010)

    not_protected = _encrypt(project_id, struct.pack("<I", 0), 0x11)
    no_password = _encrypt(project_id, b"\x00", 0x22)
    visible = _encrypt(project_id, b"\xff", 0x33)
    project = [f'ID="{project_id}"']
    project += [f"Document={name}/&H00000000" for name, kind, _ in modules if kind == "document"]
    project += [f"Module={name}" for name, kind, _ in modules if kind == "standard"]
    project += [
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        f'CMG="{not_protected}"',
        f'DPB="{no_password}"',
        f'GC="{visible}"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
        *[f"{name}=0, 0, 0, 0, C" for name, _, _ in modules],
        "",
    ]
    names = b"".join(n.encode(cp) + b"\0" + n.encode("utf-16-le") + b"\0\0" for n, _, _ in modules) + b"\0\0"

    streams = {
        "PROJECT": "\r\n".join(project).encode(cp),
        "PROJECTwm": names,
        "VBA/_VBA_PROJECT": bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00]),  # no compiled code: compile on open
        "VBA/dir": compress(dir_stream),
    }
    for name, _, source in modules:
        streams[f"VBA/{name}"] = compress(source.encode(cp))
    return compound_file(streams)


# ---------------------------------------------------------------------------------------------------------------------
# The workbook package
# ---------------------------------------------------------------------------------------------------------------------

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
STYLES = {None: 0, "bold": 1, "money": 2, "date": 3, "percent": 4}


def _col_letters(n):
    letters = ""
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _cell_ref(ref):
    letters = "".join(ch for ch in ref if ch.isalpha())
    col = 0
    for ch in letters:
        col = col * 26 + ord(ch.upper()) - 64
    return col - 1, int(ref[len(letters) :]) - 1  # zero-based column and row


def _serial(day):
    return (day - dt.date(1899, 12, 30)).days


def _sheet_xml(index, columns, rows, strings, has_buttons, styles=None, protection=""):
    """protection: a <sheetProtection> element, for a sheet the file saves protected."""
    styles = styles or STYLES
    cols = ""
    for i, column in enumerate(columns, 1):
        width, style = column if isinstance(column, tuple) else (column, None)
        formatted = f' style="{styles[style]}"' if style else ""
        cols += f'<col min="{i}" max="{i}" width="{width}"{formatted} customWidth="1"/>'
    body = []
    for r, row in enumerate(rows, 1):
        cells = []
        for c, cell in enumerate(row, 1):
            value, style = cell if isinstance(cell, tuple) else (cell, None)
            ref = f"{_col_letters(c)}{r}"
            s = f' s="{styles[style]}"' if styles[style] else ""
            if value is None:
                if s:
                    cells.append(f'<c r="{ref}"{s}/>')
            elif isinstance(value, bool):
                cells.append(f'<c r="{ref}"{s} t="b"><v>{int(value)}</v></c>')
            elif isinstance(value, str):
                strings.setdefault(value, len(strings))
                cells.append(f'<c r="{ref}"{s} t="s"><v>{strings[value]}</v></c>')
            elif isinstance(value, dt.date):
                cells.append(f'<c r="{ref}"{s}><v>{_serial(value)}</v></c>')
            else:
                cells.append(f'<c r="{ref}"{s}><v>{value!r}</v></c>')
        if cells:
            body.append(f'<row r="{r}">{"".join(cells)}</row>')
    selected = ' tabSelected="1"' if index == 1 else ""
    drawing = '<drawing r:id="rId1"/>' if has_buttons else ""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="{MAIN}" xmlns:r="{REL}">'
        f'<sheetPr codeName="Sheet{index}"/><sheetViews><sheetView workbookViewId="0"{selected}/></sheetViews>'
        f'<sheetFormatPr defaultRowHeight="15"/><cols>{cols}</cols><sheetData>{"".join(body)}</sheetData>{protection}'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        f"{drawing}</worksheet>"
    )


def _drawing_xml(buttons):
    """Buttons as Excel draws a shape with a macro assigned: a rounded rectangle that runs the macro when clicked."""
    shapes = []
    for i, (caption, macro, start, end) in enumerate(buttons):
        (c1, r1), (c2, r2) = _cell_ref(start), _cell_ref(end)
        shapes.append(
            '<xdr:twoCellAnchor editAs="absolute">'
            f"<xdr:from><xdr:col>{c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r1}</xdr:row>"
            "<xdr:rowOff>0</xdr:rowOff></xdr:from>"
            f"<xdr:to><xdr:col>{c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r2}</xdr:row>"
            "<xdr:rowOff>0</xdr:rowOff></xdr:to>"
            f'<xdr:sp macro="[0]!{macro}" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{i + 2}" '
            f'name="{escape(caption)}"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr>'
            '<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>'
            '<a:solidFill><a:srgbClr val="2F6FD0"/></a:solidFill><a:ln><a:noFill/></a:ln></xdr:spPr>'
            '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" rtlCol="0" anchor="ctr"/><a:lstStyle/>'
            '<a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1100" b="1"><a:solidFill>'
            f'<a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>{escape(caption)}</a:t></a:r></a:p></xdr:txBody>'
            "</xdr:sp><xdr:clientData/></xdr:twoCellAnchor>"
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
        f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">{"".join(shapes)}</xdr:wsDr>'
    )


STYLES_XML = (
    f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="{MAIN}">'
    '<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.00"/><numFmt numFmtId="165" '
    'formatCode="yyyy-mm-dd"/></numFmts>'
    '<fonts count="2"><font><sz val="11"/><name val="Calibri"/><family val="2"/></font>'
    '<font><b/><sz val="11"/><name val="Calibri"/><family val="2"/></font></fonts>'
    '<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    "</fills>"
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    '<xf numFmtId="9" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
)


def more_styles(formats):
    """The style sheet with more number formats, for workbooks that need them: `formats` is [(style name, one of
    Excel's own formats by its number, or a format code)], each with {"unlocked": True} after it for a style whose
    cells a protected sheet lets change. Returns the XML and every style's index by name."""
    styles, xfs, codes = dict(STYLES), [], []
    for style, number_format, *extra in formats:
        if not isinstance(number_format, int):
            code = escape(number_format, {'"': "&quot;"})
            codes.append(f'<numFmt numFmtId="{166 + len(codes)}" formatCode="{code}"/>')
            number_format = 165 + len(codes)
        styles[style] = len(STYLES) + len(xfs)
        xf = f'<xf numFmtId="{number_format}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"'
        unlocked = extra and extra[0].get("unlocked")
        xfs.append(xf + (' applyProtection="1"><protection locked="0"/></xf>' if unlocked else "/>"))
    xml = (
        STYLES_XML.replace('<cellXfs count="5">', f'<cellXfs count="{5 + len(xfs)}">')
        .replace("</cellXfs>", "".join(xfs) + "</cellXfs>")
        .replace('<numFmts count="2">', f'<numFmts count="{2 + len(codes)}">')
        .replace("</numFmts>", "".join(codes) + "</numFmts>")
    )
    return xml, styles


def build_workbook(name, spec=None, folder=None):
    """The .xlsm for example `name`: its sheets and buttons from `spec`, its VBA from the .bas files in its folder.
    A spec may bring more number formats than the examples use, as "styles": see more_styles; and defined names, as
    "names": (name, what it refers to as the file keeps it, whether it is one of the hidden names Excel adds for the
    parameters of a LAMBDA)."""
    spec = spec or EXAMPLES[name]
    folder = folder or HERE / name
    sheets = spec["sheets"]
    styles_xml, styles = more_styles(spec["styles"]) if spec.get("styles") else (STYLES_XML, STYLES)
    modules = [("ThisWorkbook", "document", document_module("ThisWorkbook", WORKBOOK_BASE))]
    modules += [(f"Sheet{i}", "document", document_module(f"Sheet{i}", SHEET_BASE)) for i in range(1, len(sheets) + 1)]
    for bas in sorted(folder.glob("*.bas")):
        source = bas.read_text(encoding="utf-8").replace("\r\n", "\n").replace("\n", "\r\n")
        modules.append((bas.stem, "standard", source))
    project_id = f"{{{sum(name.encode()):08X}-0000-4000-8000-{len(name):012X}}}"

    strings, parts = {}, {}
    overrides = [
        ("/xl/workbook.xml", "application/vnd.ms-excel.sheet.macroEnabled.main+xml"),
        ("/xl/styles.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"),
        ("/xl/sharedStrings.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"),
    ]
    workbook_rels = []
    for i, (_name, columns, rows, buttons, *options) in enumerate(sheets, 1):
        protection = options[0].get("protection", "") if options else ""
        parts[f"xl/worksheets/sheet{i}.xml"] = _sheet_xml(i, columns, rows, strings, bool(buttons), styles, protection)
        overrides.append(
            (
                f"/xl/worksheets/sheet{i}.xml",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml",
            )
        )
        workbook_rels.append((f"rId{i}", f"{REL}/worksheet", f"worksheets/sheet{i}.xml"))
        if buttons:
            parts[f"xl/drawings/drawing{i}.xml"] = _drawing_xml(buttons)
            parts[f"xl/worksheets/_rels/sheet{i}.xml.rels"] = _rels(
                [("rId1", f"{REL}/drawing", f"../drawings/drawing{i}.xml")]
            )
            overrides.append(
                (f"/xl/drawings/drawing{i}.xml", "application/vnd.openxmlformats-officedocument.drawing+xml")
            )
    n = len(sheets)
    workbook_rels += [
        (f"rId{n + 1}", f"{REL}/styles", "styles.xml"),
        (f"rId{n + 2}", f"{REL}/sharedStrings", "sharedStrings.xml"),
        (f"rId{n + 3}", "http://schemas.microsoft.com/office/2006/relationships/vbaProject", "vbaProject.bin"),
    ]
    sheet_list = "".join(
        f'<sheet name="{escape(s[0])}" sheetId="{i}"{_state(s)} r:id="rId{i}"/>' for i, s in enumerate(sheets, 1)
    )
    hidden_name = ' hidden="1" xlm="1"'
    names = "".join(
        f'<definedName name="{escape(n)}"{hidden_name if hidden else ""}>{escape(text)}</definedName>'
        for n, text, hidden in spec.get("names", ())
    )
    parts["xl/workbook.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="{MAIN}" xmlns:r="{REL}">'
        '<workbookPr codeName="ThisWorkbook"/><bookViews><workbookView/></bookViews>'
        f"<sheets>{sheet_list}</sheets>{f'<definedNames>{names}</definedNames>' if names else ''}"
        '<calcPr calcId="191029"/></workbook>'
    )
    parts["xl/_rels/workbook.xml.rels"] = _rels(workbook_rels)
    parts["xl/styles.xml"] = styles_xml
    shared = "".join(f'<si><t xml:space="preserve">{escape(s)}</t></si>' for s in strings)
    parts["xl/sharedStrings.xml"] = (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="{MAIN}" '
        f'count="{len(strings)}" uniqueCount="{len(strings)}">{shared}</sst>'
    )
    parts["xl/vbaProject.bin"] = vba_project(modules, project_id)
    parts["_rels/.rels"] = _rels([("rId1", f"{REL}/officeDocument", "xl/workbook.xml")])
    types = "".join(f'<Override PartName="{p}" ContentType="{t}"/>' for p, t in overrides)
    parts["[Content_Types].xml"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        f'<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>{types}</Types>'
    )
    return parts


def _state(sheet):
    """A sheet's state attribute: a sheet's spec may end with {"state": "hidden"} or {"state": "veryHidden"}."""
    state = sheet[4].get("state") if len(sheet) > 4 else None
    return f' state="{state}"' if state else ""


def _rels(rels):
    items = "".join(f'<Relationship Id="{i}" Type="{t}" Target="{target}"/>' for i, t, target in rels)
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="{PKG_REL}">{items}</Relationships>'


def write_workbook(path, parts):
    order = ["[Content_Types].xml", "_rels/.rels"] + sorted(
        p for p in parts if p not in ("[Content_Types].xml", "_rels/.rels")
    )
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for name in order:
            data = parts[name]
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data.encode("utf-8") if isinstance(data, str) else data)


def main(argv=None):
    names = (argv if argv is not None else sys.argv[1:]) or list(EXAMPLES)
    for name in names:
        path = HERE / name / f"{name}.xlsm"
        write_workbook(path, build_workbook(name, EXAMPLES[name]))
        print(f"wrote {path.relative_to(HERE.parent)}")


if __name__ == "__main__":
    main()
