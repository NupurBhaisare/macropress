"""MacroPress's graphics: the README's animated hero, a one-page flyer, and the GitHub social preview card.

    python docs/assets/make_graphics.py

writes hero.svg, flyer.svg and social-preview.svg beside this script, from the components below. The hero animates
with CSS, and shows its last frame when animation is off. To turn the flyer and the card into PNGs, any browser can
render them, for example headless Chrome:

    chrome --headless --hide-scrollbars --window-size=1280,640 --screenshot=social-preview.png social-preview.svg
"""

import sys
from html import escape
from pathlib import Path

OUT = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).resolve().parent

C = dict(
    bg="#0A1020",
    bg2="#131D3A",
    panel="#121C36",
    line="#27375F",
    grid="#223154",
    head="#1A2748",
    text="#EEF3FC",
    muted="#97A5C4",
    faint="#5D6C90",
    amber="#FFB547",
    amber2="#FF8A3D",
    deep="#A85715",
    mint="#3DDC97",
    kw="#82AAFF",
    fn="#C792EA",
    str="#FFCB6B",
    com="#7282A8",
)
SANS = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Inter, Roboto, Helvetica, Arial, sans-serif"
MONO = "ui-monospace, SFMono-Regular, 'SF Mono', Menlo, Consolas, 'Liberation Mono', monospace"

ITEMS = [
    ("Notebook", "3", "4.50", "13.50"),
    ("Pens", "2", "7.25", "14.50"),
    ("Stapler", "1", "12.00", "12.00"),
    ("Paper", "5", "6.40", "32.00"),
    ("Desk lamp", "1", "24.99", "24.99"),
]
COLS = [("Item", 92, "start"), ("Qty", 36, "end"), ("Price", 56, "end"), ("Amount", 66, "end")]


def defs():
    return f"""<defs>
  <linearGradient id="bg" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="{C["bg"]}"/><stop offset="1" stop-color="{C["bg2"]}"/></linearGradient>
  <linearGradient id="amber" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="{C["amber"]}"/><stop offset="1" stop-color="{C["amber2"]}"/></linearGradient>
  <radialGradient id="glow" cx="0.5" cy="0.5" r="0.5"><stop offset="0" stop-color="{C["amber"]}" stop-opacity="0.55"/><stop offset="1" stop-color="{C["amber"]}" stop-opacity="0"/></radialGradient>
  <radialGradient id="mintglow" cx="0.5" cy="0.5" r="0.5"><stop offset="0" stop-color="{C["mint"]}" stop-opacity="0.35"/><stop offset="1" stop-color="{C["mint"]}" stop-opacity="0"/></radialGradient>
  <pattern id="dots" width="22" height="22" patternUnits="userSpaceOnUse"><circle cx="2" cy="2" r="1.1" fill="#FFFFFF" fill-opacity="0.05"/></pattern>
  <marker id="head" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="7" markerHeight="7" orient="auto-start-reverse"><path d="M1 1 L9 5 L1 9 Z" fill="{C["mint"]}"/></marker>
</defs>"""


def background(w, h, r=22):
    return (
        f'<rect width="{w}" height="{h}" rx="{r}" fill="url(#bg)"/>'
        f'<rect width="{w}" height="{h}" rx="{r}" fill="url(#dots)"/>'
    )


def logo(x, y, s):
    """A spreadsheet key, ready to be pressed."""
    r, face_h, depth = s * 0.24, s * 0.86, s * 0.12
    g = [
        f'<rect x="{x}" y="{y + depth}" width="{s}" height="{face_h}" rx="{r}" fill="{C["deep"]}"/>',
        f'<rect x="{x}" y="{y}" width="{s}" height="{face_h}" rx="{r}" fill="url(#amber)"/>',
    ]
    ix, iy, iw, ih = x + s * 0.24, y + s * 0.22, s * 0.52, s * 0.42
    sw = max(s * 0.05, 1.4)
    g.append(
        f'<rect x="{ix}" y="{iy}" width="{iw}" height="{ih}" rx="{s * 0.05}" fill="none" stroke="#FFFFFF" stroke-width="{sw}"/>'
    )
    for k in (1, 2):
        g.append(
            f'<line x1="{ix + iw * k / 3}" y1="{iy}" x2="{ix + iw * k / 3}" y2="{iy + ih}" stroke="#FFFFFF" stroke-width="{sw}"/>'
        )
    g.append(
        f'<line x1="{ix}" y1="{iy + ih / 2}" x2="{ix + iw}" y2="{iy + ih / 2}" stroke="#FFFFFF" stroke-width="{sw}"/>'
    )
    return "".join(g)


def wordmark(x, y, size):
    return (
        f'<text x="{x}" y="{y}" font-family="{SANS}" font-size="{size}" font-weight="800" letter-spacing="-0.5" '
        f'fill="{C["text"]}">Macro<tspan fill="{C["amber"]}">Press</tspan></text>'
    )


def panel(x, y, w, h, title, chip=None, chip_color=None):
    out = [
        f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="14" fill="{C["panel"]}" stroke="{C["line"]}"/>',
        f'<path d="M{x} {y + 36} H{x + w}" stroke="{C["line"]}"/>',
        f'<circle cx="{x + 16}" cy="{y + 18}" r="4" fill="#FF6B6B" fill-opacity="0.8"/>',
        f'<circle cx="{x + 29}" cy="{y + 18}" r="4" fill="{C["amber"]}" fill-opacity="0.8"/>',
        f'<circle cx="{x + 42}" cy="{y + 18}" r="4" fill="{C["mint"]}" fill-opacity="0.8"/>',
        f'<text x="{x + 56}" y="{y + 23}" font-family="{MONO}" font-size="12.5" fill="{C["muted"]}">{escape(title)}</text>',
    ]
    if chip:
        cw = 9 + len(chip) * 6.6
        out.append(
            f'<rect x="{x + w - cw - 12}" y="{y + 9}" width="{cw}" height="18" rx="9" fill="{chip_color}" fill-opacity="0.16"/>'
            f'<text x="{x + w - cw / 2 - 12}" y="{y + 22}" text-anchor="middle" font-family="{SANS}" font-size="11" '
            f'font-weight="700" fill="{chip_color}">{escape(chip)}</text>'
        )
    return "".join(out)


def sheet(x, y, amounts, row_h=24, cls_prefix=None):
    """The invoice grid: a header and five items; amounts shown when `amounts` is true."""
    w = sum(c[1] for c in COLS)
    out = [f'<rect x="{x}" y="{y}" width="{w}" height="{row_h}" fill="{C["head"]}"/>']
    for r in range(len(ITEMS) + 2):
        out.append(f'<path d="M{x} {y + r * row_h} H{x + w}" stroke="{C["grid"]}"/>')
    cx = x
    for _, cw, _ in COLS:
        out.append(f'<path d="M{cx} {y} V{y + (len(ITEMS) + 1) * row_h}" stroke="{C["grid"]}"/>')
        cx += cw
    out.append(f'<path d="M{cx} {y} V{y + (len(ITEMS) + 1) * row_h}" stroke="{C["grid"]}"/>')

    def cell_text(col, row, text, color, weight="400", cls=None):
        left = x + sum(c[1] for c in COLS[:col])
        _, cw, anchor = COLS[col]
        tx = left + 7 if anchor == "start" else left + cw - 7
        ty = y + row * row_h + row_h / 2 + 4.5
        klass = f' class="{cls}"' if cls else ""
        return (
            f'<text{klass} x="{tx}" y="{ty}" text-anchor="{anchor}" font-family="{SANS}" font-size="12.5" '
            f'font-weight="{weight}" fill="{color}">{escape(text)}</text>'
        )

    for i, (name, _, _) in enumerate(COLS):
        out.append(cell_text(i, 0, name, C["muted"], "700"))
    for r, (item, qty, price, amount) in enumerate(ITEMS, 1):
        out += [cell_text(0, r, item, C["text"]), cell_text(1, r, qty, C["text"]), cell_text(2, r, price, C["text"])]
        if amounts:
            left = x + sum(c[1] for c in COLS[:3])
            klass = f"{cls_prefix}{r}" if cls_prefix else None
            bg_class = f' class="an tb {klass}-bg"' if klass else ""
            out.append(
                f'<rect{bg_class} x="{left + 1}" y="{y + r * row_h + 1}" width="{COLS[3][1] - 2}" height="{row_h - 2}" '
                f'fill="{C["mint"]}" fill-opacity="0.13"/>'
            )
            out.append(cell_text(3, r, amount, C["mint"], "700", f"an tb {klass}" if klass else None))
    return "".join(out), w, (len(ITEMS) + 1) * row_h


def button(x, y, w, h, label, face_class=""):
    return (
        f'<rect x="{x}" y="{y + 5}" width="{w}" height="{h}" rx="11" fill="{C["deep"]}"/>'
        f'<g class="{face_class}"><rect x="{x}" y="{y}" width="{w}" height="{h}" rx="11" fill="url(#amber)"/>'
        f'<text x="{x + w / 2}" y="{y + h / 2 + 5.5}" text-anchor="middle" font-family="{SANS}" font-size="15.5" '
        f'font-weight="800" fill="#2B1600">{escape(label)}</text></g>'
    )


def cursor(x, y, s=1.0):
    p = "M0 0 L0 22 L6 16.5 L10.5 26 L14.2 24.3 L9.8 15 L17.5 15 Z"
    return (
        f'<path transform="translate({x} {y}) scale({s})" d="{p}" fill="#FFFFFF" stroke="#0A1020" '
        f'stroke-width="1.6" stroke-linejoin="round"/>'
    )


def pill(cx, y, text, h=34):
    w = 46 + len(text) * 7.25
    x = cx - w / 2
    check = f"M{x + 15} {y + h / 2} l4.5 4.5 l8 -9"
    return (
        f'<rect x="{x}" y="{y}" width="{w}" height="{h}" rx="{h / 2}" fill="#FFFFFF" fill-opacity="0.05" stroke="{C["line"]}"/>'
        f'<path d="{check}" fill="none" stroke="{C["mint"]}" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>'
        f'<text x="{x + 34}" y="{y + h / 2 + 5}" font-family="{SANS}" font-size="13.5" font-weight="600" fill="{C["text"]}">{escape(text)}</text>'
    )


def code_line(x, y, parts, cls=""):
    spans = "".join(f'<tspan fill="{C[k]}">{escape(t)}</tspan>' for k, t in parts)
    klass = f' class="{cls}"' if cls else ""
    return f'<text{klass} x="{x}" y="{y}" font-family="{MONO}" font-size="12" xml:space="preserve">{spans}</text>'


VBA = [
    [("kw", "For"), ("text", " r = "), ("str", "5"), ("kw", " To"), ("text", " lastRow")],
    [("text", "  ws.Cells(r, "), ("str", '"D"'), ("text", ").Formula = …")],
    [("kw", "Next"), ("text", " r")],
]
PY = [
    [("com", "# L13: For r = 5 To lastRow")],
    [("text", "end1, step1 = "), ("fn", "rt.for_init"), ("text", "(…)")],
    [("text", "v.ws.Cells(v.r, "), ("str", "'D'"), ("text", ").Formula = …")],
]


def copy_panel(cx_, py0, cw_, ph, prefix="amt"):
    """The updated copy: the amounts filled in, the total, and the report's verdict."""
    s = []
    s.append(panel(cx_, py0, cw_, ph, "invoice-done.xlsm", "copy", C["mint"]))
    gx = cx_ + (cw_ - 250) / 2
    grid, gw, gh = sheet(gx, py0 + 50, amounts=True, cls_prefix=prefix)
    s.append(grid)
    ty = py0 + 50 + gh + 8
    s.append(
        f'<g class="an tb total"><rect x="{gx}" y="{ty}" width="{gw}" height="26" rx="6" fill="{C["mint"]}" fill-opacity="0.16"/>'
        f'<text x="{gx + 8}" y="{ty + 17.5}" font-family="{SANS}" font-size="12.5" font-weight="800" fill="{C["text"]}">Total, with 8% tax</text>'
        f'<text x="{gx + gw - 7}" y="{ty + 17.5}" text-anchor="end" font-family="{SANS}" font-size="13.5" font-weight="800" fill="{C["mint"]}">104.75</text></g>'
    )
    ry = ty + 36
    s.append(
        f'<g class="an tb report"><rect x="{gx}" y="{ry}" width="{gw}" height="30" rx="8" fill="#000000" fill-opacity="0.28" stroke="{C["line"]}"/>'
        f'<path d="M{gx + 12} {ry + 15} l4 4 l7.5 -8.5" fill="none" stroke="{C["mint"]}" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/>'
        + code_line(gx + 30, ry + 19.5, [("mint", "completed"), ("muted", "  ·  "), ("text", "11 cells changed")])
        + "</g>"
    )
    return s


def stage(py0, ph):
    """The story in three panels, at x 40 to 960: the workbook and its button, the VBA converted, the copy."""
    s = []
    # step captions
    panels = [(40, 270), (365, 270), (690, 270)]
    for (px, _pw), (n, label) in zip(
        panels, [("1", "Your workbook"), ("2", "Its VBA, converted to Python"), ("3", "A copy, updated")], strict=False
    ):
        s.append(
            f'<circle cx="{px + 10}" cy="{py0 - 17}" r="10" fill="{C["amber"]}" fill-opacity="0.16"/>'
            f'<text x="{px + 10}" y="{py0 - 12.5}" text-anchor="middle" font-family="{SANS}" font-size="12.5" '
            f'font-weight="800" fill="{C["amber"]}">{n}</text>'
            f'<text x="{px + 28}" y="{py0 - 12.5}" font-family="{SANS}" font-size="13.5" font-weight="700" '
            f'fill="{C["text"]}">{label}</text>'
        )

    # 1: the workbook and its button
    ax, aw = panels[0]
    s.append(panel(ax, py0, aw, ph, "invoice.xlsm", "VBA inside", C["amber"]))
    grid, gw, gh = sheet(ax + (aw - 250) / 2, py0 + 50, amounts=False)
    s.append(grid)
    bx, by, bw, bh = ax + 45, py0 + 206, 180, 38
    s.append(
        f'<ellipse class="an tb glow" cx="{bx + bw / 2}" cy="{by + bh / 2}" rx="125" ry="46" fill="url(#glow)" opacity="0"/>'
    )
    s.append(button(bx, by, bw, bh, "Calculate totals", "an btn"))
    click_x, click_y = bx + bw * 0.72, by + bh * 0.62
    s.append(
        f'<circle class="an tb ripple" cx="{click_x}" cy="{click_y}" r="16" fill="none" stroke="{C["amber"]}" '
        f'stroke-width="2.5" opacity="0"/>'
    )
    s.append(f'<g class="an cursor"><g class="an tb click">{cursor(click_x - 1, click_y - 1, 1.05)}</g></g>')

    # arrows between panels
    def flow(x1, x2, y, cls):
        return (
            f'<path class="an {cls}" d="M{x1} {y} H{x2}" stroke="{C["mint"]}" stroke-width="2.4" fill="none" '
            f'stroke-linecap="round" stroke-dasharray="60" marker-end="url(#head)"/>'
            f'<circle class="an tb {cls}-dot" cx="{x1}" cy="{y}" r="5" fill="{C["mint"]}" opacity="0"/>'
        )

    s.append(flow(ax + aw + 12, panels[1][0] - 12, py0 + ph / 2, "flow1"))
    s.append(flow(panels[1][0] + panels[1][1] + 12, panels[2][0] - 12, py0 + ph / 2, "flow2"))

    # 2: VBA, then Python
    bx2, bw2 = panels[1]
    s.append(panel(bx2, py0, bw2, ph, "Invoice.bas  →  .py"))
    tx = bx2 + 16
    s.append(
        f'<text x="{tx}" y="{py0 + 60}" font-family="{SANS}" font-size="10.5" font-weight="800" letter-spacing="1.4" fill="{C["faint"]}">VBA</text>'
    )
    for i, parts in enumerate(VBA):
        s.append(code_line(tx, py0 + 82 + i * 19, parts, f"an vba{i}"))
    sep_y = py0 + 150
    s.append(
        f'<g class="an conv"><path d="M{tx} {sep_y} H{bx2 + bw2 - 16}" stroke="{C["line"]}" stroke-dasharray="3 4"/>'
        f'<rect x="{bx2 + bw2 / 2 - 52}" y="{sep_y - 11}" width="104" height="22" rx="11" fill="{C["panel"]}" stroke="{C["mint"]}" stroke-opacity="0.6"/>'
        f'<text x="{bx2 + bw2 / 2}" y="{sep_y + 4}" text-anchor="middle" font-family="{SANS}" font-size="11" font-weight="800" fill="{C["mint"]}">↓  converted</text></g>'
    )
    s.append(
        f'<text x="{tx}" y="{py0 + 182}" font-family="{SANS}" font-size="10.5" font-weight="800" letter-spacing="1.4" fill="{C["faint"]}">PYTHON</text>'
    )
    for i, parts in enumerate(PY):
        s.append(code_line(tx, py0 + 204 + i * 19, parts, f"an py{i}"))
    s.append(
        f'<text class="an py2" x="{tx}" y="{py0 + ph - 18}" font-family="{SANS}" font-size="11" font-style="italic" '
        f'fill="{C["faint"]}">each VBA line kept as a comment</text>'
    )

    s += copy_panel(panels[2][0], py0, panels[2][1], ph)
    return s


# ---------------------------------------------------------------------------------------------------------------------
# The animated hero (README)
# ---------------------------------------------------------------------------------------------------------------------


def hero():
    W, H, T = 1000, 540, 12.0
    css = ["text { font-kerning: normal; } .tb { transform-box: fill-box; transform-origin: 50% 50%; }"]

    def kf(cls, steps):
        name = "k_" + cls.replace("-", "_")
        frames = " ".join(f"{t / T * 100:.3f}% {{ {props} }}" for t, props in steps)
        css.append(f"@keyframes {name} {{ {frames} }}")
        css.append(f".{cls} {{ animation: {name} {T}s cubic-bezier(.22,.8,.26,1) infinite both; }}")

    # The loop starts and ends on the finished picture, so a still of any first frame (a link preview, a GIF's
    # thumbnail) shows the whole story: it clears at CLEAR, plays, and holds the finished picture to the end.
    CLEAR = (0.3, 0.75)

    def appear(cls, t, d=0.35, extra_from="", extra_to=""):
        shown, hidden = f"opacity:1;{extra_to}", f"opacity:0;{extra_from}"
        kf(cls, [(0, shown), (CLEAR[0], shown), (CLEAR[1], hidden), (t, hidden), (t + d, shown), (T, shown)])

    s = [
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {W} {H}" width="{W}" height="{H}" role="img" '
        "aria-label=\"MacroPress: a cursor presses Calculate totals in invoice.xlsm; the workbook's VBA is converted to "
        'Python; a copy of the workbook comes back with the amounts filled in and a report: completed, 11 cells changed.">',
        "<title>MacroPress: press the buttons in Excel workbooks, without Excel</title>",
        defs(),
        "<style>__CSS__</style>",
        background(W, H),
    ]

    # header
    s.append(logo(40, 30, 40))
    s.append(wordmark(94, 62, 30))
    s.append(
        f'<text x="{W - 40}" y="60" text-anchor="end" font-family="{SANS}" font-size="17" font-weight="600" '
        f'fill="{C["muted"]}">Press the buttons in Excel workbooks. <tspan fill="{C["mint"]}">Without Excel.</tspan></text>'
    )

    py0, ph = 122, 280
    s += stage(py0, ph)

    # bottom: three promises and the command
    py_ = py0 + ph + 26
    for cx, text in zip(
        (205, 500, 795),
        ("Runs without Excel or Windows", "Stops instead of guessing", "Never touches the original"),
        strict=False,
    ):
        s.append(pill(cx, py_, text))
    cy_ = py_ + 50
    cmd_w = 610
    s.append(
        f'<rect x="{(W - cmd_w) / 2}" y="{cy_}" width="{cmd_w}" height="34" rx="9" fill="#000000" fill-opacity="0.32" stroke="{C["line"]}"/>'
    )
    s.append(
        code_line(
            (W - cmd_w) / 2 + 16,
            cy_ + 22,
            [
                ("mint", "$ "),
                ("text", "macropress press invoice.xlsm "),
                ("kw", "--button "),
                ("str", '"Calculate totals"'),
                ("kw", " --out "),
                ("text", "done.xlsm"),
            ],
        )
    )
    s.append("</svg>")

    # the timeline: the cursor comes in, presses at 2.4s, and the story follows
    away, home = "transform:translate(150px,120px)", "transform:translate(0px,0px)"
    kf(
        "cursor",
        [
            (0, f"opacity:1;{home}"),
            (CLEAR[0], f"opacity:1;{home}"),
            (CLEAR[1], f"opacity:0;{home}"),
            (CLEAR[1] + 0.05, f"opacity:0;{away}"),
            (1.2, f"opacity:1;{away}"),
            (2.25, f"opacity:1;{home}"),
            (T, f"opacity:1;{home}"),
        ],
    )
    kf("click", [(0, "transform:scale(1)"), (2.28, "transform:scale(1)"), (2.4, "transform:scale(0.84)"),
                 (2.6, "transform:scale(1)"), (T, "transform:scale(1)")])  # fmt: skip
    kf("btn", [(0, "transform:translateY(0px)"), (2.3, "transform:translateY(0px)"), (2.4, "transform:translateY(4px)"),
               (2.68, "transform:translateY(0px)"), (T, "transform:translateY(0px)")])  # fmt: skip
    kf("ripple", [(0, "opacity:0;transform:scale(0.3)"), (2.38, "opacity:0;transform:scale(0.3)"),
                  (2.43, "opacity:0.9;transform:scale(0.5)"), (3.15, "opacity:0;transform:scale(3.2)"),
                  (T, "opacity:0;transform:scale(3.2)")])  # fmt: skip
    kf("glow", [(0, "opacity:0"), (2.38, "opacity:0"), (2.58, "opacity:1"), (3.6, "opacity:0"), (T, "opacity:0")])
    for cls, t in (("flow1", 2.95), ("flow2", 5.9)):
        drawn, undrawn = "stroke-dashoffset:0", "stroke-dashoffset:60"
        kf(
            cls,
            [
                (0, f"{drawn};opacity:1"),
                (CLEAR[0], f"{drawn};opacity:1"),
                (CLEAR[1], f"{drawn};opacity:0"),
                (CLEAR[1] + 0.05, f"{undrawn};opacity:0"),
                (t, f"{undrawn};opacity:0"),
                (t + 0.05, f"{undrawn};opacity:1"),
                (t + 0.45, f"{drawn};opacity:1"),
                (T, f"{drawn};opacity:1"),
            ],
        )
        kf(
            f"{cls}-dot",
            [
                (0, "opacity:0;transform:translateX(0px)"),
                (t, "opacity:0;transform:translateX(0px)"),
                (t + 0.08, "opacity:1;transform:translateX(4px)"),
                (t + 0.5, "opacity:0;transform:translateX(34px)"),
                (T, "opacity:0;transform:translateX(34px)"),
            ],
        )
    slide = {"extra_from": "transform:translateX(-8px)", "extra_to": "transform:translateX(0px)"}
    for i in range(3):
        appear(f"vba{i}", 3.4 + i * 0.32, **slide)
    appear("conv", 4.5, d=0.4)
    for i in range(3):
        appear(f"py{i}", 4.9 + i * 0.32, **slide)
    for r in range(1, 6):
        t = 6.4 + (r - 1) * 0.3
        appear(f"amt{r}", t, d=0.3, extra_from="transform:scale(0.7)", extra_to="transform:scale(1)")
        appear(f"amt{r}-bg", t, d=0.15)
    appear("total", 8.1, d=0.4, extra_from="transform:scale(0.92)", extra_to="transform:scale(1)")
    appear("report", 8.75, d=0.45, extra_from="transform:translateY(8px)", extra_to="transform:translateY(0px)")
    css.append("@media (prefers-reduced-motion: reduce) { .an { animation: none !important; } }")
    return "\n".join(s).replace("__CSS__", "\n".join(css))


# ---------------------------------------------------------------------------------------------------------------------
# The flyer (a one-page explainer, 1080 x 1350) and the social preview card (1280 x 640): static
# ---------------------------------------------------------------------------------------------------------------------


def svg_open(w, h, label):
    return [
        f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {w} {h}" width="{w}" height="{h}" role="img" '
        f'aria-label="{escape(label)}">',
        f"<title>{escape(label)}</title>",
        defs(),
        background(w, h, 0),
    ]


def centered_brand(cx, y, logo_size, text_size):
    width = logo_size + 16 + text_size * 5.35  # the wordmark is about 5.35 times its font size wide
    x = cx - width / 2
    return logo(x, y, logo_size) + wordmark(x + logo_size + 16, y + logo_size * 0.78, text_size)


def text_c(x, y, text, size, color, weight="400", anchor="middle", family=SANS, extra=""):
    return (
        f'<text x="{x}" y="{y}" text-anchor="{anchor}" font-family="{family}" font-size="{size}" '
        f'font-weight="{weight}" fill="{color}"{extra}>{text}</text>'
    )


def flyer():
    W, H = 1080, 1350
    s = svg_open(
        W,
        H,
        "MacroPress: press the buttons in Excel workbooks, without Excel. How it works, why it is safe, "
        "and how to try it.",
    )
    s.append(centered_brand(W / 2, 84, 62, 54))
    s.append(
        text_c(
            W / 2, 236, "Press the buttons in Excel workbooks.", 44, C["text"], "800", extra=' letter-spacing="-0.6"'
        )
    )
    s.append(text_c(W / 2, 292, "Without Excel.", 44, C["mint"], "800", extra=' letter-spacing="-0.6"'))
    s.append(
        text_c(
            W / 2, 342, "MacroPress converts a workbook's VBA macros to Python and runs them on a copy,", 20, C["muted"]
        )
    )
    s.append(text_c(W / 2, 370, "on any machine with Python. No Excel, no Windows, no licence.", 20, C["muted"]))
    s.append('<g transform="translate(40 300)">' + "".join(stage(122, 280)) + "</g>")

    s.append(text_c(80, 770, "How it works", 26, C["text"], "800", "start"))
    steps = [
        ("Finds the buttons", "Form-control buttons and shapes with a", "macro, by the caption they show."),
        ("Converts the VBA", "Readable Python, one function per", "procedure, each VBA line kept as a comment."),
        (
            "Runs it on a model of Excel",
            "Cells, formulas, sorting, filters, copy and",
            "paste, and VBA's own rules for numbers.",
        ),
        ("Saves a copy and a report", "Only the changed sheets are rewritten, and", "a JSON report says what changed."),
    ]
    for i, (title, l1, l2) in enumerate(steps):
        x, y = 80 + (i % 2) * 470, 795 + (i // 2) * 136
        s.append(
            f'<rect x="{x}" y="{y}" width="450" height="118" rx="16" fill="{C["panel"]}" stroke="{C["line"]}"/>'
            f'<circle cx="{x + 34}" cy="{y + 38}" r="16" fill="{C["amber"]}" fill-opacity="0.16"/>'
            + text_c(x + 34, y + 44.5, str(i + 1), 17, C["amber"], "800")
            + text_c(x + 62, y + 45, title, 21, C["text"], "800", "start")
            + text_c(x + 62, y + 76, l1, 16, C["muted"], anchor="start")
            + text_c(x + 62, y + 98, l2, 16, C["muted"], anchor="start")
        )
    for cx, text in zip(
        (215, 540, 865),
        ("Runs without Excel or Windows", "Stops instead of guessing", "Never touches the original"),
        strict=False,
    ):
        s.append(pill(cx, 1080, text))
    tx, ty, tw = 80, 1140, 920
    s.append(
        f'<rect x="{tx}" y="{ty}" width="{tw}" height="126" rx="14" fill="#000000" fill-opacity="0.32" stroke="{C["line"]}"/>'
    )
    lines = [
        [("mint", "$ "), ("text", "pip install git+https://github.com/NupurBhaisare/macropress")],
        [("mint", "$ "), ("text", "macropress buttons invoice.xlsm")],
        [
            ("mint", "$ "),
            ("text", "macropress press invoice.xlsm "),
            ("kw", "--button "),
            ("str", '"Calculate totals"'),
            ("kw", " --out "),
            ("text", "done.xlsm"),
        ],
    ]
    for i, parts in enumerate(lines):
        spans = "".join(f'<tspan fill="{C[k]}">{escape(t)}</tspan>' for k, t in parts)
        s.append(
            f'<text x="{tx + 26}" y="{ty + 40 + i * 33}" font-family="{MONO}" font-size="17" xml:space="preserve">{spans}</text>'
        )
    s.append(
        text_c(
            W / 2,
            1312,
            "Open source  ·  MIT licence  ·  Python 3.11  ·  Linux, macOS and Windows",
            16,
            C["faint"],
            "600",
        )
    )
    s.append("</svg>")
    return "\n".join(s)


def social():
    W, H = 1280, 640
    s = svg_open(W, H, "MacroPress: press the buttons in Excel workbooks, without Excel.")
    s.append(logo(80, 92, 64))
    s.append(wordmark(162, 142, 54))
    for i, (text, color) in enumerate(
        (("Press the buttons in", C["text"]), ("Excel workbooks.", C["text"]), ("Without Excel.", C["mint"]))
    ):
        s.append(text_c(80, 262 + i * 58, text, 50, color, "800", "start", extra=' letter-spacing="-0.8"'))
    s.append(
        text_c(
            80, 452, "VBA macros converted to Python and run on a copy of the workbook,", 20, C["muted"], anchor="start"
        )
    )
    s.append(
        text_c(
            80,
            480,
            "anywhere Python runs. When it can't match Excel, it stops and says so.",
            20,
            C["muted"],
            anchor="start",
        )
    )
    for x, text in zip((80, 352, 556), ("No Excel or Windows", "Never guesses", "Original untouched"), strict=False):
        s.append(
            f'<path d="M{x} {540} l6 6 l11 -12" fill="none" stroke="{C["mint"]}" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/>'
            + text_c(x + 26, 546, text, 18, C["text"], "700", "start")
        )
    # the result, scaled up, with the button that made it
    s.append('<g transform="translate(812 64) scale(1.42)">' + "".join(copy_panel(0, 0, 270, 280)) + "</g>")
    s.append('<ellipse cx="873" cy="500" rx="118" ry="40" fill="url(#glow)"/>')
    s.append(button(775, 478, 196, 44, "Calculate totals"))
    s.append(cursor(922, 503, 1.25))
    s.append("</svg>")
    return "\n".join(s)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    for name, make in (("hero", hero), ("flyer", flyer), ("social-preview", social)):
        (OUT / f"{name}.svg").write_text(make(), encoding="utf-8")
        print("wrote", OUT / f"{name}.svg")


if __name__ == "__main__":
    main()
