"""Utilidades de diseño para la presentación 'Del ticket a producción'."""
import re
from lxml import etree
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_MARKER_STYLE
from pptx.enum.dml import MSO_LINE
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

W, H = 13.333, 7.5
MX = 0.7
CW = W - 2 * MX

INK = "1F2A30"
TEAL_D = "24414A"
TEAL = "2F6F73"
TEAL_L = "D5E6E4"
AMBER = "E0A100"
SLATE = "5B7B8A"
SAGE = "8FB5B0"
RUST = "C75B39"
GREEN = "6B8F4E"
LIGHT = "EAF0F1"
WHITE = "FFFFFF"
CODE_BG = "1B262B"

HEAD = "Cambria"
BODY = "Calibri"
MONO = "Consolas"

STATUS = {
    "visto": ("ya visto", GREEN, WHITE),
    "curso": ("en curso", AMBER, INK),
    "llega": ("os llega", SLATE, WHITE),
}

PHASES = ["Requisitos", "Diseño", "Datos", "Código", "Integración", "Entrega"]


def rgb(h):
    return RGBColor.from_string(h)


def new_prs():
    prs = Presentation()
    prs.slide_width = Emu(12192000)
    prs.slide_height = Emu(6858000)
    return prs


def _fill_paras(tf, paras, size, color, bold, font, align, italic=False, spacing=None):
    if isinstance(paras, (str, tuple)):
        paras = [paras]
    first = True
    for p in paras:
        para = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        after = 0
        popts = {}
        if isinstance(p, tuple):
            popts = p[1]
            after = popts.get("after", 0)
            runs = [(p[0], popts)]
        elif isinstance(p, dict):
            popts = p
            after = p.get("after", 0)
            runs = p["runs"]
        else:
            runs = p if isinstance(p, list) else [p]
        para.alignment = {"l": PP_ALIGN.LEFT, "c": PP_ALIGN.CENTER, "r": PP_ALIGN.RIGHT}[
            popts.get("align", align)
        ]
        if spacing:
            para.line_spacing = spacing
        para.space_after = Pt(after)
        for r in runs:
            if isinstance(r, str):
                r = (r, {})
            text, o = r
            run = para.add_run()
            run.text = text
            f = run.font
            f.size = Pt(o.get("size", size))
            f.bold = o.get("bold", bold)
            f.italic = o.get("italic", italic)
            f.name = o.get("font", font)
            f.color.rgb = rgb(o.get("color", color))


def tb(sl, x, y, w, h, paras, size=16, color=INK, bold=False, font=BODY, align="l",
       anchor="t", italic=False, spacing=None):
    box = sl.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    _fill_paras(tf, paras, size, color, bold, font, align, italic, spacing)
    return box


def _drop_style(s):
    st = s._element.find(qn("p:style"))
    if st is not None:
        s._element.remove(st)


def shape(sl, kind, x, y, w, h, fill=None, line=None, lw=1.0, r=None, dash=False):
    s = sl.shapes.add_shape(kind, Inches(x), Inches(y), Inches(w), Inches(h))
    s.shadow.inherit = False
    _drop_style(s)
    if fill:
        s.fill.solid()
        s.fill.fore_color.rgb = rgb(fill)
    else:
        s.fill.background()
    if line:
        s.line.color.rgb = rgb(line)
        s.line.width = Pt(lw)
        if dash:
            s.line.dash_style = MSO_LINE.DASH
    else:
        s.line.fill.background()
    if r is not None and kind == MSO_SHAPE.ROUNDED_RECTANGLE:
        s.adjustments[0] = r
    return s


def rr(sl, x, y, w, h, fill=None, line=None, lw=1.0, r=0.08, dash=False):
    return shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, y, w, h, fill, line, lw, r, dash)


def box(sl, x, y, w, h, text, fill=LIGHT, color=INK, size=14, bold=False, font=BODY,
        align="c", line=None, lw=1.0, kind=MSO_SHAPE.ROUNDED_RECTANGLE, r=0.12,
        anchor="m", dash=False, pad=0.08):
    s = shape(sl, kind, x, y, w, h, fill, line, lw, r, dash)
    if kind in (MSO_SHAPE.PENTAGON, MSO_SHAPE.CHEVRON):
        s.adjustments[0] = 0.3
        depth = 0.3 * h
        left = x + (0.12 if kind == MSO_SHAPE.PENTAGON else depth + 0.05)
        right = x + w - depth - 0.02
        tb(sl, left, y, right - left, h, text, size=size, color=color, bold=bold, font=font, align="c",
           anchor="m")
        return s
    tf = s.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = Inches(pad)
    tf.margin_top = tf.margin_bottom = Inches(0.04)
    tf.vertical_anchor = {"t": MSO_ANCHOR.TOP, "m": MSO_ANCHOR.MIDDLE, "b": MSO_ANCHOR.BOTTOM}[anchor]
    _fill_paras(tf, text, size, color, bold, font, align)
    return s


def line(sl, x1, y1, x2, y2, color=SLATE, w=1.75, arrow=False, dash=False, both=False):
    c = sl.shapes.add_connector(MSO_CONNECTOR.STRAIGHT, Inches(x1), Inches(y1), Inches(x2), Inches(y2))
    _drop_style(c)
    c.line.color.rgb = rgb(color)
    c.line.width = Pt(w)
    if dash:
        c.line.dash_style = MSO_LINE.DASH
    ln = c.line._get_or_add_ln()
    if both:
        etree.SubElement(ln, qn("a:headEnd"), type="triangle", w="med", len="med")
    if arrow:
        etree.SubElement(ln, qn("a:tailEnd"), type="triangle", w="med", len="med")
    return c


def set_bg(sl, color):
    sl.background.fill.solid()
    sl.background.fill.fore_color.rgb = rgb(color)


def add_title(sl, text, color=TEAL_D, y=0.42, size=32, w=None):
    t = sl.shapes.title
    t.left, t.top = Inches(MX), Inches(y)
    t.width, t.height = Inches(w or CW), Inches(0.8)
    tf = t.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    bp = tf._txBody.find(qn("a:bodyPr"))
    for ch in list(bp):
        bp.remove(ch)
    etree.SubElement(bp, qn("a:noAutofit"))
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.LEFT
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = True
    r.font.name = HEAD
    r.font.color.rgb = rgb(color)
    return t


def pill(sl, x, y, text, fill, color, size=11, h=0.3, bold=True, minw=0.0):
    w = max(minw, len(text) * 0.0725 * size / 10 + 0.34)
    box(sl, x, y, w, h, text, fill=fill, color=color, size=size, bold=bold, r=0.5, pad=0.05)
    return w


def chips_row(sl, items, y=1.3, x=MX, label="En el programa:", maxw=None):
    cx = x
    right = x + (maxw or CW)
    if label and items:
        w = len(label) * 0.075 + 0.1
        tb(sl, cx, y + 0.02, w, 0.28, label, size=11.5, color=SLATE, anchor="m")
        cx += w + 0.1
    first = cx
    for name, st in items:
        txt, fill, col = STATUS[st]
        w = len(f"{name} · {txt}") * 0.0725 * 1.1 + 0.34
        if cx + w > right and cx > first:
            cx = first
            y += 0.38
        pill(sl, cx, y, f"{name} · {txt}", fill, col, size=11)
        cx += w + 0.12
    return y


def slide_number(sl, n, color=SLATE):
    b = sl.shapes.add_textbox(Inches(W - MX - 0.6), Inches(7.0), Inches(0.6), Inches(0.28))
    tf = b.text_frame
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    fld = etree.SubElement(p._p, qn("a:fld"), id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}", type="slidenum")
    rpr = etree.SubElement(fld, qn("a:rPr"), lang="es-ES", sz="1000")
    sf = etree.SubElement(rpr, qn("a:solidFill"))
    etree.SubElement(sf, qn("a:srgbClr"), val=color)
    t = etree.SubElement(fld, qn("a:t"))
    t.text = str(n)


def footer(sl, n, phase=None, dark=False, closing=False):
    col = SAGE if dark else SLATE
    tb(sl, MX, 7.0, 3.3, 0.28, "XLIX DIM · Del ticket a producción", size=10, color=col, anchor="m")
    sw, gap = 1.02, 0.07
    total = len(PHASES) * sw + (len(PHASES) - 1) * gap
    x0 = (W - total) / 2 + 0.3
    for i, name in enumerate(PHASES):
        if closing:
            state = "done"
        elif phase is None:
            state = "future"
        else:
            state = "done" if i < phase else ("now" if i == phase else "future")
        if dark:
            fill, c = {"done": ("4C8E94", WHITE), "now": (AMBER, INK), "future": ("2E505A", SAGE)}[state]
        else:
            fill, c = {"done": (TEAL_L, TEAL), "now": (TEAL, WHITE), "future": (LIGHT, SLATE)}[state]
        box(sl, x0 + i * (sw + gap), 7.02, sw, 0.24, f"{i + 1} {name}", fill=fill, color=c,
            size=8.5, bold=(state == "now"), r=0.5, pad=0.02)
    slide_number(sl, n, col)


def banner(sl, text, y=6.05, h=0.7, fill=TEAL_D, color=WHITE, size=18, label=None):
    rr(sl, MX, y, CW, h, fill=fill, r=0.14)
    if label:
        tb(sl, MX + 0.3, y, 1.9, h, label.upper(), size=11, color=AMBER, bold=True, anchor="m")
        tb(sl, MX + 2.1, y, CW - 2.4, h, text, size=size, color=color, anchor="m")
    else:
        tb(sl, MX + 0.35, y, CW - 0.7, h, text, size=size, color=color, anchor="m")


def num_circle(sl, x, y, n, d=0.46, fill=TEAL, color=WHITE, size=15):
    box(sl, x, y, d, d, str(n), fill=fill, color=color, size=size, bold=True,
        kind=MSO_SHAPE.OVAL, pad=0)


def cards(sl, items, y=2.0, h=3.85, accent=TEAL, body_size=None, head_size=None, numbered=True,
          fill=LIGHT, gap=0.28):
    n = len(items)
    body_size = body_size or (20 if n <= 3 else 17.5)
    head_size = head_size or (23 if n <= 3 else 19)
    cw = (CW - (n - 1) * gap) / n
    for i, it in enumerate(items):
        head, body = it[0], it[1]
        ex = it[2] if len(it) > 2 else None
        x = MX + i * (cw + gap)
        rr(sl, x, y, cw, h, fill=fill, r=0.04)
        shape(sl, MSO_SHAPE.RECTANGLE, x, y, cw, 0.07, fill=accent)
        yy = y + 0.3
        if numbered:
            num_circle(sl, x + 0.25, yy, i + 1, fill=accent)
            tb(sl, x + 0.9, yy - 0.04, cw - 1.1, 0.55, head, size=head_size, bold=True, font=HEAD,
               color=TEAL_D, anchor="m")
        else:
            tb(sl, x + 0.25, yy - 0.04, cw - 0.5, 0.55, head, size=head_size, bold=True, font=HEAD,
               color=TEAL_D, anchor="m")
        tb(sl, x + 0.25, yy + 0.75, cw - 0.5, h - 1.6, body, size=body_size, color=INK, spacing=1.05)
        if ex:
            tb(sl, x + 0.25, y + h - 0.85, cw - 0.5, 0.7, ex, size=14, color=SLATE, italic=True, anchor="b")


KW = {
    "java": r"\b(public|private|protected|class|interface|extends|implements|return|new|void|final|static|import|package|int|long|String|List|if|else|for|while|record|enum|var|this|throws|try|catch|null|true|false)\b",
    "sql": r"(?i:\b(select|from|where|join|on|group by|order by|insert|into|values|update|set|delete|create|table|index|primary key|references|not null|default|check|and|or|desc|asc|sum|count|as|date|bigint|int|varchar|alter|add|column|begin|commit|rollback|explain|analyze|using|limit|offset|in|is|null)\b)",
    "docker": r"^(FROM|COPY|WORKDIR|EXPOSE|CMD|RUN|ENTRYPOINT|USER|ENV|ARG|AS)\b",
    "json": r"\b(true|false|null)\b",
    "yaml": r"\b(true|false|null)\b",
    "bash": r"\b(if|then|fi|do|done)\b",
}

C_KW, C_STR, C_ANN, C_COM, C_NUM, C_TYPE, C_DEF = "F0B429", "9FD8A9", "E58C74", "7F9AA5", "8FD3E8", "A8D8E8", "E6EEF0"


def _tokens(line_txt, lang):
    if lang == "text":
        return [(line_txt, C_DEF, False)]
    if lang == "diff":
        if line_txt.startswith("+"):
            return [(line_txt, "9FD8A9", False)]
        if line_txt.startswith("-"):
            return [(line_txt, "E58C74", False)]
        return [(line_txt, C_DEF, False)]
    if lang == "md":
        if line_txt.startswith("#"):
            return [(line_txt, C_KW, False)]
        if line_txt.startswith("|"):
            return [(line_txt, C_TYPE, False)]
        return [(line_txt, C_DEF, False)]
    parts = []
    comment = {"java": r"//.*", "sql": r"--.*", "yaml": r"#.*", "bash": r"#.*", "docker": r"#.*",
               "json": r"(?!x)x", "html": r"(?!x)x"}[lang]
    pats = [("com", comment), ("str", r'"(?:\\.|[^"\\])*"|\'[^\']*\'')]
    if lang == "java":
        pats += [("ann", r"@\w+"), ("kw", KW["java"]), ("typ", r"\b[A-Z]\w*\b"), ("num", r"\b\d+\b")]
    elif lang == "sql":
        pats += [("kw", KW["sql"]), ("num", r"\b\d+\b")]
    elif lang == "json":
        pats = [("key", r'"[^"]*"(?=\s*:)'), ("str", r'"(?:\\.|[^"\\])*"'), ("kw", KW["json"]),
                ("num", r"-?\b\d+(?:\.\d+)?\b")]
    elif lang == "yaml":
        pats = [("com", comment), ("key", r"[\w./-]+(?=:)"), ("str", pats[1][1]), ("kw", KW["yaml"]),
                ("num", r"\b\d+\b")]
    elif lang == "docker":
        pats += [("kw", KW["docker"]), ("num", r"\b\d+\b")]
    elif lang == "bash":
        pats += [("kw", KW["bash"]), ("flag", r"(?<=\s)-{1,2}[\w-]+"), ("cmd", r"^\s*\$?\s*\w[\w./-]*")]
    elif lang == "html":
        pats = [("tag", r"</?\w+|/?>"), ("str", pats[1][1]), ("ann", r"\b[\w-]+(?==)")]
    rx = re.compile("|".join(f"(?P<{n}>{p})" for n, p in pats), re.M)
    pos = 0
    colors = {"com": (C_COM, True), "str": (C_STR, False), "ann": (C_ANN, False), "kw": (C_KW, False),
              "typ": (C_TYPE, False), "num": (C_NUM, False), "key": (C_TYPE, False),
              "flag": (C_NUM, False), "cmd": (C_KW, False), "tag": (C_KW, False)}
    for m in rx.finditer(line_txt):
        if m.start() > pos:
            parts.append((line_txt[pos:m.start()], C_DEF, False))
        c, it = colors[m.lastgroup]
        parts.append((m.group(), c, it))
        pos = m.end()
    if pos < len(line_txt):
        parts.append((line_txt[pos:], C_DEF, False))
    return parts


def code_panel(sl, x, y, w, code, lang="java", size=12, h=None, fname=None, pad=0.25):
    lines = code.strip("\n").split("\n")
    lh = size * 1.28 / 72
    top = 0.5 if fname else 0.25
    hh = h or (top + len(lines) * lh + 0.25)
    rr(sl, x, y, w, hh, fill=CODE_BG, r=0.03)
    if fname:
        for i, c in enumerate([RUST, AMBER, GREEN]):
            shape(sl, MSO_SHAPE.OVAL, x + 0.22 + i * 0.2, y + 0.17, 0.12, 0.12, fill=c)
        tb(sl, x + 0.95, y + 0.1, w - 1.2, 0.28, fname, size=10.5, color=SAGE, font=MONO, anchor="m")
    b = sl.shapes.add_textbox(Inches(x + pad), Inches(y + top), Inches(w - 2 * pad), Inches(hh - top - 0.1))
    tf = b.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    for i, ln in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        p.line_spacing = 1.0
        toks = _tokens(ln, lang) or [("", C_DEF, False)]
        for text, col, it in toks:
            r = p.add_run()
            r.text = text
            r.font.size = Pt(size)
            r.font.name = MONO
            r.font.italic = it
            r.font.color.rgb = rgb(col)
        if not ln:
            r = p.add_run()
            r.text = " "
            r.font.size = Pt(size)
    return hh


def callouts(sl, x, y, w, items, gap=0.22, h=1.1, head=18, body=16):
    for i, (hd, tx) in enumerate(items):
        yy = y + i * (h + gap)
        rr(sl, x, yy, w, h, fill=LIGHT, r=0.08)
        shape(sl, MSO_SHAPE.RECTANGLE, x, yy + 0.12, 0.07, h - 0.24, fill=AMBER)
        tb(sl, x + 0.28, yy + 0.1, w - 0.45, 0.34, hd, size=head, bold=True, font=HEAD, color=TEAL_D)
        tb(sl, x + 0.28, yy + 0.46, w - 0.45, h - 0.5, tx, size=body, color=INK)


def actor(sl, cx, y, label, color=TEAL_D):
    shape(sl, MSO_SHAPE.OVAL, cx - 0.17, y, 0.34, 0.34, line=color, lw=2)
    line(sl, cx, y + 0.34, cx, y + 0.85, color=color, w=2)
    line(sl, cx - 0.32, y + 0.52, cx + 0.32, y + 0.52, color=color, w=2)
    line(sl, cx, y + 0.85, cx - 0.28, y + 1.2, color=color, w=2)
    line(sl, cx, y + 0.85, cx + 0.28, y + 1.2, color=color, w=2)
    tb(sl, cx - 0.85, y + 1.28, 1.7, 0.5, label, size=14, bold=True, color=color, align="c")
