#!/usr/bin/env python3
"""Genera 'Del ticket a producción' (XLIX Curso DIM). Uso: python3 build_deck.py [salida.pptx]"""
import math
import sys

from deck_lib import *  # noqa: F401,F403
from deck_lib import _fill_paras

V, C, L = "visto", "curso", "llega"

# Estado de cada asignatura el 9-nov-2026 según el calendario semanal (PROGRAMA V2 13/10/2026)
S = {
    "req": ("Ing. de Requisitos", V),
    "dir": ("Dirección de Proyectos", C),
    "herr": ("Herramientas colaborativas", C),
    "git": ("Git", C),
    "adoo": ("ADOO-UML", C),
    "pat": ("Patrones y pruebas", L),
    "bpmn": ("BPMN", L),
    "bbdd": ("TG Bases de Datos", L),
    "sql": ("SQL", L),
    "oracle": ("Oracle y PL/SQL", L),
    "nosql": ("NoSQL", L),
    "alg": ("Algoritmos", V),
    "poo": ("POO", C),
    "lib": ("Librerías", L),
    "ws": ("Servicios web", L),
    "ser": ("Serialización", L),
    "html": ("HTML/CSS", C),
    "ecma": ("ECMAScript", L),
    "front": ("Frontend y Backend", L),
    "post": ("Postman", L),
    "docker": ("Docker", L),
    "soserv": ("SO Servidor", C),
    "linux": ("Linux de usuario", V),
    "redes": ("Redes", V),
    "seg": ("Seguridad TIC", L),
    "ssoo": ("TG Sistemas Operativos", V),
}


def ch(*keys):
    return [S[k] for k in keys]


def fmt_clock(t):
    return f"{t // 60:02d}:{t % 60:02d}"


def N(decir, pregunta=None, historia=None, enlace=None, corto=None):
    out = [f"DECIR: {decir}"]
    if pregunta:
        out.append(f"PREGUNTA A LA SALA: {pregunta}")
    if historia:
        out.append(f"TU HISTORIA (sustitúyela por una real, sin datos del cliente): {historia}")
    if enlace:
        out.append(f"ENLACE CON EL CURSO: {enlace}")
    if corto:
        out.append(f"SI VAS JUSTA DE TIEMPO: {corto}")
    return "\n".join(out)


class Deck:
    def __init__(self):
        self.prs = new_prs()
        self.layout = self.prs.slide_layouts[5]
        self.clock = 8 * 60 + 10
        self.n = 0
        self.phase = None
        self.agenda = []
        self.closing = False

    def _new(self, bg=WHITE):
        sl = self.prs.slides.add_slide(self.layout)
        self.n += 1
        set_bg(sl, bg)
        return sl

    def _notes(self, sl, m, notes, extra=None):
        start = fmt_clock(self.clock)
        self.clock += m
        head = f"[{start} · {m} min]" if m else "[Anexo · no entra en el horario]"
        if extra:
            head += f" {extra}"
        sl.notes_slide.notes_text_frame.text = head + "\n" + (notes or "")

    def expect(self, hhmm):
        assert fmt_clock(self.clock) == hhmm, f"Reloj {fmt_clock(self.clock)} != {hhmm} (diapo {self.n})"

    # ---------- tipos de diapositiva ----------
    def title(self, m, notes):
        sl = self._new(TEAL_D)
        shape(sl, MSO_SHAPE.RECTANGLE, 0, 0, 0.35, H, fill=AMBER)
        tb(sl, 1.1, 1.2, 8, 0.4, "XLIX CURSO DIM · NOVIEMBRE 2026", size=15, color=AMBER, bold=True)
        add_title(sl, "Del ticket a producción", color=WHITE, y=1.75, size=60, w=11)
        sl.shapes.title.height = Inches(1.3)
        tb(sl, 1.1, 3.25, 10, 1.2,
           "Cómo se conectan las asignaturas del curso con el trabajo real", size=26, color=SAGE, font=HEAD)
        x = 1.1
        for i, name in enumerate(PHASES):
            w = 1.78
            kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
            box(sl, x + i * (w - 0.1), 4.85, w, 0.62, name, fill=TEAL if i % 2 == 0 else "2E6870",
                color=WHITE, size=13, bold=True, kind=kind, r=None, pad=0.18)
        tb(sl, 1.1, 5.95, 10, 0.4, "Elena Pascual · Ingeniera informática · Consultora Java y Release Manager",
           size=17, color=WHITE)
        slide_number(sl, self.n, SAGE)
        self._notes(sl, m, notes)

    def section(self, idx, title, subtitle, subjects, minutes, entry_q, notes, tag="BLOQUE"):
        self.phase = idx
        self.closing = idx is None
        sl = self._new(TEAL_D)
        start = fmt_clock(self.clock)
        end = fmt_clock(self.clock + minutes)
        tb(sl, MX, 0.7, 11, 0.35, (f"{tag} {idx + 1} · " if idx is not None else f"{tag} · ") + f"{start} – {end} · {minutes} min", size=15,
           color=AMBER, bold=True)
        add_title(sl, title, color=WHITE, y=1.15, size=50)
        sl.shapes.title.height = Inches(1.0)
        tb(sl, MX, 2.2, 11.5, 0.6, subtitle, size=22, color=SAGE, font=HEAD)
        chips_row(sl, subjects, y=3.0, label="En el programa:")
        rr(sl, MX, 4.0, CW, 1.45, fill="2E505A", r=0.08)
        box(sl, MX + 0.3, 4.35, 0.75, 0.75, "?", fill=AMBER, color=INK, size=30, bold=True, kind=MSO_SHAPE.OVAL, pad=0)
        tb(sl, MX + 1.35, 4.1, 3.0, 0.3, "ANTES DE EMPEZAR", size=11.5, color=AMBER, bold=True)
        tb(sl, MX + 1.35, 4.43, CW - 1.8, 0.9, entry_q, size=22, color=WHITE, anchor="t")
        sw = 2.05
        for i, name in enumerate(PHASES):
            kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
            fill, c = (AMBER, INK) if i == idx else (("3C7F85", WHITE) if (idx is None or i < idx) else ("2E505A", SAGE))
            box(sl, MX + i * (sw - 0.12), 5.85, sw, 0.62, f"{i + 1}  {name}", fill=fill, color=c, size=14,
                bold=(i == idx), kind=kind, r=None, pad=0.2)
        footer(sl, self.n, idx, dark=True, closing=self.closing)
        self._notes(sl, minutes and 1, notes, extra=f"Bloque completo: {start}–{end}")
        self.agenda.append((title, start, end))
        return sl

    def content(self, title, chips, m, notes, draw, label="En el programa:", bg=WHITE, title_size=32):
        sl = self._new(bg)
        add_title(sl, title, size=title_size)
        if chips:
            chips_row(sl, chips, y=1.28, label=label)
        draw(sl)
        footer(sl, self.n, self.phase, closing=self.closing)
        self._notes(sl, m, notes)
        return sl

    def statement(self, text, sub, m, notes, color_bg=TEAL_D):
        sl = self._new(color_bg)
        shape(sl, MSO_SHAPE.RECTANGLE, MX, 2.3, 0.1, 2.5, fill=AMBER)
        add_title(sl, text, color=WHITE, y=2.2, size=46, w=11.2)
        sl.shapes.title.left = Inches(MX + 0.5)
        sl.shapes.title.height = Inches(1.8)
        tb(sl, MX + 0.5, 4.2, 10.5, 1.0, sub, size=22, color=SAGE, font=HEAD)
        footer(sl, self.n, self.phase, dark=True, closing=self.closing)
        self._notes(sl, m, notes)

    def ask(self, kind, title, prompt, steps, m, notes, sub_label="Cómo lo hacemos", prompt_code=None):
        sl = self._new(TEAL)
        box(sl, MX, 0.85, 0.95, 0.95, "?", fill=AMBER, color=INK, size=44, bold=True, kind=MSO_SHAPE.OVAL, pad=0)
        tb(sl, MX + 1.2, 0.95, 5.5, 0.35, kind.upper(), size=14, color=AMBER, bold=True)
        tb(sl, MX + 1.2, 1.32, 5.3, 0.5, f"{m} min · participa toda la sala", size=15, color="D5E6E4")
        add_title(sl, title, color=WHITE, y=2.15, size=36, w=6.2)
        sl.shapes.title.height = Inches(1.95)
        if prompt_code:
            code_panel(sl, MX, 3.95, 6.3, prompt_code[0], prompt_code[1], size=12.5, h=2.65)
        else:
            tb(sl, MX, 4.3, 6.1, 1.9, prompt, size=19, color="E8F1F0", spacing=1.05)
        px, pw = 7.55, W - MX - 7.55
        rr(sl, px, 0.85, pw, 5.75, fill=TEAL_D, r=0.04)
        tb(sl, px + 0.4, 1.1, pw - 0.8, 0.4, sub_label.upper(), size=13, color=AMBER, bold=True)
        sh = min(1.15, 4.7 / len(steps))
        for i, st in enumerate(steps):
            yy = 1.65 + i * sh
            num_circle(sl, px + 0.4, yy + 0.05, i + 1, fill=AMBER, color=INK)
            tb(sl, px + 1.1, yy, pw - 1.5, sh - 0.08, st, size=17, color=WHITE, anchor="t")
        footer(sl, self.n, self.phase, dark=True, closing=self.closing)
        self._notes(sl, m, notes)

    def cards_slide(self, title, chips, items, m, notes, key=None, key_label=None, h=None, **kw):
        def draw(sl):
            hh = h or (3.6 if key else 4.5)
            cards(sl, items, y=1.95, h=hh, **kw)
            if key:
                banner(sl, key, y=1.95 + hh + 0.3, h=0.75, size=18, label=key_label)
        return self.content(title, chips, m, notes, draw)

    def flow(self, title, chips, steps, m, notes, below=None, below_label=None, key=None):
        def draw(sl):
            n = len(steps)
            sw = (CW + (n - 1) * 0.14) / n
            for i, (lab, sub) in enumerate(steps):
                x = MX + i * (sw - 0.14)
                kind = MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON
                box(sl, x, 2.05, sw, 0.95, lab, fill=TEAL if i % 2 == 0 else "3C7F85", color=WHITE,
                    size=16, bold=True, kind=kind, r=None, pad=0.3)
                tb(sl, x + 0.3, 3.15, sw - 0.6, 0.9, sub, size=14.5, color=INK, align="c")
            if below:
                rr(sl, MX, 4.1, CW, 1.6, fill=LIGHT, r=0.05)
                tb(sl, MX + 0.3, 4.2, 5, 0.3, (below_label or "¿Qué puede fallar aquí?").upper(), size=11.5,
                   color=RUST, bold=True)
                for i, t in enumerate(below):
                    x = MX + i * (sw - 0.14)
                    tb(sl, x + 0.3, 4.6, sw - 0.6, 1.2, t, size=14, color=INK, align="c")
            if key:
                banner(sl, key, y=5.95, h=0.7, size=18)
        return self.content(title, chips, m, notes, draw)

    def code(self, title, chips, code, lang, fname, bullets, m, notes, size=13, code_w=7.5, key=None):
        def draw(sl):
            hh = code_panel(sl, MX, 1.95, code_w, code, lang, size=size, fname=fname)
            callouts(sl, MX + code_w + 0.35, 1.95, CW - code_w - 0.35, bullets,
                     h=min(1.35, (4.0 - 0.22 * (len(bullets) - 1)) / len(bullets) + 0.3), gap=0.22)
            if key:
                banner(sl, key, y=6.05, h=0.68, size=17)
        return self.content(title, chips, m, notes, draw)

    def compare(self, title, chips, left, right, m, notes, key=None, lang="text", size=14, mono=False):
        def draw(sl):
            half = (CW - 0.4) / 2
            for k, (head, body, col) in enumerate([left, right]):
                x = MX + k * (half + 0.4)
                rr(sl, x, 1.95, half, 0.55, fill=col, r=0.2)
                tb(sl, x + 0.3, 1.95, half - 0.6, 0.55, head, size=18, bold=True, color=WHITE, anchor="m")
                if mono:
                    code_panel(sl, x, 2.65, half, body, lang, size=size, h=3.1)
                else:
                    rr(sl, x, 2.65, half, 3.1, fill=LIGHT, r=0.04)
                    tb(sl, x + 0.3, 2.85, half - 0.6, 2.8, body, size=17, color=INK, spacing=1.08)
            if key:
                banner(sl, key, y=6.0, h=0.7, size=18)
        return self.content(title, chips, m, notes, draw)

    def table(self, title, chips, headers, rows, widths, m, notes, status_col=None, size=14, key=None,
              row_h=0.56):
        def draw(sl):
            tw = sum(widths)
            gf = sl.shapes.add_table(len(rows) + 1, len(headers), Inches(MX), Inches(1.95), Inches(tw),
                                     Inches(row_h * (len(rows) + 1)))
            t = gf.table
            tblPr = t._tbl.tblPr
            tblPr.set("bandRow", "0")
            tblPr.set("firstRow", "0")
            for j, wd in enumerate(widths):
                t.columns[j].width = Inches(wd)
            for i in range(len(rows) + 1):
                t.rows[i].height = Inches(row_h)
            for j, hd in enumerate(headers):
                c = t.cell(0, j)
                c.fill.solid()
                c.fill.fore_color.rgb = rgb(TEAL_D)
                c.vertical_anchor = MSO_ANCHOR.MIDDLE
                _fill_paras(c.text_frame, hd, size, WHITE, True, BODY, "l")
                c.margin_left = c.margin_right = Inches(0.12)
            for i, row in enumerate(rows, 1):
                for j, val in enumerate(row):
                    c = t.cell(i, j)
                    c.fill.solid()
                    c.fill.fore_color.rgb = rgb(LIGHT if i % 2 else WHITE)
                    c.vertical_anchor = MSO_ANCHOR.MIDDLE
                    c.margin_left = c.margin_right = Inches(0.12)
                    if status_col is not None and j == status_col:
                        txt, fill, col = STATUS[val] if isinstance(val, str) else val
                        c.fill.fore_color.rgb = rgb(fill)
                        _fill_paras(c.text_frame, txt, size - 1, col, True, BODY, "c")
                    else:
                        _fill_paras(c.text_frame, val, size, INK, j == 0, BODY, "l")
            if key:
                banner(sl, key, y=6.2, h=0.6, size=17)
        return self.content(title, chips, m, notes, draw)

    def pause(self, m, notes):
        sl = self._new(TEAL_D)
        add_title(sl, "Pausa", color=WHITE, y=1.7, size=72, w=9)
        sl.shapes.title.height = Inches(1.4)
        tb(sl, MX, 3.2, 10, 0.6, f"Volvemos a las {fmt_clock(self.clock + m)}", size=34, color=AMBER, font=HEAD)
        rr(sl, MX, 4.4, 8.2, 1.3, fill="2E505A", r=0.1)
        tb(sl, MX + 0.35, 4.55, 7.6, 1.1,
           [("Mientras tanto", {"bold": True, "color": AMBER, "size": 13}),
            "Deja una pregunta en el muro de dudas: la retomamos al volver."], size=19, color=WHITE)
        footer(sl, self.n, 2, dark=True)
        self._notes(sl, m, notes)

    def save(self, path):
        self.prs.save(path)


# ======================================================================
# Diagramas reutilizables
# ======================================================================
def uc_diagram(sl):
    rr(sl, 3.1, 1.95, 7.1, 4.6, fill=None, line=SLATE, lw=2, r=0.03)
    tb(sl, 3.35, 2.03, 6, 0.35, "Sistema: Solicitudes de material", size=14, bold=True, color=SLATE)
    actor(sl, 1.6, 3.2, "Jefe de unidad")
    actor(sl, 11.7, 3.2, "Almacén")
    ovals = [("Crear solicitud", 3.55, 2.75), ("Ver solicitudes de mi unidad", 3.55, 4.15),
             ("Aprobar solicitud", 6.95, 2.75), ("Registrar entrega", 6.95, 4.15)]
    pos = {}
    for t, x, y in ovals:
        box(sl, x, y, 3.0, 1.0, t, fill=TEAL_L, color=TEAL_D, size=16, bold=True, kind=MSO_SHAPE.OVAL, line=TEAL, lw=1.5)
        pos[t] = (x, y)
    for t in ["Crear solicitud", "Ver solicitudes de mi unidad"]:
        x, y = pos[t]
        line(sl, 1.95, 3.9, x, y + 0.5, color=TEAL_D, w=1.75)
    for t in ["Aprobar solicitud", "Registrar entrega"]:
        x, y = pos[t]
        line(sl, 11.35, 3.9, x + 3.0, y + 0.5, color=TEAL_D, w=1.75)


def class_diagram(sl):
    def cls(x, y, w, name, attrs, color=TEAL):
        hh = 0.55
        bh = 0.38 * len(attrs) + 0.25
        box(sl, x, y, w, hh, name, fill=color, color=WHITE, size=17, bold=True, kind=MSO_SHAPE.RECTANGLE, r=None)
        rr(sl, x, y + hh, w, bh, fill=LIGHT, line=color, lw=1.5, r=0.0)
        tb(sl, x + 0.2, y + hh + 0.12, w - 0.4, bh, attrs, size=15, color=INK, font=MONO)
        return y + hh + bh
    cls(0.9, 2.4, 2.9, "Unidad", ["id", "nombre"])
    cls(5.1, 2.4, 3.1, "Solicitud", ["id", "cantidad", "fecha", "estado"])
    cls(9.4, 2.4, 3.0, "Material", ["id", "nombre"])
    line(sl, 3.8, 3.35, 5.1, 3.35, color=INK, w=2)
    line(sl, 8.2, 3.35, 9.4, 3.35, color=INK, w=2)
    for t, x in [("1", 3.85), ("*", 4.8), ("*", 8.25), ("1", 9.1)]:
        tb(sl, x, 3.0, 0.3, 0.3, t, size=17, bold=True, color=RUST, align="c")
    tb(sl, 3.8, 3.42, 1.3, 0.3, "hace", size=13, color=SLATE, italic=True, align="c")
    tb(sl, 8.2, 3.42, 1.2, 0.3, "pide", size=13, color=SLATE, italic=True, align="c")
    box(sl, 5.1, 5.05, 3.1, 0.5, "«enum» Estado", fill=SLATE, color=WHITE, size=15, bold=True, kind=MSO_SHAPE.RECTANGLE, r=None)
    rr(sl, 5.1, 5.55, 3.1, 0.95, fill=LIGHT, line=SLATE, lw=1.5, r=0.0)
    tb(sl, 5.3, 5.62, 2.8, 0.9, ["PENDIENTE", "APROBADA · ENTREGADA"], size=13, font=MONO)
    line(sl, 6.65, 5.05, 6.65, 4.65 + 0.0, color=SLATE, w=1.5, dash=True, arrow=True)


def seq_diagram(sl):
    names = ["Jefe de unidad", "Pantalla", "API REST", "Servicio", "Base de datos"]
    xs = [1.9, 4.5, 7.0, 9.5, 11.8]
    for nm, x in zip(names, xs):
        box(sl, x - 0.95, 1.95, 1.9, 0.5, nm, fill=TEAL_D, color=WHITE, size=14, bold=True, kind=MSO_SHAPE.RECTANGLE, r=None)
        line(sl, x, 2.45, x, 6.5, color=SAGE, w=1.5, dash=True)
    msgs = [(0, 1, "abre «Solicitudes»", False), (1, 2, "GET /unidades/7/solicitudes", False),
            (2, 3, "porUnidad(7, semana)", False), (3, 4, "SELECT … WHERE unidad_id = 7", False),
            (4, 3, "filas", True), (3, 2, "lista de DTO", True), (2, 1, "200 + JSON", True),
            (1, 0, "tabla en pantalla", True)]
    y = 2.9
    for a, b, t, ret in msgs:
        line(sl, xs[a], y, xs[b], y, color=RUST if ret else TEAL_D, w=2, arrow=True, dash=ret)
        left = min(xs[a], xs[b])
        tb(sl, left + 0.1, y - 0.3, abs(xs[a] - xs[b]) - 0.2, 0.28, t, size=12.5, color=INK, align="c", font=BODY)
        y += 0.46


def bpmn_diagram(sl):
    shape(sl, MSO_SHAPE.RECTANGLE, MX, 1.95, CW, 1.75, fill=None, line=SLATE, lw=1.5)
    shape(sl, MSO_SHAPE.RECTANGLE, MX, 3.7, CW, 1.75, fill=None, line=SLATE, lw=1.5)
    box(sl, MX, 1.95, 0.6, 1.75, "", fill=TEAL_D, kind=MSO_SHAPE.RECTANGLE, r=None)
    box(sl, MX, 3.7, 0.6, 1.75, "", fill=TEAL_D, kind=MSO_SHAPE.RECTANGLE, r=None)
    for txt, y in [("Jefe de unidad", 1.95), ("Almacén", 3.7)]:
        b = tb(sl, MX - 0.58, y + 0.85, 1.75, 0.3, txt, size=12.5, bold=True, color=WHITE, align="c", anchor="m")
        b.rotation = 270
    box(sl, 2.1, 2.625, 0.4, 0.4, "", fill=WHITE, line=GREEN, lw=3, kind=MSO_SHAPE.OVAL)
    tb(sl, 1.6, 3.1, 1.4, 0.3, "Necesita material", size=11, color=SLATE, align="c")
    line(sl, 2.5, 2.825, 2.9, 2.825, color=INK, arrow=True)
    box(sl, 2.9, 2.425, 1.7, 0.8, "Crear solicitud", fill=TEAL_L, line=TEAL, lw=1.5, size=14, bold=True)
    line(sl, 3.75, 3.225, 3.75, 4.175, color=INK, arrow=True)
    box(sl, 2.9, 4.175, 1.7, 0.8, "Revisar stock", fill=TEAL_L, line=TEAL, lw=1.5, size=14, bold=True)
    line(sl, 4.6, 4.575, 5.3, 4.575, color=INK, arrow=True)
    box(sl, 5.3, 4.075, 1.0, 1.0, "", fill=WHITE, line=AMBER, lw=3, kind=MSO_SHAPE.DIAMOND)
    tb(sl, 4.9, 5.1, 1.8, 0.3, "¿Hay stock?", size=12.5, bold=True, color=INK, align="c")
    line(sl, 6.3, 4.575, 7.3, 4.575, color=INK, arrow=True)
    tb(sl, 6.45, 4.25, 0.6, 0.3, "Sí", size=13, bold=True, color=GREEN, align="c")
    box(sl, 7.3, 4.175, 1.7, 0.8, "Preparar entrega", fill=WHITE, line=SLATE, lw=1.5, size=14, bold=True)
    line(sl, 9.0, 4.575, 9.75, 4.575, color=INK, arrow=True)
    box(sl, 9.75, 4.375, 0.4, 0.4, "", fill=WHITE, line=RUST, lw=3.5, kind=MSO_SHAPE.OVAL)
    tb(sl, 9.2, 4.85, 1.5, 0.3, "Entregado", size=11, color=SLATE, align="c")
    line(sl, 5.8, 4.075, 5.8, 3.225, color=INK, arrow=True)
    tb(sl, 5.9, 3.5, 0.6, 0.3, "No", size=13, bold=True, color=RUST)
    box(sl, 4.95, 2.425, 1.7, 0.8, "Notificar a la unidad", fill=TEAL_L, line=TEAL, lw=1.5, size=14, bold=True)
    line(sl, 6.65, 2.825, 7.5, 2.825, color=INK, arrow=True)
    box(sl, 7.5, 2.625, 0.4, 0.4, "", fill=WHITE, line=RUST, lw=3.5, kind=MSO_SHAPE.OVAL)
    tb(sl, 7.0, 3.1, 1.5, 0.3, "Replanificar", size=11, color=SLATE, align="c")
    rr(sl, 10.35, 2.1, 2.1, 1.5, fill=LIGHT, r=0.06)
    box(sl, 10.5, 2.25, 0.35, 0.28, "", fill=TEAL_L, line=TEAL, lw=1.5)
    tb(sl, 10.95, 2.22, 1.5, 0.35, "Lo hace el sistema", size=11.5, anchor="m")
    box(sl, 10.5, 2.75, 0.35, 0.28, "", fill=WHITE, line=SLATE, lw=1.5)
    tb(sl, 10.95, 2.72, 1.5, 0.35, "Lo hace una persona", size=11.5, anchor="m")
    box(sl, 10.53, 3.17, 0.3, 0.3, "", fill=WHITE, line=AMBER, lw=2, kind=MSO_SHAPE.DIAMOND)
    tb(sl, 10.95, 3.15, 1.5, 0.35, "Decisión", size=11.5, anchor="m")
    banner(sl, "Primero se acuerda el proceso. Después se decide qué parte se automatiza.", y=5.8, h=0.7, size=18)


def git_diagram(sl):
    y0, y1 = 2.35, 3.55
    line(sl, MX, y0, W - MX, y0, color=TEAL_D, w=3)
    tb(sl, MX, y0 - 0.5, 1.5, 0.3, "main", size=15, bold=True, color=TEAL_D, font=MONO)
    for x in [1.7, 3.0, 8.9, 11.0]:
        shape(sl, MSO_SHAPE.OVAL, x - 0.13, y0 - 0.13, 0.26, 0.26, fill=TEAL_D)
    line(sl, 3.0, y0, 4.0, y1, color=AMBER, w=3)
    line(sl, 4.0, y1, 7.8, y1, color=AMBER, w=3)
    line(sl, 7.8, y1, 8.9, y0, color=AMBER, w=3, arrow=True)
    tb(sl, 3.9, y1 + 0.12, 4.2, 0.3, "feature/PROJ-101-listar-solicitudes", size=13, bold=True, color="9A7000", font=MONO)
    for x in [4.7, 5.9, 7.1]:
        shape(sl, MSO_SHAPE.OVAL, x - 0.13, y1 - 0.13, 0.26, 0.26, fill=AMBER)
    box(sl, 7.55, y0 + 0.25, 1.5, 0.4, "Pull Request", fill=TEAL, color=WHITE, size=12.5, bold=True, r=0.5)
    shape(sl, MSO_SHAPE.OVAL, 10.87, y0 - 0.13, 0.26, 0.26, fill=RUST)
    box(sl, 10.4, y0 - 0.65, 1.2, 0.34, "v1.2.0", fill=RUST, color=WHITE, size=12.5, bold=True, r=0.5, font=MONO)
    tb(sl, 10.1, y0 + 0.2, 2.8, 0.3, "etiqueta = release", size=12, color=SLATE, align="c", italic=True)


def pyramid(sl, x, y):
    layers = [("E2E", "pocos · lentos", 2.2, RUST), ("Integración", "API + base de datos", 3.8, AMBER),
              ("Unitarias", "muchas · rápidas", 5.4, TEAL)]
    for i, (a, b, w, col) in enumerate(layers):
        box(sl, x + (5.4 - w) / 2, y + i * 1.1, w, 1.0, [(a, {"bold": True, "size": 16}), (b, {"size": 13})],
            fill=col, color=WHITE if col != AMBER else INK, kind=MSO_SHAPE.ROUNDED_RECTANGLE)


def checklist(sl, x, y, w, items, size=16, gap=0.5, col=TEAL):
    for i, it in enumerate(items):
        yy = y + i * gap
        shape(sl, MSO_SHAPE.ROUNDED_RECTANGLE, x, yy + 0.04, 0.28, 0.28, fill=WHITE, line=col, lw=2, r=0.2)
        tb(sl, x + 0.45, yy, w - 0.5, 0.4, it, size=size, color=INK, anchor="m")


def mini_table(sl, x, y, headers, rows, widths, size=12.5, head_fill=TEAL_D, row_h=0.4, bad_rows=()):
    xx = x
    for j, hd in enumerate(headers):
        box(sl, xx, y, widths[j], row_h, hd, fill=head_fill, color=WHITE, size=size, bold=True,
            kind=MSO_SHAPE.RECTANGLE, r=None, align="l", pad=0.1)
        xx += widths[j]
    for i, row in enumerate(rows):
        xx = x
        for j, v in enumerate(row):
            box(sl, xx, y + (i + 1) * row_h, widths[j], row_h, v, fill=LIGHT if i % 2 == 0 else WHITE, color=INK,
                size=size, kind=MSO_SHAPE.RECTANGLE, r=None, align="l", pad=0.1, line="D5DDE0", lw=0.5)
            xx += widths[j]


def build(path):
    d = Deck()

    # ------------------------------------------------------------------ APERTURA 08:10–08:30
    d.title(1, N(
        "Bienvenida. Idea fuerza: todo lo que estudiáis forma parte de una misma cadena y hoy la recorremos con una "
        "funcionalidad real, de principio a fin. No es otra asignatura: es el pegamento entre todas.",
        corto="Pasa directa a «Quién soy».",
        enlace="Horario real del lunes 9 de noviembre (calendario V2): 8:10–10:50, pausa 10:50–11:20 y 11:20–14:00, "
               "ocupado por las sesiones 36–41 de Dirección de Proyectos. Este guion respeta ese horario."))

    def quien(sl):
        cards(sl, [
            ("Consultora Java", "Servicios backend con Java y Quarkus, en IntelliJ, para BMW.",
             "Escribo y reviso código cada día"),
            ("Release manager", "Versiones y entregas con GitHub, ArgoCD y Azure.",
             "Llevo el código hasta producción"),
            ("Mi día a día", "Postgres · DB2 · Docker · Jira · Confluence",
             "Todo esto sale del programa"),
        ], y=1.55, h=3.3)
        tb(sl, MX, 5.1, 3, 0.3, "MI STACK DE CADA DÍA", size=11.5, color=SLATE, bold=True)
        groups = [("Código", ["Java", "Quarkus", "IntelliJ"], TEAL), ("Datos", ["Postgres", "DB2"], "3C7F85"),
                  ("Entrega", ["Docker", "ArgoCD", "GitHub", "Azure"], AMBER), ("Gestión", ["Jira", "Confluence"], SLATE)]
        x = MX
        for g, items, col in groups:
            tb(sl, x, 5.5, 1.2, 0.3, g, size=12.5, bold=True, color=TEAL_D)
            xx = x
            for it in items:
                w = pill(sl, xx, 5.85, it, col, INK if col == AMBER else WHITE, size=13, h=0.36)
                xx += w + 0.1
            x = xx + 0.45
    d.content("Quién soy", None, 3, N(
        "Preséntate en 2 minutos: cómo llegaste de la carrera a este trabajo y qué haces un día normal. Anticipa que cada "
        "herramienta de la tarjeta sale de una asignatura que ellos ven.",
        pregunta="¿Cuántos conocéis ya alguna de estas herramientas? (mano alzada)",
        historia="Tu primer proyecto real y qué te sorprendió respecto a lo que habías estudiado."), quien)

    d.statement("¿Para qué sirve todo esto?",
                "Esa pregunta es el hilo del día: cada asignatura, cada herramienta, cada diagrama.", 1, N(
                    "Plantea el objetivo: hoy no se enseña nada nuevo; se conecta lo que ya hacéis con por qué existe. "
                    "Invita a que alguien diga en voz alta qué asignatura le parece que «no sirve para nada»."))

    def hoy(sl):
        blocks = [("Apertura", 20, SLATE), ("Requisitos", 45, TEAL), ("Diseño", 45, "3C7F85"), ("Datos", 50, TEAL),
                  ("Pausa", 30, AMBER), ("Código", 45, TEAL), ("Integración", 50, "3C7F85"), ("Entrega", 45, TEAL),
                  ("Cierre", 20, SLATE)]
        total = 350
        x = MX
        times = ["08:10", "08:30", "09:15", "10:00", "10:50", "11:20", "12:05", "12:55", "13:40"]
        for (name, m, col), t in zip(blocks, times):
            w = CW * m / total
            box(sl, x, 2.2, w - 0.04, 1.05, [(name, {"bold": True, "size": 13 if w > 1 else 9.5})],
                fill=col, color=INK if col == AMBER else WHITE, kind=MSO_SHAPE.RECTANGLE, r=None, pad=0.02)
            tb(sl, x, 3.32, w, 0.3, t, size=12, color=SLATE, bold=True)
            x += w
        tb(sl, MX, 1.75, CW, 0.3, "LUNES 9 DE NOVIEMBRE · 08:10 – 14:00", size=12, color=SLATE, bold=True)
        cards(sl, [("Preguntad", "Interrumpid cuando queráis. Una duda vuestra vale más que mi diapositiva."),
                   ("Casos rotos", "Tres veces os doy un fallo y decidís por dónde empezar a mirar."),
                   ("Muro de dudas", "Lo que no demos tiempo a responder queda apuntado y lo retomamos.")],
              y=3.95, h=2.3, accent=AMBER)
    d.content("Cómo funciona hoy", None, 2, N(
        "Explica el mapa del día con los horarios reales: dos mitades de 2 h 40 min con una pausa de 30 min. Pacta las "
        "reglas: se puede interrumpir, hay tres «casos rotos» y un muro (pizarra o folio) para dudas que quedan pendientes.",
        corto="Enseña solo la barra de tiempo."), hoy)

    d.ask("Pregunta abierta", "¿En qué proyecto estáis trabajando?",
          "Cada persona cuenta en un minuto qué hace y con qué herramientas.",
          ["Un minuto por persona, sin preparación.",
           "Yo apunto las herramientas en la pizarra, en cuatro columnas: Código · Datos · Entrega · Gestión.",
           "Volveremos a este mapa a las 13:40, al cerrar."], 7, N(
              "Rompehielos para saber con quién hablas: nivel, herramientas conocidas y tono de la sala.",
              pregunta="¿Y qué tecnología os suena pero todavía no sabéis para qué sirve?",
              enlace="Lo que apuntes será el hilo para el mapa final de las 13:40.",
              corto="Limita a 30 segundos por persona."))

    def donde(sl):
        cols = [("Ya lo habéis visto", GREEN, WHITE,
                 ["Arquitectura de computadores", "Redes de datos", "Algoritmos y estructuras de datos",
                  "Linux de usuario", "Ingeniería de requisitos", "Teoría de sistemas operativos"]),
                ("Lo estáis viendo", AMBER, INK,
                 ["Dirección de proyectos", "ADOO · UML", "Herramientas colaborativas (Git)", "HTML / CSS",
                  "POO (empieza esta semana)", "SO de servidor"]),
                ("Os llega", SLATE, WHITE,
                 ["Bases de datos: modelado, SQL, Oracle, NoSQL", "Patrones y pruebas",
                  "Servicios web, serialización", "ECMAScript, frontend y backend", "Docker, Postman, Gradle",
                  "Seguridad TIC · BPMN"])]
        cw = (CW - 0.5) / 3
        for i, (hd, fill, col, items) in enumerate(cols):
            x = MX + i * (cw + 0.25)
            rr(sl, x, 1.55, cw, 0.6, fill=fill, r=0.2)
            tb(sl, x + 0.3, 1.55, cw - 0.6, 0.6, hd, size=19, bold=True, color=col, font=HEAD, anchor="m")
            rr(sl, x, 2.25, cw, 3.55, fill=LIGHT, r=0.04)
            tb(sl, x + 0.3, 2.45, cw - 0.55, 3.3, [(t, {"after": 7}) for t in items], size=17, color=INK)
        banner(sl, "Más de la mitad del programa aún está por delante: hoy os cuento para qué.", y=6.0, h=0.7,
               size=18)
    d.content("Dónde estáis en el curso", None, 3, N(
        "Sitúa al grupo: llevan unas diez semanas y lo ya visto es la base (algoritmos, redes, Linux, requisitos). "
        "Lo que más afecta a un proyecto real, como bases de datos, servicios web, Docker o pruebas, todavía no ha llegado: "
        "la charla funciona como adelanto para que lleguen a esas asignaturas sabiendo para qué sirven.",
        pregunta="¿Alguna de estas asignaturas la habéis visto ya en otra formación?",
        enlace="Estado según el calendario semanal de 13/10/2026 (hasta el 18/12 solo aparecen Dirección de Proyectos, "
               "ADOO, Git, HTML/CSS, POO y SO de servidor). Confirma con el tutor si ha cambiado algo.",
        corto="Lee solo los títulos de las tres columnas."), donde)

    def caso(sl):
        rr(sl, MX, 1.55, CW, 1.25, fill=TEAL_D, r=0.08)
        tb(sl, MX + 0.4, 1.62, 3, 0.3, "LA HISTORIA", size=11.5, color=AMBER, bold=True)
        tb(sl, MX + 0.4, 1.93, CW - 0.8, 0.8,
           [[("Como ", {}), ("jefe de unidad", {"bold": True, "color": AMBER}),
             (" quiero ver las solicitudes de material de mi unidad de esta semana, para priorizar la reposición.", {})]],
           size=21, color=WHITE, font=HEAD)
        cards(sl, [("Unidad", "Quién pide. Tiene un nombre y un jefe."),
                   ("Solicitud", "Qué pide, cuánto, cuándo y en qué estado está."),
                   ("Material", "El catálogo: lo que se puede pedir.")],
              y=3.1, h=2.2, accent=AMBER, numbered=False)
        banner(sl, "Pequeño a propósito: cabe en una pantalla pero toca todas las capas.", y=5.65, h=0.7, size=18)
    d.content("El caso que seguiremos todo el día", None, 2, N(
        "Presenta el hilo conductor: una funcionalidad sencilla (listar las solicitudes de material de una unidad) que "
        "atravesará requisitos, diseño, datos, código, integración y entrega. Es un ejemplo genérico, sin datos reales ni "
        "clasificados.",
        pregunta="¿Qué os imagináis que puede salir mal en algo tan simple?",
        corto="Lee la historia y pasa al mapa."), caso)

    def viaje(sl):
        steps = [("Requisitos", "Qué se pide"), ("Diseño", "Cómo se piensa"), ("Datos", "Qué se guarda"),
                 ("Código", "Cómo se escribe"), ("Integración", "Unir las capas"), ("Entrega", "Llegar a producción")]
        sw = (CW + 5 * 0.14) / 6
        for i, (a, b) in enumerate(steps):
            x = MX + i * (sw - 0.14)
            box(sl, x, 1.95, sw, 1.0, a, fill=TEAL if i % 2 == 0 else "3C7F85", color=WHITE, size=16, bold=True,
                kind=MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON, r=None, pad=0.3)
            tb(sl, x + 0.25, 3.05, sw - 0.5, 0.4, b, size=14, color=SLATE, align="c")
        subs = [["Ing. de Requisitos", "Dirección de Proyectos"], ["ADOO · UML", "Patrones · BPMN"],
                ["Bases de datos", "SQL · Oracle · NoSQL"], ["POO · Algoritmos", "Librerías · Git"],
                ["Servicios web", "HTML · ECMAScript · Postman"], ["Docker · Linux", "Redes · Seguridad"]]
        for i, s in enumerate(subs):
            x = MX + i * (sw - 0.14) + 0.12
            rr(sl, x, 3.75, sw - 0.38, 1.7, fill=LIGHT, r=0.06)
            tb(sl, x + 0.12, 3.85, sw - 0.62, 0.3, "ASIGNATURAS", size=9.5, color=SLATE, bold=True, align="c")
            tb(sl, x + 0.12, 4.2, sw - 0.62, 1.2, [(t, {"after": 6}) for t in s], size=14, color=INK, align="c")
        banner(sl, "Seis bloques, seis pasos. Este mapa nos acompaña abajo en cada diapositiva.", y=5.85, h=0.7, size=18)
    d.content("El viaje de una funcionalidad", None, 1, N(
        "Mapa del día. Cada paso es un bloque y debajo ves qué asignaturas del programa lo cubren. Señala la barra de "
        "progreso del pie de cada diapositiva: siempre sabréis en qué paso estamos."), viaje)
    d.expect("08:30")

    # ------------------------------------------------------------------ B1 REQUISITOS 08:30–09:15
    d.section(0, "Requisitos y planificación", "Aquí nace todo", ch("req", "dir", "herr"), 45,
              "¿Alguien ha trabajado con Jira, Trello o algo parecido? ¿Qué pasaba cuando un ticket no se entendía?",
              N("Abre el bloque. Aquí empieza todo: sin un buen requisito, el mejor código resuelve el problema equivocado.",
                enlace="Ingeniería de requisitos ya la habéis dado. Dirección de proyectos la estáis dando ahora; lo que "
                       "veamos es cómo se ve en un tablero real."), tag="BLOQUE")

    def jerarquia(sl):
        box(sl, 3.9, 1.95, 5.5, 0.85, [("ÉPICA", {"size": 10.5, "bold": True, "color": AMBER}),
                                       ("Gestión de solicitudes de material", {"bold": True, "size": 17})],
            fill=TEAL_D, color=WHITE, kind=MSO_SHAPE.ROUNDED_RECTANGLE, r=0.12)
        hist = [("Ver solicitudes de mi unidad", TEAL, WHITE), ("Crear una solicitud", "3C7F85", WHITE),
                ("Aprobar una solicitud", "3C7F85", WHITE)]
        xs = [0.7, 4.75, 8.8]
        for (t, f, c), x in zip(hist, xs):
            line(sl, 6.65, 2.8, x + 1.9, 3.35, color=SLATE, w=1.75)
            box(sl, x, 3.35, 3.8, 0.85, [("HISTORIA", {"size": 10.5, "bold": True, "color": AMBER}), (t, {"bold": True, "size": 16})],
                fill=f, color=c, kind=MSO_SHAPE.ROUNDED_RECTANGLE, r=0.12)
        tasks = ["API: GET solicitudes", "Consulta + índice", "Pantalla del listado", "Pruebas"]
        line(sl, 2.6, 4.2, 2.6, 4.55, color=SLATE, w=1.5)
        for i, t in enumerate(tasks):
            xx = 0.7 + (i % 2) * 1.95
            yy = 4.65 + (i // 2) * 0.62
            box(sl, xx, yy, 1.85, 0.52, t, fill=LIGHT, line=TEAL, lw=1.25, size=12.5, bold=True, r=0.2, pad=0.05)
        defs = [("Épica", "el objetivo de negocio"), ("Historia", "lo que necesita el usuario y por qué"),
                ("Tarea", "trabajo concreto, con responsable y estado")]
        for i, (a, c) in enumerate(defs):
            yy = 4.6 + i * 0.6
            shape(sl, MSO_SHAPE.RECTANGLE, 5.0, yy + 0.05, 0.07, 0.4, fill=AMBER)
            tb(sl, 5.25, yy, 7.3, 0.5, [[(a + ": ", {"bold": True, "color": TEAL_D, "font": HEAD}), (c, {})]],
               size=17, color=INK, anchor="m")
        tb(sl, 0.7, 6.02, 4, 0.3, "TAREAS DE LA PRIMERA HISTORIA", size=10.5, bold=True, color=SLATE)
    d.content("Todo empieza con un ticket", ch("dir", "herr"), 4, N(
        "Recorre la jerarquía de arriba abajo con nuestro caso. Una épica se divide en historias y cada historia en "
        "tareas que alguien puede terminar en uno o dos días. Muestra cómo se refleja en el tablero.",
        pregunta="¿Qué diferencia hay entre una historia y una tarea?",
        historia="Enseña una épica real de tu proyecto con los nombres cambiados y cuántos tickets generó."), jerarquia)

    def pedir(sl):
        rr(sl, MX, 1.95, 4.6, 2.3, fill=LIGHT, r=0.05)
        tb(sl, MX + 0.3, 2.05, 4, 0.3, "LO QUE PIDEN", size=12, bold=True, color=RUST)
        tb(sl, MX + 0.3, 2.45, 4.0, 1.7, "«Quiero un informe de solicitudes.»", size=24, color=INK, font=HEAD, italic=True)
        shape(sl, MSO_SHAPE.RIGHT_ARROW, 5.5, 2.75, 0.8, 0.7, fill=AMBER)
        rr(sl, 6.55, 1.95, W - MX - 6.55, 2.3, fill=TEAL_L, r=0.05)
        tb(sl, 6.85, 2.05, 4, 0.3, "LO QUE NECESITAN", size=12, bold=True, color=TEAL)
        tb(sl, 6.85, 2.45, W - MX - 7.15, 1.7,
           "Ver las solicitudes de mi unidad de esta semana, ordenadas por fecha, y poder exportarlas.", size=19,
           color=INK, spacing=1.05)
        tb(sl, MX, 4.5, CW, 0.3, "CINCO PREGUNTAS QUE LO DESBLOQUEAN", size=12, bold=True, color=SLATE)
        qs = ["¿Quién lo usa?", "¿Qué decisión toma con él?", "¿Con qué datos?", "¿Cada cuánto?", "¿Cómo sabremos que está bien?"]
        w = (CW - 4 * 0.2) / 5
        for i, q in enumerate(qs):
            box(sl, MX + i * (w + 0.2), 4.9, w, 1.2, q, fill=TEAL_D, color=WHITE, size=17, bold=True, font=HEAD)
        tb(sl, MX, 6.3, CW, 0.4, "Si no hay respuesta a la última, el requisito no está terminado.", size=16, color=SLATE, italic=True)
    d.content("Pedir no es necesitar", ch("req"), 4, N(
        "Un requisito vago produce trabajo repetido. Compara la frase inicial con la versión verificable y enseña "
        "las cinco preguntas que haces para llegar de una a otra.",
        pregunta="¿Qué os pedirían a vosotros con una frase así? ¿Qué preguntaríais primero?",
        historia="Un requisito de tu proyecto que parecía claro y escondía dos o tres suposiciones.",
        enlace="Ingeniería de requisitos: criterios de aceptación y requisitos verificables."), pedir)

    d.ask("Ejercicio en parejas", "Convierte esta petición en una historia",
          "«Necesito un panel de entregas.»\n\nTenéis 8 minutos. Hoy no hay respuesta correcta única: importa el razonamiento.",
          ["Decid quién lo usa y para qué decisión (2 minutos).",
           "Escribid la historia: «Como… quiero… para…».",
           "Redactad tres criterios de aceptación comprobables.",
           "Apuntad una cosa que NO entra en esta versión."], 8, N(
              "Dales la petición vaga y deja que trabajen en parejas. Pasea por la sala y escucha qué suposiciones aparecen. "
              "Después, que dos parejas lean su historia en voz alta.",
              pregunta="¿Qué habéis tenido que suponer? ¿A quién se lo habríais preguntado?",
              corto="Hazlo en 5 minutos y solo pide la historia, sin criterios.",
              enlace="Es el ejercicio de Ingeniería de requisitos aplicado a una petición real."))

    def historia(sl):
        rr(sl, MX, 1.95, CW, 0.95, fill=TEAL_D, r=0.1)
        tb(sl, MX + 0.4, 1.95, CW - 0.8, 0.95,
           [[("Como ", {}), ("jefe de unidad", {"bold": True, "color": AMBER}), (" quiero ", {}),
             ("ver las solicitudes de mi unidad de esta semana", {"bold": True, "color": AMBER}),
             (" para priorizar la reposición.", {})]], size=20, color=WHITE, font=HEAD, anchor="m")
        crit = [("Dado que soy jefe de la unidad 7", "cuando abro «Solicitudes»", "entonces veo solo las de mi unidad"),
                ("Dado que no elijo semana", "cuando se abre el listado", "entonces veo la semana actual"),
                ("Dado que no hay ninguna", "cuando abro el listado", "entonces veo «Sin solicitudes esta semana»")]
        cw = (CW - 0.5) / 3
        for i, (a, b, c) in enumerate(crit):
            x = MX + i * (cw + 0.25)
            rr(sl, x, 3.2, cw, 2.45, fill=LIGHT, r=0.05)
            tb(sl, x + 0.25, 3.3, cw, 0.3, f"CRITERIO {i + 1}", size=11.5, bold=True, color=TEAL)
            tb(sl, x + 0.25, 3.7, cw - 0.5, 1.9,
               [[("Dado ", {"bold": True, "color": TEAL}), (a.replace("Dado ", ""), {})],
                [("Cuando ", {"bold": True, "color": TEAL}), (b.replace("cuando ", ""), {})],
                [("Entonces ", {"bold": True, "color": TEAL}), (c.replace("entonces ", ""), {})]], size=16, color=INK, spacing=1.15)
        banner(sl, "Cada criterio se convierte, más tarde, en una prueba automática.", y=5.95, h=0.7, size=18)
    d.content("Así queda nuestra historia", ch("req"), 4, N(
        "Es el modelo para comparar con lo que han hecho en el ejercicio: historia de una frase y criterios en formato "
        "Dado / Cuando / Entonces. Resalta que los criterios hacen que dos personas distintas entiendan lo mismo y que, "
        "más adelante, se convierten en pruebas.",
        pregunta="¿Qué criterio añadiríais vosotros?"), historia)

    def tablero(sl):
        cols = [("Por hacer", SLATE, ["PROJ-102 Exportar a CSV", "PROJ-105 Mensaje sin resultados"]),
                ("En curso", AMBER, ["PROJ-103 Índice (unidad, fecha)"]),
                ("En revisión", TEAL, ["PROJ-101 API: listar solicitudes"]),
                ("Hecho", GREEN, ["PROJ-100 Modelo de datos"])]
        cw = (CW - 3 * 0.2) / 4
        for i, (nm, col, cards_) in enumerate(cols):
            x = MX + i * (cw + 0.2)
            rr(sl, x, 1.95, cw, 3.65, fill=LIGHT, r=0.04)
            box(sl, x, 1.95, cw, 0.5, nm, fill=col, color=INK if col == AMBER else WHITE, size=16, bold=True,
                kind=MSO_SHAPE.RECTANGLE, r=None)
            for k, t in enumerate(cards_):
                rr(sl, x + 0.15, 2.65 + k * 1.0, cw - 0.3, 0.85, fill=WHITE, line="C9D4D7", lw=1, r=0.06)
                tb(sl, x + 0.3, 2.72 + k * 1.0, cw - 0.6, 0.72, t, size=13.5, color=INK, anchor="m")
        banner(sl, "Alcance, estimación, sprint y seguimiento: el tablero es la verdad compartida del equipo.", y=5.85, h=0.7, size=18)
    d.content("Planificar: un sprint, un tablero", ch("dir"), 4, N(
        "Enseña un tablero real con las columnas habituales. Explica el alcance (qué entra y qué no), la estimación y por qué "
        "las entregas cortas se pueden revisar. Recalca que el estado visible para todos evita preguntas del tipo «¿cómo va esto?».",
        pregunta="¿Cómo se coordinan ahora en vuestros trabajos de clase: quién hace qué y cómo se enteran los demás?",
        historia="Cómo organiza tu equipo los sprints y qué ceremonia te parece más útil."), tablero)

    def estimar(sl):
        tb(sl, MX, 1.85, 6, 0.3, "LO QUE PLANIFICAMOS", size=12, bold=True, color=SLATE)
        tb(sl, MX, 3.15, 6, 0.3, "LO QUE PASA", size=12, bold=True, color=SLATE)
        box(sl, MX, 2.2, 5.0, 0.7, "Programar", fill=TEAL, color=WHITE, size=16, bold=True, kind=MSO_SHAPE.RECTANGLE, r=None)
        parts = [("Programar", 2.3, TEAL), ("Integrar", 1.1, AMBER), ("Pruebas", 1.0, "3C7F85"), ("Esperas y revisiones", 1.3, RUST)]
        x = MX
        for nm, w, col in parts:
            box(sl, x, 3.5, w, 0.7, nm, fill=col, color=INK if col == AMBER else WHITE, size=12.5, bold=True,
                kind=MSO_SHAPE.RECTANGLE, r=None, pad=0.03)
            x += w
        tb(sl, MX, 4.35, 5.7, 0.3, "Ejemplo ilustrativo.", size=12, color=SLATE, italic=True)
        callouts(sl, 7.1, 1.95, W - MX - 7.1, [("Estima en rangos", "«Entre 3 y 5 días», no «el jueves»."),
                                               ("Divide hasta tareas de ≤ 1 día", "Lo pequeño se estima mejor."),
                                               ("Cuenta lo que no es código", "Integración, pruebas, revisiones y esperas.")],
                 h=1.15, gap=0.2)
        banner(sl, "Las estimaciones fallan por lo que no cuentas, no por lo que programas mal.", y=5.85, h=0.7, size=18)
    d.content("Por qué fallan las estimaciones", ch("dir"), 4, N(
        "Una estimación casi nunca falla en el tiempo de programar: falla en lo que se olvida (integración, pruebas, "
        "revisiones, accesos pendientes). Las barras son ilustrativas, no estadísticas.",
        pregunta="¿Alguna vez un trabajo os ha llevado el doble de lo previsto? ¿Dónde se fue el tiempo?",
        historia="Un caso en el que una estimación falló y qué se aprendió; sin nombres ni cifras del cliente.",
        corto="Enseña solo la barra y el mensaje final."), estimar)

    def dod(sl):
        half = (CW - 0.4) / 2
        for k, (hd, col, items) in enumerate([
            ("Listo para empezar", TEAL, ["Se entiende quién lo usa y para qué", "Tiene criterios de aceptación", "Sabemos de qué depende",
                                          "Es lo bastante pequeña para un sprint", "Alguien la ha estimado"]),
            ("Hecho de verdad", GREEN, ["Cumple todos los criterios", "Tiene pruebas automáticas", "Otra persona ha revisado el código",
                                        "Está documentada donde toca", "Se puede desplegar sin pasos manuales"])]):
            x = MX + k * (half + 0.4)
            rr(sl, x, 1.95, half, 0.6, fill=col, r=0.2)
            tb(sl, x + 0.3, 1.95, half - 0.6, 0.6, hd, size=19, bold=True, color=WHITE, font=HEAD, anchor="m")
            rr(sl, x, 2.65, half, 3.1, fill=LIGHT, r=0.04)
            checklist(sl, x + 0.35, 2.9, half - 0.7, items, size=17, gap=0.58, col=col)
        banner(sl, "Si el equipo no está de acuerdo en qué es «hecho», cada uno entrega algo distinto.", y=5.95, h=0.7, size=18)
    d.content("¿Listo para empezar? ¿Hecho de verdad?", ch("dir"), 3, N(
        "Dos listas de comprobación que acuerda el equipo: una antes de empezar y otra antes de dar algo por terminado. "
        "Evitan los «casi está» que nunca terminan.",
        pregunta="¿Qué añadiríais a la lista de «hecho»?",
        corto="Lee solo los tres primeros puntos de cada lista."), dod)

    d.ask("Pregunta abierta", "¿Cuál fue el último requisito mal entendido?",
          "Contad un caso: qué pasó y cómo se descubrió.",
          ["Voluntarios: dos o tres historias, un minuto cada una.",
           "Yo apunto el patrón común en la pizarra: falta de ejemplos, usuario ausente, supuestos.",
           "Al final, una regla que podamos aplicar la próxima vez."], 7, N(
              "Dinámica abierta. Si nadie habla, empieza tú con un caso propio y pregunta si les suena parecido.",
              historia="Un requisito malentendido de tu trabajo, qué costó rehacerlo y qué se cambió en el proceso.",
              corto="Quédate con dos historias y el patrón común."))

    def documentar(sl):
        items = [("README", "Cómo arrancar y probar el proyecto.", "Lo primero que abre un recién llegado"),
                 ("Confluence", "Decisiones, diseño y guías de operación.", "Lo que explica el porqué"),
                 ("En el código", "Nombres claros y comentarios que explican el porqué.", "No el qué, que ya se ve"),
                 ("El contrato", "OpenAPI: qué ofrece cada servicio.", "Documentación que se puede probar")]
        w = (CW - 3 * 0.25) / 4
        for i, (a, b, c) in enumerate(items):
            x = MX + i * (w + 0.25)
            rr(sl, x, 1.95, w, 3.2, fill=LIGHT, r=0.04)
            shape(sl, MSO_SHAPE.RECTANGLE, x, 1.95, w, 0.07, fill=AMBER)
            tb(sl, x + 0.25, 2.2, w - 0.5, 0.5, a, size=21, bold=True, font=HEAD, color=TEAL_D)
            tb(sl, x + 0.25, 2.85, w - 0.5, 1.4, b, size=16.5, color=INK)
            tb(sl, x + 0.25, 4.3, w - 0.5, 0.8, c, size=13, color=SLATE, italic=True, anchor="b")
        banner(sl, "Si alguien pregunta dos veces lo mismo, se documenta.", y=5.45, h=0.9, size=24, label="Mi regla")
    d.content("Documentar: dónde vive el saber", ch("herr", "dir"), 6, N(
        "Es el tema que más suele costar a los equipos. Muestra los cuatro sitios donde vive el conocimiento y propón la regla "
        "de «dos preguntas = documentar». Tienes una plantilla de README en el anexo del final.",
        pregunta="¿Qué documentáis hoy en vuestro proyecto y dónde? ¿Podría entenderlo alguien nuevo en un día?",
        historia="Un README o página de Confluence que te salvó (o que echaste de menos) al entrar en un proyecto.",
        enlace="Herramientas colaborativas incluye 3 sesiones de Markdown: es exactamente el formato del README."), documentar)
    d.expect("09:15")

    # ------------------------------------------------------------------ B2 DISEÑO 09:15–10:00
    d.section(1, "Diseño", "Pensar antes de teclear", ch("adoo", "pat", "bpmn"), 45,
              "¿Cuándo dibujáis algo antes de programar y cuándo os lanzáis directamente al código?",
              N("Abre el bloque. Pensar un poco antes de escribir evita rehacer mucho después. El diseño aquí es una "
                "herramienta de conversación, no un entregable.",
                enlace="ADOO-UML la estáis viendo ahora (mitad del temario). Patrones y BPMN os llegan más adelante."))

    d.content("Quién hace qué: casos de uso", ch("adoo"), 4, N(
        "Casos de uso con nuestro ejemplo: dos actores y cuatro funciones. Es la forma más rápida de comprobar con el "
        "usuario que hablamos del mismo sistema.",
        pregunta="¿Falta algún actor? ¿Quién más toca estos datos?",
        corto="Enseña solo los actores y las cuatro funciones."), uc_diagram)
    d.content("Qué cosas hay: diagrama de clases", ch("adoo", "poo"), 4, N(
        "El modelo del dominio con tres clases y un enumerado. Es la fotografía de «qué cosas existen» y cómo se relacionan. "
        "Anticipa que más adelante este mismo modelo será una tabla y una clase Java.",
        pregunta="¿Dónde pondríais la fecha: en la solicitud o en el material? ¿Por qué?"), class_diagram)
    d.content("Quién llama a quién: secuencia", ch("adoo"), 4, N(
        "Sigue la petición de arriba abajo. Es el diagrama que más ayuda a depurar: cada flecha es un sitio donde algo "
        "puede fallar. Lo retomaremos en el bloque de integración.",
        pregunta="¿En qué flecha pondríais primero un log?",
        corto="Cuenta solo el camino de ida y el de vuelta."), seq_diagram)

    def zoom(sl):
        cards(sl, [("Contexto", "¿Quién usa el sistema y con qué otros sistemas habla?", "Una caja y flechas"),
                   ("Contenedores", "¿Qué piezas se despliegan por separado: front, API, base de datos?", "Lo que desplegamos"),
                   ("Componentes", "¿Qué hay dentro de cada pieza?", "Lo que programamos")],
              y=1.95, h=3.55, accent=TEAL)
        banner(sl, "El diagrama es una herramienta de conversación, no un entregable.", y=5.75, h=0.75, size=20)
    d.content("Dibuja a tres niveles de zoom", ch("adoo"), 3, N(
        "Modelo C4 simplificado. Cuando dibujo una arquitectura empiezo por el contexto y solo bajo al detalle si lo "
        "necesita la conversación. Un diagrama perfecto que nadie mira vale menos que uno a mano que alinea a todo el equipo.",
        pregunta="¿Qué nivel os piden normalmente en la documentación de un proyecto?"), zoom)

    d.content("BPMN: el proceso primero", ch("bpmn"), 5, N(
        "Un proceso mal entendido produce un sistema que no sirve. Con BPMN ves quién hace cada paso, qué decide y qué se "
        "automatiza. Pide a la sala un proceso militar cotidiano (por ejemplo, una solicitud de material o un parte de "
        "incidencias) y dibújalo en la pizarra en dos minutos.",
        pregunta="¿Qué parte de un proceso real de vuestra unidad se podría automatizar y cuál conviene que siga siendo manual?",
        enlace="BPMN y gestión por procesos es una asignatura que os llega más adelante."), bpmn_diagram)

    d.cards_slide("Soluciones conocidas a problemas repetidos", ch("pat", "poo"), [
        ("Inyección de dependencias", "Una clase pide lo que necesita y el framework se lo da.", "En Quarkus: @Inject"),
        ("Repositorio", "El acceso a datos queda detrás de una interfaz.", "Cambia la BD sin tocar la lógica"),
        ("DTO", "Lo que viaja por la red no es lo que guarda la base de datos.", "Evita exponer tu modelo interno"),
        ("Estrategia", "Distintas reglas intercambiables según el caso.", "Sustituye un «if» gigante")],
        4, N(
            "No hace falta memorizar los 23 patrones clásicos. Enseña los cuatro que ves cada semana. El valor es el "
            "vocabulario común: decir «eso es un repositorio» ahorra una explicación.",
            pregunta="¿Cuál de estos habéis usado sin saber cómo se llamaba?",
            historia="El patrón que más has usado en Quarkus y un caso en el que te evitó un problema.",
            enlace="Patrones y pruebas os llega más adelante. Esto es el porqué."),
        key="Un patrón es un nombre compartido para una solución que ya funciona.", h=3.55)

    d.code("Decisiones: por qué elegimos A y no B", ch("herr"), """# ADR-007: Postgres para solicitudes

Estado: aceptada · 2026-11-09

## Contexto
Necesitamos consultas por unidad y semana.

## Decisión
PostgreSQL, con índice (unidad_id, fecha).

## Alternativas descartadas
Fichero plano · NoSQL de documentos.

## Consecuencias
Hay que mantener migraciones de esquema.""", "md", "docs/adr/007-postgres.md", [
        ("Una página, no un tratado", "Contexto, decisión, alternativas y consecuencias."),
        ("Responde a «¿por qué así?»", "Dentro de un año nadie lo recordará."),
        ("Se versiona con el código", "Va en el repositorio, junto al cambio.")], 4, N(
        "Un registro de decisiones de arquitectura (ADR) cabe en media página. Es documentación barata y muy útil cuando "
        "entra alguien nuevo. Usa este ejemplo para explicar cómo documentas elecciones como Postgres frente a otros motores.",
        pregunta="¿Qué decisión de vuestro proyecto de clase habríais escrito así?",
        historia="Una decisión que hoy nadie sabe explicar y que un ADR habría resuelto.",
        enlace="Se escribe en Markdown, que veis en Herramientas colaborativas."), size=13, code_w=7.3)

    d.ask("Dibuja con tu pareja", "Dibujad la arquitectura de vuestro proyecto",
          "Cajas y flechas, sin UML perfecto. Incluid todo lo que, si falla, rompe el sistema.",
          ["Cinco minutos por parejas, en folio o pizarra.",
           "Incluid: usuario, front, back, base de datos, sistemas externos y dónde se despliega.",
           "Dos parejas enseñan su dibujo; el resto busca qué falta."], 12, N(
              "Dinámica. Ten preparada una pizarra o folios. Cuando enseñen, busca qué falta casi siempre: la base de datos, las "
              "dependencias externas, la autenticación y el despliegue. No corrijas el dibujo: pregunta «¿qué pasa si esto se cae?».",
              pregunta="¿Qué caja os ha costado más dibujar?",
              corto="Hazlo en 8 minutos con una sola pareja que enseña."), sub_label="Reglas")

    def faltan(sl):
        items = ["La base de datos", "Los sistemas externos de los que depende", "Cómo se autentica la gente",
                 "Dónde se despliega y con qué entornos", "Qué pasa cuando algo falla"]
        rr(sl, MX, 1.95, 6.6, 3.9, fill=LIGHT, r=0.04)
        tb(sl, MX + 0.35, 2.1, 6, 0.3, "LO QUE CASI SIEMPRE FALTA EN EL DIBUJO", size=12, bold=True, color=RUST)
        checklist(sl, MX + 0.35, 2.6, 6, items, size=17.5, gap=0.62, col=RUST)
        callouts(sl, 7.7, 1.95, W - MX - 7.7, [("Los dibujos que no incluyen datos", "suelen romperse en la base de datos."),
                                               ("Los que no incluyen dependencias", "suelen romperse en la integración."),
                                               ("Los que no incluyen el despliegue", "suelen romperse en producción.")],
                 h=1.18, gap=0.18)
        banner(sl, "Dibujar bien es dibujar lo que da miedo, no lo que ya controlas.", y=6.05, h=0.65, size=18)
    d.content("Lo que casi siempre falta", ch("adoo"), 4, N(
        "Cierra la dinámica con el patrón que has visto en los dibujos. Vuelve sobre los dibujos de la sala y marca qué cajas "
        "faltaban. Anticipa que los tres siguientes bloques cubren justo lo que suele faltar.",
        pregunta="¿Qué añadiríais a vuestro dibujo después de ver esta lista?"), faltan)
    d.expect("10:00")

    # ------------------------------------------------------------------ B3 DATOS 10:00–10:50
    d.section(2, "Datos", "Los datos viven más que el código", ch("bbdd", "sql", "oracle", "nosql"), 50,
              "¿Quién ha escrito alguna vez una consulta SQL? ¿Y quién ha tenido que arreglar datos mal guardados?",
              N("Abre el bloque. Aquí la mayoría aún no ha dado nada de bases de datos, así que es un adelanto: qué os van "
                "a enseñar y por qué importa tanto.",
                enlace="Todo el módulo de bases de datos (modelado, SQL, Oracle y PL/SQL, NoSQL) todavía no tiene fecha en el calendario."))

    d.statement("Una aplicación se reescribe. Los datos se migran.",
                "Cuando cambia el front o el framework, los datos de hace diez años siguen ahí.", 1, N(
                    "Idea fuerza del bloque: el código se puede rehacer, los datos no. Por eso se diseñan con más cuidado."))

    def modelo(sl):
        def ent(x, y, w, name, attrs):
            box(sl, x, y, w, 0.5, name, fill=TEAL, color=WHITE, size=15, bold=True, kind=MSO_SHAPE.RECTANGLE, r=None)
            rr(sl, x, y + 0.5, w, 0.33 * len(attrs) + 0.2, fill=LIGHT, line=TEAL, lw=1.25, r=0.0)
            tb(sl, x + 0.15, y + 0.58, w - 0.3, 0.33 * len(attrs), attrs, size=12.5, font=MONO)
        ent(0.7, 2.0, 1.9, "unidad", ["id PK", "nombre"])
        ent(3.3, 2.0, 2.2, "solicitud", ["id PK", "unidad_id FK", "material_id FK", "cantidad", "fecha", "estado"])
        ent(0.7, 3.9, 1.9, "material", ["id PK", "nombre"])
        line(sl, 3.3, 2.75, 2.6, 2.75, color=INK, w=1.75, arrow=True)
        line(sl, 3.3, 4.2, 2.6, 4.65, color=INK, w=1.75, arrow=True)
        code_panel(sl, 5.7, 1.95, 6.9, """CREATE TABLE solicitud (
  id          BIGINT PRIMARY KEY,
  unidad_id   BIGINT NOT NULL REFERENCES unidad(id),
  material_id BIGINT NOT NULL REFERENCES material(id),
  cantidad    INT    NOT NULL CHECK (cantidad > 0),
  fecha       DATE   NOT NULL,
  estado      VARCHAR(12) NOT NULL DEFAULT 'PENDIENTE'
);

CREATE INDEX idx_solicitud_unidad_fecha
  ON solicitud (unidad_id, fecha);""", "sql", size=12.5, fname="V1__solicitudes.sql")
        banner(sl, "Las reglas del dominio viven también en la base de datos: claves, NOT NULL y CHECK.", y=5.95, h=0.7, size=17)
    d.content("Del modelo a la tabla", ch("bbdd", "sql"), 5, N(
        "Pasa del diagrama de clases a las tablas: cada entidad una tabla, cada relación una clave foránea. Fíjate en que la base "
        "de datos también valida (cantidad > 0, NOT NULL): es la última línea de defensa si el código falla.",
        pregunta="¿Qué pasaría si no pusiéramos la restricción de cantidad > 0?",
        historia="Un dato corrupto en una base real y cuánto costó limpiarlo.",
        enlace="Teoría general de bases de datos y modelado; SQL."), modelo)

    def normal(sl):
        tb(sl, MX, 1.85, 6, 0.3, "SIN NORMALIZAR", size=12, bold=True, color=RUST)
        mini_table(sl, MX, 2.2, ["unidad", "material", "cant."],
                   [["U7", "Baterías AA", "12"], ["Unidad 7", "baterías aa", "5"], ["U-7", "Baterías AA ", "3"]],
                   [1.6, 2.2, 1.0], size=13, head_fill=RUST, row_h=0.5)
        tb(sl, MX, 4.35, 5.3, 0.9, "¿Cuántas baterías pidió la unidad 7? Depende de cómo cuentes.", size=17, color=INK, italic=True)
        shape(sl, MSO_SHAPE.RIGHT_ARROW, 6.2, 2.8, 0.7, 0.6, fill=AMBER)
        tb(sl, 7.2, 1.85, 6, 0.3, "NORMALIZADO", size=12, bold=True, color=GREEN)
        mini_table(sl, 7.2, 2.2, ["unidad_id", "material_id", "cant."], [["7", "3", "12"], ["7", "3", "5"], ["7", "3", "3"]],
                   [1.55, 1.75, 1.0], size=13, head_fill=GREEN, row_h=0.5)
        mini_table(sl, 7.2, 4.35, ["id", "nombre"], [["3", "Baterías AA"]], [0.8, 3.0], size=13, head_fill=GREEN, row_h=0.45)
        banner(sl, "Un dato, un sitio. Lo demás son referencias.", y=5.85, h=0.7, size=20)
    d.content("Normalizar: un dato, un sitio", ch("bbdd"), 4, N(
        "Compara la tabla donde cada fila escribe el nombre a su manera con la versión con claves. Es el motivo de la "
        "normalización: poder contar, buscar y corregir en un solo punto.",
        pregunta="¿Alguna vez habéis tenido que unir dos hojas de Excel con nombres escritos de forma distinta?",
        enlace="TG Bases de datos y modelado."), normal)

    d.code("SQL en el día a día", ch("sql"), """SELECT m.nombre,
       SUM(s.cantidad) AS total
FROM   solicitud s
JOIN   material  m ON m.id = s.material_id
WHERE  s.unidad_id = 7
  AND  s.fecha >= DATE '2026-11-02'
  AND  s.fecha <  DATE '2026-11-09'
GROUP BY m.nombre
ORDER BY total DESC;""", "sql", "consulta-semanal.sql", [
        ("Filtros y agregados", "WHERE reduce, GROUP BY resume."),
        ("JOIN", "Une tablas por sus claves."),
        ("Rangos de fecha medio abiertos", "Evitan errores en el último día.")], 5, N(
        "Esta es la consulta que responde a nuestra historia. Recórrela línea a línea: de dónde sale cada dato. Resalta el rango "
        "de fechas con >= y <, que evita el clásico error de perder las del último día.",
        pregunta="¿Qué devolvería si quitáramos el GROUP BY?",
        historia="Una consulta tuya real (anonimizada) y cómo la mejoraste.",
        enlace="Lenguaje SQL. Postgres y DB2 aceptan casi todo esto; cada motor tiene sus matices."), size=14, code_w=7.0)

    def indices(sl):
        half = (CW - 0.4) / 2
        for k, (hd, col, txt, note) in enumerate([
            ("Sin índice", RUST, """Seq Scan on solicitud
  Filter: (unidad_id = 7 AND
           fecha >= '2026-11-02')
  Rows Removed by Filter: 9 999 000""", "Lee toda la tabla"),
            ("Con índice (unidad_id, fecha)", GREEN, """Index Scan using
  idx_solicitud_unidad_fecha
  Index Cond: (unidad_id = 7 AND
               fecha >= '2026-11-02')""", "Va directo a las filas")]):
            x = MX + k * (half + 0.4)
            rr(sl, x, 1.95, half, 0.55, fill=col, r=0.2)
            tb(sl, x + 0.3, 1.95, half - 0.6, 0.55, hd, size=18, bold=True, color=WHITE, anchor="m")
            code_panel(sl, x, 2.65, half, txt, "text", size=13, h=1.9)
            tb(sl, x, 4.7, half, 0.4, note, size=18, bold=True, color=col, font=HEAD)
        tb(sl, MX, 5.2, CW, 0.3, "Salida de EXPLAIN en Postgres, ilustrativa y abreviada.", size=12, color=SLATE, italic=True)
        banner(sl, "Antes de optimizar, mide: pide siempre el plan de ejecución.", y=5.7, h=0.8, size=19, label="Regla")
    d.content("Índices: rendimiento con volumen", ch("bbdd", "sql"), 5, N(
        "Con cien filas todo funciona; con millones, no. Enseña la diferencia entre leer toda la tabla y usar el índice. "
        "Matiza: un índice acelera las lecturas pero ralentiza las escrituras y ocupa espacio. Se crea para consultas "
        "concretas, no «por si acaso».",
        pregunta="¿En qué columnas pondríais un índice para nuestra consulta y por qué en ese orden?",
        historia="Una consulta lenta de tu proyecto y cómo diste con la causa."), indices)

    d.code("Transacciones: todo o nada", ch("bbdd"), """@Transactional
public void aprobar(Long id) {
    Solicitud s = Solicitud.findById(id);
    stock.descontar(s.material, s.cantidad);
    s.estado = Estado.APROBADA;
}""", "java", "SolicitudService.java", [
        ("Atomicidad", "O se hacen las dos cosas o ninguna."),
        ("Consistencia y aislamiento", "Nadie ve datos a medias."),
        ("Durabilidad", "Lo confirmado sobrevive a una caída.")], 5, N(
        "Aprobar una solicitud implica descontar stock y cambiar su estado. Si falla a mitad, no puede quedar descontado "
        "pero sin aprobar. Eso es una transacción. En Quarkus basta con @Transactional; las propiedades ACID son lo que "
        "hay debajo.",
        pregunta="¿Qué pasaría en este código si no hubiera transacción y la segunda línea lanzara un error?",
        historia="Un fallo real por no usar una transacción (o por abrirla demasiado grande).",
        enlace="Teoría general de bases de datos: ACID y concurrencia."), size=14.5, code_w=7.0)

    def migra(sl):
        steps = [("1. Añadir", "Columna nueva, opcional."), ("2. Desplegar", "El código escribe en las dos."),
                 ("3. Rellenar", "Se copian los datos antiguos."), ("4. Limpiar", "Se quita lo viejo.")]
        w = (CW - 3 * 0.25) / 4
        for i, (a, b) in enumerate(steps):
            x = MX + i * (w + 0.25)
            box(sl, x, 1.95, w, 0.8, a, fill=TEAL if i % 2 == 0 else "3C7F85", color=WHITE, size=18, bold=True,
                kind=MSO_SHAPE.PENTAGON, r=None, font=HEAD)
            tb(sl, x + 0.1, 2.9, w - 0.2, 0.8, b, size=16, color=INK)
        code_panel(sl, MX, 3.85, 6.5, """-- V3__solicitud_prioridad.sql
ALTER TABLE solicitud
  ADD COLUMN prioridad VARCHAR(10);""", "sql", size=13, fname="src/main/resources/db/migration")
        callouts(sl, 7.5, 3.85, W - MX - 7.5, [("Versionadas", "Como el código: una migración por cambio."),
                                               ("Repetibles", "Cualquier entorno llega al mismo esquema.")],
                 h=0.98, gap=0.15)
        banner(sl, "Cambiar el esquema sin romper a nadie: en dos despliegues, no en uno.", y=6.0, h=0.65, size=17)
    d.content("Migraciones: cambiar el esquema sin romper", ch("sql"), 5, N(
        "Un cambio de esquema se despliega por fases para que la versión vieja y la nueva convivan. Herramientas como Flyway o "
        "Liquibase versionan las migraciones junto al código (Quarkus las integra).",
        pregunta="¿Qué pasa con el código antiguo si una migración elimina una columna que aún usa?",
        historia="Una migración que salió mal y cómo se hizo marcha atrás."), migra)

    def motor(sl):
        rows = [["PostgreSQL", "Aplicaciones nuevas, código abierto, soporta JSON", "Operación y ajuste a tu cargo"],
                ["DB2", "Entornos corporativos y mainframe con mucho volumen", "Licencias y dialecto propio"],
                ["Oracle y PL/SQL", "Sistemas con mucha lógica dentro de la BD", "Lógica difícil de versionar y probar"],
                ["NoSQL", "Datos sin esquema fijo o lecturas masivas", "Sin joins ni transacciones complejas (según motor)"]]
        mini_table(sl, MX, 1.95, ["Motor", "Encaja cuando…", "Ojo con…"], rows, [2.5, 5.0, 4.4], size=15, row_h=0.74)
        banner(sl, "Cada motor encaja con un caso. Elegir es entender el caso, no defender una marca.", y=6.0, h=0.7, size=18)
    d.content("Elegir el motor", ch("bbdd", "oracle", "nosql"), 4, N(
        "Mantén un tono neutral. Cada motor encaja con un caso: explica por qué usas Postgres y DB2 en tu día a día y cuándo "
        "has visto elegir otro. Anticipa que en el curso verán Oracle y PL/SQL y NoSQL.",
        pregunta="¿Cuál elegiríais para un registro con millones de lecturas al día y casi ninguna escritura?",
        historia="Por qué en tu proyecto se eligió cada motor.",
        enlace="Oracle y PL/SQL, NoSQL y modelado están en el módulo de Bases de datos."), motor)

    def quiz(sl, answers=False):
        qs = [("Un índice siempre acelera las consultas.", "F", "También ralentiza las escrituras y no sirve si no se usa."),
              ("Si el servidor se cae a mitad de una transacción, los datos quedan a medias.", "F", "La atomicidad lo impide: se deshace entera."),
              ("«SELECT *» es inocuo en consultas de producción.", "F", "Trae columnas que no necesitas y rompe si cambia la tabla."),
              ("NoSQL significa que no hace falta diseñar los datos.", "F", "Hay que diseñar igual, solo que de otra forma.")]
        for i, (q, a, e) in enumerate(qs):
            y = 1.95 + i * 1.18
            rr(sl, MX, y, CW, 1.05, fill=LIGHT, r=0.08)
            num_circle(sl, MX + 0.25, y + 0.28, i + 1, fill=TEAL)
            tb(sl, MX + 1.0, y + 0.08, 8.0 if not answers else 6.4, 0.9, q, size=17.5, color=INK, anchor="m")
            if answers:
                box(sl, MX + 7.6, y + 0.27, 0.9, 0.5, "FALSO" if a == "F" else "VERDAD", fill=RUST if a == "F" else GREEN,
                    color=WHITE, size=13, bold=True, r=0.5)
                tb(sl, MX + 8.65, y + 0.08, CW - 8.85, 0.9, e, size=14, color=SLATE, anchor="m")
            else:
                for k, lab in enumerate(["V", "F"]):
                    box(sl, W - MX - 1.9 + k * 0.95, y + 0.25, 0.8, 0.55, lab, fill=WHITE, line=TEAL, lw=2, color=TEAL,
                        size=20, bold=True, r=0.3)
    d.content("¿Verdadero o falso?", ch("bbdd", "sql"), 3, N(
        "Votación a mano alzada, afirmación por afirmación. Pide que alguien defienda su voto antes de revelar la respuesta.",
        pregunta="¿Quién ha votado distinto al resto? ¿Por qué?",
        corto="Haz solo las dos primeras."), lambda sl: quiz(sl, False))
    d.content("Las respuestas", ch("bbdd", "sql"), 2, N(
        "Las cuatro son falsas a propósito: sirven para desmontar mitos habituales. Entra en la que más debate haya tenido."),
        lambda sl: quiz(sl, True))

    d.ask("Caso roto", "La consulta lenta",
          "Funciona con 100 filas y falla con 10 millones.\n\n¿Dónde miráis primero?",
          ["Cada grupo propone una hipótesis y dice cómo la comprobaría.",
           "Después comparamos con el orden razonable.",
           "Pista: no vale «añadir más máquina» como primera respuesta."], 10, N(
              "Que propongan hipótesis en voz alta. Después enseña el orden que sigues: plan de ejecución, índices, volumen de "
              "datos, bloqueos y, solo al final, hardware.",
              pregunta="¿Qué mediríais antes de tocar nada?",
              historia="Una lentitud real que diagnosticaste y qué resultó ser.",
              corto="Pide una hipótesis por grupo y cuenta la solución.",
              enlace="Es el puente entre SQL, modelado y Linux (mirar recursos)."))
    d.expect("10:50")

    d.pause(30, N(
        "Pausa de 30 minutos (10:50–11:20), la del calendario. Aprovecha para mirar el muro de dudas, recoger las que "
        "más veces se hayan repetido y tomar algo. Al volver, retoma una duda pendiente como calentamiento de 2 minutos."))
    d.expect("11:20")

    # ------------------------------------------------------------------ B4 CÓDIGO 11:20–12:05
    d.section(3, "Código", "Ahora sí se programa", ch("poo", "alg", "git", "lib"), 45,
              "¿En qué lenguaje trabajáis y cómo repartís el código cuando sois varios?",
              N("Abre la segunda mitad. El código es solo una parte del trabajo, pero es donde se juntan todas las decisiones "
                "anteriores.",
                enlace="Algoritmos casi terminada, POO empieza esta semana, Git en marcha. Librerías y pruebas os llegan más adelante."))

    def formas(sl):
        pw, ph = (CW - 0.3) / 2, 1.95
        pans = [("tabla  ·  SQL", "sql", "CREATE TABLE solicitud (\n  cantidad INT NOT NULL\n);"),
                ("entidad  ·  Java", "java", "@Entity\npublic class Solicitud\n        extends PanacheEntity {\n    public int cantidad;\n}"),
                ("JSON  ·  lo que viaja", "json", '{ "id": 42,\n  "material": "Baterías AA",\n  "cantidad": 12 }'),
                ("formulario  ·  HTML", "html", '<input name="cantidad"\n       type="number" min="1">')]
        for i, (nm, lang, c) in enumerate(pans):
            x = MX + (i % 2) * (pw + 0.3)
            y = 1.85 + (i // 2) * (ph + 0.15)
            code_panel(sl, x, y, pw, c, lang, size=13, h=ph, fname=nm)
        banner(sl, "Un mismo concepto, cuatro formas: ahí nacen los problemas de integración.", y=6.2, h=0.55, size=16)
    d.content("Una idea, cuatro formas", ch("poo", "html", "sql"), 5, N(
        "La «cantidad» de una solicitud existe como columna, atributo Java, campo JSON y campo de formulario. Cada capa la "
        "llama, la valida y la formatea a su manera. Los errores de integración nacen casi siempre en este paso.",
        pregunta="¿Dónde validaríais que la cantidad sea mayor que cero? ¿En una capa o en todas?",
        historia="Un error clásico por un nombre o tipo distinto entre capas."), formas)

    def capas(sl):
        layers = [("Resource", "Recibe la petición HTTP", TEAL), ("Service", "Aplica las reglas de negocio", "3C7F85"),
                  ("Repository", "Habla con la base de datos", TEAL)]
        for i, (a, b, col) in enumerate(layers):
            y = 1.95 + i * 1.3
            box(sl, MX, y, 2.2, 0.95, a, fill=col, color=WHITE, size=17, bold=True, font=MONO)
            tb(sl, MX + 2.35, y, 2.3, 0.95, b, size=15, color=INK, anchor="m")
            if i < 2:
                shape(sl, MSO_SHAPE.DOWN_ARROW, MX + 0.9, y + 0.98, 0.4, 0.3, fill=AMBER)
        code_panel(sl, 5.7, 1.95, 6.9, """@ApplicationScoped
public class SolicitudService {

    private final SolicitudRepository repo;

    SolicitudService(SolicitudRepository repo) {
        this.repo = repo;
    }

    public List<Solicitud> deLaSemana(long unidad) {
        return repo.porUnidadYFecha(unidad, inicio(), fin());
    }
}""", "java", size=12.5, fname="SolicitudService.java")
        banner(sl, "Cada capa hace una cosa: más fácil de leer, de probar y de cambiar.", y=6.0, h=0.7, size=18)
    d.content("POO en la práctica: capas e inyección", ch("poo", "pat"), 5, N(
        "Pasa de la teoría a un ejemplo: tres clases con una responsabilidad cada una. El servicio no crea su repositorio, "
        "se lo dan (inyección de dependencias), y por eso se puede probar con uno falso. Es la POO que más se usa en un servicio.",
        pregunta="¿Qué cambiaría si el servicio hiciera «new SolicitudRepository()» dentro?",
        enlace="POO fundamentos empieza esta semana. Patrones os llega más adelante."), capas)

    def alg(sl):
        ns = list(range(10, 101, 10))
        cd = CategoryChartData()
        cd.categories = [str(n) for n in ns]
        cd.add_series("n (una pasada)", ns)
        cd.add_series("n · log n (ordenar)", [round(n * math.log2(n)) for n in ns])
        cd.add_series("n² (comparar todo con todo)", [n * n for n in ns])
        gf = sl.shapes.add_chart(XL_CHART_TYPE.LINE, Inches(MX), Inches(1.9), Inches(6.9), Inches(4.0), cd)
        chart = gf.chart
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.include_in_layout = False
        chart.legend.font.size = Pt(12)
        for s, col in zip(chart.plots[0].series, [TEAL, AMBER, RUST]):
            s.format.line.color.rgb = rgb(col)
            s.format.line.width = Pt(3.5)
            s.smooth = False
            s.marker.style = XL_MARKER_STYLE.NONE
        chart.value_axis.tick_labels.font.size = Pt(11)
        chart.category_axis.tick_labels.font.size = Pt(11)
        chart.value_axis.major_gridlines.format.line.color.rgb = rgb("D5DDE0")
        chart.category_axis.has_title = True
        chart.category_axis.axis_title.text_frame.text = "Número de elementos"
        chart.category_axis.axis_title.text_frame.paragraphs[0].runs[0].font.size = Pt(11)
        code_panel(sl, 7.85, 1.95, W - MX - 7.85, """// Lista: recorre todo en cada búsqueda
List<String> vistos = new ArrayList<>();
vistos.contains(id);   // O(n)

// Conjunto: salta directo
Set<String> vistos = new HashSet<>();
vistos.contains(id);   // O(1) de media""", "java", size=11.5, h=2.55, fname="Vistos.java")
        callouts(sl, 7.85, 4.7, W - MX - 7.85, [("Elegir bien lista, mapa o conjunto", "Es la decisión que más rendimiento da.")],
                 h=1.2)
        banner(sl, "No hace falta calcular: hace falta saber qué crece rápido y qué no.", y=6.1, h=0.62, size=17)
    d.content("Algoritmos que importan", ch("alg"), 4, N(
        "La gráfica enseña por qué importa la complejidad: con pocos datos todo es rápido; con muchos, lo cuadrático se "
        "dispara. Un ejemplo real es buscar duplicados en una lista con muchos elementos: pasar a un conjunto convierte "
        "minutos en milisegundos. Tono práctico, no teórico.",
        pregunta="¿Cuántas comparaciones hace «comparar todo con todo» con mil elementos?",
        historia="Un problema de rendimiento causado por una mala estructura de datos.",
        enlace="Algoritmos y estructuras de datos: casi terminada."), alg)

    d.compare("Código que se entiende en seis meses", ch("poo", "alg"),
              ("Difícil de leer", """public List<Object[]> g(long u, int s) {
    int d = 7; // días
    ...
}""", RUST),
              ("Fácil de leer", """// Urgentes primero, por orden del jefe
public List<SolicitudDto> deLaSemana(
        long unidadId, Semana semana) {
    int DIAS_POR_SEMANA = 7;
    ...
}""", GREEN), 4, N(
        "Compara los dos fragmentos: nombres que explican, tipos del dominio y un comentario que cuenta el porqué, no el qué. "
        "La persona que lo lea dentro de seis meses (probablemente tú) lo agradecerá.",
        pregunta="¿Qué os cuesta más: entender código ajeno o acordaros del vuestro después de un mes?"),
        key="Los comentarios explican el porqué. El código ya cuenta el qué.", lang="java", size=15, mono=True)

    def tests(sl):
        pyramid(sl, MX, 1.95)
        code_panel(sl, 6.6, 1.95, W - MX - 6.6, """@QuarkusTest
class SolicitudResourceTest {

    @Test
    void listaSoloLasDeMiUnidad() {
        given().when()
            .get("/unidades/7/solicitudes")
            .then().statusCode(200)
            .body("unidad", everyItem(is("U7")));
    }
}""", "java", size=12, fname="SolicitudResourceTest.java")
        banner(sl, "Cada criterio de aceptación se convierte en una prueba que se ejecuta sola.", y=6.0, h=0.7, size=17)
    d.content("Pruebas: la red de seguridad", ch("pat"), 4, N(
        "Pirámide: muchas pruebas pequeñas y rápidas, algunas de integración y pocas de extremo a extremo. El ejemplo es el "
        "criterio de aceptación de nuestra historia convertido en código. Gracias a ellas se puede cambiar código sin miedo.",
        pregunta="¿Qué pasa en un proyecto sin pruebas cuando alguien cambia una línea?",
        historia="Una prueba automática que te salvó de romper algo en producción.",
        enlace="Técnicas y prácticas: pruebas, patrones y refactorización, os llega más adelante."), tests)

    d.cards_slide("No reinventar la rueda", ch("lib", "herr"), [
        ("Dependencias", "Maven o Gradle declaran qué usas y con qué versión.", "Gradle lo veréis en Herramientas colaborativas"),
        ("Frameworks", "Quarkus: arranque rápido, servicios REST, inyección y configuración.", "Resuelve lo repetido"),
        ("Antes de añadir una", "¿Está mantenida? ¿Qué licencia tiene? ¿Tiene vulnerabilidades conocidas? ¿La necesito?", None)],
        4, N(
            "Reutilizar es bueno, con criterio. Cada dependencia es código ajeno que mantienes y una puerta de entrada de "
            "vulnerabilidades. Menciona herramientas de revisión automática de dependencias.",
            pregunta="¿Cuántas dependencias creéis que tiene un servicio típico? (adivinad una cifra)",
            historia="Una dependencia que causó un problema (versión, licencia, vulnerabilidad).",
            enlace="Implementación y reutilización de librerías; Gradle forma parte de Herramientas colaborativas."),
        key="Cada librería que añades es código ajeno que ahora también es tuyo.", h=3.55)

    def gitslide(sl):
        git_diagram(sl)
        items = [("Ramas", "Cada cambio, aislado."), ("Pull request", "Revisión por otra persona."),
                 ("Commits claros", "Una historia que se entiende."), ("Releases", "Versiones etiquetadas y trazables.")]
        w = (CW - 3 * 0.2) / 4
        for i, (a, b) in enumerate(items):
            x = MX + i * (w + 0.2)
            rr(sl, x, 4.5, w, 1.75, fill=LIGHT, r=0.06)
            shape(sl, MSO_SHAPE.RECTANGLE, x, 4.5, w, 0.06, fill=AMBER)
            tb(sl, x + 0.22, 4.7, w - 0.4, 0.4, a, size=18, bold=True, font=HEAD, color=TEAL_D)
            tb(sl, x + 0.22, 5.2, w - 0.4, 1.0, b, size=15, color=INK)
    d.content("Git en equipo", ch("git"), 5, N(
        "Recorre el dibujo: main siempre estable, una rama por cambio, una pull request para revisar y una etiqueta para marcar "
        "una versión. Es el flujo que, tal cual, usarás en el trabajo.",
        pregunta="¿Qué problema habéis tenido ya con Git (conflictos, pérdidas de cambios)?",
        historia="El flujo de ramas de tu equipo y cómo se revisa una PR en la práctica.",
        enlace="Git son 22 de las 50 sesiones de Herramientas colaborativas y se está dando ahora."), gitslide)

    d.compare("Commits que cuentan una historia", ch("git"),
              ("Historial que no ayuda", """fix
cambios
arreglo definitivo
asdf
ahora sí""", RUST),
              ("Historial que se lee", """PROJ-101 Lista solicitudes por unidad
PROJ-103 Añade índice (unidad_id, fecha)
PROJ-130 Corrige zona horaria en fechas
PROJ-131 Prueba del caso sin resultados""", GREEN), 3, N(
        "El historial de commits es documentación. Un buen mensaje enlaza con el ticket y dice qué cambia. Más adelante, "
        "gracias a eso, se genera el listado de cambios de cada versión.",
        pregunta="¿Cuál de los dos historiales os gustaría heredar?",
        enlace="Git y Markdown en Herramientas colaborativas."),
        key="Un buen commit enlaza con el ticket y dice qué cambia.", lang="text", size=14, mono=True)

    d.ask("Revisión de código", "Revisa este Pull Request", "",
          prompt_code=("""+ public List<Solicitud> listar(Long unidad) {
+   String sql = "SELECT * FROM solicitud"
+       + " WHERE unidad_id = " + unidad;
+   // TODO: arreglar
+   log.info("password=" + config.password);
+   return em.createNativeQuery(sql)
+            .getResultList();
+ }""", "diff"), steps=["Por parejas, buscad todos los problemas que veáis.", "Apuntad el más grave.",
       "Cinco minutos. Luego los repasamos juntos."], m=7, notes=N(
          "Proyecta el código. Problemas que deberían salir: inyección SQL por concatenar; contraseña en el log; SELECT *; "
          "devuelve la entidad directamente; sin paginar; TODO sin ticket; sin pruebas; el nombre «listar» es genérico.",
          pregunta="¿Cuál es el más grave y por qué?",
          historia="Algo que se te ha colado en una revisión y cómo lo detectaste después.",
          corto="Pide solo que encuentren el problema de seguridad."), sub_label="Reglas")
    d.content("Lo que debería detectar la revisión", ch("git", "seg"), 3, N(
        "Repasa los cinco puntos con la sala. La seguridad (inyección y contraseña en el log) va primero porque es lo que no se "
        "puede arreglar después. Comenta que una PR pequeña es mucho más fácil de revisar.",
        pregunta="¿Cuántos habíais encontrado? ¿Cuál se os escapó?"),
        lambda sl: (cards(sl, [("Seguridad", "Concatenar SQL abre la puerta a inyección. La contraseña, al log."), ("Rendimiento", "SELECT * y sin paginar: con millones de filas no escala."),
                               ("Mantenimiento", "TODO sin ticket, nombre genérico, sin pruebas.")], y=1.95, h=3.7, accent=RUST),
                    banner(sl, "Una revisión sirve para lo que las pruebas no ven.", y=5.9, h=0.7, size=19)))
    d.expect("12:05")

    # ------------------------------------------------------------------ B5 INTEGRACIÓN 12:05–12:55
    d.section(4, "Unir las capas", "El punto que más cuesta", ch("ws", "ser", "html", "ecma", "front", "post"), 50,
              "¿Qué os ha costado más al unir pantalla, API y base de datos? ¿Cómo lo descubristeis?",
              N("Este es el bloque estrella: lo que las encuestas del equipo suelen señalar como más difícil. Dedícale tiempo real.",
                enlace="Servicios web, serialización, ECMAScript y frontend/backend aún no se han dado. HTML/CSS sí, y Postman llegará en Herramientas colaborativas."))

    def peticion(sl):
        steps = [("Pantalla", "El usuario pulsa"), ("API REST", "Ruta y verbo"), ("Servicio", "Lógica de negocio"),
                 ("Base de datos", "Consulta SQL"), ("Respuesta", "JSON a la pantalla")]
        fails = ["Botón sin acción o validación", "404 o 405: ruta o verbo", "Excepción sin controlar: 500",
                 "Consulta lenta o sin datos", "JSON distinto del acordado"]
        n = 5
        sw = (CW + (n - 1) * 0.14) / n
        for i, (a, b) in enumerate(steps):
            x = MX + i * (sw - 0.14)
            box(sl, x, 1.95, sw, 0.95, a, fill=TEAL if i % 2 == 0 else "3C7F85", color=WHITE, size=16.5, bold=True,
                kind=MSO_SHAPE.PENTAGON if i == 0 else MSO_SHAPE.CHEVRON, r=None, pad=0.3)
            tb(sl, x + 0.3, 3.0, sw - 0.6, 0.4, b, size=14.5, color=INK, align="c")
        rr(sl, MX, 3.6, CW, 1.65, fill=LIGHT, r=0.05)
        tb(sl, MX + 0.3, 3.7, 5, 0.3, "¿QUÉ PUEDE FALLAR EN CADA SALTO?", size=11.5, bold=True, color=RUST)
        for i, f in enumerate(fails):
            x = MX + i * (sw - 0.14)
            tb(sl, x + 0.3, 4.2, sw - 0.6, 1.0, f, size=14.5, color=INK, align="c")
        code_panel(sl, MX, 5.45, CW, "GET /unidades/7/solicitudes?semana=2026-W45   →   200 OK   {...}", "text", size=13.5, h=0.55)
    d.content("El viaje de una petición", ch("ws", "front", "html"), 5, N(
        "Sigue la petición con el ejemplo concreto de abajo. En cada salto pregunta qué puede fallar. Es el mismo recorrido del "
        "diagrama de secuencia, pero ahora desde el punto de vista de quien depura.",
        pregunta="Si la pantalla sale vacía, ¿en qué salto empezaríais a mirar?",
        enlace="Servicios web y frontend/backend os llegan más adelante; hoy son el mapa para entenderlos."), peticion)

    d.code("El contrato manda", ch("ws", "ser"), """/unidades/{id}/solicitudes:
  get:
    parameters:
      - name: semana
        in: query
        schema: { type: string, example: "2026-W45" }
    responses:
      "200":
        description: Lista de solicitudes""", "yaml", "openapi.yaml", [
        ("Se acuerda antes", "Front y back trabajan en paralelo."),
        ("Se prueba", "El contrato no es un PDF: es ejecutable."),
        ("Cambia con versión", "Romper el contrato es romper al otro equipo.")], 5, N(
        "Un contrato OpenAPI define rutas, parámetros y respuestas. Si front y back lo acuerdan antes de programar, pueden "
        "trabajar a la vez y simular uno de los lados. Quarkus lo genera a partir del código con una extensión.",
        pregunta="¿Quién debería ser el dueño del contrato: el back, el front o los dos?",
        historia="Una vez que el contrato cambió y rompió al otro equipo (y cómo se evitó después)."), size=13, code_w=7.2)

    def verbos(sl):
        grupos = [("2xx · Todo bien", GREEN, ["200  OK", "201  Creado", "204  Sin contenido"]),
                  ("4xx · Has pedido mal", AMBER, ["400  Datos no válidos", "401  Sin autenticar", "403  Sin permiso", "404  No existe", "409  Conflicto"]),
                  ("5xx · He fallado yo", RUST, ["500  Error del servidor", "503  No disponible"])]
        ws_ = [3.6, 4.9, 3.2]
        x = MX
        for (hd, col, items), w in zip(grupos, ws_):
            rr(sl, x, 1.95, w, 0.6, fill=col, r=0.2)
            tb(sl, x + 0.25, 1.95, w - 0.5, 0.6, hd, size=18, bold=True, color=INK if col == AMBER else WHITE, anchor="m", font=HEAD)
            rr(sl, x, 2.65, w, 2.95, fill=LIGHT, r=0.04)
            tb(sl, x + 0.3, 2.85, w - 0.5, 2.7, [(t, {"after": 8}) for t in items], size=17, color=INK, font=MONO)
            x += w + 0.1
        banner(sl, "GET lee · POST crea · PUT reemplaza · PATCH modifica · DELETE borra", y=5.85, h=0.7, size=18, label="Verbos")
    d.content("Verbos y códigos de estado", ch("ws"), 4, N(
        "Los códigos de estado son un idioma común: el 4xx dice que el cliente se ha equivocado y el 5xx que se ha equivocado el "
        "servidor. Saber esto ahorra la mitad del tiempo de diagnóstico.",
        pregunta="Si el front recibe un 404, ¿quién debe arreglarlo? ¿Y si recibe un 500?",
        corto="Enseña solo los grupos 2xx, 4xx y 5xx."), verbos)

    def postman(sl):
        rr(sl, MX, 1.95, 7.6, 3.95, fill=LIGHT, r=0.03)
        box(sl, MX + 0.3, 2.2, 0.95, 0.5, "GET", fill=GREEN, color=WHITE, size=14, bold=True, r=0.15)
        tb(sl, MX + 1.4, 2.2, 6, 0.5, "{{baseUrl}}/unidades/7/solicitudes?semana=2026-W45", size=14, font=MONO, anchor="m")
        code_panel(sl, MX + 0.3, 3.0, 7.0, """pm.test("200 y lista no vacía", () => {
  pm.response.to.have.status(200);
  const lista = pm.response.json();
  pm.expect(lista.length).to.be.above(0);
});""", "java", size=12.5, h=2.2, fname="Tests")
        tb(sl, MX + 0.3, 5.35, 7, 0.4, "Respuesta: 200 OK · 84 ms · 1,2 KB", size=14, color=SLATE, font=MONO)
        callouts(sl, 8.65, 1.95, W - MX - 8.65, [("Probar el back sin el front", "Un cliente genérico para cualquier API."),
                                                 ("Colecciones y entornos", "Mismas peticiones en dev, test y pre."),
                                                 ("Documentación viva", "Se pueden ejecutar en CI.")], h=1.2, gap=0.18)
        banner(sl, "Si falla en Postman, el problema está en el back; si no, mira el front.", y=6.05, h=0.65, size=17)
    d.content("Postman: probar sin pantalla", ch("post", "ws"), 4, N(
        "Una petición de Postman aísla el back: si responde bien aquí, el fallo está en el front o en cómo se llama. "
        "Enseña una colección real: variables de entorno, tests y cómo se ejecuta desde la línea de comandos en CI.",
        pregunta="¿Alguien lo ha usado o ha usado curl?",
        historia="Cómo organiza tu equipo las colecciones y para qué las usáis.",
        enlace="Postman son 10 sesiones dentro de Herramientas colaborativas: os llegarán."), postman)

    def tipos(sl):
        rows = [["Fechas", "2026-11-05T23:30:00Z", "El front muestra 6 de noviembre por la zona horaria"],
                ["Números grandes", "9007199254740993", "JavaScript pierde precisión (límite 2^53)"],
                ["Vacío", "null · ausente · \"\"", "Tres cosas distintas que se mezclan"],
                ["Nombres", "fecha_solicitud / fechaSolicitud", "Cada capa tiene su convención"]]
        mini_table(sl, MX, 1.95, ["Dato", "Lo que envía el back", "Lo que puede pasar"], rows, [2.4, 4.0, 5.5], size=15, row_h=0.72)
        banner(sl, "Fechas en ISO-8601 y en UTC. Los ids largos, como texto.", y=5.95, h=0.7, size=19, label="Regla")
    d.content("Fechas, tipos y nulos", ch("ser", "ecma"), 4, N(
        "Cada capa entiende los datos a su manera. La fecha con zona horaria y los ids de más de 15 dígitos son los dos fallos "
        "que más veo. Pregunta quién ha visto una fecha un día desplazada.",
        pregunta="¿Alguna vez una fecha os ha salido un día antes o después?",
        enlace="Serialización en texto plano (JSON, XML) y ECMAScript os llegan más adelante."), tipos)

    def cors(sl):
        box(sl, MX, 2.05, 3.2, 1.0, [("Navegador", {"bold": True, "size": 17}), ("app.ejemplo.es", {"size": 13, "font": MONO})],
            fill=TEAL, color=WHITE)
        box(sl, W - MX - 3.2, 2.05, 3.2, 1.0, [("API", {"bold": True, "size": 17}), ("api.ejemplo.es", {"size": 13, "font": MONO})],
            fill="3C7F85", color=WHITE)
        steps = [("1  OPTIONS /solicitudes  (preflight)", True, 3.3), ("2  Access-Control-Allow-Origin: app.ejemplo.es", False, 3.9),
                 ("3  GET /solicitudes", True, 4.5)]
        for t, fwd, y in steps:
            xa, xb = (MX + 1.6, W - MX - 1.6) if fwd else (W - MX - 1.6, MX + 1.6)
            line(sl, xa, y + 0.4, xb, y + 0.4, color=TEAL_D if fwd else RUST, w=2.5, arrow=True, dash=not fwd)
            tb(sl, MX + 1.8, y, CW - 3.6, 0.35, t, size=15, color=INK, align="c", font=MONO)
        banner(sl, "Quien bloquea es el navegador, no el servidor. Por eso en Postman funciona.", y=5.65, h=0.9, size=19, label="La clave")
    d.content("CORS en una imagen", ch("html", "front"), 4, N(
        "El error más típico al unir front y back. Los dos están en orígenes distintos, así que el navegador pregunta antes (preflight) "
        "si el servidor lo permite. Se arregla en el back, configurando qué orígenes se aceptan. Postman no lo muestra "
        "porque no es un navegador.",
        pregunta="¿Quién ha visto este error en rojo en la consola?",
        historia="La primera vez que te encontraste CORS y cuánto tardaste en entenderlo."), cors)

    d.code("Errores que el front entiende", ch("ws", "ser"), """HTTP/1.1 404 Not Found

{
  "codigo":  "SOLICITUD_NO_ENCONTRADA",
  "mensaje": "No existe la solicitud 42",
  "traceId": "a1b2c3d4"
}""", "json", "respuesta-error.json", [
        ("El front sabe qué ha fallado", "Mensaje claro, no «Error 500»."),
        ("El traceId une los logs", "Del error en pantalla a la línea exacta."),
        ("Nunca la traza interna", "Nada de nombres de clases ni de tablas.")], 3, N(
        "Un error sin contrato obliga al front a adivinar. Con un cuerpo estándar se puede mostrar un mensaje útil y, con el "
        "identificador de traza, localizar el problema en los logs de producción.",
        pregunta="¿Qué mostraríais al usuario cuando la respuesta es un 500?"), size=14, code_w=7.0)

    def nplus1(sl):
        code_panel(sl, MX, 1.95, 6.6, """SELECT * FROM solicitud WHERE unidad_id = 7;
-- 1 consulta, 200 filas

SELECT * FROM material WHERE id = 3;
SELECT * FROM material WHERE id = 8;
SELECT * FROM material WHERE id = 3;
-- ... 200 consultas más""", "sql", size=12.5, fname="log de consultas", h=2.9)
        code_panel(sl, 7.55, 1.95, W - MX - 7.55, """SELECT s FROM Solicitud s
  JOIN FETCH s.material
  WHERE s.unidad.id = ?1""", "sql", size=12.5, fname="Una sola consulta", h=1.8)
        tb(sl, 7.55, 4.0, W - MX - 7.55, 0.9, "Con 200 filas son 201 viajes a la base de datos; con un JOIN, uno.", size=16, color=INK)
        banner(sl, "Si una pantalla va lenta, cuenta cuántas consultas hace antes de culpar al servidor.", y=5.3, h=0.9, size=19, label="Pista")
    d.content("Consultas de más: el problema N+1", ch("bbdd", "sql"), 4, N(
        "El fallo clásico del acceso a datos con ORM: una consulta para la lista y otra por cada fila para sus relaciones. Se "
        "detecta mirando el log de consultas y se corrige trayendo todo con un JOIN.",
        pregunta="¿Cómo notaríais, sin ver el código, que estáis haciendo cientos de consultas?",
        historia="Un N+1 que encontraste en un proyecto real y cuánto mejoró el tiempo de respuesta."), nplus1)

    d.code("Paginar siempre", ch("ws"), """GET /solicitudes?page=2&size=50

{
  "items": [ { "id": 101, "cantidad": 12 }, ... ],
  "page":  2,
  "size":  50,
  "total": 1832
}""", "json", "respuesta-paginada.json", [
        ("Nunca «devuélvelo todo»", "Hoy son 50 filas, mañana 5 millones."),
        ("El total permite paginar", "El front sabe cuántas páginas hay."),
        ("Orden estable", "Sin ORDER BY, las páginas se repiten.")], 3, N(
        "Un listado sin paginar es una bomba de relojería: funciona con datos de prueba y se cae en producción. Siempre "
        "página, límite y orden.",
        pregunta="¿Qué pasa si la página 3 se pide mientras alguien inserta filas?",
        corto="Enseña la respuesta y el mensaje final."), size=14, code_w=7.0)

    d.ask("Caso roto", "¿Front, back o base de datos?",
          "La pantalla sale vacía.\n\nDecidid dónde miráis y en qué orden.",
          ["Grupos de tres: proponed un orden de investigación.",
           "Justificad cada paso: ¿qué descarta?",
           "Después comparamos con mi orden: consola del navegador → respuesta de la API → logs del servicio → consulta a la BD."], 13, N(
              "Dinámica en grupos. Orden razonable: consola del navegador (errores y red), respuesta de la API (con Postman o el "
              "panel de red), logs del servicio con el traceId y, por último, la consulta a la BD. Que sientan que cada paso descarta "
              "una capa.",
              pregunta="¿Qué os habría gustado tener a mano para decidirlo en 30 segundos?",
              historia="Un caso real de pantalla vacía, qué era y cuánto tardaste.",
              corto="Haz el ejercicio en 8 minutos y cuenta tú el orden razonable."), sub_label="Cómo lo hacemos")
    d.expect("12:55")

    # ------------------------------------------------------------------ B6 ENTREGA 12:55–13:40
    d.section(5, "Entrega y operación", "Un código que no llega a producción no sirve", ch("git", "docker", "soserv", "linux", "redes", "seg"), 45,
              "¿Cómo creéis que llega el código de un programador a un servidor real? Dibujad el camino en el aire.",
              N("Aquí está mi día a día como release manager: pasar de «funciona en mi máquina» a «funciona en producción».",
                enlace="Docker, seguridad TIC y la parte de despliegue aún no se han dado. Linux de usuario y redes ya los conocéis: aquí se ve para qué."))

    d.flow("Del commit a producción", ch("git", "docker"), [
        ("Pull request", "Revisión"), ("Build y tests", "Automáticos"), ("Imagen Docker", "Empaquetado"),
        ("Despliegue", "Desde Git, con ArgoCD"), ("Entorno", "Pruebas y producción")], 5, N(
        "Es el recorrido completo que haces cuando entregas: cada paso es una puerta que se puede cerrar. Enseña tu flujo real, "
        "sin datos sensibles, y marca dónde se puede parar una entrega.",
        pregunta="¿Cuál de estas puertas os parece más importante?",
        historia="El flujo real de tu equipo y dónde se corta una release.",
        enlace="Git y Docker se dan en el curso; el despliegue automático lo veréis aquí por primera vez."),
        below=["Revisión sin hacer", "Fallo en las pruebas", "Imagen que no arranca", "Configuración distinta", "Datos diferentes"],
        below_label="Dónde se frena una entrega", key="Cada paso es una puerta: si falla, no se pasa.")

    d.code("Integración continua: que lo compruebe una máquina", ch("git"), """name: build
on: pull_request
jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-java@v4
        with: { distribution: temurin, java-version: 21 }
      - run: ./mvnw verify""", "yaml", ".github/workflows/build.yml", [
        ("Se ejecuta en cada PR", "Nadie fusiona con las pruebas rotas."),
        ("Mismo resultado para todos", "No depende de la máquina de nadie."),
        ("Un fichero, versionado", "La definición vive con el código.")], 4, N(
        "Un flujo de GitHub Actions mínimo. Cada pull request lanza el build y las pruebas; sin verde no se fusiona. Si usas "
        "otra herramienta en tu proyecto, menciónala: la idea es la misma.",
        pregunta="¿Qué pruebas pondríais obligatorias antes de fusionar?",
        historia="Un día en que el pipeline paró algo que habría llegado a producción."), size=12.5, code_w=7.0)

    d.code("Docker: lo empaqueta todo", ch("docker"), """FROM eclipse-temurin:21-jre
WORKDIR /app
COPY target/quarkus-app/ ./
EXPOSE 8080
CMD ["java", "-jar", "quarkus-run.jar"]""", "docker", "Dockerfile", [
        ("Imagen", "La aplicación y sus dependencias, juntas."),
        ("Contenedor", "Mismo comportamiento en cualquier sitio."),
        ("Fin del «en mi máquina funciona»", "Lo que se prueba es lo que se despliega.")], 4, N(
        "Un Dockerfile mínimo de un servicio Quarkus. La imagen se construye una vez y viaja por todos los entornos sin "
        "cambiar. Es la base de cualquier despliegue moderno.",
        pregunta="¿Alguna vez un programa os ha funcionado en un ordenador y en otro no?",
        historia="Un problema real que Docker os ha resuelto.",
        enlace="Docker son 10 sesiones previstas más adelante en el programa."), size=14, code_w=7.0)

    def gitops(sl):
        box(sl, MX, 2.0, 2.9, 1.0, [("Repo de código", {"bold": True, "size": 16}), ("Java · Dockerfile", {"size": 12.5})], fill=TEAL, color=WHITE)
        box(sl, 4.0, 2.0, 2.9, 1.0, [("Registro", {"bold": True, "size": 16}), ("imágenes con versión", {"size": 12.5})], fill="3C7F85", color=WHITE)
        box(sl, MX, 4.1, 2.9, 1.0, [("Repo de configuración", {"bold": True, "size": 16}), ("YAML de despliegue", {"size": 12.5})], fill=TEAL, color=WHITE)
        box(sl, 4.0, 4.1, 2.9, 1.0, [("ArgoCD", {"bold": True, "size": 17}), ("compara y sincroniza", {"size": 12.5})], fill=AMBER, color=INK)
        box(sl, 7.4, 4.1, 2.4, 1.0, [("Clúster", {"bold": True, "size": 17}), ("Azure", {"size": 12.5})], fill=TEAL_D, color=WHITE)
        line(sl, 3.6, 2.5, 4.0, 2.5, color=INK, arrow=True)
        line(sl, 5.45, 3.0, 2.15, 4.1, color=SLATE, arrow=True, dash=True)
        tb(sl, 5.7, 3.4, 2.5, 0.3, "actualiza la versión", size=12, color=SLATE, italic=True)
        line(sl, 3.6, 4.6, 4.0, 4.6, color=INK, arrow=True, both=True)
        line(sl, 6.9, 4.6, 7.4, 4.6, color=INK, arrow=True)
        callouts(sl, 10.1, 1.95, W - MX - 10.1, [("Git manda", "Lo que está en Git es lo que corre."),
                                                 ("Se auto-corrige", "Si alguien toca a mano, ArgoCD lo revierte."),
                                                 ("Todo es auditable", "Quién, cuándo, por qué.")], h=1.0, gap=0.1, head=15, body=13)
        banner(sl, "Git dice cómo debe estar el sistema. ArgoCD hace que esté así.", y=5.65, h=0.85, size=21, label="GitOps")
    d.content("Despliegue desde Git: GitOps", ch("git", "docker"), 5, N(
        "El repositorio de configuración es la verdad: ArgoCD vigila ese repo, lo compara con lo que corre en el clúster y lo "
        "sincroniza. Los cambios en producción se hacen con una pull request, no con comandos a mano. Es el núcleo de mi trabajo.",
        pregunta="¿Qué ventaja tiene que cualquier cambio en producción pase por una PR?",
        historia="Un día en que ArgoCD detectó una diferencia manual y qué pasó.",
        corto="Cuenta solo la frase final y la ventaja de auditar."), gitops)

    def entornos(sl):
        envs = [("DEV", "Cada PR", "Datos de prueba"), ("TEST", "Pruebas funcionales", "Datos sintéticos"),
                ("PRE", "Ensayo general", "Copia del entorno real"), ("PRO", "Aprobación y ventana", "Datos reales")]
        cols = [SAGE, "3C7F85", TEAL, RUST]
        w = (CW - 3 * 0.55) / 4
        for i, ((n_, a, b), col) in enumerate(zip(envs, cols)):
            x = MX + i * (w + 0.55)
            box(sl, x, 2.05, w, 0.8, n_, fill=col, color=WHITE if i else INK, size=24, bold=True, font=HEAD)
            rr(sl, x, 2.95, w, 1.9, fill=LIGHT, r=0.05)
            tb(sl, x + 0.2, 3.1, w - 0.4, 0.9, a, size=18, bold=True, color=TEAL_D, font=HEAD)
            tb(sl, x + 0.2, 4.0, w - 0.4, 0.7, b, size=15, color=INK)
            if i < 3:
                shape(sl, MSO_SHAPE.RIGHT_ARROW, x + w + 0.08, 2.2, 0.4, 0.5, fill=AMBER)
        banner(sl, "El mismo artefacto recorre todos los entornos. Lo único que cambia es la configuración.", y=5.2, h=0.9, size=19, label="Regla")
    d.content("Entornos: de dev a producción", ch("docker", "seg"), 3, N(
        "El mismo paquete avanza por entornos, cada uno más parecido a producción. Explica qué se comprueba en cada uno y quién "
        "aprueba el salto a PRO. Insiste en que nunca se reconstruye para otro entorno.",
        pregunta="¿Qué os daría más miedo: desplegar en PRE o en PRO?",
        corto="Cuenta solo DEV y PRO."), entornos)

    def release(sl):
        box(sl, MX, 1.95, 3.2, 1.35, [("v2.4.1", {"size": 44, "bold": True, "font": MONO})], fill=TEAL_D, color=WHITE)
        for i, (a, b, c) in enumerate([("2", "MAJOR", "Rompe compatibilidad"), ("4", "MINOR", "Añade funcionalidad"), ("1", "PATCH", "Corrige un error")]):
            y = 1.95 + i * 0.5
            box(sl, 4.2, y + 0.05, 0.4, 0.4, a, fill=AMBER, color=INK, size=16, bold=True, kind=MSO_SHAPE.OVAL, pad=0, font=MONO)
            tb(sl, 4.75, y + 0.02, 1.3, 0.45, b, size=14, bold=True, color=TEAL_D, anchor="m", font=MONO)
            tb(sl, 6.0, y + 0.02, 3.2, 0.45, c, size=14, color=INK, anchor="m")
        cards(sl, [("Trazabilidad", "De la versión a los commits y de ahí al ticket."), ("Calendario", "Cuándo se entrega y con qué contenido."),
                   ("Reversión", "Un plan para volver atrás sin improvisar.")], y=3.65, h=2.1, accent=AMBER, numbered=False)
        callouts(sl, 9.6, 1.95, W - MX - 9.6, [("Notas de versión", "Qué cambia y qué riesgo tiene.")], h=1.35)
        banner(sl, "Una release es una promesa: qué lleva, cuándo sale y cómo se deshace.", y=5.95, h=0.7, size=18)
    d.content("Release management", ch("git", "dir"), 5, N(
        "Tu rol. Una versión con número que significa algo (MAJOR.MINOR.PATCH), trazabilidad hasta el ticket, un calendario y "
        "siempre un plan de vuelta atrás. Cuéntales cómo organizas una release de principio a fin.",
        pregunta="Si veis la versión 3.0.0 después de la 2.9.4, ¿qué esperáis que haya cambiado?",
        historia="Una entrega que salió mal y lo que cambió después.",
        enlace="Git (etiquetas y versiones) y Dirección de proyectos (calendario y riesgos)."), release)

    def checks(sl):
        items = ["Las pruebas automáticas están en verde", "Las notas de versión están escritas", "Las migraciones están probadas y son reversibles",
                 "Hay un plan de vuelta atrás", "La monitorización está lista", "Hay ventana de despliegue y responsables", "Los usuarios están avisados"]
        half = CW / 2 + 0.2
        rr(sl, MX, 1.95, CW, 3.95, fill=LIGHT, r=0.04)
        checklist(sl, MX + 0.4, 2.2, half - 0.5, items[:4], size=17.5, gap=0.88, col=TEAL)
        checklist(sl, MX + half + 0.1, 2.2, CW - half - 0.5, items[4:], size=17.5, gap=0.88, col=TEAL)
        banner(sl, "Si no puedes volver atrás, no estás listo para ir adelante.", y=6.05, h=0.65, size=19)
    d.content("Antes de liberar: la lista", ch("git", "dir"), 4, N(
        "La lista que repaso antes de cada entrega. No es burocracia: cada punto sale de algo que falló una vez. Pregunta qué "
        "añadiría la sala.",
        pregunta="¿Qué punto de esta lista os parece innecesario? ¿Y cuál falta?",
        historia="El punto que se añadió a tu lista tras un incidente."), checks)

    def falla(sl):
        rows = [["Connection refused", "Nadie escucha: proceso caído o puerto equivocado"], ["Timeout", "Red, cortafuegos o proxy en medio"],
                ["404", "Ruta o versión desplegada incorrecta"], ["500", "Mirad los logs del servicio con el traceId"],
                ["Muy lento", "Base de datos, recursos o consultas de más"]]
        mini_table(sl, MX, 1.95, ["Síntoma", "Primera comprobación"], rows, [2.8, 4.3], size=14, row_h=0.62)
        code_panel(sl, 8.05, 1.95, W - MX - 8.05, """$ curl -v https://api.ejemplo.es/health
$ nslookup api.ejemplo.es
$ docker logs --tail 100 solicitudes
$ kubectl logs deploy/solicitudes
$ df -h && free -m
$ journalctl -u solicitudes -n 50""", "bash", size=11.5, h=3.2, fname="terminal")
        banner(sl, "Linux, redes y logs: lo que se mira cuando algo falla. Ahí se ve para qué sirven esas asignaturas.", y=5.9, h=0.75, size=17)
    d.content("Cuando falla en producción: qué miro", ch("linux", "redes", "soserv"), 5, N(
        "Un fallo en producción casi siempre se resuelve con lo básico: ¿responde?, ¿llega la red?, ¿qué dice el log?, ¿qué "
        "cambió? Enlaza con SSOO, redes y seguridad. Es el mejor argumento para que se tomen en serio Linux y redes.",
        pregunta="Si un servicio da «connection refused», ¿qué es lo primero que miráis?",
        historia="Un incidente real y cómo lo localizaste (sin datos del cliente).",
        enlace="Linux de usuario y redes ya los habéis dado; SO de servidor lo estáis empezando."), falla)

    d.cards_slide("Seguridad: lo que se pasa por alto", ch("seg", "linux"), [
        ("Secretos fuera del repo", "Contraseñas y claves en un gestor de secretos, nunca en Git."),
        ("Mínimo privilegio", "Cada persona y cada servicio, solo los permisos que necesita."),
        ("Dependencias vigiladas", "Las vulnerabilidades conocidas se detectan antes de entregar."),
        ("Logs limpios", "Sin contraseñas ni datos sensibles en los registros.")],
        3, N(
            "Cuatro hábitos que evitan los incidentes más habituales. Dado que el público es militar, conecta con la ciberdefensa: "
            "una contraseña en un repositorio es un incidente aunque el repositorio sea privado.",
            pregunta="¿Qué pasaría si alguien clonara hoy vuestro repositorio de clase?",
            enlace="Seguridad en tecnologías de la información y nuevas tendencias en seguridad os llegan más adelante."),
        key="La seguridad no se añade al final: se decide en cada paso del viaje.", h=3.55)

    d.ask("Anécdota compartida", "Una entrega que salió mal",
          "Contad una entrega o un cambio que salió mal, en vuestro trabajo o en prácticas.\n\nSin culpables: buscamos qué cambió después.",
          ["Voluntarios: dos o tres historias.",
           "Armamos juntos un mini informe: qué pasó, qué impacto tuvo y cuál fue la causa.",
           "Cerramos con una cosa que cambiaríamos en el proceso."], 6, N(
              "Postmortem sin culpables en miniatura. Si nadie cuenta, cuenta tú una entrega que salió mal. Usa la estructura: qué "
              "pasó, impacto, causa raíz, qué cambiamos.",
              historia="Una entrega que salió mal, qué cambió después y qué aprendisteis.",
              pregunta="¿Qué mecanismo habría evitado esto: una prueba, una revisión o un plan de vuelta atrás?",
              corto="Cuenta tú sola una historia y pide solo la lección."))
    d.expect("13:40")

    # ------------------------------------------------------------------ CIERRE 13:40–14:00
    d.section(None, "Cierre", "Qué hemos recorrido y qué os lleváis", [], 20,
              "Volvamos a la pizarra de las 8:10: ¿qué herramienta de la lista tiene ahora un sentido distinto?",
              N("Cierra el círculo. Vuelve a la lista de herramientas de la apertura y pide que alguien la reinterprete con lo que ha "
                "visto hoy."), tag="CIERRE")

    d.table("Mapa: asignatura y práctica real", None,
            ["Fase", "Asignaturas del programa", "Lo que uso yo", "Estado"],
            [["Requisitos y planificación", "Ing. de Requisitos · Dirección de Proyectos", "Jira · Confluence", ("Visto y en curso", GREEN, WHITE)],
             ["Diseño", "ADOO-UML · Patrones · BPMN", "Diagramas · ADR", ("En curso y por llegar", AMBER, INK)],
             ["Datos", "TG BBDD · SQL · Oracle · NoSQL", "Postgres · DB2", ("Por llegar", SLATE, WHITE)],
             ["Código", "POO · Algoritmos · Librerías · Pruebas", "Java · Quarkus · IntelliJ", ("En curso y por llegar", AMBER, INK)],
             ["Integración", "Servicios web · HTML/ECMAScript · Postman", "REST · JSON · Postman", ("En curso y por llegar", AMBER, INK)],
             ["Entrega y operación", "Git · Docker · SO · Redes · Seguridad", "GitHub · ArgoCD · Azure", ("De todo un poco", AMBER, INK)]],
            [2.9, 4.3, 2.6, 2.1], 5, N(
                "El mapa final: cada fila une las asignaturas con la herramienta del trabajo real y con el momento del ciclo. "
                "La columna «Estado» indica el progreso global de cada fase (la que mezcla vistas y por llegar se marca como en curso). "
                "Pide que completen con lo que apuntaron en la pizarra al principio.",
                pregunta="¿Qué fila os ha hecho cambiar de opinión sobre una asignatura?",
                enlace="Estados según el calendario semanal V2 (13/10/2026)."),
            status_col=3, size=14, row_h=0.55, key="Cada asignatura es una pieza del mismo viaje.")

    d.cards_slide("Tres consejos para el curso", None, [
        ("Documentad pronto", "Un README ahorra horas, a otros y a vosotros mismos."),
        ("Automatizad", "Pruebas y despliegues repetibles: lo manual falla."),
        ("Preguntad el porqué", "Cada asignatura tiene un uso real. Buscadlo.")], 3, N(
        "Tres consejos prácticos sacados de todo lo anterior. Adáptalos a lo que hayas oído durante el día. Si quieres, cuenta "
        "qué hubieras hecho distinto cuando empezaste tú.",
        pregunta="¿Con cuál empezaréis esta semana?"), accent=AMBER, h=3.4,
        key="Si solo recordáis una cosa: preguntad el porqué.")

    d.cards_slide("Para seguir por vuestra cuenta", None, [
        ("Pruebas automáticas", "Qué probar, cuánto y cómo integrarlo en el pipeline."),
        ("Observabilidad", "Logs, métricas y trazas para entender un sistema vivo."),
        ("Revisión de código", "Cómo dar y recibir feedback sin conflicto."),
        ("Documentación como código", "Markdown, diagramas en texto y decisiones versionadas.")], 3, N(
        "Temas que aparecen constantemente en proyectos reales y que se pueden reforzar por cuenta propia. No son críticas al "
        "programa, que cubre una base sólida: son ideas por si quieren profundizar. Otras ideas por si hay tiempo: inglés técnico, "
        "estimación y comunicación con usuarios, y cultura de seguridad desde el diseño.",
        pregunta="¿Cuál de ellos os gustaría ver con más detalle en una sesión futura?"),
        key="Ideas, no obligaciones: elegid una y probadla en vuestro proyecto.", h=3.5, head_size=16.5, body_size=16)

    d.ask("Ronda final", "¿Qué os lleváis hoy?", "Una idea que aplicaréis en vuestro proyecto del curso.",
          ["Ronda rápida: una frase por persona.",
           "Dejad una pregunta que se haya quedado en el aire.",
           "Votad qué bloque os habría gustado ampliar."], 5, N(
              "Ronda rápida. Apunta las ideas y las preguntas pendientes y decide si merece la pena enviarles un resumen o una chuleta después.",
              corto="Una frase por persona, sin comentarios."))

    d.statement("Gracias", "Resto del tiempo para preguntas, dudas de carrera o casos concretos.", 3, N(
        "Tiempo abierto para preguntas, dudas de carrera o casos concretos. Si sobra tiempo, retoma una pregunta del muro de dudas. "
        "Ofrece compartir las plantillas del anexo (README, comandos útiles)."))
    d.expect("14:00")

    # ------------------------------------------------------------------ ANEXO
    d.code("Anexo · Plantilla de README", None, """# solicitudes-api

Servicio REST para consultar solicitudes de material.

## Arrancar
./mvnw quarkus:dev

## Probar
./mvnw verify
Colección Postman en /docs/postman

## Configuración
| Variable | Qué es |
| DB_URL   | Conexión a Postgres |

## Contacto
Equipo: #solicitudes""", "md", "README.md", [
        ("Arrancar", "Un comando, no diez pasos."),
        ("Probar", "Cómo saber que funciona."),
        ("Configurar", "Qué variables hacen falta.")], 0, N(
        "Plantilla de apoyo para quien la pida. No se presenta en el horario."), size=13, code_w=7.0)

    d.code("Anexo · Chuleta de comandos", None, """# Git
git switch -c feature/PROJ-101-listar
git commit -m "PROJ-101 Lista por unidad"
git push -u origin HEAD

# Docker
docker build -t solicitudes:1.2.0 .
docker run -p 8080:8080 solicitudes:1.2.0

# Kubernetes
kubectl get pods
kubectl logs deploy/solicitudes

# Red
curl -v https://api.ejemplo.es/health""", "bash", "chuleta.sh", [
        ("Lo que más uso", "Los comandos de un día normal."),
        ("Se aprende usando", "Mejor con un proyecto propio."),
        ("Sin memorizar", "Con la ayuda --help basta.")], 0, N(
        "Para compartir después. No se presenta en el horario."), size=12.5, code_w=7.0)

    d.save(path)
    return d


if __name__ == "__main__":
    out = sys.argv[1] if len(sys.argv) > 1 else "Del_ticket_a_produccion_XLIX_DIM_v2.pptx"
    deck = build(out)
    print(f"{deck.n} diapositivas -> {out}")
