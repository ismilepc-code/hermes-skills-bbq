# ШАБЛОН GRILI.UA — генератор PDF. Дані беруться з grili_data.py, логотип — grili_logo.png
import os
LOGO = os.environ.get("GRILI_LOGO", "grili_logo.png")
OUT = os.environ.get("GRILI_OUT", ".")
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Table, TableStyle,
                                Paragraph, Spacer, KeepTogether)
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from grili_data import DATA, COMPANY, GRAPHITE, ORANGE, GREEN

import glob
def _font(name):
    # шукаємо шрифт у системних папках; Carlito — метрична заміна Calibri
    for p in ["/usr/share/fonts/truetype/crosextra/" + name] + glob.glob("/usr/share/fonts/**/" + name, recursive=True):
        if os.path.exists(p):
            return p
    return None
_reg, _bold = _font("Carlito-Regular.ttf"), _font("Carlito-Bold.ttf")
if not (_reg and _bold):
    print("УВАГА: шрифт Carlito не знайдено, використовую DejaVu Sans")
    _reg, _bold = _font("DejaVuSans.ttf"), _font("DejaVuSans-Bold.ttf")
pdfmetrics.registerFont(TTFont("C", _reg))
pdfmetrics.registerFont(TTFont("CB", _bold))
from reportlab.pdfbase.pdfmetrics import registerFontFamily
registerFontFamily("C", normal="C", bold="CB")

G = colors.HexColor("#" + GRAPHITE)
O = colors.HexColor("#" + ORANGE)
GR = colors.HexColor("#" + GREEN)
LIGHT = colors.HexColor("#EEF1F1")  # світлий відтінок графітового для зебри

def fmt(n):
    return f"{int(round(n)):,}".replace(",", " ")

def fmt_q(q):
    return (f"{q:g}").replace(".", ",")

W, H = A4
LM = RM = 14 * mm
logo = ImageReader(LOGO)
lw, lh = logo.getSize()

def header_footer(c, doc):
    c.saveState()
    # логотип
    w = 62 * mm
    h = w * lh / lw
    top = H - 10 * mm
    c.drawImage(logo, LM, top - h, w, h, mask="auto")
    # контакти праворуч
    c.setFillColor(G)
    x = W - RM
    y = top - 2.5 * mm
    c.setFont("CB", 8.5)
    c.drawRightString(x, y, "   ".join(COMPANY["phones"]))
    c.setFont("C", 8.5)
    c.drawRightString(x, y - 4.2 * mm, COMPANY["email"])
    c.drawRightString(x, y - 8.4 * mm, "Магазини: " + ";  ".join(COMPANY["shops"]))
    # помаранчева лінія
    c.setStrokeColor(O)
    c.setLineWidth(1.6)
    c.line(LM, top - h - 3 * mm, W - RM, top - h - 3 * mm)
    # низ сторінки
    c.setFont("C", 7.5)
    c.setFillColor(G)
    c.drawString(LM, 8 * mm, f"Пропозиція №{DATA['number']}  •  GRILI.UA — все для барбекю")
    c.drawRightString(W - RM, 8 * mm, f"Сторінка {doc.page}")
    c.restoreState()

def st(**k):
    base = dict(fontName="C", fontSize=9, leading=11.5, textColor=G)
    base.update(k)
    if "fontSize" in k and "leading" not in k:
        base["leading"] = k["fontSize"] * 1.28
    return ParagraphStyle("x", **base)
s_cell = st()
s_cell_r = st(alignment=2)
s_cell_c = st(alignment=1)
s_head = st(fontName="CB", textColor=colors.white, alignment=1)
s_sec = st(fontName="CB", textColor=colors.white, fontSize=9.5)
s_sub = st(fontName="CB", alignment=2)
s_title = st(fontName="CB", fontSize=15, leading=19)
s_sub_title = st(fontSize=10.5, leading=13)
s_info = st(fontSize=9.5, leading=13, alignment=2)
s_cond = st(fontSize=8.8, leading=11.5, leftIndent=4*mm, bulletIndent=0, bulletFontName="CB", bulletColor=O)

story = []
info = Paragraph(
    f"<b>Місце монтажу:</b> {DATA['place']}<br/><b>Замовник:</b> {DATA['customer']}"
    f"<br/><b>Конт. тел.:</b> {DATA['phone'] or '__________'}", s_info)
title = [Paragraph(f"Пропозиція №{DATA['number']}", s_title),
         Spacer(1, 1.5 * mm),
         Paragraph(DATA["title"], s_sub_title),
         Paragraph(f"від {DATA['date']}", st(fontSize=9, textColor=O))]
t0 = Table([[title, info]], colWidths=[105 * mm, None])
t0.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
story += [t0, Spacer(1, 5 * mm)]

cw = [9 * mm, None, 25 * mm, 19 * mm, 29 * mm]
rows = [[Paragraph(x, s_head) for x in ["№", "Найменування", "Вартість, грн", "Кількість", "Кінцевий прорахунок, грн"]]]
style = [
    ("BACKGROUND", (0, 0), (-1, 0), G),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("LINEBELOW", (0, 1), (-1, -1), 0.4, colors.HexColor("#C9CFD1")),
    ("TOPPADDING", (0, 0), (-1, -1), 3.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
]
n = 0
grand = 0
s_old = st(fontSize=7.5, leading=9, alignment=2)
for sec, items in DATA["sections"]:
    if not items:
        continue
    r = len(rows)
    rows.append([Paragraph(sec, s_sec), "", "", "", ""])
    style += [("SPAN", (0, r), (-1, r)), ("BACKGROUND", (0, r), (-1, r), GR)]
    sub = 0
    priced = False
    for i, it in enumerate(items):
        n += 1
        name, price, qty = it[0], it[1], it[2]
        extra = it[3] if len(it) > 3 else None
        rr = len(rows)
        if price is None:
            # рядок без суми: текст замість вартості, у підсумки не входить
            rows.append([Paragraph(str(n), s_cell_c), Paragraph(name, s_cell),
                         Paragraph(extra or "", s_cell_c), "", ""])
            style.append(("SPAN", (2, rr), (4, rr)))
        else:
            priced = True
            total = round(price * qty)
            sub += total
            if extra:
                pcell = [Paragraph(f"<strike>{fmt(extra)}</strike>", s_old), Paragraph(fmt(price), s_cell_r)]
            else:
                pcell = Paragraph(fmt(price), s_cell_r)
            rows.append([Paragraph(str(n), s_cell_c), Paragraph(name, s_cell),
                         pcell, Paragraph(fmt_q(qty), s_cell_c),
                         Paragraph(fmt(total), s_cell_r)])
        if i % 2 == 1:
            style.append(("BACKGROUND", (0, rr), (-1, rr), LIGHT))
    if priced:
        r = len(rows)
        rows.append(["", Paragraph(f"Разом за розділ «{sec}»:", s_sub), "", "",
                     Paragraph(fmt(sub), st(fontName="CB", alignment=2, textColor=O))])
        style += [("SPAN", (1, r), (3, r)), ("LINEABOVE", (0, r), (-1, r), 0.8, G),
                  ("LINEBELOW", (0, r), (-1, r), 0.8, G)]
    grand += sub

tbl = Table(rows, colWidths=cw, repeatRows=1)
tbl.setStyle(TableStyle(style))
story.append(tbl)
story.append(Spacer(1, 4 * mm))

usd = round(grand / DATA["rate"])
rate_s = f"{DATA['rate']:.2f}".replace(".", ",")
tot = Table([
    [Paragraph("Разом, грн:", st(fontName="CB", fontSize=11, alignment=2)),
     Paragraph(fmt(grand), st(fontName="CB", fontSize=12, alignment=2, textColor=colors.white))],
    [Paragraph(f"Разом, USD (курс НБУ на {DATA['rate_date']}: {rate_s}):", st(fontSize=10, alignment=2)),
     Paragraph(fmt(usd), st(fontName="CB", fontSize=11, alignment=2))],
], colWidths=[None, 40 * mm])
tot.setStyle(TableStyle([
    ("BACKGROUND", (1, 0), (1, 0), O),
    ("LINEBELOW", (1, 1), (1, 1), 1, O),
    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ("TOPPADDING", (0, 0), (-1, -1), 5), ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
]))
story.append(tot)
story.append(Spacer(1, 6 * mm))

conds = [Paragraph("Умови", st(fontName="CB", fontSize=11, textColor=GR)), Spacer(1, 1.5 * mm)]
for ctext in DATA["conditions"]:
    conds.append(Paragraph(ctext.format(**DATA), s_cond, bulletText="•"))
    conds.append(Spacer(1, 1.2 * mm))
story.append(KeepTogether(conds))
story.append(Spacer(1, 10 * mm))

# --- описи товарів (необов'язково): DATA["descriptions"] = [(заголовок, [(параметр, значення), ...]), ...]
for dtitle, drows in DATA.get("descriptions", []):
    blk = [Paragraph(dtitle, st(fontName="CB", fontSize=11, textColor=G)), Spacer(1, 2 * mm)]
    trows = [[Paragraph(f"<b>{a}</b>", st(fontSize=8.8, leading=11)), Paragraph(b, st(fontSize=8.8, leading=11))]
             for a, b in drows]
    dt = Table(trows, colWidths=[45 * mm, None])
    dst = [("VALIGN", (0, 0), (-1, -1), "TOP"), ("TOPPADDING", (0, 0), (-1, -1), 3),
           ("BOTTOMPADDING", (0, 0), (-1, -1), 3), ("LINEBELOW", (0, 0), (-1, -1), 0.4, LIGHT)]
    for i in range(1, len(trows), 2):
        dst.append(("BACKGROUND", (0, i), (-1, i), LIGHT))
    dt.setStyle(TableStyle(dst))
    blk.append(dt)
    story.append(KeepTogether(blk))
    story.append(Spacer(1, 10 * mm))

# --- креслення (необов'язково): DATA["drawings"] = [(заголовок, шлях_до_зображення), ...]
from reportlab.platypus import Image as RLImage
for gtitle, gpath in DATA.get("drawings", []):
    ir = ImageReader(gpath)
    iw, ih = ir.getSize()
    maxw = W - LM - RM
    dw = maxw
    dh = dw * ih / iw
    story.append(KeepTogether([Paragraph(gtitle, st(fontName="CB", fontSize=11, textColor=G)),
                               Spacer(1, 3 * mm), RLImage(gpath, dw, dh), Spacer(1, 10 * mm)]))

sig = Table([
    [Paragraph("<b>Менеджер:</b> ______________________________", st(fontSize=10)),
     Paragraph("<b>Підпис:</b> ____________________", st(fontSize=10))],
    [Paragraph("<b>Телефон:</b> ______________________________", st(fontSize=10)),
     Paragraph("<b>Дата:</b> ____________________", st(fontSize=10))],
], colWidths=[None, 70 * mm])
sig.setStyle(TableStyle([("TOPPADDING", (0, 0), (-1, -1), 6), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
story.append(KeepTogether([sig]))

top_margin = 10 * mm + 62 * mm * lh / lw + 8 * mm
doc = BaseDocTemplate(f"{OUT}/Пропозиція №{DATA['number']} — {DATA['work_type']} — {DATA['customer']}.pdf",
                      pagesize=A4, leftMargin=LM, rightMargin=RM, topMargin=top_margin,
                      bottomMargin=15 * mm, title=f"Пропозиція №{DATA['number']}", author="GRILI.UA")
frame = Frame(LM, 15 * mm, W - LM - RM, H - top_margin - 15 * mm, id="f",
              leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
doc.addPageTemplates([PageTemplate(id="p", frames=[frame], onPage=header_footer)])
doc.build(story)
print("grand", grand, "usd", usd)
