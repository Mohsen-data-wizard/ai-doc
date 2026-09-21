# -*- coding: utf-8 -*-
"""
ساخت PDF پیوست روش‌شناسی از فایل Markdown — با پشتیبانی راست‌به‌چپ فارسی.
اجرا:  python3 ساخت-پیوست-PDF.py
خروجی: یادداشت-روش‌شناسی-پیوست-مقاله.pdf

نیازمندی: reportlab · arabic-reshaper · python-bidi  ·  فونت Vazirmatn
نکته: برای قلم رسمی نشریه (ب لوتوس)، مسیر فونت را در FONT_REG / FONT_BOLD بدهید.
"""
import re
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                KeepTogether)
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from arabic_reshaper import reshape
from bidi.algorithm import get_display

import sys
SRC = sys.argv[1] if len(sys.argv)>1 else "یادداشت-روش‌شناسی-پیوست-مقاله.md"
OUT = sys.argv[2] if len(sys.argv)>2 else "یادداشت-روش‌شناسی-پیوست-مقاله.pdf"
FONT_REG  = "fonts/Vazirmatn-Regular.ttf"
FONT_BOLD = "fonts/Vazirmatn-Bold.ttf"      # اگر نبود، همان قلم عادی به کار می‌رود

# ── ثبت فونت ────────────────────────────────────────────────────────────────
pdfmetrics.registerFont(TTFont("Fa", FONT_REG))
try:
    pdfmetrics.registerFont(TTFont("FaB", FONT_BOLD))
    BOLD = "FaB"
except Exception:
    BOLD = "Fa"

def _esc(t: str) -> str:
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def fa(t: str) -> str:
    """چینش راست‌به‌چپ متن ساده (بدون نشانه‌گذاری)"""
    return get_display(reshape(_esc(t)), base_dir="R")

def fa_rich(t: str) -> str:
    """متن فارسی با پشتیبانی **پررنگ**: هر قطعه جداگانه شکل‌دهی و ترتیب قطعه‌ها
    برای نمایش راست‌به‌چپ معکوس می‌شود؛ بنابراین تگ‌ها خام چاپ نمی‌شوند."""
    parts = re.split(r"(\*\*.+?\*\*|`.+?`)", _esc(t))
    out = []
    for seg in parts:
        if not seg:
            continue
        if seg.startswith("**") and seg.endswith("**"):
            out.append(f"<b>{get_display(reshape(seg[2:-2]), base_dir="R")}</b>")
        elif seg.startswith("`") and seg.endswith("`"):
            out.append(get_display(reshape(seg[1:-1]), base_dir="R"))
        else:
            out.append(get_display(reshape(seg), base_dir="R"))
    return "".join(reversed(out))

# ── سبک‌ها ─────────────────────────────────────────────────────────────────
S = {
 "title": ParagraphStyle("t",  fontName=BOLD, fontSize=15, leading=23, alignment=2,
                         textColor=colors.HexColor("#12303f"), spaceAfter=4),
 "sub":   ParagraphStyle("s",  fontName="Fa", fontSize=9.5, leading=16, alignment=2,
                         textColor=colors.HexColor("#5a6a73"), spaceAfter=2),
 "h2":    ParagraphStyle("h2", fontName=BOLD, fontSize=11.5, leading=19, alignment=2,
                         textColor=colors.HexColor("#12303f"), spaceBefore=12, spaceAfter=4),
 "p":     ParagraphStyle("p",  fontName="Fa", fontSize=9.6, leading=17, alignment=2, spaceAfter=3),
 "quote": ParagraphStyle("q",  fontName="Fa", fontSize=9.2, leading=16, alignment=2,
                         textColor=colors.HexColor("#3c4a52"), rightIndent=14, spaceAfter=3),
 "li":    ParagraphStyle("l",  fontName="Fa", fontSize=9.6, leading=17, alignment=2,
                         rightIndent=10, spaceAfter=2),
 "cell":  ParagraphStyle("c",  fontName="Fa", fontSize=8.2, leading=13, alignment=2),
}


def is_ltr(line: str) -> bool:
    lat = len(re.findall(r"[A-Za-z]", line)); allc = len(re.findall(r"\S", line)) or 1
    return lat / allc > 0.6

# ── خواندن و تجزیه ──────────────────────────────────────────────────────────
lines = open(SRC, encoding="utf-8").read().split("\n")
story, i = [], 0
while i < len(lines):
    ln = lines[i].rstrip()
    if not ln.strip():
        i += 1; continue
    if ln.startswith("|"):                     # جدول
        rows = []
        while i < len(lines) and lines[i].startswith("|"):
            cells = [c.strip() for c in lines[i].strip().strip("|").split("|")]
            if not all(re.fullmatch(r"[-: ]*", c) for c in cells):
                rows.append(cells)
            i += 1
        ncol = max(len(r) for r in rows)
        rows = [r + [""] * (ncol - len(r)) for r in rows]
        rows = [list(reversed(r)) for r in rows]   # جدول راست‌به‌چپ: ستون اول سمت راست
        data = [[Paragraph(fa_rich(c) if not is_ltr(c) else _esc(c), S["cell"])
                 for c in r] for r in rows]
        w = 17.4 * cm / ncol
        t = Table(data, colWidths=[w] * ncol, hAlign="RIGHT", repeatRows=1)
        t.setStyle(TableStyle([
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#b9c4ca")),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#eef4f7")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
            ("TOPPADDING", (0, 0), (-1, -1), 2.5), ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5),
        ]))
        story += [Spacer(1, 4), t, Spacer(1, 6)]
        continue
    if ln.startswith("# "):
        story += [Paragraph(fa(ln[2:]), S["title"])]
    elif ln.startswith("## "):
        story += [Paragraph(fa(ln[3:]), S["h2"])]
    elif ln.startswith("### "):
        story += [Paragraph(fa(ln[4:]), S["h2"])]
    elif ln.startswith("> "):
        story += [Paragraph(fa_rich(ln[2:]), S["quote"])]
    elif re.match(r"^[-*] ", ln):
        story += [Paragraph("• " + fa_rich(ln[2:]), S["li"])]
    elif re.match(r"^\d+[.)] ", ln):
        story += [Paragraph(fa_rich(ln), S["li"])]
    elif re.fullmatch(r"[-*_]{3,}", ln.strip()):
        story += [Spacer(1, 5)]
    else:
        st = S["p"]
        story += [Paragraph(fa_rich(ln) if not is_ltr(ln) else inline(ln), st)]
    i += 1

doc = SimpleDocTemplate(OUT, pagesize=A4, rightMargin=2.4*cm, leftMargin=2.4*cm,
                        topMargin=2.0*cm, bottomMargin=1.9*cm,
                        title="پیوست روش‌شناسی و داده", author="")

def footer(canv, d):
    canv.saveState()
    canv.setFont("Fa", 7.6); canv.setFillColor(colors.HexColor("#7a8a92"))
    canv.drawCentredString(A4[0]/2, 1.05*cm, fa(f"{sys.argv[3] if len(sys.argv)>3 else 'پیوست روش‌شناسی و داده'} — صفحه‌ی {d.page}"))
    canv.restoreState()

doc.build(story, onFirstPage=footer, onLaterPages=footer)
print("✅ PDF ساخته شد:", OUT)
