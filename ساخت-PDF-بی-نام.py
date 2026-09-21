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

SRC  = "مقاله-بی-نام-نسخه-سابمیت.md"
OUT  = "مقاله-بی-نام-نسخه-سابمیت.pdf"
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

def _norm(t: str) -> str:
    return (t.replace("⟦", "[ ").replace("⟧", " ]")
             .replace("⟪", "[ ").replace("⟫", " ]"))

def fa(t: str) -> str:
    """چینش راست‌به‌چپ متن ساده (بدون نشانه‌گذاری)"""
    return get_display(reshape(_esc(_norm(t))), base_dir="R")

def fa_rich(t: str) -> str:
    """متن فارسی با پشتیبانی **پررنگ**: هر قطعه جداگانه شکل‌دهی و ترتیب قطعه‌ها
    برای نمایش راست‌به‌چپ معکوس می‌شود؛ بنابراین تگ‌ها خام چاپ نمی‌شوند."""
    parts = re.split(r"(\*\*.+?\*\*|`.+?`)", _esc(_norm(t)))
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
 "en":    ParagraphStyle("en", fontName="Fa", fontSize=9.6, leading=17, alignment=0, spaceAfter=3),
}

def inline(t: str) -> str:
    """پاراگراف لاتین: بدون بازچینش، با حفظ نشانه‌گذاری پررنگ/ایتالیک."""
    t = _esc(_norm(t))
    t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
    t = re.sub(r"\*(.+?)\*", r"<i>\1</i>", t)
    return t

def fa_sup(t: str) -> str:
    """پشتیبانی از <sup>N</sup> در متن فارسی با حفظ بالانویس."""
    t = _esc(_norm(t))
    parts = re.split(r"(&lt;sup&gt;.*?&lt;/sup&gt;)", t)
    out = []
    for seg in parts:
        if seg.startswith("&lt;sup&gt;"):
            inner = seg[len("&lt;sup&gt;"):-len("&lt;/sup&gt;")]
            out.append("<super>" + get_display(reshape(inner), base_dir="R") + "</super>")
        else:
            out.append(fa_rich(seg.replace("&amp;","&")) if seg else "")
    return "".join(out)


def wrap_fa(text: str, maxc: int):
    """پاراگراف فارسی را پیش از شکل‌دهی به خطوط می‌شکند تا ترتیب خطوط درست بماند."""
    words = text.split(' ')
    strip = lambda w: len(re.sub(r"\*\*|`|<sup>|</sup>", "", w))
    lines, cur, n = [], "", 0
    for w in words:
        l = strip(w)
        if cur and n + l + 1 > maxc:
            lines.append(cur); cur, n = w, l
        else:
            cur = (cur + " " + w) if cur else w; n += l + (1 if len(cur) > l else 0)
    if cur: lines.append(cur)
    return lines

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
    if ln.startswith("مأخذ: افشاهای شرکتی") or ln.startswith("مأخذ: افشاهای کدال"):
        i += 1; continue   # عنوان و مأخذ داخل خود تصویر شکل درج شده است
    if ln.startswith("**شکل ۱.") and "نردبان" in ln:
        try:
            from reportlab.platypus import Image
            W = 15.6 * cm
            img = Image("شکل-۱-نردبان-دامنه‌ها.png", width=W, height=W * 2060.0 / 4451.0)
            story += [Spacer(1, 3), img, Spacer(1, 6)]
        except Exception as e:
            print("⚠️ تصویر شکل درج نشد:", e)
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
        for k, ch in enumerate(wrap_fa(ln[2:], 62)):
            stl = ParagraphStyle(f"ti{k}", parent=S["title"], spaceAfter=(S["title"].spaceAfter if True else 0))
            story += [Paragraph(fa(ch), stl)]
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
        if is_ltr(ln):
            story += [Paragraph(inline(ln), S["en"])]
        else:
            base = st.fontSize or 9.6
            maxc = int(17.0 * cm / (0.505 * base))
            chunks = wrap_fa(ln, maxc)
            for k, ch in enumerate(chunks):
                stl = ParagraphStyle(f"{st.name}{k}", parent=st,
                                     spaceAfter=(st.spaceAfter if k == len(chunks) - 1 else 0))
                story += [Paragraph(fa_sup(ch) if "<sup>" in ch else fa_rich(ch), stl)]
    i += 1

doc = SimpleDocTemplate(OUT, pagesize=A4, rightMargin=2.4*cm, leftMargin=2.4*cm,
                        topMargin=2.0*cm, bottomMargin=1.9*cm,
                        title="پیوست روش‌شناسی و داده", author="")

def footer(canv, d):
    canv.saveState()
    canv.setFont("Fa", 7.6); canv.setFillColor(colors.HexColor("#7a8a92"))
    canv.drawCentredString(A4[0]/2, 1.05*cm, fa(f"اقتصاد سنجش خسارت انرژی جنگ — نسخه‌ی داوری (بدون مشخصات نویسندگان) — صفحه‌ی {d.page}"))
    canv.restoreState()

doc.build(story, onFirstPage=footer, onLaterPages=footer)
print("✅ PDF ساخته شد:", OUT)
