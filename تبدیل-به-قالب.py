#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
تبدیل مقاله-EENR-نسخه-ارسال.md به قالب.docx
- بدون هیچ تغییر در محتوا (تمام متن، جدول، شکل، پانویس حفظ می‌شود)
- حفظ استایل‌های قالب (فونت B Zar/B Lotus/Times etc.) و هدر/فوتر
- اجرا: python3 تبدیل-به-قالب.py
- خروجی: مقاله-EENR-نسخه-ارسال-قالب.docx  (+ رونوشت روی قالب-پرشده.docx برای بررسی)
"""

import re
import pathlib
from docx import Document
from docx.shared import Pt, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SRC_MD = "مقاله-EENR-نسخه-ارسال.md"
TEMPLATE = "قالب.docx"
OUT_DOCX = "مقاله-EENR-نسخه-ارسال-قالب.docx"
OUT_DOCX2 = "قالب-پرشده.docx"  # نام دوم برای اطمینان
FIG_PNG = "شکل-۱-نردبان-دامنه‌ها.png"

# ---------- helper: تشخیص زبان ----------
def is_ltr(text: str) -> bool:
    # اگر بیش از 40% حروف لاتین باشد -> LTR (برای تشخیص انگلیسی)
    lat = len(re.findall(r'[A-Za-z]', text))
    tot = len(re.findall(r'\S', text)) or 1
    # همچنین اگر حاوی حروف فارسی نباشد -> LTR
    has_fa = bool(re.search(r'[\u0600-\u06FF]', text))
    if not has_fa and lat > 0:
        return True
    return (lat / tot) > 0.45

def is_mostly_persian(text: str) -> bool:
    return not is_ltr(text)

# ---------- helper: RTL تنظیم ----------
def set_paragraph_rtl(paragraph, rtl: bool):
    pPr = paragraph._p.get_or_add_pPr()
    # حذف قبلی
    for b in pPr.findall(qn('w:bidi')):
        pPr.remove(b)
    if rtl:
        bidi = OxmlElement('w:bidi')
        bidi.set(qn('w:val'), '1')
        pPr.append(bidi)
        # همچنین rtlGutter? نه نیازی نیست
    # برای bidiVisual جدول نیز بعدا

def set_paragraph_spacing(paragraph, before=0, after=120, line=259, line_rule='auto'):
    pPr = paragraph._p.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    spacing.set(qn('w:before'), str(before))
    spacing.set(qn('w:after'), str(after))
    spacing.set(qn('w:line'), str(line))
    spacing.set(qn('w:lineRule'), line_rule)

# ---------- helper: عدد اعشاری با / برای RTL ----------
def _wrap_slash_numbers(s: str) -> str:
    """
    ژورنال ممیز را با / می‌خواهد (مثل ۵۸/۵). در پاراگراف RTL، / خنثی است و ۵۸/۵ به ۵/۵۸ برمی‌گردد.
    با احاطه کردن هر عدد اعشاریِ شامل / با LRM (U+200E) نظم چپ‌به‌راست داخل RTL حفظ می‌شود.
    الگو: رقم‌های فارسی/انگلیسی + / + رقم‌ها  (پشتیبانی از ۰-۹ و 0-9)
    """
    # قبل/بعد عددِ اسلش‌دار، LRM می‌گذاریم تا «جزیرهٔ LTR» بسازد
    # از رجی‌اکس برای فارسی (۰-۹ = \u06F0-\u06F9) و عربی-هندی (\u0660-\u0669) و لاتین استفاده می‌کنیم
    # نکته: برای جلوگیری از LRM مضاعف، اول LRM‌های موجود را پاک می‌کنیم
    s = s.replace('\u200e','').replace('\u200f','')
    # هر رخداد مثل ۵۸/۵ یا ۰/۴۲ یا 58/5 را بیاب
    # شامل حالت‌های چند اسلش مثل ۰/۲ تا ۱/۴ هر کدام جداگانه
    return re.sub(r'([\d\u06F0-\u06F9\u0660-\u0669]+/[\d\u06F0-\u06F9\u0660-\u0669]+)', '\u200e\\1\u200e', s)

# ---------- helper: runها با فرمت ----------
def add_formatted_runs(paragraph, text, is_rtl_context: bool, base_font=None):
    """
    متن را با پشتیبانی **پررنگ** *ایتالیک* `کد` و <sup> </sup> به run تبدیل می‌کند.
    برای انگلیسی فونت Times New Roman می‌گذارد.
    نکتهٔ ممیز: اعداد اعشاری با / قبل از پردازش با LRM محصور می‌شوند.
    """
    # حفظ نظم / در محیط RTL
    text = _wrap_slash_numbers(text)
    # ابتدا تگ‌های html sup/sub را موقت جدا می‌کنیم
    # الگوی کلی برای بلوک‌های پررنگ، کد، sup/sub
    # برای جلوگیری از تداخل ** با * ، اول ** را استخراج می‌کنیم
    token_pat = re.compile(r'(\*\*.*?\*\*|`[^`]*`|<sup>.*?</sup>|<sub>.*?</sub>)', re.DOTALL)
    parts = token_pat.split(text)
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**') and len(part) >=4:
            inner = part[2:-2]
            # داخل پررنگ ممکن است شامل <sup> یا *ایتالیک* باشد؛ بازگشتی هندل می‌کنیم
            # تقسیم inner به قطعات sup/sub و متن عادی، همه با bold
            # الگو برای sup داخل bold
            sub_parts = re.split(r'(<sup>.*?</sup>|<sub>.*?</sub>)', inner)
            for sp in sub_parts:
                if not sp:
                    continue
                if sp.startswith('<sup>'):
                    inner_sup = re.sub(r'^<sup>|</sup>$', '', sp)
                    run = paragraph.add_run(inner_sup)
                    run.bold = True
                    run.font.superscript = True
                    if not is_rtl_context:
                        run.font.name = 'Times New Roman'
                elif sp.startswith('<sub>'):
                    inner_sub = re.sub(r'^<sub>|</sub>$', '', sp)
                    run = paragraph.add_run(inner_sub)
                    run.bold = True
                    run.font.subscript = True
                    if not is_rtl_context:
                        run.font.name = 'Times New Roman'
                else:
                    # متن عادی داخل bold ممکن است حاوی *ایتالیک* باشد (نادر)
                    # ساده به عنوان bold
                    # اگر نیاز به ایتالیک داخل bold بود، می‌توان اضافه کرد
                    # برای حفظ سادگی، کل sp را bold می‌کنیم، اما ایتالیک داخلی را هم هندل
                    # تقسیم به ایتالیک
                    italic_pat_inner = re.compile(r'\*([^*]+?)\*')
                    last_inner = 0
                    for m2 in italic_pat_inner.finditer(sp):
                        before2 = sp[last_inner:m2.start()]
                        if before2:
                            run2 = paragraph.add_run(before2)
                            run2.bold = True
                            if not is_rtl_context:
                                run2.font.name = 'Times New Roman'
                        inner_it = m2.group(1)
                        run2 = paragraph.add_run(inner_it)
                        run2.bold = True
                        run2.italic = True
                        if not is_rtl_context:
                            run2.font.name = 'Times New Roman'
                        last_inner = m2.end()
                    remain2 = sp[last_inner:]
                    if remain2:
                        run2 = paragraph.add_run(remain2)
                        run2.bold = True
                        if not is_rtl_context:
                            run2.font.name = 'Times New Roman'
            continue
        elif part.startswith('`') and part.endswith('`'):
            inner = part[1:-1]
            run = paragraph.add_run(inner)
            run.font.name = 'Courier New'
            run.font.size = Pt(9)
            continue
        elif part.startswith('<sup>'):
            inner = re.sub(r'^<sup>|</sup>$', '', part)
            run = paragraph.add_run(inner)
            run.font.superscript = True
            if not is_rtl_context:
                run.font.name = 'Times New Roman'
            continue
        elif part.startswith('<sub>'):
            inner = re.sub(r'^<sub>|</sub>$', '', part)
            run = paragraph.add_run(inner)
            run.font.subscript = True
            if not is_rtl_context:
                run.font.name = 'Times New Roman'
            continue
        else:
            # متن عادی: ممکن است حاوی *ایتالیک* باشد
            # پردازش ایتالیک تک‌ستاره
            # الگو: *text*  اما نه ** (چون قبلا جدا شد)
            # از finditer استفاده می‌کنیم
            italic_pat = re.compile(r'\*([^*]+?)\*')
            last = 0
            for m in italic_pat.finditer(part):
                before = part[last:m.start()]
                if before:
                    run = paragraph.add_run(before)
                    if not is_rtl_context:
                        run.font.name = 'Times New Roman'
                inner = m.group(1)
                run = paragraph.add_run(inner)
                run.italic = True
                if not is_rtl_context:
                    run.font.name = 'Times New Roman'
                last = m.end()
            remain = part[last:]
            if remain:
                # همچنین لینک‌های http را به صورت هایپرلینک ساده متن نگه می‌داریم (بدون اکتیو)
                # اما هایپرلینک واقعی نیاز به رلیشن دارد -> فعلا متن خالی
                run = paragraph.add_run(remain)
                if not is_rtl_context:
                    run.font.name = 'Times New Roman'
                # اگر متن حاوی انگلیسی و فارسی مخلوط بود، فونت بهری هردو؟ فارسی B Zar برای RTL و انگلیسی Times برای LTR.
                # با تعیین فونت کلی Times برای کل پاراگراف LTR، فارسی داخلش کمی ناخوانا می‌شود
                # ولی قالب انگلیسی همه Times است، ایرادی ندارد.
            continue

def add_hyperlink(paragraph, url, display_text=None):
    # low-level hyperlink addition (اگر نیاز شد)
    # فعلا استفاده نمی‌شود؛ متن URL به صورت ساده اضافه می‌شود
    pass

# ---------- helper: جدول ----------
def create_table(doc, rows, is_rtl_table=True):
    """
    rows: list of list of strings (already split cells)
    is_rtl_table: اگر جدول فارسی است bidiVisual فعال می‌شود (بدون معکوس‌سازی اضافی)
    """
    if not rows:
        return None
    # برای فارسی نیازی به معکوس‌سازی دستی نیست وقتی bidiVisual تنظیم می‌شود؛
    # نگاشت اصلی حفظ می‌شود تا ستون اول (سمت راست در RTL) همان باشد.
    # اگر قبلاً معکوس می‌کردیم + bidiVisual = دو بار معکوس = برعکس
    # پس فقط bidiVisual کافیست.
    ncols = max(len(r) for r in rows)
    # Normalize rows to same column count
    for r in rows:
        while len(r) < ncols:
            r.append('')
    tbl = doc.add_table(rows=0, cols=ncols)
    tbl.style = 'Table Grid'
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = True
    # تنظیم bidiVisual برای جدول فارسی
    if is_rtl_table:
        tblPr = tbl._tbl.find(qn('w:tblPr'))
        if tblPr is not None:
            bidiVis = OxmlElement('w:bidiVisual')
            tblPr.append(bidiVis)
    # سطرها
    for i, row_cells in enumerate(rows):
        row = tbl.add_row()
        row.height_rule = None
        for j, cell_text in enumerate(row_cells):
            cell = row.cells[j]
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            # شید برای هدر
            if i == 0:
                shading = OxmlElement('w:shd')
                shading.set(qn('w:val'), 'clear')
                shading.set(qn('w:color'), 'auto')
                shading.set(qn('w:fill'), 'EEF4F7')  # همان #eef4f7
                tcPr = cell._tc.get_or_add_tcPr()
                tcPr.append(shading)
                # border? style Table Grid خود دارد
            # جلوگیری از شکست
            # متن سلول: پاراگراف داخل سلول
            # سلول‌ها به صورت پیش‌فرض یک پاراگراف خالی دارند؛ از آن استفاده می‌کنیم
            para = cell.paragraphs[0]
            # تنظیم جهت
            cell_is_ltr = is_ltr(cell_text) and not re.search(r'[\u0600-\u06FF]', cell_text)
            # اما هدرها اغلب فارسی‌اند
            set_paragraph_rtl(para, not cell_is_ltr)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER if i==0 else (WD_ALIGN_PARAGRAPH.CENTER if cell_is_ltr else WD_ALIGN_PARAGRAPH.CENTER)
            # در سلول‌ها تراز وسط برای متن جدول (مطابق قالب: style a1 centered)
            # اضافه کردن متن با فرمت
            # خالی نباشد
            if cell_text.strip() == '':
                para.add_run('')
            else:
                # برای سلول، فونت B Lotus برای فارسی، Times برای انگلیسی
                # اما با trick: اگر فارسی است، run فونت خودکار B Lotus از استایل a1 خواهد آمد
                # ولی چون سلول داخل Table Grid است، فونت پیش‌فرض Calibri است؛ پس باید دستی ست کنیم
                # ما runهای داخل متن جدول را با فونت مناسب می‌سازیم
                # برای فارسی، B Lotus (یا B Zar) و برای انگلیسی Times
                # اینجا add_formatted_runs را با is_rtl مناسب صدا می‌زنیم
                add_formatted_runs(para, cell_text.strip(), not cell_is_ltr)
                # sizing: برای هدر bold
                for run in para.runs:
                    if i == 0:
                        run.bold = True
                        run.font.size = Pt(9.5)
                    else:
                        run.font.size = Pt(9)
                    if cell_is_ltr:
                        run.font.name = 'Times New Roman'
                    else:
                        # Persian cell: B Lotus
                        # ست فونت cs?
                        # در python-docx font.name فقط ascii می‌گذارد؛ برای فارسی باید rFonts cs تنظیم شود
                        # با دستی XML:
                        r = run._element
                        rPr = r.find(qn('w:rPr'))
                        if rPr is None:
                            rPr = OxmlElement('w:rPr')
                            r.insert(0, rPr)
                        rFonts = rPr.find(qn('w:rFonts'))
                        if rFonts is None:
                            rFonts = OxmlElement('w:rFonts')
                            rPr.append(rFonts)
                        # برای فارسی فقط cs را B Lotus می‌کنیم
                        rFonts.set(qn('w:cs'), 'B Lotus')
                        rFonts.set(qn('w:hAnsi'), 'Times New Roman')
                        rFonts.set(qn('w:ascii'), 'Times New Roman')
            # padding کم
            tcPr = cell._tc.get_or_add_tcPr()
            tcMar = OxmlElement('w:tcMar')
            for m in ['top','left','bottom','right']:
                el = OxmlElement(f'w:{m}')
                el.set(qn('w:w'), '80')
                el.set(qn('w:type'), 'dxa')
                tcMar.append(el)
            tcPr.append(tcMar)
    # تنظیم عرض ستون‌ها مساوی
    # عمدا auto
    return tbl

# ---------- parsing Markdown ----------
def parse_markdown(md_text: str):
    """
    خروجی: لیستی از بلوک‌ها (type, data)
    type ∈ 'heading','paragraph','table','ulist','olist','hr','blockquote','figure_placeholder'
    """
    lines = md_text.splitlines()
    blocks = []
    i = 0
    # برای حفظ --- افقی
    while i < len(lines):
        line = lines[i]
        stripped = line.strip()
        # skip empty
        if stripped == '':
            i += 1
            continue
        # HR: خطی فقط شامل --- یا *** یا ___ (با فاصله)
        if re.fullmatch(r'[-*_]{3,}\s*', stripped) or stripped == '---':
            # در فایل ما "---" به معنی جدا کننده بخش است
            # به عنوان hr ثبت می‌کنیم (بعدا به عنوان خط جداکننده نمایش می‌دهیم)
            blocks.append(('hr', None))
            i += 1
            continue
        # Table: خط با |
        if stripped.startswith('|'):
            rows = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                rows.append(lines[i].strip())
                i += 1
            # تجزیه rows به سلول‌ها
            parsed_rows = []
            for r in rows:
                # حذف | اول و آخر
                inner = r.strip().strip('|')
                cells = [c.strip() for c in inner.split('|')]
                # خط جداکننده --- را حذف کن
                if all(re.fullmatch(r'[-:\s]*', c) for c in cells):
                    continue
                parsed_rows.append(cells)
            if parsed_rows:
                blocks.append(('table', parsed_rows))
            continue
        # Heading: #...
        m = re.match(r'^(#{1,6})\s*(.*)', line)
        if m:
            level = len(m.group(1))
            text = m.group(2).strip()
            # اگر # است و متن خالی؟ بگذر
            blocks.append(('heading', (level, text)))
            i += 1
            continue
        # Blockquote: > 
        if line.lstrip().startswith('> '):
            # جمع‌آوری بلوک نقل‌قول چندخطی اگر باشد
            q_lines = []
            while i < len(lines) and lines[i].lstrip().startswith('> '):
                q_lines.append(lines[i].lstrip()[2:])
                i += 1
            blocks.append(('blockquote', "\n".join(q_lines)))
            continue
        # Image markdown: ![alt](path)
        if stripped.startswith('!['):
            blocks.append(('image_md', stripped))
            i += 1
            continue
        # List unordered: - یا * + فاصله
        if re.match(r'^\s*[-*]\s+', line):
            items = []
            while i < len(lines) and re.match(r'^\s*[-*]\s+', lines[i]):
                items.append(re.match(r'^\s*[-*]\s+(.*)', lines[i]).group(1).strip())
                i += 1
            blocks.append(('ulist', items))
            continue
        # List ordered: number. 
        if re.match(r'^\s*\d+[.)]\s+', line):
            items = []
            while i < len(lines) and re.match(r'^\s*\d+[.)]\s+', lines[i]):
                items.append(re.match(r'^\s*\d+[.)]\s+(.*)', lines[i]).group(1).strip())
                i += 1
            blocks.append(('olist', items))
            continue
        # Paragraph: جمع‌آوری تا خط خالی یا بلوک بعدی
        # در این فایل پاراگراف‌ها غالبا یک خط‌اند ولی برای اطمینان چندخط را جمع می‌کنیم
        # شرط توقف: خط بعدی خالی یا شروع بلوک جدید (heading/table/list/quote/hr)
        para_lines = [line.rstrip()]
        i += 1
        # نگاه به آینده: اگر خط بعدی خالی نیست و شروع بلوک نیست، آن را به پاراگراف فعلی بچسبان (برای پاراگراف‌های چندخطی)
        # در این مطبوعات پاراگراف‌های چندخطی وجود ندارد، پس یک خط کافیست؛ اما برای امنیت:
        while i < len(lines):
            nxt = lines[i]
            nxt_stripped = nxt.strip()
            if nxt_stripped == '':
                break
            if re.match(r'^\s*#{1,6}\s', nxt) or nxt_stripped.startswith('|') or nxt.lstrip().startswith('> ') or re.fullmatch(r'[-*_]{3,}\s*', nxt_stripped) or re.match(r'^\s*[-*]\s+', nxt) or re.match(r'^\s*\d+[.)]\s+', nxt) or nxt_stripped.startswith('!['):
                break
            # همچنین اگر nxt_stripped شروع به **جدول یا **شکل کند و para_lines قبلی نیز ** ... باشد؟ جدا نگه‌دار
            # به طور کلی اگر para فعلی یک خط bold کامل است (شروع و پایان **)، آن را جداگانه نگه‌دار
            # فعلا هر خط را جداگانه پاراگراف در نظر می‌گیریم
            break
        text = " ".join(p.strip() for p in para_lines if p.strip() != '')
        # تشخیص فیگور placeholder که در متن آمده با **شکل 1...**
        # حفظ به عنوان paragraph عادی، بعدا به style caption نگاشت می‌شود
        blocks.append(('paragraph', text))
        # i already advanced
    return blocks

# ---------- تولید docx ----------
def build_docx():
    md_text = pathlib.Path(SRC_MD).read_text(encoding='utf-8')
    blocks = parse_markdown(md_text)
    print(f"📄 پارس شد: {len(blocks)} بلوک")
    # شمارش
    from collections import Counter
    cnt = Counter([b[0] for b in blocks])
    print(cnt)

    doc = Document(TEMPLATE)

    # پاک‌سازی بدنه ولی حفظ sectPr و هدر/فوتر
    body = doc.element.body
    # جمع‌آوری sectPr‌های بدنه (مستقیم یا داخل pPr آخرین پاراگراف)
    # رویکرد امن با python-docx: حذف همه پاراگراف‌ها و جدول‌ها
    # پاراگراف‌ها
    for p in list(doc.paragraphs):
        # مراقب: هدربار پاراگراف‌ها داخل هدرها جزو doc.paragraphs نیست؟ در python-docx فقط بدنه را می‌دهد
        p_elem = p._element
        parent = p_elem.getparent()
        if parent is body:
            parent.remove(p_elem)
    for tbl in list(doc.tables):
        tbl_elem = tbl._tbl
        parent = tbl_elem.getparent()
        if parent is body:
            parent.remove(tbl_elem)
    # اکنون body باید فقط sectPr داشته باشد (اگر وجود داشت)
    # اگر sectPr حذف شد، آن را دوباره اضافه می‌کنیم
    if body.find(qn('w:sectPr')) is None:
        # سعی کن از قالب اصلی sectPr را بخوانی
        import zipfile
        z = zipfile.ZipFile(TEMPLATE)
        doc_xml = z.read('word/document.xml').decode()
        m = re.search(r'<w:sectPr.*?</w:sectPr>', doc_xml, re.DOTALL)
        if m:
            from docx.oxml import parse_xml
            sectPr_elem = parse_xml(m.group(0).encode('utf-8'))
            body.append(sectPr_elem)
            print("⚠️ sectPr بازسازی شد")
    print(f"بعد پاکسازی: پاراگراف‌ها {len(doc.paragraphs)} جدول‌ها {len(doc.tables)}")

    # متغیر وضعیت بخش: برای منابع/چکیده
    current_section = ""
    # برای تشخیص اینکه بعد از چکیده هستیم (برای استایل چکیده)
    # همچنین برای ذخیره caption قبلی (جدول/شکل)
    # اضافه کردن محتوا

    # فونت‌های قالب: برای راحتی نام استایل‌ها را به فارسی نگه می‌داریم
    # نگاشت سطح هدینگ به استایل
    def add_heading_block(level, text):
        nonlocal current_section
        is_rtl = is_mostly_persian(text)
        # به‌روزرسانی current_section
        # نرمالایز متن برای مقایسه
        norm = re.sub(r'[\s\u200c]+', ' ', text).strip()
        current_section = norm
        # انتخاب استایل
        if level == 1:
            style_name = 'عنوان'  # a8
            para = doc.add_paragraph(style=style_name)
            set_paragraph_rtl(para, is_rtl)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            # سایز بزرگ: قالب 15 ضخیم، استایل خود این را دارد
            add_formatted_runs(para, text, is_rtl)
            # فاصله
            set_paragraph_spacing(para, before=180, after=180, line=360, line_rule='auto')
        elif level == 2:
            # اگر متن انگلیسی است (Abstract, References, etc.) از Heading 1/2 انگلیسی
            if not is_rtl:
                # انگلیسی: Heading 1 با Times
                style_name = 'Heading 1'
                para = doc.add_paragraph(style=style_name)
                set_paragraph_rtl(para, False)
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                # متن را با Times و Bold
                # Heading 1 استایل خود bold دارد
                add_formatted_runs(para, text, False)
                set_paragraph_spacing(para, before=240, after=120, line=276, line_rule='auto')
            else:
                style_name = 'تیتر'  # a3
                para = doc.add_paragraph(style=style_name)
                set_paragraph_rtl(para, True)
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # یا RIGHT?
                # تیتر معمولا راست؟ اما قالب jc both دارد
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                add_formatted_runs(para, text, True)
                # برای هدینگ سطح 2، کمی فاصله بیشتر
                set_paragraph_spacing(para, before=240, after=120, line=276, line_rule='auto')
        else: # level 3
            if not is_rtl:
                style_name = 'Heading 2'
                para = doc.add_paragraph(style=style_name)
                set_paragraph_rtl(para, False)
                para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                add_formatted_runs(para, text, False)
                set_paragraph_spacing(para, before=200, after=100, line=259, line_rule='auto')
            else:
                # برای زیرعنوان فارسی، از تیتر با سایز کمی کوچکتر اما همان استایل
                style_name = 'تیتر'
                para = doc.add_paragraph(style=style_name)
                set_paragraph_rtl(para, True)
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                add_formatted_runs(para, text, True)
                set_paragraph_spacing(para, before=200, after=100, line=259, line_rule='auto')
        # برای بخش‌های خاص، بعد از هدینگ، ممکن است نیاز به خط جداکننده نباشد
        return para

    def add_paragraph_block(text):
        # تشخیص نوع پاراگراف خاص
        stripped = text.strip()
        # خطوطی که فقط ** هستند؟ نه
        # تشخیص عنوان جدول/شکل
        is_table_caption = bool(re.match(r'^\*\*جدول\s*[۱۲۳۴0-9]', stripped) or re.match(r'^جدول\s*[۱۲۳۴0-9]', stripped) or 'جدول ' in stripped[:20] and 'رررر' in stripped)
        # الگوی جدول/شکل در فایل: "**جدول ۱. ..." یا "**شکل ۱. ..."
        if re.match(r'^\*\*جدول\s*[\d۰-۹]+', stripped) or re.match(r'^\*\*شکل\s*[۱۲۳۴0-9]', stripped) or stripped.startswith('**جدول') or stripped.startswith('**شکل'):
            # کپشن شکل/جدول
            is_rtl = True  # فارسی
            # حذف ** اطراف
            clean = re.sub(r'^\*\*', '', stripped)
            clean = re.sub(r'\*\*$', '', clean)
            # اگر شامل (قلم بی لوتوس، سایز 11) است، آن توضیح را حذف کنیم؟ در قالب آن فقط راهنماست، در مقاله واقعی فقط عنوان کافیست
            # ولی برای حفظ محتوا، توضیح قلم را نمایش نمی‌دهیم اگر فقط راهنماست
            # در فایل md واقعی عنوان شکل کامل است و توضیح قلم ندارد؛ فقط قالب آن را داشت
            # پس clean را نگه می‌داریم
            para = doc.add_paragraph(style='عنوان ج')  # a2 : B Lotus Bold center
            set_paragraph_rtl(para, True)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_formatted_runs(para, clean.strip(), True)
            set_paragraph_spacing(para, before=120, after=60, line=216, line_rule='auto')
            return para
        if stripped.startswith('مأخذ:') or stripped.startswith('**مأخذ:'):
            # منبع جدول/شکل
            clean = re.sub(r'^\*\*', '', stripped)
            clean = re.sub(r'\*\*$', '', clean)
            is_rtl = is_mostly_persian(clean)
            para = doc.add_paragraph(style='عنوان شکل و جدول') # a0 : B Lotus 10 center
            set_paragraph_rtl(para, is_rtl)
            para.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_rtl else WD_ALIGN_PARAGRAPH.LEFT
            # سایز کوچکتر (قالب a0 سایز 10 دارد)
            add_formatted_runs(para, clean.strip(), is_rtl)
            # italic کمی؟
            para.runs and setattr(para.runs[0], 'italic', False)
            set_paragraph_spacing(para, before=40, after=120, line=200, line_rule='auto')
            return para
        if stripped.startswith('**کلیدواژه') or stripped.startswith('**طبقه') or stripped.startswith('کلیدواژه') or stripped.startswith('طبقه'):
            is_rtl = True
            para = doc.add_paragraph(style='متن')
            set_paragraph_rtl(para, True)
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            add_formatted_runs(para, stripped, True)
            set_paragraph_spacing(para, before=60, after=60, line=240, line_rule='auto')
            return para
        if stripped.startswith('**Keywords') or stripped.startswith('**JEL') or stripped.startswith('Keywords') or stripped.startswith('JEL Classification'):
            is_rtl = False
            para = doc.add_paragraph(style='Normal')
            set_paragraph_rtl(para, False)
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            add_formatted_runs(para, stripped, False)
            set_paragraph_spacing(para, before=60, after=60, line=240, line_rule='auto')
            return para
        # تشخیص بخش منابع: اگر current_section شامل "منابع" یا "References" است، استایل متفاوت
        # این تشخیص بر اساس current_section که از هدینگ آخر به‌روز شده
        if 'منابع' in current_section and not 'References' in current_section:
            # منابع فارسی
            is_rtl = True
            para = doc.add_paragraph(style='منابع') # a4 با hanging
            set_paragraph_rtl(para, True)
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            # hanging indent 1cm ≈ 567 twips
            pPr = para._p.get_or_add_pPr()
            ind = pPr.find(qn('w:ind'))
            if ind is None:
                ind = OxmlElement('w:ind')
                pPr.append(ind)
            ind.set(qn('w:left'), '567')
            ind.set(qn('w:hanging'), '567')
            # راست‌به‌چپ: برای hanging فارسی باید right؟
            add_formatted_runs(para, stripped, True)
            set_paragraph_spacing(para, before=40, after=40, line=240, line_rule='auto')
            # فونت B Lotus سایز 10؟ استایل a4 سایز 12 دارد اما در قالب گفته 12 برای فارسی hanging
            return para
        if 'References' in current_section or 'REFERENCE' in current_section.upper():
            is_rtl = False
            para = doc.add_paragraph(style='منابع ل') # a5 : Times 10 hanging
            set_paragraph_rtl(para, False)
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT
            pPr = para._p.get_or_add_pPr()
            ind = pPr.find(qn('w:ind'))
            if ind is None:
                ind = OxmlElement('w:ind')
                pPr.append(ind)
            ind.set(qn('w:left'), '567')
            ind.set(qn('w:hanging'), '567')
            add_formatted_runs(para, stripped, False)
            set_paragraph_spacing(para, before=40, after=40, line=240, line_rule='auto')
            return para
        # چکیده: تشخیص بخش چکیده
        if current_section in ['چکیده', 'چکیده '] or 'چکیده' == current_section.strip():
            # پاراگراف چکیده باید توجیه both با فونت B Zar
            # اگر متن انگلیسی بود (Abstract) نه
            is_rtl = is_mostly_persian(stripped)
            if is_rtl:
                para = doc.add_paragraph(style='چکیده') # a6
                set_paragraph_rtl(para, True)
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                add_formatted_runs(para, stripped, True)
                set_paragraph_spacing(para, before=80, after=80, line=259, line_rule='auto')
                return para
            else:
                # انگلیسی چکیده: Normal Times justify
                para = doc.add_paragraph(style='Normal')
                set_paragraph_rtl(para, False)
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                add_formatted_runs(para, stripped, False)
                set_paragraph_spacing(para, before=60, after=60, line=259, line_rule='auto')
                return para
        # تشخیص "عنوان فارسی:" یا "عنوان انگلیسی:" یا نویسندگان
        if stripped.startswith('**عنوان فارسی:**') or stripped.startswith('**عنوان انگلیسی:**') or stripped.startswith('**نویسندگان:**') or stripped.startswith('**نویسنده‌ی مسئول'):
            is_rtl = is_mostly_persian(stripped)  # ولی شامل انگلیسی هم هست؛ تشخیص بر فارسی
            # برای خطوط دو زبانه، جهت RTL نگه می‌داریم اما انگلیسی داخل آن LTR خواهد ماند
            # استایل متن با bold label
            para = doc.add_paragraph(style='متن')
            set_paragraph_rtl(para, True)
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            add_formatted_runs(para, stripped, True)
            set_paragraph_spacing(para, before=60, after=60, line=240, line_rule='auto')
            return para
        # تشخیص شکل تصویر placeholder
        if 'شکل-۱' in stripped or 'شکل ۱.' in stripped and 'نردبان' in stripped:
            # این همان caption شکل است که بالا هندل شد؛ اگر به اینجا رسید یعنی pattern نگرفته
            pass
        # حالت عادی: paragraph معمولی (متن اصلی)
        is_rtl = is_mostly_persian(stripped)
        # اگر در بخش Abstract انگلیسی هستیم (current_section Abstract) و متن انگلیسی است، LTR
        if current_section.strip().lower() == 'abstract' and not is_rtl:
            para = doc.add_paragraph(style='Normal')
            set_paragraph_rtl(para, False)
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            add_formatted_runs(para, stripped, False)
            set_paragraph_spacing(para, before=80, after=80, line=259, line_rule='auto')
            return para
        # پیش‌فرض فارسی
        if is_rtl:
            para = doc.add_paragraph(style='متن')
            set_paragraph_rtl(para, True)
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            # برای پاراگراف اصلی، تورفتگی اول خط 0.5cm?
            # قالب متن دارای firstLine 565? در P44 دیدیم firstLine 565 (0.99cm) برای پاراگراف‌های دارای تصویر
            # برای متن اصلی معمولا اول خط تورفتگی ندارد، اما می‌توانیم بگذاریم
            add_formatted_runs(para, stripped, True)
            set_paragraph_spacing(para, before=60, after=60, line=259, line_rule='auto')
            return para
        else:
            # انگلیسی
            para = doc.add_paragraph(style='Normal')
            set_paragraph_rtl(para, False)
            para.alignment = WD_ALIGN_PARAGRAPH.LEFT if 'References' not in current_section else WD_ALIGN_PARAGRAPH.LEFT
            # justify برای انگلیسی abstract؟ می‌گذاریم justify
            if current_section.lower().strip() in ['abstract', '1. introduction / context', '2. problem statement']:
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            add_formatted_runs(para, stripped, False)
            set_paragraph_spacing(para, before=60, after=60, line=259, line_rule='auto')
            return para

    # ---- پیمایش بلوک‌ها ----
    for idx, (btype, data) in enumerate(blocks):
        if btype == 'heading':
            level, text = data
            add_heading_block(level, text)
        elif btype == 'paragraph':
            # اگر پاراگراف خالی؟ بگذر
            if not data.strip():
                continue
            # اگر پاراگراف مربوط به شکل باشد و بعد تصویر می‌آید، هندل ویژه
            # تشخیص اینکه این پاراگراف کپشن شکل است و باید تصویر بعدش بیاید
            # در md بعد از "**شکل ۱. ..." معمولا خط بعدی مأخذ است و سپس تصویر در docx باید بیاید
            # ما تصویر را بعد از caption اضافه می‌کنیم
            para = add_paragraph_block(data)
            # اگر این کپشن شکل بود، تصویر را بعدش اضافه کن
            if data.strip().startswith('**شکل ۱.') and 'نردبان' in data:
                # اضافه کردن تصویر
                try:
                    # فاصله قبل تصویر
                    p = doc.add_paragraph()
                    pPr = p._p.get_or_add_pPr()
                    jc = OxmlElement('w:jc')
                    jc.set(qn('w:val'), 'center')
                    pPr.append(jc)
                    set_paragraph_rtl(p, True)
                    run = p.add_run()
                    # سایز تصویر: عرض 14cm تقریبا یا 6 اینچ
                    # A4 usable width 12cm, پس 12cm
                    from docx.shared import Cm
                    run.add_picture(FIG_PNG, width=Cm(12.0))
                    # در صورت نبود فایل، خطا نده
                    # caption زیر تصویر دیگر لازم نیست چون مأخذ جدا می‌آید
                    print(f"✅ تصویر شکل ۱ درج شد در اندیس {idx}")
                except Exception as e:
                    print(f"⚠️ خطا در درج تصویر: {e}")
                    # fallback به متن
                    err_para = doc.add_paragraph(style='عنوان شکل و جدول')
                    set_paragraph_rtl(err_para, True)
                    err_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    add_formatted_runs(err_para, "[تصویر شکل ۱ — فایل PNG یافت نشد]", True)
                # بعد مأخذ خواهد آمد به عنوان paragraph بعدی
        elif btype == 'table':
            # تشخیص RTL بودن جدول: اگر اولین سلول فارسی است
            sample = " ".join([" ".join(r) for r in data[:2]])
            is_rtl_tbl = is_mostly_persian(sample)
            # اگر جدول مربوط به معادل‌های لاتین است (ستون دوم لاتین)، باز هم فارسی محسوب می‌شود اما ستون‌ها فارسی-لاتین
            # برای آن جدول خاص، ستون اول فارسی، دوم لاتین -> معکوس کردن ستون‌ها همچنان درست است
            create_table(doc, data, is_rtl_tbl)
            # فاصله بعد جدول
            doc.add_paragraph().add_run().add_break()  # فاصله کوچک
        elif btype == 'ulist':
            for item in data:
                is_rtl = is_mostly_persian(item)
                para = doc.add_paragraph(style='List Paragraph')
                set_paragraph_rtl(para, is_rtl)
                if is_rtl:
                    para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    # bullet فارسی با • و راست‌چین
                    # python-docx به صورت خودکار bullet نمی‌دهد، ما دستی •
                    # برای لیست راست‌به‌چپ، bullet در راست
                    add_formatted_runs(para, "• " + item, True)
                else:
                    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    add_formatted_runs(para, "• " + item, False)
                set_paragraph_spacing(para, before=40, after=40, line=240, line_rule='auto')
        elif btype == 'olist':
            for idx_i, item in enumerate(data, 1):
                is_rtl = is_mostly_persian(item)
                para = doc.add_paragraph(style='List Paragraph')
                set_paragraph_rtl(para, is_rtl)
                # اعداد فارسی
                fa_num = str(idx_i)
                # تبدیل به فارسی
                fa_digits = str.maketrans("0123456789", "۰۱۲۳۴۵۶۷۸۹")
                fa_num = fa_num.translate(fa_digits) if is_rtl else str(idx_i)
                prefix = f"{fa_num}. "
                add_formatted_runs(para, prefix + item, is_rtl)
                set_paragraph_spacing(para, before=40, after=40, line=240, line_rule='auto')
        elif btype == 'blockquote':
            is_rtl = is_mostly_persian(data)
            para = doc.add_paragraph(style='متن')
            set_paragraph_rtl(para, is_rtl)
            para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            # نقل قول: ایتالیک + تورفتگی
            pPr = para._p.get_or_add_pPr()
            ind = pPr.find(qn('w:ind'))
            if ind is None:
                ind = OxmlElement('w:ind')
                pPr.append(ind)
            ind.set(qn('w:right'), '720')  # 0.5 inch indent برای RTL
            ind.set(qn('w:left'), '720')
            add_formatted_runs(para, data, is_rtl)
            for run in para.runs:
                run.italic = True
            set_paragraph_spacing(para, before=60, after=60, line=240, line_rule='auto')
        elif btype == 'hr':
            # خط جداکننده: یک پاراگراف با border bottom یا سه ستاره
            para = doc.add_paragraph()
            pPr = para._p.get_or_add_pPr()
            pBdr = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '4')
            bottom.set(qn('w:space'), '1')
            bottom.set(qn('w:color'), 'B9C4CA')
            pBdr.append(bottom)
            pPr.append(pBdr)
            # خالی
            para.add_run()
            set_paragraph_spacing(para, before=120, after=120, line=240, line_rule='auto')
        elif btype == 'image_md':
            # اگر در md تصویر با ![]()
            # استخراج مسیر
            m = re.search(r'!\[.*?\]\((.*?)\)', data)
            if m:
                img_path = m.group(1).strip()
                try:
                    para = doc.add_paragraph()
                    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = para.add_run()
                    from docx.shared import Cm
                    run.add_picture(img_path, width=Cm(12.0))
                except Exception as e:
                    print(f"⚠️ تصویر md درج نشد {img_path}: {e}")

    # ---- تنظیمات نهایی سند (فونت پیش‌فرض، زبان) ----
    # اطمینان از وجود فونت‌های فارسی در document
    # ذخیره
    doc.save(OUT_DOCX)
    doc.save(OUT_DOCX2)
    print(f"✅ فایل‌های موقت ساخته شد، در حال ترمیم سکشن و هدرها…")
    # --- ترمیم سکشن: اطمینان از حفظ هدر/فوترهای قالب (6 مرجع) ---
    def patch_section(path):
        import zipfile, shutil, os
        # خواندن sectPr کامل از قالب
        with zipfile.ZipFile(TEMPLATE) as zt:
            t_doc = zt.read("word/document.xml").decode()
            sectprs = re.findall(r'<w:sectPr.*?</w:sectPr>', t_doc, re.DOTALL)
            # انتخاب کامل‌ترین (بیشترین headerReference)
            target = max(sectprs, key=lambda s: s.count("headerReference"))
            # اطمینان از داشتن bidi و ... (target دارد)
        tmp = path + ".patched.tmp"
        with zipfile.ZipFile(path, 'r') as zi:
            with zipfile.ZipFile(tmp, 'w', compression=zipfile.ZIP_DEFLATED) as zo:
                for item in zi.infolist():
                    data = zi.read(item.filename)
                    if item.filename == "word/document.xml":
                        content = data.decode()
                        # حذف تمام sectPrهای موجود (داخل pPr و مستقیم)
                        content_no = re.sub(r'<w:sectPr[^>]*>.*?</w:sectPr>', '', content, flags=re.DOTALL)
                        # تمیزکاری: ممکن است sectPr داخل w:pPr خالی مانده باشد، اشکالی ندارد
                        # افزودن target قبل از </w:body> به عنوان سکشن پایانی (مستقیم زیر body)
                        if '</w:body>' in content_no:
                            content_new = content_no.replace('</w:body>', target + '</w:body>')
                        else:
                            content_new = content_no + target
                        data = content_new.encode('utf-8')
                    zo.writestr(item, data)
        shutil.move(tmp, path)
        print(f"  🔧 {path} ترمیم شد (sectPr کامل)")
    patch_section(OUT_DOCX)
    patch_section(OUT_DOCX2)
    print(f"✅ فایل‌های خروجی نهایی:\n - {OUT_DOCX}\n - {OUT_DOCX2}\nاندازه: {pathlib.Path(OUT_DOCX).stat().st_size/1024:.1f} KB")

if __name__ == "__main__":
    build_docx()
