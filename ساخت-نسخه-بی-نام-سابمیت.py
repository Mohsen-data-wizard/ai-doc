# -*- coding: utf-8 -*-
"""ساخت «فایل نخست» ارسال: نسخه‌ی کامل مقاله، بدون مشخصات نویسندگان (داوری دوسوناشناس).

اجرا:  python3 ساخت-نسخه-بی-نام-سابمیت.py
ورودی: مقاله-نسخه-ارسال-EENR.md   (نسخه‌ی با نام — برای بسته‌ی داخلی)
خروجی: مقاله-بی-نام-نسخه-سابمیت.md

قاعده‌ها:
 ۱. هیچ نشانی از هویت نویسنده (نام، ایمیل، وابستگی، ORCID) در خروجی نمی‌ماند.
 ۲. هیچ قلم ⟦…⟧ باز نمی‌ماند؛ جای هر قلم، جمله‌ی کامل جایگزین می‌شود.
 ۳. متن علمی، ارجاع‌ها و پانویس‌ها دست‌نخورده می‌مانند.
"""
import re
import sys

SRC = 'مقاله-نسخه-ارسال-EENR.md'
OUT = 'مقاله-بی-نام-نسخه-سابمیت.md'

lines = open(SRC, encoding='utf-8').read().split('\n')

BLIND_NOTE = ('**نسخه‌ی داوری دوسوناشناس:** مطابق راهنمای نشریه، این فایل «کامل مقاله بدون '
              'مشخصات نویسندگان» است؛ چکیده‌ی فارسی و لاتین و کلیدواژه‌ها در همین فایل ضمیمه‌اند. '
              'مشخصات کامل نویسندگان و شناسه‌های ORCID، مطابق الزام سامانه، در فایل جداگانه بارگذاری شده است.')

ACK = ('این پژوهش بر پایه‌ی داده‌های عمومی و اسناد منتشرشده انجام شده است و در جریان آن از '
       'حمایت مالی نهاد بیرونی استفاده نشده است.')

out, done = [], {'head': 0, 'orcid': 0, 'ack': 0, 'cite': 0}
i = 0
while i < len(lines):
    ln = lines[i]

    # ۱) بلوک مشخصات نویسندگان در صفحه‌ی نخست → یادداشت کورسازی
    if ln.startswith('**نویسندگان:**'):
        out.append(BLIND_NOTE)
        while i + 1 < len(lines) and (lines[i + 1].startswith('**نویسنده‌ی مسئول')
                                      or lines[i + 1].startswith('**پانویس صفحه‌ی نخست')):
            i += 1
        done['head'] = 1
        i += 1
        continue

    # ۲) جدول ORCID → یادداشت کورسازی
    if ln.strip() == '| نویسنده | ORCID |':
        out.append('برای رعایت داوری دوسوناشناس، شناسه‌های ORCID در این نسخه درج نشده‌اند؛'
                   ' در فایل مشخصات نویسندگان ارائه شده‌اند.')
        out.append('')
        while i < len(lines) and (lines[i].startswith('|') or not lines[i].strip()):
            i += 1
            if i < len(lines) and lines[i].startswith('## '):
                break
        done['orcid'] = 1
        continue

    # ۳) قلم تقدیر و سپاسگزاری
    if ln.startswith('⟦یک پاراگراف کوتاه برای تقدیر'):
        out.append(ACK)
        done['ack'] = 1
        i += 1
        continue

    # ۴) قطعه‌ی «استناد به این مقاله» (وابسته به DOI و شماره‌ی آینده)
    if 'استناد به این مقاله' in ln and ln.startswith('## '):
        out.append(ln)
        out.append('')
        out.append('*این بخش پس از تخصیص شناسه‌ی دیجیتال (DOI) و تعیین شماره‌ی نشریه تکمیل می‌شود.*')
        out.append('')
        i += 1
        while i < len(lines) and not lines[i].startswith('## ') and not lines[i].startswith('---'):
            i += 1
        done['cite'] = 1
        continue

    out.append(ln)
    i += 1

t = '\n'.join(out)

# ── کنترل‌ها ────────────────────────────────────────────────────────────────
assert all(done.values()), f'جایگزینی ناقص: {done}'
assert '⟦' not in t and '⟧' not in t, 'قلم باز در فایل بی‌نام مانده است'
leaks = [p for p in ('اسکندری', 'Eskandaripour', 'vru.ac.ir', 'orcid.org',
                     '0000-0001-9550-1972', 'رفسنجان', 'Rafsanjan', '0913')
         if p in t]
assert not leaks, f'نشانی هویت در فایل مانده: {leaks}'
for sec in ('## چکیده', '## Abstract', '## ۱. مقدمه', '## منابع', '## References', '## پانویس‌ها'):
    assert sec in t, f'بخش گم‌شده: {sec}'
wc = len(re.sub(r'\s+', ' ', t).split())
open(OUT, 'w', encoding='utf-8').write(t)
print(f'✅ {OUT} ساخته شد · {len(t.split(chr(10)))} سطر · ≈{wc} توکن · صفر قلم باز · صفر نشانی هویت')
