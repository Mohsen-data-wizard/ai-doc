# -*- coding: utf-8 -*-
"""ساخت HTML راست‌به‌چپ از نسخه‌ی ارسال — با فونت و تصویر جاسازی‌شده (بدون منبع بیرونی)."""
import re, base64, html
SRC='مقاله-بی-نام-نسخه-سابمیت.md'; OUT='مقاله-بی-نام-نسخه-سابمیت.html'
def b64(p): return base64.b64encode(open(p,'rb').read()).decode()
F_REG=b64('fonts/Vazirmatn-Regular.ttf'); F_BOLD=b64('fonts/Vazirmatn-Bold.ttf')
FIG=b64('شکل-۱-نردبان-دامنه‌ها.png')
def norm(t): return t.replace('⟦','[ ').replace('⟧',' ]')
def inl(t):
    t=html.escape(norm(t))
    t=re.sub(r'\*\*(.+?)\*\*',r'<strong>\1</strong>',t)
    t=re.sub(r'(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)',r'<em>\1</em>',t)
    t=re.sub(r'`(.+?)`',r'<code>\1</code>',t)
    return t
def is_ltr(s):
    lat=len(re.findall(r'[A-Za-z]',s)); allc=len(re.findall(r'\S',s)) or 1
    return lat/allc>0.6
lines=open(SRC,encoding='utf-8').read().split('\n')
body=[]; i=0
while i<len(lines):
    ln=lines[i].rstrip()
    if not ln.strip(): i+=1; continue
    if ln.startswith('|'):
        rows=[]
        while i<len(lines) and lines[i].startswith('|'):
            cells=[c.strip() for c in lines[i].strip().strip('|').split('|')]
            if not all(re.fullmatch(r'[-: ]*',c) for c in cells): rows.append(cells)
            i+=1
        ncol=max(len(r) for r in rows); rows=[r+['']*(ncol-len(r)) for r in rows]
        t=['<table><thead><tr>']+[f'<th>{inl(c)}</th>' for c in rows[0]]+['</tr></thead><tbody>']
        for r in rows[1:]: t.append('<tr>'+''.join(f'<td>{inl(c)}</td>' for c in r)+'</tr>')
        t.append('</tbody></table>'); body.append(''.join(t)); continue
    if ln.startswith('**شکل ۱.') and 'نردبان' in ln:
        body.append(f'<figure><img src="data:image/png;base64,{FIG}" alt="شکل ۱"></figure>')
        i+=1
        while i<len(lines) and (lines[i].startswith('مأخذ:') or not lines[i].strip()): i+=1
        continue
    if ln.startswith('# '):   body.append(f'<h1>{inl(ln[2:])}</h1>')
    elif ln.startswith('## '):body.append(f'<h2>{inl(ln[3:])}</h2>')
    elif ln.startswith('### '):body.append(f'<h3>{inl(ln[4:])}</h3>')
    elif ln.startswith('> '): body.append(f'<blockquote>{inl(ln[2:])}</blockquote>')
    elif re.match(r'^[-*] ',ln): body.append(f'<li>{inl(ln[2:])}</li>')
    elif re.fullmatch(r'[-*_]{3,}',ln.strip()): body.append('<hr>')
    elif is_ltr(ln): body.append(f'<p class="en" dir="ltr">{inl(ln)}</p>')
    else: body.append(f'<p>{inl(ln)}</p>')
    i+=1
html_out=f'''<!DOCTYPE html>
<html lang="fa" dir="rtl"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>اقتصاد سنجش خسارت انرژی جنگ — نسخه‌ی ارسال</title>
<style>
@font-face{{font-family:'Vazir';src:url(data:font/ttf;base64,{F_REG}) format('truetype');font-weight:400}}
@font-face{{font-family:'Vazir';src:url(data:font/ttf;base64,{F_BOLD}) format('truetype');font-weight:700}}
*{{box-sizing:border-box}}
body{{font-family:'Vazir',Tahoma,sans-serif;line-height:2;max-width:52rem;margin:0 auto;padding:2rem 1.6rem;color:#16232c;background:#fff;font-size:15.5px}}
h1{{font-size:1.6rem;color:#12303f;border-bottom:3px solid #12303f;padding-bottom:.5rem;margin:2rem 0 1rem}}
h2{{font-size:1.25rem;color:#12303f;margin:1.8rem 0 .6rem}}
h3{{font-size:1.08rem;color:#1d4b63}}
p{{text-align:justify;margin:.55rem 0}}
p.en{{direction:ltr;text-align:left;font-size:.95em;color:#22303a}}
blockquote{{background:#f2f7fa;border-right:4px solid #12303f;margin:.9rem 0;padding:.7rem 1rem;border-radius:6px;color:#2b3b45}}
table{{width:100%;border-collapse:collapse;margin:1rem 0;font-size:.86em}}
th,td{{border:1px solid #b9c4ca;padding:.45rem .55rem;text-align:right;vertical-align:middle}}
th{{background:#eef4f7;color:#12303f}}
tr:nth-child(even) td{{background:#fafcfd}}
li{{margin:.35rem 1.2rem 0 0}}
figure{{margin:1.2rem 0;text-align:center}}
figure img{{max-width:100%;height:auto;border:1px solid #e3eaee;border-radius:6px}}
hr{{border:0;border-top:1px dashed #cbd6dc;margin:1.6rem 0}}
strong{{color:#0f2b38}}
code{{background:#f1f5f7;padding:.07em .3em;border-radius:4px;font-size:.9em}}
sup{{font-size:.72em}}
@media print{{body{{max-width:none;font-size:12.5px}} h1{{page-break-before:auto}} figure{{page-break-inside:avoid}}}}
</style></head><body>
{''.join(body)}
</body></html>'''
open(OUT,'w',encoding='utf-8').write(html_out)
print('✅ HTML ساخته شد:', OUT, '|', round(len(html_out)/1024), 'KB')
