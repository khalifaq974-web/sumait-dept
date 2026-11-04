#!/usr/bin/env python3
"""يولّد gas/Index.html (صفحة الموقع داخل Apps Script في حساب الوزارة) من index.html.

لماذا: HtmlService يعدّل نص السكربت المضمّن (مثل قطع «https://») فيتعطل الموقع كله بصمت.
الحل: كل سكربت مضمّن يُرمَّز base64 ويُفكّ عند التحميل (atob + TextDecoder) ويُضاف عنصر <script>.
- يحذف <script src="config.js"> (لا يوجد في Apps Script؛ الاتصال عبر google.script.run).
- يحذف الوسوم التي يتولاها doGet (viewport) — يبقى الباقي كما هو.
الاستعمال: python3 tools/build_gas_index.py   (من جذر المستودع)
"""
import base64, re, sys, pathlib

root = pathlib.Path(__file__).resolve().parent.parent
src = (root / "index.html").read_text(encoding="utf-8")

# رابط الموقع في حساب الوزارة (يظهر في رسائل الواتساب)؛ الصفحة داخل Apps Script تعمل في إطار عنوانه مختلف
SITE_URL = "https://script.google.com/a/macros/moe.edu.kw/s/AKfycbx36T9qahEPemxQzSmZ8-nav6LGuTavEgUFahlbHzX5IATKbF-ayeUGwE6fUkEAZbvK/exec"
src = src.replace('<script src="config.js"></script>', '<script>window.API_URL = ""; window.SITE_URL = "%s";</script>' % SITE_URL)
src = re.sub(r'<meta name="viewport"[^>]*>\s*', "", src, count=1)

LOADER = ('<script>(function(){var b="%s";var s=document.createElement("script");'
          's.text=new TextDecoder("utf-8").decode(Uint8Array.from(atob(b),function(c){return c.charCodeAt(0);}));'
          'document.currentScript.parentNode.insertBefore(s,document.currentScript.nextSibling);})();</script>')

n = 0
def enc(m):
    global n
    attrs, body = m.group(1), m.group(2)
    if "src=" in attrs:
        return m.group(0)
    n += 1
    return LOADER % base64.b64encode(body.encode("utf-8")).decode("ascii")

out = re.sub(r"<script([^>]*)>(.*?)</script>", enc, src, flags=re.S)
if "https://" in "".join(re.findall(r"<script>(.*?)</script>", out, flags=re.S)):
    sys.exit("خطأ: بقي نص https:// داخل سكربت مضمّن")
(root / "gas" / "Index.html").write_text(out, encoding="utf-8")
print(f"gas/Index.html جاهز: {n} سكربت مرمّز، الحجم {len(out.encode('utf-8'))//1024} ك.ب")
