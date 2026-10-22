"""
قالب PDF المعتمد لتقارير قسم المواد الشرعية (تصميم نموذج «تقرير زيارة رئيس القسم»).
الاستعمال:
  python3 templates/build_report_pdf.py report.json out.pdf
report.json:
  {
    "box": "تقرير نشاط",                     # نص الصندوق الكحلي تحت «التوجيه الفني للعلوم الشرعية»
    "title": "عنوان التقرير",
    "info": [["المعلم", "أ. ..."], ["التاريخ", "..."], ["نوع النشاط", "..."], ["الفئة المستهدفة", "..."]],
    "text": "## مقدمة\nفقرة...\n- نقطة",       # نفس صيغة عمود «التقرير» في الشيت (يُتجاهل سطر «العنوان:»)
    "photos": ["photo1.jpg"],                  # اختياري
    "signs": [["رئيس القسم", "أ. خليفة القصار"], ["مدير المعهد", "د. حامد غازي العتيبي"]],
    "landscape": false
  }
يحتاج: الخطوط في fonts/ (NotoNaskhArabic[wght].ttf و NotoKufiArabic[wght].ttf) والشعارات من index.html، و playwright.
"""
import base64, html, json, os, re, sys
from playwright.sync_api import sync_playwright

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def logos():
    s = open(os.path.join(ROOT, "index.html"), encoding="utf-8").read()
    get = lambda n: re.search(r'const %s = "([^"]+)"' % n, s).group(1)
    return get("MOE_B64"), get("GUID_B64"), get("LOGO_B64")

def body_html(text):
    out, items = [], []
    def flush():
        if items: out.append("<ul>" + "".join(f"<li>{html.escape(x)}</li>" for x in items) + "</ul>"); items.clear()
    for raw in str(text or "").split("\n"):
        l = raw.strip()
        if not l or re.match(r"^العنوان\s*:", l): continue
        if l.startswith("#"): flush(); out.append(f"<h3>{html.escape(l.lstrip('#').strip())}</h3>")
        elif re.match(r"^[-•*]\s+", l): items.append(re.sub(r"^[-•*]\s+", "", l))
        else: flush(); out.append(f"<p>{html.escape(l)}</p>")
    flush()
    return "".join(out)

def build(spec, out_pdf, fonts_dir):
    moe, guid, inst = logos()
    land = spec.get("landscape")
    photos = "".join(f'<img src="data:image/jpeg;base64,{base64.b64encode(open(p, "rb").read()).decode()}">' for p in spec.get("photos", []))
    info = spec.get("info", [])
    rows = "".join("<tr>" + "".join(f'<td class="k">{html.escape(k)}:</td><td>{html.escape(v)}</td>' for k, v in info[i:i + 2]) + "</tr>" for i in range(0, len(info), 2))
    signs = "".join(f'<div><b>{html.escape(r)}</b><div class="n">{html.escape(n)}</div></div>' for r, n in spec.get("signs", []))
    F = "file://" + os.path.abspath(fonts_dir) + "/"
    doc = f'''<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><style>
@font-face{{font-family:Naskh;src:url("{F}NotoNaskhArabic[wght].ttf");font-weight:100 900}}
@font-face{{font-family:Kufi;src:url("{F}NotoKufiArabic[wght].ttf");font-weight:100 900}}
@page{{size:A4 {"landscape" if land else "portrait"};margin:0}}
*{{box-sizing:border-box}} body{{margin:0;font-family:Naskh,serif;color:#1f2a33}}
.page{{width:{"297mm" if land else "210mm"};height:{"210mm" if land else "297mm"};padding:8mm}}
.frame{{border:3px double #1E3A5F;height:100%;padding:5mm 8mm 4mm;display:flex;flex-direction:column}}
.head{{display:grid;grid-template-columns:34mm 1fr 34mm;align-items:center}}
.head .moe{{width:27mm;justify-self:center}} .head .inst{{width:28mm;justify-self:center}}
.mid{{text-align:center}} .mid .g{{width:10mm}} .mid .t{{font-family:Kufi;color:#1E3A5F;font-weight:700;font-size:12pt;margin:1mm 0 1.5mm}}
.box{{display:inline-block;background:#1E3A5F;color:#fff;font-family:Kufi;font-weight:700;font-size:13pt;padding:1.2mm 12mm;border-radius:2.5mm}}
.rule{{border-bottom:3px double #1E3A5F;margin:3mm 0 3mm}}
h1{{font-family:Kufi;color:#1E3A5F;text-align:center;font-size:14pt;margin:0 0 2.5mm;font-weight:800}}
table.info{{width:100%;border-collapse:collapse;font-size:10pt;margin-bottom:2mm}}
table.info td{{border:1px solid #C9D6E3;padding:1.1mm 2.5mm}} table.info td.k{{background:#EEF3F8;color:#1E3A5F;font-weight:700;white-space:nowrap;width:1%}}
.body{{flex:1;font-size:10.2pt;line-height:1.62;text-align:justify}}
h3{{font-family:Kufi;color:#1E3A5F;font-size:10.5pt;margin:1.8mm 0 .7mm;font-weight:700;background:#EEF3F8;padding:.4mm 2mm;border-radius:1mm}}
p{{margin:0 0 1mm}} ul{{margin:0;padding-right:5mm}} li::marker{{color:#1E3A5F}}
.photos{{float:left;width:44mm;margin:1.5mm 4mm 1mm 0;display:grid;gap:2mm}} .photos img{{width:44mm;border-radius:1.5mm;border:1px solid #C9D6E3}}
.sign{{display:grid;grid-template-columns:repeat({max(1, len(spec.get("signs", [])))},1fr);text-align:center;font-size:10.5pt;margin-top:2mm}}
.sign b{{display:block;color:#1E3A5F;font-family:Kufi}} .sign .n{{font-weight:700;margin-bottom:9mm}}
.foot{{text-align:center;font-size:7.5pt;color:#6B7782}}
</style></head><body><div class="page"><div class="frame">
<div class="head"><img class="moe" src="data:image/png;base64,{moe}"><div class="mid"><img class="g" src="data:image/png;base64,{guid}"><div class="t">التوجيه الفني للعلوم الشرعية</div><span class="box">{html.escape(spec.get("box", ""))}</span></div><img class="inst" src="data:image/png;base64,{inst}"></div>
<div class="rule"></div>
<h1>{html.escape(spec.get("title", ""))}</h1>
{f'<table class="info">{rows}</table>' if rows else ""}
<div class="body">{f'<div class="photos">{photos}</div>' if photos else ""}{spec.get("html") or body_html(spec.get("text"))}</div>
<div class="sign">{signs}</div>
<div class="foot">معهد عبدالرحمن السميط الديني · قسم المواد الشرعية · رئيس القسم: أ. خليفة حمد القصار · مدير المعهد: د. حامد غازي العتيبي</div>
</div></div></body></html>'''
    tmp = os.path.abspath(out_pdf) + ".html"
    open(tmp, "w", encoding="utf-8").write(doc)
    with sync_playwright() as p:
        b = p.chromium.launch(); pg = b.new_page(); pg.goto("file://" + tmp); pg.wait_for_timeout(700)
        over = pg.evaluate('document.querySelector(".frame").scrollHeight - document.querySelector(".frame").clientHeight')
        pg.pdf(path=out_pdf, prefer_css_page_size=True, print_background=True); b.close()
    os.remove(tmp)
    return over  # > 0 يعني أن المحتوى تجاوز الصفحة: اختصر النص أو صغّر الصور

if __name__ == "__main__":
    spec = json.load(open(sys.argv[1], encoding="utf-8"))
    fonts = sys.argv[3] if len(sys.argv) > 3 else os.path.join(ROOT, "fonts")
    print("overflow_px", build(spec, sys.argv[2], fonts))
