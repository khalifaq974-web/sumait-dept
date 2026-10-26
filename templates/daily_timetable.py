"""
جدول الحصص اليومي لمعلمي القسمين بالديباجة (شعار الوزارة والمعهد فقط) — يُرسل لرئيس الشعبة كل صباح.
python3 templates/daily_timetable.py REPO OUT.pdf YYYY-MM-DD [covers.json]
"""
# covers.json (اختياري): {"absent": ["اسم"], "covers": [{"period": 3, "absent": "اسم", "sub": "اسم البديل", "swap": false}], "returns": [{"teacher": "صاحب الحصة", "period": 2, "by": "من يردّها"}]}
import sys, html, json, re, datetime
REPO, OUT, ISO = sys.argv[1], sys.argv[2], sys.argv[3]
COV = json.load(open(sys.argv[4], encoding="utf-8")) if len(sys.argv) > 4 else {}
src = open(REPO + "/templates/build_report_pdf.py", encoding="utf-8").read()
src = src.replace('<img class="g" src="data:image/png;base64,{guid}">', '')
src = src.replace('التوجيه الفني للعلوم الشرعية', 'معهد عبدالرحمن السميط الديني')
src = src.replace('font-size:9pt;margin:1mm 0 2mm', 'font-size:11pt;margin:1mm 0 2mm')
src = src.replace('margin-bottom:9mm', 'margin-bottom:6mm').replace('margin:3mm 0 3mm', 'margin:2mm 0 2mm')
src = src.replace('ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))', 'ROOT = %r' % REPO)
tpl = {}; exec(compile(src.split('if __name__')[0], "tpl", "exec"), tpl)

idx = open(REPO + "/index.html", encoding="utf-8").read()
TT = json.loads(re.search(r"const TIMETABLE = (\{.*?\n\});", idx, re.S).group(1).replace("\n", ""))
d = datetime.date.fromisoformat(ISO)
DAY = (d.weekday() + 1) % 7  # 0 الأحد … 4 الخميس
DAYS = ["الأحد", "الإثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"]
MON = ["يناير","فبراير","مارس","أبريل","مايو","يونيو","يوليو","أغسطس","سبتمبر","أكتوبر","نوفمبر","ديسمبر"]
HD = ["أحمد صادق", "أحمد صلاح", "عماد", "كرم"]  # معلمو قسم الحديث، والباقون قرآن (رئيس الشعبة مستثنى)
teachers = [t for t in TT if not t.startswith("خليفة")]
DEPTS = [("قسم الحديث الشريف", "hd", [t for t in teachers if any(t.startswith(h) for h in HD)]),
         ("قسم القرآن الكريم", "qr", [t for t in teachers if not any(t.startswith(h) for h in HD)])]
SUBJ = {"hd": "حديث", "qr": "قرآن"}
PN = ["", "الأولى", "الثانية", "الثالثة", "الرابعة", "الخامسة", "السادسة", "السابعة"]
absent = set(COV.get("absent", []))
subs = {(c["absent"], int(c["period"])): (c.get("sub", ""), c.get("swap")) for c in COV.get("covers", [])}
rets = {(r["teacher"], int(r["period"])): r["by"] for r in COV.get("returns", [])}
short = lambda n: " ".join(n.split()[:2])

def cell(t, dept, p, v):
    if v is None: return '<td class="free"></td>'
    if v in ("R", "RF"): return f'<td class="res">{"احتياط" if v == "R" else "احتياط فنية"}</td>'
    s, c = v.split(":"); g, n = c.split("-")
    tag = f'<span class="sub">{SUBJ[s]}</span>' if s != dept else ""
    if t in absent:
        sb, sw = subs.get((t, p), ("", False))
        tag += f'<span class="cov">{"تبديل" if sw else "بديل"}: {html.escape(short(sb))}</span>' if sb else '<span class="cov none">غير مغطاة</span>'
    elif (t, p) in rets:
        tag += f'<span class="cov ret">يدخلها: {html.escape(short(rets[(t, p)]))} (رد تبديل)</span>'
    return f'<td class="cls"><b>{g}/{n}</b>{tag}</td>'

head = '<tr><th class="nm">المعلم</th>' + "".join(f"<th>{PN[p]}</th>" for p in range(1, 8)) + '<th>المجموع</th></tr>'
rows = []
for title, dp, ts in DEPTS:
    rows.append(f'<tr class="dept"><td colspan="9">{title}</td></tr>')
    for t in ts:
        day = {p: c for (dd, p, c) in TT[t] if dd == DAY}
        n = sum(1 for c in day.values() if c not in ("R", "RF"))
        ab = ' <span class="ab">(غائب)</span>' if t in absent else ""
        cls = ' class="abs"' if t in absent else ""
        rows.append(f'<tr{cls}><td class="nm">أ. {html.escape(t)}{ab}</td>' + "".join(cell(t, dp, p, day.get(p)) for p in range(1, 8)) + f'<td class="cnt">{n}</td></tr>')
style = '''<style>
table.grid th,table.grid td{text-align:center;vertical-align:middle;height:8mm}
table.grid th.nm,table.grid td.nm{text-align:right;width:58mm;padding-right:3mm;font-weight:700}
table.grid tr.dept td{text-align:right;padding-right:3mm;height:7mm;font-family:Kufi}
td.cls{padding-top:.3mm;padding-bottom:.3mm;line-height:1.15} td.cls b{font-size:12.5pt;color:#1E3A5F} td.cls .sub{display:block;font-size:8pt;color:#B4382B;font-weight:700;line-height:1.1}
td.cls .cov{display:block;font-size:7.5pt;color:#2F7D4F;font-weight:700;line-height:1.1} td.cls .cov.none{color:#B4382B} td.cls .cov.ret{color:#6B4FB5}
tr.abs td.nm{color:#B4382B} .ab{font-size:9pt}
td.res{background:#FFF6E5;color:#9A6400;font-weight:700;font-size:9.5pt} td.free{background:#FAFBFC}
td.cnt{font-weight:800;color:#1E3A5F;background:#EEF3F8}
.legend{font-size:9pt;color:#5A6670;margin-top:1mm}
</style>'''
body = style + f'<table class="grid">{head}{"".join(rows)}</table>' + \
  '<div class="legend">الخلية الفارغة = لا توجد حصة · «احتياط» حصص الاحتياط المسندة في الجدول ولا تُحسب في المجموع · الكلمة الحمراء تحت الفصل = مادة غير مادة القسم.</div>'
spec = {"box": "جدول الحصص اليومي", "title": "جدول حصص معلمي قسم الحديث الشريف والقرآن الكريم",
        "info": [["اليوم", DAYS[DAY]], ["التاريخ", f"{d.day} {MON[d.month-1]} {d.year}م"]],
        "html": body, "landscape": True,
        "signs": [["رئيس الشعبة", "أ. خليفة القصار"], ["مدير المعهد", "د. حامد غازي العتيبي"]]}
print("overflow", tpl["build"](spec, OUT, REPO + "/fonts"))
