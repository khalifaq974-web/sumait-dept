/**
 * موقع قسم المواد الشرعية (الحديث الشريف · القرآن الكريم) — معهد عبدالرحمن السميط الديني
 * الأركان: ما قطع من المنهج · الإنجازات · محاضر الاجتماعات · الزيارات الصفية
 * بعد أي تحديث للكود: شغّل setup مرة واحدة، ثم حدّث النشر (Manage deployments ← ✏️ ← New version).
 */

const SHEET_CODES = "الرموز";
const SHEET_DATA = "الإدخالات";
const SHEET_LOG = "السجل";
const SHEET_ACH = "الإنجازات";
const SHEET_NSH = "النشرات والتعاميم";
const SHEET_MRP = "تقارير إنجازات القسم";
const SHEET_ACTLOG = "الإشعارات";
const SHEET_NSHACK = "اطلاع النشرات";
const SHEET_MIN = "المحاضر";
const SHEET_SIG = "التواقيع";
const SHEET_MYSIG = "توقيعات المعلمين";
const SHEET_VIS = "الزيارات";
const SHEET_MAIL = "سجل التنبيهات";
const SHEET_ABS = "الغياب";
const SHEET_COV = "الاحتياط";

const ST_PENDING = "بانتظار الصياغة";
const ST_READY = "جاهز";

// التنبيهات بالإيميل (تصل لرئيس الشعبة فقط)
const HEAD_EMAIL = "khalifaq974@gmail.com";
const SITE_URL = "https://khalifaq974-web.github.io/sumait-dept/";
const SITE_NAME = "موقع قسم المواد الشرعية";

// القسمان: hd = الحديث الشريف، qr = القرآن الكريم
const SUBJ = { hd: "الحديث الشريف", qr: "القرآن الكريم" };
const MAX_DONE = { hd: 22, qr: 22 };

// المعلمون وفصولهم من الجدول العام. كل فصل يُكتب "القسم:الصف-الشعبة"، مثل "hd:11-1" = الحديث الشريف ١١/١
const TEACHERS = [
  ["خليفة حمد خليفة القصار", ["hd:11-1"], true],
  ["أحمد صادق أحمد أبوسعدة", ["hd:11-2", "hd:11-3", "hd:11-4", "hd:11-5"]],
  ["أحمد صلاح زيدان القناوي", ["hd:9-1", "hd:9-2", "hd:9-5"]],
  ["عماد صادق علي العايق", ["hd:10-3", "hd:10-5", "hd:12-1", "hd:12-3"]],
  ["كرم عبدالحميد أبوزيد أحمد", ["hd:10-1", "hd:10-2", "hd:10-4", "hd:12-4", "hd:12-5"]],
  ["منذر خالد مطلق الحربي", ["qr:12-1", "qr:12-2", "qr:12-3", "qr:12-4", "qr:12-5", "hd:9-3", "hd:9-4", "hd:12-2"]],
  ["معاذ وليد مبارك الشطي", ["qr:9-1", "qr:9-2", "qr:9-3", "qr:9-4", "qr:9-5"]],
  ["يوسف عدنان عبدالكريم الشطي", ["qr:10-1", "qr:10-2", "qr:10-3", "qr:10-4", "qr:10-5"]],
  ["عبدالله أحمد عبدالمجد الماجد", ["qr:11-1", "qr:11-2", "qr:11-3", "qr:11-4", "qr:11-5"]]
];

// الحسابات التجريبية: تجرّب كل شيء، وبياناتها معزولة ولا تظهر في أي تقرير أو متابعة
const DEMOS = { "تجريبي ١": ["hd:11-1", "qr:12-1"], "تجريبي ٢": ["hd:9-1", "qr:10-1"] };
const DEMO_MINUTE = "demo-minute";
function isDemo_(name) { return Object.prototype.hasOwnProperty.call(DEMOS, name); }
function demoPrefix_(me) { return me.isDemo ? "DEMO|" + me.name + "|" : ""; }
function classesOf_(name) { const t = TEACHERS.find(x => x[0] === name); return t ? t[1] : (isDemo_(name) ? DEMOS[name] : null); }

function doGet() {
  return ContentService.createTextOutput("خادم " + SITE_NAME + " يعمل. افتح الموقع من: " + SITE_URL);
}

// ======================= واجهة الموقع الخارجي (GitHub Pages) =======================
// الصفحة المستضافة خارج جوجل ترسل الطلبات هنا، فلا يتأثر المعلم بتعدد حسابات جوجل في متصفحه.
const API = { login, saveEntry, headData, submitAchievement, myAchievements, editAchievement, deleteAchievement, allAchievements, saveReport,
  getPhoto, listMinutes, getMinute, saveMinute, deleteMinute, signMinute, mySignature, saveFileToDrive, demoAccounts, demoAction,
  saveVisit, listVisits, getVisit, signVisit, deleteVisit, reserveData, saveAbsence, deleteAbsence, setSwapReturn, requestSwap, listSwaps, cancelSwap, classesOverview, classSheet, saveGrades,
  listStudentReports, saveStudentReport, seenStudentReport, deleteStudentReport, saveRemedialPlan,
  myProfile, saveProfile, listPermissions, getPermission, submitPermission, mailPermissionPdf, decidePermission, deletePermission,
  listCirculars, addCircular, getCircularFile, ackCircular, circularAcks, deleteCircular,
  listMonthlyReports, saveMonthlyReport, deleteMonthlyReport, listActivity, markActivitySeen, markSectionSeen,
  listTasks, saveTask, deleteTask, getTaskFile, respondTask, saveMyPhone, teacherPhones, phonesList,
  listAlerts, saveAlert, deleteAlert, tickAlert, approveMyReport };

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const f = Object.prototype.hasOwnProperty.call(API, req.fn) ? API[req.fn] : null;
    if (!f) throw new Error("طلب غير معروف.");
    const args = Array.isArray(req.args) ? req.args : [];
    out = { ok: true, result: f.apply(null, args) };
    if (ACT_LOG[req.fn]) { try { logAct_(req.fn, args, out.result); } catch (x) {} }
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// ======================= الإعداد =======================

/** يُشغَّل بعد كل تحديث للكود. لا يحذف شيئًا، يضيف الناقص فقط */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let codes = ss.getSheetByName(SHEET_CODES);
  if (!codes) {
    codes = ss.insertSheet(SHEET_CODES);
    codes.getRange("B:B").setNumberFormat("@");
    codes.appendRow(["المعلم", "الرمز", "رئيس الشعبة", "ملاحظة"]);
    codes.setRightToLeft(true);
    codes.setFrozenRows(1);
  }
  codes.getRange("B:B").setNumberFormat("@");
  // يضيف أي معلم أو حساب تجريبي غير موجود (لا يغيّر رموز الموجودين)
  const existing = codes.getRange(1, 1, Math.max(codes.getLastRow(), 1), 1).getDisplayValues().map(r => r[0]);
  TEACHERS.forEach(t => { if (existing.indexOf(t[0]) === -1) writeRow_(codes, codes.getLastRow() + 1, [t[0], newCode_(), t[2] ? "نعم" : "", ""]); });
  Object.keys(DEMOS).forEach(n => { if (existing.indexOf(n) === -1) writeRow_(codes, codes.getLastRow() + 1, [n, newCode_(), "", "تجريبي"]); });
  codes.autoResizeColumns(1, 4);
  sheet_(SHEET_DATA, ["المفتاح", "الفصل", "القسم", "المعلم", "الحصص المنجزة", "آخر ما أُنجز", "ملاحظة", "آخر تحديث"], true);
  sheet_(SHEET_LOG, ["الوقت", "المعلم", "الفصل", "القسم", "الحصص المنجزة", "ملاحظة"], true);
  sheet_(SHEET_ACH, ["المعرف", "المعلم", "التاريخ", "النوع", "العنوان", "الفئة المستهدفة", "وصف المعلم", "الصور", "الحالة", "التقرير", "تاريخ الرفع", "تاريخ الصياغة", "السجل", "النتائج والفائزون", "القسم"], true);
  sheet_(SHEET_MIN, ["المعرف", "العنوان", "التاريخ", "الوقت", "المكان", "الحضور", "جدول الأعمال", "ما دار في الاجتماع", "التوصيات", "تاريخ الإنشاء"], true);
  sheet_(SHEET_SIG, ["المحضر", "المعلم", "وقت التوقيع", "التوقيع"], true);
  sheet_(SHEET_MYSIG, ["المعلم", "التوقيع", "آخر تحديث"], true);
  sheet_(SHEET_VIS, ["المعرف", "المعلم", "الفصل", "القسم", "التاريخ", "الحصة", "نوع الزيارة", "الموضوع", "رقم الزيارة", "التقديرات",
    "الإيجابيات", "الملاحظات والتوصيات", "توقيع رئيس الشعبة", "تاريخ الإنشاء", "توقيع المعلم", "وقت توقيع المعلم"], true);
  sheet_(SHEET_MAIL, ["الوقت", "إلى", "الموضوع", "النتيجة"], true);
  sheet_(SHEET_ABS, ["المعرف", "التاريخ", "المعلم الغائب", "ملاحظة", "تاريخ التسجيل", "آخر تعديل"], true);
  covSheet_();
  swpSheet_();
  sheet_(SHEET_STU, ["الصف", "م", "(بلا أسماء)", "ملاحظة"], true);
  gradesSheet_();
  anonymizeStudents_();
  srepSheet_();
  profSheet_();
  permSheet_();
  tskSheet_(); tskStSheet_(); alrSheet_(); alrStSheet_();
  ensureDemoMinute_();
  folder_();
  const first = ss.getSheets()[0];
  if (first.getName() !== SHEET_CODES && first.getLastRow() === 0 && ss.getSheets().length > 3) ss.deleteSheet(first);
}

function sheet_(name, headers, asText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    if (asText) sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat("@");
    sh.appendRow(headers);
    sh.setRightToLeft(true);
    sh.setFrozenRows(1);
  }
  return sh;
}

/** مجلد Drive لصور الإنجازات والملفات المحفوظة */
function folder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty("FOLDER_ID");
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) {} }
  const f = DriveApp.createFolder("موقع قسم المواد الشرعية - الملفات");
  props.setProperty("FOLDER_ID", f.getId());
  return f;
}

function newCode_() {
  const chars = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789";
  let s = "";
  for (let i = 0; i < 8; i++) s += chars[Math.floor(Math.random() * chars.length)];
  return s;
}
function newId_() { return Utilities.getUuid().replace(/-/g, "").slice(0, 10); }
function nowText_() { return Utilities.formatDate(new Date(), "Asia/Kuwait", "yyyy-MM-dd HH:mm"); }
function sh_(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }
function rows_(name) { return sh_(name).getDataRange().getDisplayValues(); }

/** يكتب صفًا كنص (حتى لا تحوّل الورقة التواريخ تلقائيًا) */
function writeRow_(sh, rowIndex, values) {
  const rg = sh.getRange(rowIndex, 1, 1, values.length);
  rg.setNumberFormat("@");
  rg.setValues([values.map(v => v == null ? "" : String(v))]);
}
function appendText_(sh, values) { writeRow_(sh, sh.getLastRow() + 1, values); }
function findRow_(sh, col, value) {
  const vals = sh.getRange(1, col, Math.max(sh.getLastRow(), 1), 1).getDisplayValues();
  for (let i = 1; i < vals.length; i++) if (vals[i][0] === value) return i + 1;
  return -1;
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/** يتحقق من الرمز ويعيد المعلم، أو يرمي خطأ */
function who_(code) {
  code = String(code || "").trim().toUpperCase();
  if (code.length < 6) throw new Error("الرمز غير صحيح.");
  const rows = sh_(SHEET_CODES).getDataRange().getDisplayValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim().toUpperCase() === code) {
      const name = String(rows[i][0]).trim();
      if (isDemo_(name)) {
        if (String(rows[i][3]).indexOf("موقوف") !== -1) throw new Error("هذا الحساب التجريبي موقوف.");
        return { name, classes: DEMOS[name], isHead: false, isDemo: true };
      }
      const t = TEACHERS.find(x => x[0] === name);
      if (!t) throw new Error("الاسم في ورقة الرموز لا يطابق قائمة المعلمين.");
      return { name, classes: t[1], isHead: String(rows[i][2]).trim() === "نعم", isDemo: false };
    }
  }
  Utilities.sleep(800); // يبطّئ محاولات التخمين
  throw new Error("الرمز غير صحيح.");
}
function head_(code) {
  const me = who_(code);
  if (!me.isHead) throw new Error("هذه الصلاحية لرئيس الشعبة فقط.");
  return me;
}

// ======================= ما قطع من المنهج =======================
// مفتاح الإدخال: "الصف-الشعبة__القسم" مثل 11-1__hd

/** prefix فارغ = بيانات القسم الرسمية فقط، وإلا بيانات حساب تجريبي واحد */
function readEntries_(prefix) {
  prefix = prefix || "";
  const rows = rows_(SHEET_DATA);
  const out = {};
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const k = String(r[0] || "");
    if (!k) continue;
    if (prefix ? k.indexOf(prefix) !== 0 : k.indexOf("DEMO|") === 0) continue;
    out[k.slice(prefix.length)] = { done: Number(r[4]), note: String(r[6] || ""), teacher: String(r[3]), updatedAt: String(r[7] || "") };
  }
  return out;
}

function login(code) {
  const me = who_(code);
  const all = readEntries_(demoPrefix_(me));
  const mine = {};
  me.classes.forEach(a => { const p = a.split(":"), k = p[1] + "__" + p[0]; if (all[k]) mine[k] = all[k]; });
  return { name: me.name, classes: me.classes, isHead: me.isHead, isDemo: me.isDemo, entries: mine, all: me.isHead ? all : null, phone: phones_()[me.name] || "" };
}

function saveEntry(code, classId, subject, done, note, label) {
  const me = who_(code);
  if (!MAX_DONE[subject]) throw new Error("القسم غير معروف.");
  if (me.classes.indexOf(subject + ":" + classId) === -1) throw new Error("هذا الفصل ليس من فصولك.");
  done = Math.floor(Number(done));
  if (!(done >= 0 && done <= MAX_DONE[subject])) throw new Error("قيمة غير صحيحة.");
  note = String(note || "").slice(0, 300);
  label = String(label || "").slice(0, 200);
  const key = demoPrefix_(me) + classId + "__" + subject;
  return withLock_(() => {
    const sh = sh_(SHEET_DATA);
    const now = nowText_();
    const row = [key, classId, SUBJ[subject], me.name, done, label, note, now];
    const idx = findRow_(sh, 1, key);
    if (idx > 0) writeRow_(sh, idx, row); else appendText_(sh, row);
    if (!me.isDemo) appendText_(sh_(SHEET_LOG), [now, me.name, classId, SUBJ[subject], done, note]);
    return { updatedAt: now };
  });
}

function headData(code) { head_(code); return readEntries_(); }

// ======================= الإنجازات =======================
// أعمدة ورقة الإنجازات:
// 0 المعرف | 1 المعلم | 2 التاريخ | 3 النوع | 4 العنوان | 5 الفئة المستهدفة | 6 وصف المعلم
// 7 الصور (معرفات Drive مفصولة بفاصلة) | 8 الحالة | 9 التقرير | 10 تاريخ الرفع | 11 تاريخ الصياغة
// 12 السجل (pd = التنمية المهنية وإنجاز الموظف، act = الأنشطة والمسابقات) | 13 النتائج والفائزون | 14 القسم (hd/qr)
const ACH_KINDS = { pd: "التنمية المهنية وإنجاز الموظف", act: "الأنشطة والمسابقات" };

function achObj_(r, withReport) {
  const o = { id: r[0], teacher: r[1], date: r[2], type: r[3], title: r[4], target: r[5], desc: r[6],
    photos: r[7] ? r[7].split(",").filter(Boolean) : [], status: r[8] || ST_PENDING, createdAt: r[10], reportedAt: r[11],
    kind: r[12] === "act" ? "act" : "pd", results: r[13] || "", dept: r[14] === "qr" || r[14] === "hd" ? r[14] : achDept_(r[1]), editedAt: r[15] || "", teacherOkAt: r[16] || "" };
  if (withReport) o.report = r[9];
  return o;
}

/** قسم الإنجاز إن لم يُحدَّد: من مادة المعلم (إن كان يدرّس مادة واحدة)، وإلا الحديث الشريف */
function achDept_(name) { const cl = classesOf_(name) || []; const d = cl.map(c => c.split(":")[0]); return d.length && d.every(x => x === "qr") ? "qr" : "hd"; }
function submitAchievement(code, a) {
  const me = who_(code);
  const title = String(a.title || "").trim().slice(0, 150);
  const desc = String(a.desc || "").trim().slice(0, 6000);
  if (!title) throw new Error("اكتب عنوان النشاط.");
  if (desc.length < 10) throw new Error("اكتب وصفًا للنشاط ولو بسطرين.");
  const photos = (a.photos || []).slice(0, 4);
  const folder = folder_();
  const id = newId_();
  const blobs = [];
  const ids = photos.map((b64, i) => {
    const blob = Utilities.newBlob(Utilities.base64Decode(String(b64).replace(/^data:image\/\w+;base64,/, "")), "image/jpeg", `${id}-${i + 1}.jpg`);
    blobs.push(blob);
    return folder.createFile(blob).getId();
  });
  const kind = a.kind === "act" ? "act" : "pd", results = kind === "act" ? String(a.results || "").trim().slice(0, 1500) : "";
  const dept = a.dept === "qr" || a.dept === "hd" ? a.dept : achDept_(me.name);
  const row = [id, me.name, String(a.date || "").slice(0, 10), String(a.type || "").slice(0, 60), title,
    String(a.target || "").slice(0, 150), desc, ids.join(","), ST_PENDING, "", nowText_(), "", kind, results, dept];
  withLock_(() => appendText_(sh_(SHEET_ACH), row));
  notify_(me, `${kind === "act" ? "نشاط/مسابقة" : "تنمية مهنية"} من أ. ${me.name}: ${title}`, `
    <p>رفع <b>أ. ${esc_(me.name)}</b> ${kind === "act" ? "نشاطًا" : "إنجازًا"} جديدًا في سجل «${ACH_KINDS[kind]}»، وهو بانتظار الصياغة.</p>
    <table cellpadding="6" style="border-collapse:collapse;border:1px solid #d6e2ec">
      <tr><td style="background:#e8f1f8"><b>العنوان</b></td><td>${esc_(title)}</td></tr>
      <tr><td style="background:#e8f1f8"><b>القسم</b></td><td>${esc_(SUBJ[dept])}</td></tr>
      <tr><td style="background:#e8f1f8"><b>النوع</b></td><td>${esc_(row[3])}</td></tr>
      <tr><td style="background:#e8f1f8"><b>التاريخ</b></td><td>${esc_(row[2])}</td></tr>
      <tr><td style="background:#e8f1f8"><b>الفئة المستهدفة</b></td><td>${esc_(row[5] || "—")}</td></tr>
    </table>
    <p><b>وصف المعلم:</b><br>${esc_(desc).replace(/\n/g, "<br>")}</p>${results ? `<p><b>النتائج والفائزون:</b><br>${esc_(results).replace(/\n/g, "<br>")}</p>` : ""}
    <p>عدد الصور: ${ids.length} (مرفقة بهذه الرسالة).</p>`, blobs);
  return { id };
}

/** إنجازات المعلم نفسه — بدون نص التقرير */
function myAchievements(code) {
  const me = who_(code);
  // التقرير المصاغ يظهر لصاحبه فور صياغته ليراجعه ويعتمده بنفسه، ثم يطبعه PDF
  return rows_(SHEET_ACH).slice(1).filter(r => r[0] && r[1] === me.name).map(r => { const o = achObj_(r, false); if (r[9]) o.report = r[9]; return o; }).reverse();
}

/** تعديل المعلم لما رفعه (النص والصور): يعود «بانتظار الاعتماد» ويصل رئيسَ الشعبة إشعار */
function editAchievement(code, id, a) {
  const me = who_(code);
  const title = String(a.title || "").trim().slice(0, 150), desc = String(a.desc || "").trim().slice(0, 6000);
  if (!title) throw new Error("اكتب العنوان.");
  if (desc.length < 10) throw new Error("اكتب وصفًا ولو بسطرين.");
  const res = withLock_(() => {
    const sh = sh_(SHEET_ACH), row = findRow_(sh, 1, String(id));
    if (row < 0) throw new Error("غير موجود.");
    const r = sh.getRange(row, 1, 1, 15).getDisplayValues()[0];
    if (r[1] !== me.name && !me.isHead) throw new Error("لا يمكنك تعديل هذا.");
    const old = (r[7] || "").split(",").filter(Boolean);
    const keep = (a.keep || []).map(String).filter(x => old.indexOf(x) !== -1);
    old.filter(x => keep.indexOf(x) === -1).forEach(fid => { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} });
    const folder = folder_(), stamp = Date.now();
    const added = (a.photos || []).slice(0, Math.max(0, 4 - keep.length)).map((b64, i) => folder.createFile(Utilities.newBlob(
      Utilities.base64Decode(String(b64).replace(/^data:image\/\w+;base64,/, "")), "image/jpeg", `${r[0]}-e${stamp}-${i + 1}.jpg`)).getId());
    const ids = keep.concat(added);
    if (!ids.length) throw new Error("أبقِ صورة واحدة على الأقل.");
    const kind = r[12] === "act" ? "act" : "pd";
    r[2] = String(a.date || r[2]).slice(0, 10); r[3] = String(a.type || r[3]).slice(0, 60); r[4] = title; r[5] = String(a.target || "").slice(0, 150); r[6] = desc;
    r[7] = ids.join(","); r[13] = kind === "act" ? String(a.results || "").trim().slice(0, 1500) : "";
    if (a.dept === "hd" || a.dept === "qr") r[14] = a.dept;
    writeRow_(sh, row, r);
    if (!me.isHead) { sh.getRange(row, 16).setNumberFormat("@").setValue(nowText_()); sh.getRange(row, 17).setValue(""); } // التعديل يحتاج اعتمادًا جديدًا للتقرير
    return { kind, photos: ids, status: r[8] };
  });
  if (!me.isHead) notify_(me, `تعديل ${res.kind === "act" ? "نشاط" : "إنجاز"} من أ. ${me.name}: ${title}`,
    `<p>عدّل <b>أ. ${esc_(me.name)}</b> ${res.kind === "act" ? "نشاطه" : "إنجازه"} «${esc_(title)}» في سجل «${ACH_KINDS[res.kind]}».</p>`);
  return res;
}

/** المعلم يراجع التقرير المصاغ (ويعدّله إن أراد) ثم يعتمده بنفسه، فيصل رئيسَ الشعبة إشعار */
function approveMyReport(code, id, text) {
  const me = who_(code);
  text = String(text || "").trim().slice(0, 20000);
  if (text.length < 30) throw new Error("نص التقرير قصير جدًا.");
  return withLock_(() => {
    const sh = sh_(SHEET_ACH), row = findRow_(sh, 1, String(id));
    if (row < 0) throw new Error("غير موجود.");
    const r = sh.getRange(row, 1, 1, 10).getDisplayValues()[0];
    if (r[1] !== me.name) throw new Error("هذا ليس تقريرك.");
    if (!r[9]) throw new Error("لم يُصغ التقرير بعد.");
    if (!sh.getRange(1, 17).getDisplayValue()) sh.getRange(1, 17).setValue("اعتماد المعلم");
    const at = nowText_();
    sh.getRange(row, 10).setValue(text); sh.getRange(row, 9).setValue(ST_READY);
    sh.getRange(row, 17).setNumberFormat("@").setValue(at);
    return { teacherOkAt: at, report: text, kind: sh.getRange(row, 13).getDisplayValue() === "act" ? "act" : "pd", title: r[4] };
  });
}

function deleteAchievement(code, id) {
  const me = who_(code);
  return withLock_(() => {
    const sh = sh_(SHEET_ACH), row = findRow_(sh, 1, id);
    if (row < 0) throw new Error("النشاط غير موجود.");
    const r = sh.getRange(row, 1, 1, 12).getDisplayValues()[0];
    if (!me.isHead && r[1] !== me.name) throw new Error("لا يمكن حذف هذا النشاط.");
    (r[7] || "").split(",").filter(Boolean).forEach(fid => { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} });
    sh.deleteRow(row);
    return true;
  });
}

/** لرئيس الشعبة: كل الإنجازات مع التقارير */
function allAchievements(code) {
  head_(code);
  // تظهر رفوعات الحسابات التجريبية لرئيس الشعبة بعلامة «تجريبي»، ولا تدخل في السجلات المطبوعة
  return rows_(SHEET_ACH).slice(1).filter(r => r[0]).map(r => Object.assign(achObj_(r, true), { demo: isDemo_(r[1]) })).reverse();
}

function saveReport(code, id, text) {
  head_(code);
  text = String(text || "").slice(0, 20000);
  return withLock_(() => {
    const sh = sh_(SHEET_ACH), row = findRow_(sh, 1, id);
    if (row < 0) throw new Error("النشاط غير موجود.");
    const r = sh.getRange(row, 1, 1, 12).getDisplayValues()[0];
    r[9] = text; r[8] = text.trim() ? ST_READY : ST_PENDING; r[11] = text.trim() ? nowText_() : "";
    writeRow_(sh, row, r);
    return { status: r[8], reportedAt: r[11] };
  });
}

/** يعيد صورة إنجاز: لصاحب الإنجاز أو رئيس الشعبة فقط */
function getPhoto(code, fileId) {
  const me = who_(code);
  const r = rows_(SHEET_ACH).slice(1).find(x => (x[7] || "").split(",").indexOf(fileId) !== -1);
  if (!r || (!me.isHead && r[1] !== me.name)) throw new Error("لا تملك صلاحية عرض هذه الصورة.");
  const blob = DriveApp.getFileById(fileId).getBlob();
  return "data:image/jpeg;base64," + Utilities.base64Encode(blob.getBytes());
}

// ======================= محاضر الاجتماعات =======================
// أعمدة المحاضر: 0 المعرف | 1 العنوان | 2 التاريخ | 3 الوقت | 4 المكان | 5 الحضور (JSON)
// 6 جدول الأعمال | 7 ما دار في الاجتماع | 8 التوصيات | 9 تاريخ الإنشاء

function minObj_(r) {
  let att = [];
  try { att = JSON.parse(r[5] || "[]"); } catch (e) {}
  return { id: r[0], title: r[1], date: r[2], time: r[3], place: r[4], attendees: att, agenda: r[6], discussion: r[7], decisions: r[8], createdAt: r[9] };
}
function sigRows_() { return rows_(SHEET_SIG).slice(1).filter(r => r[0]); }

function listMinutes(code) {
  const me = who_(code);
  const sigs = sigRows_();
  return rows_(SHEET_MIN).slice(1).filter(r => r[0] && (me.isDemo ? r[0] === DEMO_MINUTE : r[0] !== DEMO_MINUTE)).map(r => {
    const m = minObj_(r);
    const mine = sigs.find(s => s[0] === m.id && s[1] === me.name);
    return { id: m.id, title: m.title, date: m.date, attendees: m.attendees,
      signedCount: sigs.filter(s => s[0] === m.id).length, mySignedAt: mine ? mine[2] : "" };
  }).reverse();
}

/** المحضر كاملًا. رئيس الشعبة يرى صور كل التواقيع، والمعلم يرى توقيعه فقط */
function getMinute(code, id) {
  const me = who_(code);
  const r = rows_(SHEET_MIN).slice(1).find(x => x[0] === id);
  if (!r || (me.isDemo !== (id === DEMO_MINUTE))) throw new Error("المحضر غير موجود.");
  const m = minObj_(r);
  m.signatures = sigRows_().filter(s => s[0] === id).map(s => ({ teacher: s[1], signedAt: s[2], image: (me.isHead || s[1] === me.name) ? s[3] : "" }));
  return m;
}

function saveMinute(code, m) {
  head_(code);
  const title = String(m.title || "").trim().slice(0, 150);
  if (!title) throw new Error("اكتب عنوان الاجتماع.");
  const names = TEACHERS.map(t => t[0]);
  const att = (m.attendees || []).filter(n => names.indexOf(n) !== -1);
  if (!att.length) throw new Error("اختر الحضور.");
  return withLock_(() => {
    const sh = sh_(SHEET_MIN);
    const id = m.id || newId_();
    const row = m.id ? findRow_(sh, 1, m.id) : -1;
    if (m.id && row < 0) throw new Error("المحضر غير موجود.");
    const created = row > 0 ? sh.getRange(row, 10).getDisplayValue() : nowText_();
    const values = [id, title, String(m.date || "").slice(0, 10), String(m.time || "").slice(0, 20), String(m.place || "").slice(0, 100),
      JSON.stringify(att), String(m.agenda || "").slice(0, 5000), String(m.discussion || "").slice(0, 15000), String(m.decisions || "").slice(0, 8000), created];
    if (row > 0) writeRow_(sh, row, values); else appendText_(sh, values);
    return { id };
  });
}

function deleteMinute(code, id) {
  head_(code);
  if (id === DEMO_MINUTE) throw new Error("لا يمكن حذف المحضر التجريبي.");
  return withLock_(() => {
    const sh = sh_(SHEET_MIN), row = findRow_(sh, 1, id);
    if (row > 0) sh.deleteRow(row);
    const sig = sh_(SHEET_SIG), vals = sig.getDataRange().getDisplayValues();
    for (let i = vals.length - 1; i >= 1; i--) if (vals[i][0] === id) sig.deleteRow(i + 1);
    return true;
  });
}

function signMinute(code, id, image) {
  const me = who_(code);
  image = String(image || ""); // التوقيع الإلكتروني أُلغي (قرار رئيس الشعبة ٢٥/١٠/٢٠٢٦): يكفي تأكيد الاطلاع، والتوقيع يدوي على الورق
  if (image && !validSig_(image)) throw new Error("التوقيع غير صالح، ارسمه مرة أخرى.");
  const r = rows_(SHEET_MIN).slice(1).find(x => x[0] === id);
  if (!r || (me.isDemo !== (id === DEMO_MINUTE))) throw new Error("المحضر غير موجود.");
  if (minObj_(r).attendees.indexOf(me.name) === -1) throw new Error("اسمك ليس ضمن حضور هذا الاجتماع.");
  return withLock_(() => {
    const sh = sh_(SHEET_SIG), vals = sh.getDataRange().getDisplayValues();
    for (let i = 1; i < vals.length; i++) if (vals[i][0] === id && vals[i][1] === me.name) throw new Error("أكّدت اطلاعك على هذا المحضر من قبل.");
    const at = nowText_();
    appendText_(sh, [id, me.name, at, image]);
    if (image) saveMySig_(me.name, image, at);
    // عند اكتمال توقيع جميع الحضور يصل إيميل لرئيس الشعبة
    const m = minObj_(r), signed = vals.slice(1).filter(v => v[0] === id).map(v => [v[1], v[2]]).concat([[me.name, at]]);
    if (m.attendees.every(n => signed.some(x => x[0] === n))) {
      notify_(me, `اكتمل اطلاع الحضور على محضر: ${m.title}`, `
        <p>اطّلع جميع الحضور (${m.attendees.length}) على محضر <b>${esc_(m.title)}</b> بتاريخ ${esc_(m.date)}.</p>
        <ul>${signed.map(x => `<li>${esc_(x[0])} — ${esc_(x[1])}</li>`).join("")}</ul>
        <p>يمكنك الآن تنزيل المحضر وورد بالتواقيع من ركن «محاضر الاجتماعات».</p>`);
    }
    return { signedAt: at };
  });
}

function saveMySig_(name, image, at) {
  const my = sh_(SHEET_MYSIG), row = findRow_(my, 1, name);
  if (row > 0) writeRow_(my, row, [name, image, at]); else appendText_(my, [name, image, at]);
}

function mySignature(code) {
  const me = who_(code);
  const r = rows_(SHEET_MYSIG).slice(1).find(x => x[0] === me.name);
  return r ? r[1] : "";
}

// ======================= حفظ الملفات في Drive =======================

function saveFileToDrive(code, name, b64) {
  head_(code);
  name = String(name || "ملف.docx").replace(/[\\/:*?"<>|]/g, "-").slice(0, 150);
  const blob = Utilities.newBlob(Utilities.base64Decode(String(b64).replace(/^data:[^,]+,/, "")),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document", name);
  const f = folder_().createFile(blob);
  return f.getUrl();
}

// ======================= الحسابات التجريبية =======================

function ensureDemoMinute_() {
  const sh = sh_(SHEET_MIN);
  if (findRow_(sh, 1, DEMO_MINUTE) > 0) return;
  appendText_(sh, [DEMO_MINUTE, "محضر تجريبي للعرض", Utilities.formatDate(new Date(), "Asia/Kuwait", "yyyy-MM-dd"), "الحصة الخامسة", "غرفة قسم المواد الشرعية",
    JSON.stringify(Object.keys(DEMOS)), "مثال على بند من جدول الأعمال\nمثال على بند ثانٍ",
    "هذا محضر تجريبي يوضح طريقة الاطلاع على المحاضر والتوقيع عليها إلكترونيًا.", "مثال على توصية\nمثال على توصية ثانية", nowText_()]);
}

/** لرئيس الشعبة: حالة الحسابات التجريبية */
function demoAccounts(code) {
  head_(code);
  const rows = sh_(SHEET_CODES).getDataRange().getDisplayValues();
  const entries = rows_(SHEET_DATA).map(r => r[0]);
  const ach = rows_(SHEET_ACH).slice(1), sigs = sigRows_(), vis = visRows_();
  return Object.keys(DEMOS).map(n => {
    const r = rows.find(x => x[0] === n) || [];
    return { name: n, code: r[1] || "", active: String(r[3] || "").indexOf("موقوف") === -1,
      entries: entries.filter(k => k.indexOf("DEMO|" + n + "|") === 0).length,
      achievements: ach.filter(a => a[1] === n).length, signatures: sigs.filter(x => x[1] === n).length,
      visits: vis.filter(v => v[1] === n).length };
  });
}

/** action: toggle (إيقاف/تفعيل) · wipe (مسح بياناته) · newcode (رمز جديد) */
function demoAction(code, name, action) {
  head_(code);
  if (!isDemo_(name)) throw new Error("ليس حسابًا تجريبيًا.");
  return withLock_(() => {
    const codes = sh_(SHEET_CODES), row = findRow_(codes, 1, name);
    if (row < 0) throw new Error("الحساب غير موجود. شغّل setup.");
    if (action === "toggle") {
      const note = codes.getRange(row, 4).getDisplayValue();
      codes.getRange(row, 4).setValue(note.indexOf("موقوف") !== -1 ? "تجريبي" : "تجريبي - موقوف");
    } else if (action === "newcode") {
      codes.getRange(row, 2).setNumberFormat("@").setValue(newCode_());
    } else if (action === "wipe") {
      const data = sh_(SHEET_DATA), dv = data.getDataRange().getDisplayValues();
      for (let i = dv.length - 1; i >= 1; i--) if (dv[i][0].indexOf("DEMO|" + name + "|") === 0) data.deleteRow(i + 1);
      const ach = sh_(SHEET_ACH), av = ach.getDataRange().getDisplayValues();
      for (let i = av.length - 1; i >= 1; i--) if (av[i][1] === name) {
        (av[i][7] || "").split(",").filter(Boolean).forEach(fid => { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} });
        ach.deleteRow(i + 1);
      }
      const grd = sh_(SHEET_GRD);
      [SHEET_PERM, SHEET_PROF].forEach((sn, j) => { const sh = sh_(sn); if (!sh) return; const v = sh.getDataRange().getDisplayValues(); for (let i = v.length - 1; i >= 1; i--) if (v[i][j === 0 ? 2 : 0] === name) sh.deleteRow(i + 1); });
      const srp = sh_(SHEET_SREP);
      if (srp) { const sv = srp.getDataRange().getDisplayValues(); for (let i = sv.length - 1; i >= 1; i--) if (sv[i][2] === name) srp.deleteRow(i + 1); }
      if (grd) { const gv = grd.getDataRange().getDisplayValues(); for (let i = gv.length - 1; i >= 1; i--) if (gv[i][0].indexOf("DEMO|" + name + "|") === 0) grd.deleteRow(i + 1); }
      const vis = sh_(SHEET_VIS), vv = vis.getDataRange().getDisplayValues();
      for (let i = vv.length - 1; i >= 1; i--) if (vv[i][1] === name) vis.deleteRow(i + 1);
      [SHEET_SIG, SHEET_MYSIG].forEach((sn, j) => {
        const sh = sh_(sn), v = sh.getDataRange().getDisplayValues();
        for (let i = v.length - 1; i >= 1; i--) if (v[i][j === 0 ? 1 : 0] === name) sh.deleteRow(i + 1);
      });
    } else throw new Error("إجراء غير معروف.");
    return true;
  });
}

// ======================= التنبيهات بالإيميل =======================

function esc_(t) { return String(t == null ? "" : t).replace(/[&<>"]/g, c => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c])); }

/** يرسل تنبيهًا لرئيس الشعبة. الحسابات التجريبية تصل بعلامة [تجريبي].
 * لا يوقف العملية إن فشل الإرسال، لكن يسجّل كل محاولة ونتيجتها في ورقة «سجل التنبيهات» */
function notify_(me, subject, html, attachments) {
  const demo = me && me.isDemo;
  subject = (demo ? "[تجريبي] " : "") + subject;
  let result = "أُرسل";
  try {
    MailApp.sendEmail({
      to: HEAD_EMAIL,
      subject: subject,
      name: SITE_NAME,
      htmlBody: `<div dir="rtl" style="font-family:Tahoma,Arial,sans-serif;font-size:14px;line-height:1.8;color:#2b2b2b">
        ${demo ? '<p style="background:#fff4d6;padding:6px 10px;border-radius:6px">هذا تنبيه من <b>حساب تجريبي</b>، ولا يظهر في التقارير الرسمية.</p>' : ""}
        ${html}
        <p><a href="${SITE_URL}" style="color:#1f6fa8">فتح موقع القسم</a></p></div>`,
      attachments: attachments || []
    });
  } catch (e) { result = "فشل: " + (e && e.message || e); console.error("notify failed: " + e); }
  try { const sh = sh_(SHEET_MAIL) || sheet_(SHEET_MAIL, ["الوقت", "إلى", "الموضوع", "النتيجة"], true); appendText_(sh, [nowText_(), HEAD_EMAIL, subject, result]); } catch (e) {}
  return result;
}

/** لتجربة الإيميل من المحرر */
function testEmail() { Logger.log(notify_(null, "تجربة تنبيهات موقع القسم", "<p>وصلتك هذه الرسالة، يعني التنبيهات تعمل.</p>") + " — المتبقي اليوم: " + MailApp.getRemainingDailyQuota()); }

// ======================= الزيارات الصفية =======================
// نموذج «تقرير زيارة رئيس الشعبة» — التوجيه الفني للعلوم الشرعية (١٧ بندًا، نفس النموذج للقسمين)
// أعمدة الزيارات: 0 المعرف | 1 المعلم | 2 الفصل (hd:11-1) | 3 القسم | 4 التاريخ | 5 الحصة | 6 نوع الزيارة | 7 الموضوع
// 8 رقم الزيارة | 9 التقديرات (JSON: [[التقدير، الملاحظة], ...]) | 10 الإيجابيات | 11 الملاحظات والتوصيات
// 12 توقيع رئيس الشعبة | 13 تاريخ الإنشاء | 14 توقيع المعلم | 15 وقت توقيع المعلم
const VIS_COLS = 16;
const VIS_ITEMS = 17;
const VIS_LEVELS = ["ممتاز", "جيد جدًا", "جيد", "ضعيف"];

function visObj_(r, withImages) {
  let ratings = [];
  try { ratings = JSON.parse(r[9] || "[]"); } catch (e) {}
  const o = { id: r[0], teacher: r[1], classId: r[2], dept: r[3], date: r[4], period: r[5], visitType: r[6], topic: r[7],
    number: r[8], ratings, positives: r[10], recommendations: r[11], createdAt: r[13], teacherSignedAt: r[15], demo: isDemo_(r[1]) };
  if (withImages) { o.headSig = r[12]; o.teacherSig = r[14]; }
  return o;
}
function visRows_() { return rows_(SHEET_VIS).slice(1).filter(r => r[0]); }
function validSig_(img) { return /^data:image\/png;base64,/.test(img) && img.length <= 48000; }

/** لرئيس الشعبة: إنشاء زيارة أو تعديلها (قبل توقيع المعلم) */
function saveVisit(code, v) {
  const me = head_(code);
  const teacher = String(v.teacher || "");
  const classes = classesOf_(teacher);
  if (!classes) throw new Error("اختر المعلم.");
  if (classes.indexOf(v.classId) === -1) throw new Error("هذا الفصل ليس من فصول المعلم.");
  const dept = String(v.classId).split(":")[0];
  const ratings = (Array.isArray(v.ratings) ? v.ratings : []).slice(0, VIS_ITEMS).map(x => {
    const lv = String((x && x[0]) || "");
    return [VIS_LEVELS.indexOf(lv) !== -1 ? lv : "", String((x && x[1]) || "").slice(0, 300)];
  });
  const headSig = String(v.headSig || "");
  if (headSig && !validSig_(headSig)) throw new Error("توقيعك غير صالح، ارسمه مرة أخرى.");
  return withLock_(() => {
    const sh = sh_(SHEET_VIS);
    const row = v.id ? findRow_(sh, 1, v.id) : -1;
    let old = null;
    if (v.id) {
      if (row < 0) throw new Error("الزيارة غير موجودة.");
      old = sh.getRange(row, 1, 1, VIS_COLS).getDisplayValues()[0];
      if (old[15]) throw new Error("اطّلع المعلم على هذه الزيارة، فلا يمكن تعديلها.");
    }
    const number = old && old[1] === teacher ? old[8] : String(visRows_().filter(r => r[1] === teacher).length + 1);
    const values = [old ? old[0] : newId_(), teacher, v.classId, SUBJ[dept] || "", String(v.date || "").slice(0, 10), String(v.period || "").slice(0, 10),
      String(v.visitType || "").slice(0, 60), String(v.topic || "").slice(0, 200), number, JSON.stringify(ratings),
      String(v.positives || "").slice(0, 3000), String(v.recommendations || "").slice(0, 3000),
      headSig || (old ? old[12] : ""), old ? old[13] : nowText_(), "", ""];
    if (row > 0) writeRow_(sh, row, values); else appendText_(sh, values);
    if (headSig) saveMySig_(me.name, headSig, nowText_());
    return { id: values[0], number };
  });
}

/** رئيس الشعبة يرى كل الزيارات، والمعلم يرى زياراته فقط */
function listVisits(code) {
  const me = who_(code);
  return visRows_().filter(r => me.isHead || r[1] === me.name).map(r => visObj_(r, false)).reverse();
}

function getVisit(code, id) {
  const me = who_(code);
  const r = visRows_().find(x => x[0] === id);
  if (!r || (!me.isHead && r[1] !== me.name)) throw new Error("الزيارة غير موجودة.");
  return visObj_(r, true);
}

/** توقيع المعلم على أنه اطّلع على زيارته */
function signVisit(code, id, image) {
  const me = who_(code);
  image = String(image || "");
  if (image && !validSig_(image)) throw new Error("التوقيع غير صالح، ارسمه مرة أخرى.");
  return withLock_(() => {
    const sh = sh_(SHEET_VIS), row = findRow_(sh, 1, id);
    if (row < 0) throw new Error("الزيارة غير موجودة.");
    const r = sh.getRange(row, 1, 1, VIS_COLS).getDisplayValues()[0];
    if (r[1] !== me.name) throw new Error("هذه الزيارة ليست لك.");
    if (r[15]) throw new Error("أكّدت اطلاعك على هذه الزيارة من قبل.");
    const at = nowText_();
    r[14] = image; r[15] = at;
    writeRow_(sh, row, r);
    if (image) saveMySig_(me.name, image, at);
    notify_(me, `اطّلع أ. ${me.name} على الزيارة الصفية رقم ${r[8]}`, `
      <p>اطّلع <b>أ. ${esc_(me.name)}</b> على الزيارة الصفية رقم ${esc_(r[8])} (${esc_(r[3])}) بتاريخ ${esc_(r[4])} في ${esc_(at)}.</p>
      <p>يمكنك تنزيل تقرير الزيارة من ركن «الزيارات الصفية».</p>`);
    return { signedAt: at };
  });
}

function deleteVisit(code, id) {
  head_(code);
  return withLock_(() => {
    const sh = sh_(SHEET_VIS), row = findRow_(sh, 1, id);
    if (row > 0) sh.deleteRow(row);
    return true;
  });
}

// ======================= الاحتياط (لرئيس الشعبة فقط) =======================
// الجدول العام للقسم: لكل معلم [اليوم (0 الأحد … 4 الخميس)، الحصة، الفصل]
// الفصل "R" = حصة احتياط في الجدول، "RF" = احتياط فنية: حصة للمعلم تحتاج تغطية عند غيابه، وهو فيها متاح للتغطية عن غيره
// يجب أن يطابق TIMETABLE في index.html
const TIMETABLE = {
  "خليفة حمد خليفة القصار": [[3,2,"hd:11-1"],[4,5,"hd:11-1"]],
  "أحمد صادق أحمد أبوسعدة": [[0,1,"hd:11-4"],[0,2,"hd:11-5"],[1,3,"hd:11-3"],[1,6,"hd:11-2"],[2,2,"hd:11-3"],[2,5,"RF"],[2,6,"hd:11-5"],[3,4,"hd:11-2"],[3,6,"R"],[4,4,"RF"],[4,7,"hd:11-4"]],
  "أحمد صلاح زيدان القناوي": [[0,1,"hd:9-1"],[1,2,"hd:9-5"],[1,5,"hd:9-1"],[2,6,"R"],[3,1,"hd:9-2"],[4,4,"hd:9-2"],[4,5,"hd:9-5"]],
  "عماد صادق علي العايق": [[0,5,"hd:12-1"],[1,4,"hd:12-3"],[1,7,"hd:12-1"],[2,3,"hd:10-5"],[2,6,"hd:12-3"],[3,7,"hd:10-3"],[4,1,"hd:10-5"],[4,2,"hd:10-3"]],
  "كرم عبدالحميد أبوزيد أحمد": [[0,2,"hd:12-4"],[0,4,"hd:12-5"],[1,2,"hd:10-1"],[2,5,"hd:12-4"],[2,7,"hd:10-2"],[3,1,"hd:10-1"],[3,4,"hd:10-4"],[3,7,"hd:12-5"],[4,3,"hd:10-2"],[4,6,"hd:10-4"]],
  "منذر خالد مطلق الحربي": [[0,1,"qr:12-2"],[0,3,"qr:12-3"],[0,4,"hd:12-2"],[0,6,"qr:12-1"],[1,2,"qr:12-4"],[1,4,"hd:9-4"],[1,7,"qr:12-5"],[2,2,"qr:12-5"],[2,5,"qr:12-2"],[2,7,"hd:9-3"],[3,3,"hd:12-2"],[3,5,"hd:9-4"],[3,6,"qr:12-4"],[4,2,"qr:12-3"],[4,4,"qr:12-1"],[4,6,"hd:9-3"]],
  "معاذ وليد مبارك الشطي": [[0,4,"qr:9-2"],[0,7,"qr:9-5"],[1,3,"qr:9-4"],[1,6,"qr:9-3"],[2,3,"qr:9-1"],[2,7,"qr:9-2"],[3,4,"qr:9-3"],[3,6,"qr:9-1"],[4,1,"qr:9-4"],[4,2,"qr:9-5"]],
  "يوسف عدنان عبدالكريم الشطي": [[0,4,"qr:10-2"],[0,6,"qr:10-3"],[1,3,"qr:10-3"],[1,5,"qr:10-5"],[1,6,"R"],[2,5,"qr:10-1"],[2,7,"qr:10-4"],[3,4,"qr:10-5"],[3,7,"qr:10-2"],[4,1,"qr:10-1"],[4,3,"qr:10-4"]],
  "عبدالله أحمد عبدالمجد الماجد": [[0,2,"qr:11-2"],[0,4,"qr:11-1"],[1,5,"qr:11-3"],[2,2,"qr:11-5"],[2,3,"qr:11-1"],[2,6,"RF"],[3,1,"qr:11-3"],[3,3,"qr:11-4"],[3,6,"qr:11-5"],[4,5,"qr:11-4"],[4,6,"qr:11-2"]]
};
const COV_DONE = "مغطاة", COV_NONE = "غير مغطاة", COV_PRESENT = "حضر المعلم";
// التبديل: معلم المادة الأخرى في الفصل نفسه يدخل حصة الغائب، ثم يردّها الغائب بدخول حصة ذلك المعلم في الفصل نفسه في يوم آخر
const COV_SWAP = "تبديل", RET_PENDING = "لم يُرد", RET_DONE = "تم الرد";
const COV_HEAD = ["المعرف", "معرف الغياب", "التاريخ", "الحصة", "الفصل", "المعلم الغائب", "المعلم البديل", "الحالة", "تاريخ التسجيل",
  "تاريخ الرد", "حصة الرد", "فصل الرد", "حالة الرد"];
function covSheet_() {
  const sh = sheet_(SHEET_COV, COV_HEAD, true), n = COV_HEAD.length;
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
  if (sh.getRange(1, 10).getDisplayValue() !== COV_HEAD[9]) {
    sh.getRange(1, 10, sh.getMaxRows(), n - 9).setNumberFormat("@");
    sh.getRange(1, 1, 1, n).setValues([COV_HEAD]);
  }
  return sh;
}
function isCov_(status) { return status === COV_DONE || status === COV_SWAP; }

function dayOf_(iso) {
  const m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error("التاريخ غير صحيح.");
  return new Date(+m[1], +m[2] - 1, +m[3], 12).getDay(); // 0 الأحد … 6 السبت
}
function lessonAt_(teacher, day, period) {
  return (TIMETABLE[teacher] || []).find(x => x[0] === day && x[1] === period) || null;
}
// ======================= تبديل الحصص (بطلب المعلم من حسابه فقط) =======================
// المعلم يبدّل حصته مع معلم المادة الأخرى في الفصل نفسه: الآخر يدخل حصته، ويردّها هو بدخول حصة الآخر في الفصل نفسه في يوم الرد.
// لا يُنشئ رئيس الشعبة تبديلًا بنفسه (قراره ٢٦/١٠/٢٠٢٦)؛ يُطبع «نموذج تبديل حصة» الإداري ويوقَّع ورقيًا.
const SHEET_SWP = "التبديل";
const SWP_HEAD = ["المعرف", "التاريخ", "الحصة", "الفصل", "المعلم", "المعلم الآخر", "تاريخ الرد", "حصة الرد", "فصل الرد", "الحالة", "حالة الرد", "تاريخ الطلب", "طلبه"];
const SWP_ACTIVE = "معتمد", SWP_CANCEL = "ملغى";
function swpSheet_() { return sheet_(SHEET_SWP, SWP_HEAD, true); }
function swpObj_(r) {
  return { id: r[0], date: r[1], period: Number(r[2]), cls: r[3], teacher: r[4], other: r[5], retDate: r[6], retPeriod: Number(r[7]), retCls: r[8],
    status: r[9], retStatus: r[10], createdAt: r[11], by: r[12] };
}
function swpRows_() { const sh = sh_(SHEET_SWP); return sh ? sh.getDataRange().getDisplayValues().slice(1).filter(r => r[0]).map(swpObj_) : []; }
// أحداث التبديل المعتمدة: في كل حدث يدخل «in» حصة «out» في الفصل cls
function swapEvents_(swaps) {
  const ev = [];
  swaps.filter(x => x.status === SWP_ACTIVE).forEach(x => {
    ev.push({ date: x.date, period: x.period, cls: x.cls, out: x.teacher, in: x.other, id: x.id });
    ev.push({ date: x.retDate, period: x.retPeriod, cls: x.retCls, out: x.other, in: x.teacher, id: x.id });
  });
  return ev;
}
// حصة المعلم الفعلية في تاريخ معين بعد احتساب التبديلات المعتمدة
function effLesson_(t, date, period, ev) {
  const hit = (ev || []).find(e => e.date === date && e.period === period && (e.in === t || e.out === t));
  if (hit) return hit.in === t ? [dayOf_(date), period, hit.cls] : null;
  return lessonAt_(t, dayOf_(date), period);
}
const isLes_ = l => !!l && l[2] !== "R" && l[2] !== "RF";
function todayIso_() { return Utilities.formatDate(new Date(), "Asia/Kuwait", "yyyy-MM-dd"); }

function listSwaps(code) {
  const me = who_(code);
  const all = swpRows_();
  return me.isHead && !me.isDemo ? all.reverse() : all.filter(x => x.teacher === me.name || x.other === me.name).reverse();
}
function requestSwap(code, q) {
  const me = who_(code);
  if (me.isDemo) throw new Error("تبديل الحصص غير متاح في الحساب التجريبي.");
  const teacher = me.name;
  if (!TIMETABLE[teacher]) throw new Error("لا يوجد لك جدول حصص في الموقع.");
  const date = String(q.date || "").slice(0, 10), p = Math.floor(Number(q.period)), other = String(q.other || "");
  const rd = String(q.retDate || "").slice(0, 10), rp = Math.floor(Number(q.retPeriod));
  if (dayOf_(date) > 4 || dayOf_(rd) > 4) throw new Error("أحد اليومين عطلة (جمعة أو سبت).");
  const today = todayIso_();
  if (date < today || rd < today) throw new Error("لا يكون التبديل في يوم مضى.");
  if (date === rd) throw new Error("يوم الرد يكون غير يوم الحصة.");
  return withLock_(() => {
    const ev = swapEvents_(swpRows_());
    const les = effLesson_(teacher, date, p, ev);
    if (!isLes_(les)) throw new Error("ليست لك حصة في هذا الوقت.");
    const k = les[2].split(":")[1], otherCls = (les[2].split(":")[0] === "hd" ? "qr" : "hd") + ":" + k;
    if (!TIMETABLE[other] || other === teacher) throw new Error("اختر المعلم الذي ستبدّل معه.");
    if (isLes_(effLesson_(other, date, p, ev))) throw new Error("أ. " + other + " عنده حصة في هذا الوقت.");
    const rl = effLesson_(other, rd, rp, ev);
    if (!rl || rl[2] !== otherCls) throw new Error("حصة الرد ليست حصة أ. " + other + " في الفصل نفسه.");
    if (isLes_(effLesson_(teacher, rd, rp, ev))) throw new Error("عندك حصة في وقت الرد.");
    const abs = rows_(SHEET_ABS).slice(1).filter(r => r[0]);
    if (abs.some(r => (r[1] === date || r[1] === rd) && r[2] === other && r[6] !== ABS_EXCUSE)) throw new Error("أ. " + other + " مسجّل غائبًا في أحد اليومين.");
    const id = newId_();
    appendText_(swpSheet_(), [id, date, String(p), les[2], teacher, other, rd, String(rp), otherCls, SWP_ACTIVE, RET_PENDING, nowText_(), me.name]);
    const cl = c => esc_(c.split(":")[1].replace("-", "/")) + " · " + SUBJ[c.split(":")[0]];
    notify_(me, `تبديل حصة: أ. ${teacher} مع أ. ${other}`, `
      <p>سجّل <b>أ. ${esc_(teacher)}</b> تبديل حصة مع <b>أ. ${esc_(other)}</b> في الفصل ${esc_(k.replace("-", "/"))}:</p>
      <ul><li>${esc_(date)} الحصة ${p}: يدخل أ. ${esc_(other)} حصة ${cl(les[2])}</li>
      <li>${esc_(rd)} الحصة ${rp}: يدخل أ. ${esc_(teacher)} حصة ${cl(otherCls)}</li></ul>
      <p>نموذج التبديل يُطبع من الموقع ويوقَّع ورقيًا.</p>`);
    return { id };
  });
}
function cancelSwap(code, id) {
  const me = who_(code);
  return withLock_(() => {
    const sh = swpSheet_(), row = findRow_(sh, 1, String(id || ""));
    if (row < 0) throw new Error("التبديل غير موجود.");
    const r = sh.getRange(row, 1, 1, SWP_HEAD.length).getDisplayValues()[0];
    if (r[4] !== me.name && r[12] !== me.name && !(me.isHead && !me.isDemo)) throw new Error("لا يمكنك إلغاء هذا التبديل.");
    r[9] = SWP_CANCEL; writeRow_(sh, row, r);
    return true;
  });
}

// «عذر طارئ»: المعلم حاضر في الدوام لكن لا يستطيع دخول حصص محددة، فتُبدَّل تلك الحصص فقط (عمود G = النوع)
const ABS_EXCUSE = "عذر طارئ";
function absObj_(r) { return { id: r[0], date: r[1], teacher: r[2], note: r[3], createdAt: r[4], updatedAt: r[5], kind: r[6] === ABS_EXCUSE ? ABS_EXCUSE : "" }; }
function covObj_(r) {
  return { id: r[0], absenceId: r[1], date: r[2], period: Number(r[3]), classId: r[4], absent: r[5], substitute: r[6], status: r[7],
    retDate: r[9] || "", retPeriod: Number(r[10]) || 0, retClass: r[11] || "", retStatus: r[12] || "" };
}

/** كل بيانات الغياب والاحتياط */
function reserveData(code) {
  head_(code);
  return { absences: rows_(SHEET_ABS).slice(1).filter(r => r[0]).map(absObj_), covers: covSheet_().getDataRange().getDisplayValues().slice(1).filter(r => r[0]).map(covObj_), swaps: swpRows_(), swap: true, excuse: true };
}

/** تسجيل غياب معلم (أو تعديله) مع توزيع حصصه على البدلاء */
function saveAbsence(code, a) {
  head_(code);
  const teacher = String(a.teacher || ""), date = String(a.date || "").slice(0, 10);
  if (!TIMETABLE[teacher]) throw new Error("اختر المعلم الغائب.");
  const day = dayOf_(date);
  if (day > 4) throw new Error("التاريخ يوم عطلة (جمعة أو سبت).");
  const items = (Array.isArray(a.items) ? a.items : []).slice(0, 7), excuse = a.kind === ABS_EXCUSE;
  if (excuse && !items.length) throw new Error("اختر الحصة التي لا يستطيع المعلم دخولها.");
  return withLock_(() => {
    const absSh = sh_(SHEET_ABS), covSh = covSheet_(), ev = swapEvents_(swpRows_());
    if (absSh.getRange(1, 7).getDisplayValue() !== "النوع") absSh.getRange(1, 7).setValue("النوع");
    const absRows = absSh.getDataRange().getDisplayValues();
    let rowIdx = -1, id = String(a.id || "");
    for (let i = 1; i < absRows.length; i++) {
      if (id ? absRows[i][0] === id : (absRows[i][1] === date && absRows[i][2] === teacher)) { rowIdx = i + 1; id = absRows[i][0]; break; }
    }
    if (a.id && rowIdx < 0) throw new Error("الغياب غير موجود.");
    if (!id) id = newId_();
    const absentToday = absRows.slice(1).filter(r => r[0] && r[1] === date && r[0] !== id && r[6] !== ABS_EXCUSE).map(r => r[2]);
    if (absentToday.indexOf(teacher) !== -1) throw new Error("هذا المعلم مسجّل غائبًا في هذا اليوم.");
    const covRows = covSh.getDataRange().getDisplayValues();
    // حصص أصحاب الأعذار الطارئة الأخرى في اليوم: لا يُختارون بدلاء فيها
    const excIds = absRows.slice(1).filter(r => r[0] && r[0] !== id && r[6] === ABS_EXCUSE).map(r => r[0]);
    const excBusy = covRows.slice(1).filter(r => r[0] && r[2] === date && excIds.indexOf(r[1]) !== -1).map(r => r[5] + "|" + r[3]);
    const myPeriods = items.map(it => String(Math.floor(Number(it.period))));
    const asSub = covRows.slice(1).filter(r => r[0] && r[1] !== id && r[2] === date && isCov_(r[7]) && r[6] === teacher && (!excuse || myPeriods.indexOf(r[3]) !== -1));
    if (asSub.length) throw new Error("أ. " + teacher + " مسجّل بديلًا في هذا اليوم (الحصة " + asSub.map(r => r[3]).join("، ") + " بدلًا من أ. " + asSub[0][5] + "). غيّر البديل هناك أولًا.");
    const busy = covRows.slice(1).filter(r => r[0] && r[1] !== id && r[2] === date && isCov_(r[7])).map(r => r[6] + "|" + r[3]);
    // ساعات رد التبديل المحجوزة في تبديلات أخرى: "المعلم|التاريخ|الحصة"
    const retBusy = [];
    covRows.slice(1).filter(r => r[0] && r[1] !== id && r[7] === COV_SWAP && r[9]).forEach(r => { retBusy.push(r[5] + "|" + r[9] + "|" + r[10], r[6] + "|" + r[9] + "|" + r[10]); });
    const absentOn = d => absRows.slice(1).filter(r => r[0] && r[1] === d && r[0] !== id && r[6] !== ABS_EXCUSE).map(r => r[2]);
    const out = [], seen = {};
    items.forEach(it => {
      const period = Math.floor(Number(it.period));
      const les = effLesson_(teacher, date, period, ev);
      // حصص الاحتياط والاحتياط الفنية في الجدول العام حصص للمعلم، وتحتاج تغطية عند غيابه (قرار رئيس الشعبة ٢٦/١٠/٢٠٢٦)
      if (!les) throw new Error("الحصة " + period + " ليست من حصص المعلم في هذا اليوم.");
      if (seen[period]) return; seen[period] = 1;
      let status = [COV_DONE, COV_NONE, COV_PRESENT].indexOf(it.status) !== -1 ? it.status : COV_NONE;
      let sub = String(it.substitute || ""), ret = ["", "", "", ""];
      if (isCov_(status)) {
        if (!TIMETABLE[sub] || sub === teacher) throw new Error("اختر المعلم البديل للحصة " + period + ".");
        const l2 = effLesson_(sub, date, period, ev);
        if (l2 && l2[2] !== "R" && l2[2] !== "RF") throw new Error("أ. " + sub + " عنده حصة في الحصة " + period + ".");
        if (absentToday.indexOf(sub) !== -1) throw new Error("أ. " + sub + " غائب في هذا اليوم.");
        if (busy.indexOf(sub + "|" + period) !== -1) throw new Error("أ. " + sub + " مكلّف باحتياط آخر في الحصة " + period + ".");
        if (excBusy.indexOf(sub + "|" + period) !== -1) throw new Error("أ. " + sub + " عنده عذر طارئ في الحصة " + period + ".");
        if (status === COV_SWAP) ret = swapReturn_(teacher, sub, les[2], date, it, retBusy, absentOn, period);
      } else sub = "";
      out.push([newId_(), id, date, String(period), les[2], teacher, sub, status, nowText_()].concat(ret));
    });
    const now = nowText_();
    const absVals = [id, date, teacher, String(a.note || "").slice(0, 300), rowIdx > 0 ? absRows[rowIdx - 1][4] : now, now, excuse ? ABS_EXCUSE : ""];
    if (rowIdx > 0) writeRow_(absSh, rowIdx, absVals); else appendText_(absSh, absVals);
    for (let i = covRows.length - 1; i >= 1; i--) if (covRows[i][1] === id) covSh.deleteRow(i + 1);
    out.forEach(r => appendText_(covSh, r));
    return { id };
  });
}

/** يتحقق من حصة الرد في التبديل ويعيد [تاريخ الرد، حصة الرد، فصل الرد، حالة الرد] */
function swapReturn_(teacher, sub, cls, date, it, retBusy, absentOn, period) {
  const k = cls.split(":")[1], other = (cls.split(":")[0] === "hd" ? "qr" : "hd") + ":" + k;
  if (!(TIMETABLE[sub] || []).some(x => x[2] === other)) throw new Error("التبديل في الحصة " + period + " يكون مع معلم يدرّس الفصل نفسه في المادة الأخرى.");
  const rd = String(it.retDate || "").slice(0, 10), rp = Math.floor(Number(it.retPeriod));
  const rday = dayOf_(rd);
  if (rd <= date || rday > 4) throw new Error("اختر يوم رد صحيحًا للتبديل في الحصة " + period + ".");
  const l = lessonAt_(sub, rday, rp);
  if (!l || l[2] !== other) throw new Error("حصة الرد للتبديل في الحصة " + period + " ليست حصة أ. " + sub + " في الفصل نفسه.");
  const mine = lessonAt_(teacher, rday, rp);
  if (mine && mine[2] !== "R" && mine[2] !== "RF") throw new Error("أ. " + teacher + " عنده حصة في وقت الرد (الحصة " + period + ").");
  const off = absentOn(rd);
  if (off.indexOf(teacher) !== -1 || off.indexOf(sub) !== -1) throw new Error("أحد المعلمَين غائب في يوم الرد (الحصة " + period + ").");
  if (retBusy.indexOf(teacher + "|" + rd + "|" + rp) !== -1 || retBusy.indexOf(sub + "|" + rd + "|" + rp) !== -1) throw new Error("وقت الرد للحصة " + period + " محجوز في تبديل آخر.");
  retBusy.push(teacher + "|" + rd + "|" + rp, sub + "|" + rd + "|" + rp);
  return [rd, String(rp), other, RET_PENDING];
}

/** تعليم رد التبديل: تم أو لم يتم (رئيس الشعبة) */
function setSwapReturn(code, id, done) {
  head_(code);
  return withLock_(() => {
    const sh = swpSheet_(), row = findRow_(sh, 1, String(id || ""));
    if (row < 0) throw new Error("التبديل غير موجود.");
    sh.getRange(row, 11).setNumberFormat("@").setValue(done ? RET_DONE : RET_PENDING);
    return true;
  });
}

function deleteAbsence(code, id) {
  head_(code);
  return withLock_(() => {
    const absSh = sh_(SHEET_ABS), row = findRow_(absSh, 1, id);
    if (row > 0) absSh.deleteRow(row);
    const covSh = covSheet_(), v = covSh.getDataRange().getDisplayValues();
    for (let i = v.length - 1; i >= 1; i--) if (v[i][1] === id) covSh.deleteRow(i + 1);
    return true;
  });
}

// ======================= كشوف متابعة المتعلمين (الدرجات) =======================
// قوائم الطلاب في ورقة «الطلاب» (الصف-الشعبة، م، —، ملاحظة) بالأرقام فقط، والدرجات في ورقة «الدرجات».
// قرار رئيس الشعبة (٢٥/١٠/٢٠٢٦): لا تُحفظ أسماء المتعلمين في الموقع ولا في البيانات، للخصوصية والتزامًا بلوائح وزارة التربية؛ المتعلم يُعرَّف برقمه في كشف الصف.
// توزيع الدرجات حسب خطة التوجيه الفني ٢٠٢٦/٢٠٢٧ (واحد للحديث والقرآن):
// السلوك ٣+٣ · المشاركة ٣+٣ · الواجبات ٣+٣ · المشروع ٣٫٥+٣٫٥ (الفترة الأولى + الفترة الثانية) · اختبار قصير ١٥ = المجموع ٤٠
// قرار رئيس الشعبة: لا يُرصد في الكشف غير هذه البنود (لا امتحان نهائي)
const SHEET_STU = "الطلاب";
const SHEET_GRD = "الدرجات";
const GRD_MAX = [3, 3, 3, 3, 3, 3, 3.5, 3.5, 15];
const GRD_HEAD = ["المفتاح", "القسم", "الصف", "الفصل الدراسي", "م", "(بلا أسماء)", "السلوك ق١", "السلوك ق٢", "المشاركة ق١", "المشاركة ق٢",
  "الواجبات ق١", "الواجبات ق٢", "المشروع ق١", "المشروع ق٢", "الاختبار القصير", "المجموع (٤٠)", "المعلم", "آخر تحديث"];
const DEMO_STUDENTS = 12;

function teacherOf_(cls) { const t = TEACHERS.find(x => x[1].indexOf(cls) !== -1); return t ? t[0] : ""; }
function canSeeClass_(me, cls) { return me.isHead && !me.isDemo ? !!teacherOf_(cls) : me.classes.indexOf(cls) !== -1; }
function gradesSheet_() { return sheet_(SHEET_GRD, GRD_HEAD, true); }
function studentsOf_(me, cls) {
  if (me.isDemo) return Array.from({ length: DEMO_STUDENTS }, (_, i) => ({ n: i + 1, note: "" }));
  const sec = String(cls).split(":")[1];
  const sh = sh_(SHEET_STU);
  if (!sh) return [];
  return sh.getDataRange().getDisplayValues().slice(1).filter(r => r[0] === sec && Number(r[1]) > 0)
    .map(r => ({ n: Number(r[1]), note: String(r[3] || "") })).sort((a, b) => a.n - b.n);
}

/** حذف أسماء المتعلمين من البيانات (مرة واحدة، ثم لا يعمل): الطلاب، الدرجات، تقارير المتعلمين، الإشعارات.
 *  الدرجات تُربط برقم المتعلم بدل اسمه، والتقارير تحمل رقمه. لا تُحذف أي درجة. */
function anonymizeStudents_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("anonStudents") === "1") return;
  withLock_(() => {
    if (props.getProperty("anonStudents") === "1") return;
    const num = {}, byName = {};
    const stu = sh_(SHEET_STU);
    if (stu && stu.getLastRow() > 1) {
      const v = stu.getRange(2, 1, stu.getLastRow() - 1, 4).getDisplayValues();
      v.forEach(r => { const nm = String(r[2]).trim(); if (nm && r[0]) { num[r[0] + "|" + nm] = r[1]; if (!byName[nm]) byName[nm] = r[1]; } });
      stu.getRange(2, 3, v.length, 1).clearContent();
      stu.getRange(1, 3).setValue("(بلا أسماء)");
    }
    const grd = sh_(SHEET_GRD);
    if (grd && grd.getLastRow() > 1) {
      const n = grd.getLastRow() - 1, v = grd.getRange(2, 1, n, 6).getDisplayValues();
      const keys = v.map(r => { const k = String(r[0]); if (!k || !r[4]) return [k]; return [k.slice(0, k.lastIndexOf("|") + 1) + r[4]]; });
      grd.getRange(2, 1, n, 1).setNumberFormat("@").setValues(keys);
      grd.getRange(2, 6, n, 1).clearContent();
      grd.getRange(1, 6).setValue("(بلا أسماء)");
    }
    const srp = sh_(SHEET_SREP);
    if (srp && srp.getLastRow() > 1) {
      const n = srp.getLastRow() - 1, v = srp.getRange(2, 1, n, 7).getDisplayValues();
      srp.getRange(2, 7, n, 1).setNumberFormat("@").setValues(v.map(r => {
        const nm = String(r[6]).trim(); if (/^\d+$/.test(nm)) return [nm];
        return [num[String(r[3]).split(":")[1] + "|" + nm] || ""];
      }));
      srp.getRange(1, 7).setValue("رقم المتعلم");
    }
    const act = sh_(SHEET_ACTLOG);
    if (act && act.getLastRow() > 1) {
      const n = act.getLastRow() - 1, rg = act.getRange(2, 4, n, 1), v = rg.getDisplayValues();
      let changed = false;
      const out = v.map(([t]) => { const m = String(t).match(/^((?:أرسل|عدّل) تقرير [^:]+): (.+) \((.+)\)$/);
        if (!m || /^المتعلم رقم/.test(m[2])) return [t]; changed = true;
        return [`${m[1]}: المتعلم رقم ${arNum_(byName[m[2].trim()] || "؟")} (${m[3]})`]; });
      if (changed) rg.setNumberFormat("@").setValues(out);
    }
    props.setProperty("anonStudents", "1");
  });
}
function sum_(vals) { return Math.round(vals.reduce((a, v) => a + Number(v), 0) * 100) / 100; }
function totals_(v) { return [v.every(x => x !== "") ? String(sum_(v)) : ""]; }

/** نظرة عامة على صفوفي (أو كل الصفوف لرئيس الشعبة): عدد المتعلمين ومدى اكتمال الرصد في كل فصل دراسي */
function classesOverview(code) {
  const me = who_(code);
  const list = me.isHead && !me.isDemo ? TEACHERS.reduce((a, t) => a.concat(t[1]), []) : me.classes;
  const counts = {};
  anonymizeStudents_();
  if (me.isDemo) list.forEach(c => counts[c.split(":")[1]] = DEMO_STUDENTS);
  else if (sh_(SHEET_STU)) rows_(SHEET_STU).slice(1).forEach(r => { if (r[0] && Number(r[1]) > 0) counts[r[0]] = (counts[r[0]] || 0) + 1; });
  const pre = demoPrefix_(me), stat = {};
  if (sh_(SHEET_GRD)) rows_(SHEET_GRD).slice(1).forEach(r => {
    const k = String(r[0] || "");
    if (pre ? k.indexOf(pre) !== 0 : k.indexOf("DEMO|") === 0) return;
    const p = k.slice(pre.length).split("|"), id = p[0] + "|" + p[1];
    const s = stat[id] || (stat[id] = { any: 0, form: 0, low: 0, updatedAt: "" });
    if (r.slice(6, 15).some(x => x !== "")) s.any++;
    if (r[15] !== "") { s.form++; if (Number(r[15]) < 20) s.low++; }
    if (r[17] > s.updatedAt) s.updatedAt = r[17];
  });
  return list.map(c => ({ cls: c, teacher: me.isDemo ? me.name : teacherOf_(c), students: counts[c.split(":")[1]] || 0,
    t1: stat[c + "|1"] || null, t2: stat[c + "|2"] || null }));
}

/** كشف صف واحد في فصل دراسي: المتعلمون ودرجاتهم */
function classSheet(code, cls, term) {
  const me = who_(code);
  cls = String(cls || ""); term = Number(term) === 2 ? 2 : 1;
  if (!canSeeClass_(me, cls)) throw new Error("هذا الصف ليس من صفوفك.");
  anonymizeStudents_();
  const pre = demoPrefix_(me) + cls + "|" + term + "|";
  const grades = {}; let updatedAt = "";
  if (sh_(SHEET_GRD)) rows_(SHEET_GRD).slice(1).forEach(r => {
    if (String(r[0]).indexOf(pre) !== 0) return;
    grades[String(r[0]).slice(pre.length)] = r.slice(6, 15);
    if (r[17] > updatedAt) updatedAt = r[17];
  });
  return { cls, term, teacher: me.isDemo ? me.name : teacherOf_(cls), students: studentsOf_(me, cls), grades, updatedAt };
}

/** حفظ درجات صف: data = { "رقم المتعلم": [٩ قيم نصية أو فارغة] } */
function saveGrades(code, cls, term, data) {
  const me = who_(code);
  cls = String(cls || ""); term = Number(term) === 2 ? 2 : 1;
  if (!canSeeClass_(me, cls)) throw new Error("هذا الصف ليس من صفوفك.");
  const studs = studentsOf_(me, cls);
  if (!studs.length) throw new Error("لا توجد قائمة متعلمين لهذا الصف.");
  const p = cls.split(":"), dept = SUBJ[p[0]], teacher = me.isDemo ? me.name : teacherOf_(cls);
  const pre = demoPrefix_(me) + cls + "|" + term + "|", now = nowText_();
  const rows = [];
  studs.forEach(s => {
    const v = (data && data[String(s.n)]) || [];
    const vals = GRD_MAX.map((mx, i) => {
      const x = String(v[i] == null ? "" : v[i]).trim();
      if (x === "") return "";
      if (!/^\d{1,2}(\.\d{1,2})?$/.test(x) || Number(x) > mx) throw new Error(`درجة غير صحيحة للمتعلم رقم ${s.n}: «${x}» (العظمى ${mx}).`);
      return String(Number(x));
    });
    rows.push([pre + s.n, dept, p[1], term === 1 ? "الأول" : "الثاني", String(s.n), ""].concat(vals, totals_(vals), [teacher, now]));
  });
  return withLock_(() => {
    const sh = gradesSheet_();
    const keys = sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 1).getDisplayValues().map(r => r[0]);
    const idx = rows.map(r => keys.indexOf(r[0]) + 1);
    const contiguous = idx[0] > 0 && idx.every((x, i) => x === idx[0] + i);
    if (contiguous) {
      const rg = sh.getRange(idx[0], 1, rows.length, GRD_COLS_());
      rg.setNumberFormat("@"); rg.setValues(rows);
    } else {
      const fresh = [];
      rows.forEach((r, i) => { if (idx[i] > 0) writeRow_(sh, idx[i], r); else fresh.push(r); });
      if (fresh.length) { const rg = sh.getRange(sh.getLastRow() + 1, 1, fresh.length, GRD_COLS_()); rg.setNumberFormat("@"); rg.setValues(fresh); }
    }
    return { updatedAt: now };
  });
}
function GRD_COLS_() { return GRD_HEAD.length; }

// ======================= تقارير المتعلمين (الفائق والمتعثر) =======================
// يرسلها المعلم لمتعلم من صفوفه، فتصل رئيس الشعبة بالإيميل وتظهر عنده في ركن الصفوف
const SHEET_SREP = "تقارير المتعلمين";
const SREP_TYPES = { excel: "متعلم فائق", weak: "متعلم متعثر" };
const SREP_HEAD = ["المعرف", "النوع", "المعلم", "الصف", "القسم", "الفصل الدراسي", "رقم المتعلم", "المجموع في الكشف", "البيانات", "تاريخ الإرسال", "اطّلع رئيس الشعبة", "آخر تعديل"];
function srepSheet_() { return sheet_(SHEET_SREP, SREP_HEAD, true); }
function srepObj_(r) {
  let data = {}; try { data = JSON.parse(r[8] || "{}"); } catch (e) {}
  let plan = null; try { plan = r[12] ? JSON.parse(r[12]) : null; } catch (e) {}
  return { id: r[0], type: r[1], teacher: r[2], cls: r[3], term: Number(r[5]) || 1, student: r[6], score: r[7], data, createdAt: r[9], seenAt: r[10], updatedAt: r[11],
    plan, planAt: r[13] || "" };
}

// ======================= الخطة العلاجية التفصيلية للمتعلم المتعثر =======================
// اختيارية للمعلم، تُبنى على تقرير «متعلم متعثر» (المظاهر والأسباب منه)، وتُحفظ في عمودين بعد أعمدة التقرير
const PLAN_HEAD = ["الخطة العلاجية", "آخر تحديث للخطة"];
function srepPlanSheet_() {
  const sh = srepSheet_(), c = SREP_HEAD.length, n = c + PLAN_HEAD.length;
  if (sh.getMaxColumns() < n) sh.insertColumnsAfter(sh.getMaxColumns(), n - sh.getMaxColumns());
  if (sh.getRange(1, c + 1).getDisplayValue() !== PLAN_HEAD[0]) {
    sh.getRange(1, c + 1, sh.getMaxRows(), PLAN_HEAD.length).setNumberFormat("@");
    sh.getRange(1, c + 1, 1, PLAN_HEAD.length).setValues([PLAN_HEAD]);
  }
  return sh;
}
function cleanPlan_(p) {
  p = p || {};
  const S = (v, n) => String(v == null ? "" : v).trim().slice(0, n || 300);
  const rows = (a, keys, max) => (Array.isArray(a) ? a : []).slice(0, max).map(r => { const o = {}; keys.forEach(k => o[k] = S(r && r[k])); return o; }).filter(o => keys.some(k => o[k]));
  return { start: S(p.start, 10), weeks: S(p.weeks, 2), tools: cleanList_(p.tools), preQuiz: S(p.preQuiz, 6), preTotal: S(p.preTotal, 6),
    goals: rows(p.goals, ["g", "crit"], 6), strat: cleanList_(p.strat), actions: rows(p.actions, ["a", "tool", "time", "by"], 8),
    follow: rows(p.follow, ["week", "date", "done", "level"], 8), parent: rows(p.parent, ["date", "how", "agreed"], 6),
    postQuiz: S(p.postQuiz, 6), postTotal: S(p.postTotal, 6), result: S(p.result, 30), recs: S(p.recs, 1500) };
}
function saveRemedialPlan(code, id, plan) {
  const me = who_(code);
  const sh = srepPlanSheet_(), c = SREP_HEAD.length;
  return withLock_(() => {
    const row = findRow_(sh, 1, String(id || ""));
    if (row < 0) throw new Error("التقرير غير موجود.");
    const old = sh.getRange(row, 1, 1, c + PLAN_HEAD.length).getDisplayValues()[0];
    if (old[1] !== SREP_TYPES.weak) throw new Error("الخطة العلاجية للمتعلم المتعثر فقط.");
    if (old[2] !== me.name && !(me.isHead && !me.isDemo)) throw new Error("لا يمكنك تعديل خطة غيرك.");
    const P = cleanPlan_(plan);
    if (!P.goals.length) throw new Error("اكتب هدفًا واحدًا على الأقل للخطة.");
    let prev = {}; try { prev = old[c] ? JSON.parse(old[c]) : {}; } catch (e) {}
    const now = nowText_(), rg = sh.getRange(row, c + 1, 1, PLAN_HEAD.length);
    rg.setNumberFormat("@"); rg.setValues([[JSON.stringify(P), now]]);
    const first = !old[c], done = P.result && P.result !== prev.result, p = old[3].split(":");
    if (first || done) notify_(me, `${first ? "أعدّ" : "أنهى"} أ. ${me.name} الخطة العلاجية للمتعلم رقم ${arNum_(old[6])} (${SUBJ[p[0]]})`, `
      <p>${first ? "أعدّ" : "سجّل"} <b>أ. ${esc_(me.name)}</b> ${first ? "خطة علاجية تفصيلية" : "التقويم الختامي للخطة العلاجية"} لـ<b>المتعلم رقم ${arNum_(old[6])}</b> في الصف ${esc_(p[1].replace("-", "/"))} · ${esc_(SUBJ[p[0]])}.</p>
      ${done ? `<p><b>النتيجة:</b> ${esc_(P.result)}</p>` : `<p><b>الأهداف:</b></p><ul>${P.goals.map(g => "<li>" + esc_(g.g) + "</li>").join("")}</ul>`}`);
    return { planAt: now };
  });
}
function cleanList_(a) { return (Array.isArray(a) ? a : []).slice(0, 20).map(x => String(x).trim().slice(0, 200)).filter(Boolean); }

function listStudentReports(code) {
  const me = who_(code);
  anonymizeStudents_();
  const sh = sh_(SHEET_SREP);
  if (!sh) return [];
  return sh.getDataRange().getDisplayValues().slice(1).filter(r => r[0])
    .filter(r => me.isHead && !me.isDemo ? !isDemo_(r[2]) : r[2] === me.name)
    .map(srepObj_).reverse();
}

function saveStudentReport(code, rep) {
  const me = who_(code);
  const cls = String(rep.cls || "");
  if (!canSeeClass_(me, cls)) throw new Error("هذا الصف ليس من صفوفك.");
  if (!SREP_TYPES[rep.type]) throw new Error("نوع التقرير غير معروف.");
  const student = String(rep.student || "").trim();
  if (!studentsOf_(me, cls).some(s => String(s.n) === student)) throw new Error("اختر رقم المتعلم من قائمة الصف.");
  const d = rep.data || {};
  const data = { a: cleanList_(d.a), b: cleanList_(d.b), c: cleanList_(d.c),
    desc: String(d.desc || "").trim().slice(0, 2000), need: String(d.need || "").trim().slice(0, 1000),
    rec: String(d.rec || "").trim().slice(0, 1500), other: String(d.other || "").trim().slice(0, 1000) };
  if (!data.a.length) throw new Error(rep.type === "excel" ? "اختر مظهرًا واحدًا على الأقل من مظاهر التفوق." : "اختر مظهرًا واحدًا على الأقل من مظاهر التعثر.");
  if (data.desc.length < 10) throw new Error("اكتب وصفًا موجزًا لحالة المتعلم ولو بسطر.");
  const term = Number(rep.term) === 2 ? 2 : 1, now = nowText_();
  const teacher = me.name;
  const p = cls.split(":");
  const sh = srepSheet_();
  let id = String(rep.id || ""), created = now, row = -1;
  if (id) {
    row = findRow_(sh, 1, id);
    if (row < 0) throw new Error("التقرير غير موجود.");
    const old = sh.getRange(row, 1, 1, SREP_HEAD.length).getDisplayValues()[0];
    if (old[2] !== me.name && !(me.isHead && !me.isDemo)) throw new Error("لا يمكنك تعديل تقرير غيرك.");
    created = old[9];
  } else id = newId_();
  const vals = [id, SREP_TYPES[rep.type], teacher, cls, SUBJ[p[0]], String(term), student, String(rep.score || "").slice(0, 40), JSON.stringify(data), created, "", row > 0 ? now : ""];
  withLock_(() => { if (row > 0) writeRow_(sh, row, vals); else appendText_(sh, vals); });
  const li = a => a.length ? "<ul>" + a.map(x => "<li>" + esc_(x) + "</li>").join("") + "</ul>" : "—";
  const t = rep.type === "excel";
  notify_(me, `${row > 0 ? "عدّل" : "أرسل"} أ. ${me.name} تقرير ${SREP_TYPES[rep.type]} (${SUBJ[p[0]]}): المتعلم رقم ${arNum_(student)}`, `
    <p>${row > 0 ? "عدّل" : "أرسل"} <b>أ. ${esc_(me.name)}</b> تقرير <b>${SREP_TYPES[rep.type]}</b> في <b>${SUBJ[p[0]]}</b> عن <b>المتعلم رقم ${arNum_(student)}</b>.</p>
    <table cellpadding="6" style="border-collapse:collapse;border:1px solid #d6e2ec">
      <tr><td style="background:#e8f1f8"><b>الصف</b></td><td>${esc_(p[1].replace("-", "/"))} · ${esc_(SUBJ[p[0]])}</td></tr>
      <tr><td style="background:#e8f1f8"><b>الفصل الدراسي</b></td><td>${term === 1 ? "الأول" : "الثاني"}</td></tr>
      <tr><td style="background:#e8f1f8"><b>المجموع في الكشف</b></td><td>${esc_(vals[7] || "لم يكتمل رصده بعد")}</td></tr>
    </table>
    <p><b>${t ? "مظاهر التفوق" : "مظاهر التعثر"}:</b></p>${li(data.a)}
    ${t ? "" : `<p><b>الأسباب المحتملة:</b></p>${li(data.b)}`}
    <p><b>${t ? "أساليب الرعاية والإثراء" : "الإجراءات العلاجية المتخذة"}:</b></p>${li(t ? data.b : data.c)}
    <p><b>الوصف:</b><br>${esc_(data.desc).replace(/\n/g, "<br>")}</p>
    ${data.need ? `<p><b>الدعم المطلوب من رئيس الشعبة أو الإدارة:</b><br>${esc_(data.need).replace(/\n/g, "<br>")}</p>` : ""}
    ${data.rec ? `<p><b>توصية المعلم:</b><br>${esc_(data.rec).replace(/\n/g, "<br>")}</p>` : ""}`);
  return { id };
}

function seenStudentReport(code, id) {
  const me = head_(code);
  if (me.isDemo) return true;
  return withLock_(() => {
    const sh = sh_(SHEET_SREP), row = sh ? findRow_(sh, 1, String(id)) : -1;
    if (row > 0 && !sh.getRange(row, 11).getDisplayValue()) sh.getRange(row, 11).setNumberFormat("@").setValue(nowText_());
    return true;
  });
}

function deleteStudentReport(code, id) {
  const me = who_(code);
  return withLock_(() => {
    const sh = sh_(SHEET_SREP), row = sh ? findRow_(sh, 1, String(id)) : -1;
    if (row < 0) throw new Error("التقرير غير موجود.");
    const owner = sh.getRange(row, 3).getDisplayValue();
    if (owner !== me.name && !(me.isHead && !me.isDemo)) throw new Error("لا يمكنك حذف تقرير غيرك.");
    sh.deleteRow(row);
    return true;
  });
}

// ======================= سجل الاستئذان (إذن خروج — نموذج ٥م) =======================
// المعلم يحفظ بياناته (الوظيفة، الرقم المدني، القسم) مرة واحدة، ثم يطلب إذن الخروج ويوقّع،
// فيصل رئيس الشعبة إيميل ويعتمده بتوقيعه. توقيع المدير المساعد ومدير المعهد وجزء التأخير على الورقة المطبوعة.
// الرقم المدني في الشيت فقط (لا في المستودع)، ولا يراه إلا صاحبه ورئيس الشعبة.
const SHEET_PROF = "بيانات المعلمين";
const SHEET_PERM = "الاستئذان";
const PERM_TYPES = ["بداية الدوام", "أثناء الدوام", "نهاية الدوام"];
const PERM_PENDING = "بانتظار الاعتماد", PERM_OK = "معتمد", PERM_NO = "مرفوض";
const PERM_HEAD = ["المعرف", "المسلسل", "المعلم", "نوع الاستئذان", "القسم", "الوظيفة", "الرقم المدني", "سبب التصريح", "التاريخ", "من", "إلى",
  "المعلم متفرغ / بديل", "اسم البديل", "توقيع المعلم", "وقت الإرسال", "الحالة", "توقيع رئيس الشعبة", "وقت الاعتماد", "ملاحظة رئيس الشعبة"];
function profSheet_() { return sheet_(SHEET_PROF, ["المعلم", "الوظيفة", "الرقم المدني", "القسم", "آخر تحديث"], true); }
function permSheet_() { return sheet_(SHEET_PERM, PERM_HEAD, true); }
function cleanProfile_(p) {
  const job = String(p && p.job || "").trim().slice(0, 60), dept = String(p && p.dept || "").trim().slice(0, 80);
  const civil = String(p && p.civil || "").replace(/[٠-٩]/g, d => "٠١٢٣٤٥٦٧٨٩".indexOf(d)).replace(/\s/g, "");
  if (!job) throw new Error("اكتب الوظيفة.");
  if (!/^\d{12}$/.test(civil)) throw new Error("الرقم المدني يتكون من ١٢ رقمًا.");
  if (!dept) throw new Error("اكتب القسم.");
  return { job, civil, dept };
}
function myProfile(code) {
  const me = who_(code), sh = sh_(SHEET_PROF);
  const r = sh ? sh.getDataRange().getDisplayValues().slice(1).find(x => x[0] === me.name) : null;
  return r ? { job: r[1], civil: r[2], dept: r[3] } : null;
}
function saveProfile_(name, p) {
  const sh = profSheet_(), row = findRow_(sh, 1, name), vals = [name, p.job, p.civil, p.dept, nowText_()];
  if (row > 0) writeRow_(sh, row, vals); else appendText_(sh, vals);
}
function saveProfile(code, p) { const me = who_(code), c = cleanProfile_(p); withLock_(() => saveProfile_(me.name, c)); return c; }

function permObj_(r, withImages) {
  const o = { id: r[0], serial: r[1], teacher: r[2], type: r[3], dept: r[4], job: r[5], civil: r[6], reason: r[7], date: r[8], from: r[9], to: r[10],
    mode: r[11], substitute: r[12], sentAt: r[14], status: r[15], decidedAt: r[17], headNote: r[18], editedAt: r[19] || "" };
  if (withImages) { o.teacherSig = r[13]; o.headSig = r[16]; }
  return o;
}
function permRows_() { const sh = sh_(SHEET_PERM); return sh ? sh.getDataRange().getDisplayValues().slice(1).filter(r => r[0]) : []; }
function canSeePerm_(me, r) { return r[2] === me.name || (me.isHead && !me.isDemo && !isDemo_(r[2])); }

function listPermissions(code) {
  const me = who_(code);
  return permRows_().filter(r => canSeePerm_(me, r)).map(r => permObj_(r, false)).reverse();
}
function getPermission(code, id) {
  const me = who_(code), r = permRows_().find(x => x[0] === String(id));
  if (!r || !canSeePerm_(me, r)) throw new Error("الاستئذان غير موجود.");
  return permObj_(r, true);
}
function submitPermission(code, p) {
  const me = who_(code), prof = cleanProfile_(p);
  const type = String(p.type || ""); if (PERM_TYPES.indexOf(type) === -1) throw new Error("اختر نوع الاستئذان: بداية الدوام أو أثناءه أو نهايته.");
  const reason = String(p.reason || "").trim().slice(0, 300); if (reason.length < 3) throw new Error("اكتب سبب الاستئذان.");
  const date = String(p.date || ""); const day = dayOf_(date); if (day > 4) throw new Error("التاريخ ليس يوم دوام.");
  const tm = t => { const m = String(t || "").match(/^(\d{2}):(\d{2})$/); if (!m) throw new Error("اكتب وقت الاستئذان (من وإلى)."); return +m[1] * 60 + +m[2]; };
  const toEnd = p.to === "نهاية الدوام";
  tm(p.from);
  if (!toEnd && tm(p.to) <= tm(p.from)) throw new Error("وقت النهاية لازم يكون بعد وقت البداية.");
  const mode = p.mode === "بديل" ? "بديل" : "متفرغ";
  const sub = mode === "بديل" ? String(p.substitute || "").trim().slice(0, 80) : "";
  if (mode === "بديل" && !sub) throw new Error("اكتب اسم المعلم البديل.");
  p.sig = String(p.sig || ""); if (p.sig && !validSig_(p.sig)) throw new Error("التوقيع غير صالح.");
  if (p.id) {
    const ed = withLock_(() => {
      const sh = permSheet_(), row = findRow_(sh, 1, String(p.id));
      if (row < 0) throw new Error("الاستئذان غير موجود.");
      const r = sh.getRange(row, 1, 1, PERM_HEAD.length).getDisplayValues()[0];
      if (r[2] !== me.name) throw new Error("ليس طلبك.");
      saveProfile_(me.name, prof);
      const now = nowText_();
      writeRow_(sh, row, [r[0], r[1], r[2], type, prof.dept, prof.job, prof.civil, reason, date, p.from, p.to, mode, sub, p.sig || r[13], r[14], r[15], r[16], r[17], r[18], now]);
      if (p.sig) saveMySig_(me.name, p.sig, now);
      return r[1];
    });
    notify_(me, `تعديل طلب استئذان من أ. ${me.name} (${date})`, `<p>عدّل <b>أ. ${esc_(me.name)}</b> طلب الاستئذان رقم ${esc_(ed)}.</p>
      <p>${esc_(type)} · ${esc_(date)} · من ${esc_(p.from)} ${p.to === "نهاية الدوام" ? "حتى نهاية الدوام" : "إلى " + esc_(p.to)} · ${esc_(reason)}</p>`);
    return { id: String(p.id), serial: ed, edited: true };
  }
  const id = newId_(), now = nowText_(), year = date.slice(0, 4);
  const serial = withLock_(() => {
    saveProfile_(me.name, prof);
    const rows = permRows_().filter(r => r[8].slice(0, 4) === year && (me.isDemo ? r[2] === me.name : !isDemo_(r[2])));
    const n = rows.reduce((mx, r) => Math.max(mx, Number(String(r[1]).split("/")[0]) || 0), 0) + 1;
    const s = n + "/" + year;
    appendText_(permSheet_(), [id, s, me.name, type, prof.dept, prof.job, prof.civil, reason, date, p.from, p.to, mode, sub, p.sig, now, PERM_PENDING, "", "", ""]);
    if (p.sig) saveMySig_(me.name, p.sig, now);
    return s;
  });
  if (!p.deferMail) permMail_(me, { serial, type, date, from: p.from, to: p.to, reason, mode, sub }, null);
  return { id, serial };
}
function permMail_(me, q, pdf) {
  const toEnd = q.to === "نهاية الدوام";
  notify_(me, `طلب استئذان من أ. ${me.name} (${q.date})`, `
    <p>أرسل <b>أ. ${esc_(me.name)}</b> طلب استئذان${pdf ? "، والنموذج مرفق PDF جاهز للطباعة والتوقيع" : ""}.</p>
    <table cellpadding="6" style="border-collapse:collapse;border:1px solid #d6e2ec">
      <tr><td style="background:#e8f1f8"><b>المسلسل</b></td><td>${esc_(q.serial)}</td></tr>
      <tr><td style="background:#e8f1f8"><b>النوع</b></td><td>${esc_(q.type)}</td></tr>
      <tr><td style="background:#e8f1f8"><b>التاريخ</b></td><td>${esc_(q.date)}</td></tr>
      <tr><td style="background:#e8f1f8"><b>الوقت</b></td><td>من ${esc_(q.from)} ${toEnd ? "حتى نهاية الدوام" : "إلى " + esc_(q.to)}</td></tr>
      <tr><td style="background:#e8f1f8"><b>السبب</b></td><td>${esc_(q.reason)}</td></tr>
      <tr><td style="background:#e8f1f8"><b>الحصص</b></td><td>${q.mode === "بديل" ? "بديل المعلم: " + esc_(q.sub) : "المعلم متفرغ"}</td></tr>
    </table>`, pdf ? [pdf] : []);
}
/** بعد الإرسال: يرسل المتصفح نموذج الإذن PDF فيصل رئيس الشعبة مرفقًا بالإيميل (مرة واحدة لكل طلب) */
function mailPermissionPdf(code, id, b64) {
  const me = who_(code);
  const sh = permSheet_(), row = findRow_(sh, 1, String(id));
  if (row < 0) throw new Error("الاستئذان غير موجود.");
  const r = sh.getRange(row, 1, 1, PERM_HEAD.length).getDisplayValues()[0];
  if (r[2] !== me.name) throw new Error("ليس طلبك.");
  const props = PropertiesService.getScriptProperties(), key = "permMail_" + r[0];
  if (props.getProperty(key)) return { sent: false };
  props.setProperty(key, "1");
  let pdf = null;
  const data = String(b64 || "").replace(/^data:application\/pdf;base64,/, "");
  if (data && data.length < 8000000) pdf = Utilities.newBlob(Utilities.base64Decode(data), "application/pdf", `إذن خروج ${r[1].replace("/", "-")} - ${r[2]}.pdf`);
  permMail_(me, { serial: r[1], type: r[3], date: r[8], from: r[9], to: r[10], reason: r[7], mode: r[11], sub: r[12] }, pdf);
  return { sent: true, attached: !!pdf };
}
function decidePermission(code, id, approve, sig, note) {
  const me = head_(code);
  return withLock_(() => {
    const sh = permSheet_(), row = findRow_(sh, 1, String(id));
    if (row < 0) throw new Error("الاستئذان غير موجود.");
    const r = sh.getRange(row, 1, 1, PERM_HEAD.length).getDisplayValues()[0];
    if (isDemo_(r[2]) && !me.isDemo) throw new Error("هذا طلب حساب تجريبي.");
    const now = nowText_();
    if (approve) {
      if (!validSig_(sig)) throw new Error("وقّع لاعتماد الاستئذان.");
      sh.getRange(row, 16, 1, 4).setNumberFormat("@").setValues([[PERM_OK, sig, now, String(note || "").slice(0, 300)]]);
      saveMySig_(me.name, sig, now);
    } else {
      sh.getRange(row, 16, 1, 4).setNumberFormat("@").setValues([[PERM_NO, "", now, String(note || "").slice(0, 300)]]);
    }
    return { status: approve ? PERM_OK : PERM_NO, decidedAt: now };
  });
}
function deletePermission(code, id) {
  const me = who_(code);
  return withLock_(() => {
    const sh = permSheet_(), row = findRow_(sh, 1, String(id));
    if (row < 0) throw new Error("الاستئذان غير موجود.");
    const r = sh.getRange(row, 1, 1, PERM_HEAD.length).getDisplayValues()[0];
    const head = me.isHead && !me.isDemo;
    if (!(head || r[2] === me.name)) throw new Error("لا يمكن حذف هذا الاستئذان.");
    sh.deleteRow(row);
    return true;
  });
}


// ======================= سجل النشرات والتعاميم =======================
// رئيس الشعبة يرفع النشرة/التعميم (ملف PDF أو صورة)، والمعلمون المعنيون يطّلعون ويوقّعون.
// أعمدة «النشرات والتعاميم»: 0 المعرف | 1 النوع | 2 الرقم | 3 التاريخ | 4 العنوان | 5 جهة الإصدار | 6 المعنيون (all/hd/qr)
// 7 ملاحظة | 8 الملف (Drive) | 9 اسم الملف | 10 نوع الملف | 11 وقت الرفع
// أعمدة «اطلاع النشرات»: 0 النشرة | 1 المعلم | 2 وقت الاطلاع | 3 التوقيع
const NSH_HEAD = ["المعرف", "النوع", "الرقم", "التاريخ", "العنوان", "جهة الإصدار", "المعنيون", "ملاحظة", "الملف", "اسم الملف", "نوع الملف", "وقت الرفع"];
const NSH_AUD = { all: "جميع معلمي الشعبة", hd: "معلمو الحديث الشريف", qr: "معلمو القرآن الكريم" };
function nshSheet_() { return sheet_(SHEET_NSH, NSH_HEAD, true); }
function nshAckSheet_() { return sheet_(SHEET_NSHACK, ["النشرة", "المعلم", "وقت الاطلاع", "التوقيع"], true); }
function nshRows_() { return nshSheet_().getDataRange().getDisplayValues().slice(1).filter(r => r[0]); }
function nshAcks_() { return nshAckSheet_().getDataRange().getDisplayValues().slice(1).filter(r => r[0]); }
/** المعلمون المعنيون بالنشرة (بدون رئيس الشعبة) */
function nshTargets_(aud) {
  return TEACHERS.filter(t => !t[2] && (aud === "all" || t[1].some(c => c.indexOf(aud + ":") === 0))).map(t => t[0]);
}
function nshObj_(r, me, acks) {
  const targets = nshTargets_(r[6]), mine = acks.filter(a => a[0] === r[0] && targets.indexOf(a[1]) !== -1);
  const o = { id: r[0], kind: r[1], number: r[2], date: r[3], title: r[4], source: r[5], aud: r[6], audLabel: NSH_AUD[r[6]] || "", note: r[7],
    fileName: r[9], mime: r[10], createdAt: r[11], targets, ackCount: mine.length };
  if (me.isHead) o.acks = mine.map(a => ({ teacher: a[1], at: a[2] }));
  else { const a = acks.find(x => x[0] === r[0] && x[1] === me.name); o.myAckAt = a ? a[2] : ""; o.forMe = targets.indexOf(me.name) !== -1; }
  return o;
}
function listCirculars(code) {
  const me = who_(code);
  if (me.isDemo) return [];
  const acks = nshAcks_();
  return nshRows_().map(r => nshObj_(r, me, acks)).filter(o => me.isHead || o.forMe)
    .sort((a, b) => (b.date + b.createdAt).localeCompare(a.date + a.createdAt));
}
function addCircular(code, c) {
  const me = head_(code);
  const kind = c.kind === "تعميم" ? "تعميم" : "نشرة";
  const title = String(c.title || "").trim().slice(0, 200); if (title.length < 3) throw new Error("اكتب عنوان " + kind + ".");
  const date = String(c.date || "").slice(0, 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error("اختر التاريخ.");
  const aud = NSH_AUD[c.aud] ? c.aud : "all";
  const mime = String(c.mime || ""); if (!/^(application\/pdf|image\/(jpeg|png))$/.test(mime)) throw new Error("الملف لازم يكون PDF أو صورة.");
  const data = String(c.data || "").replace(/^data:[^,]+,/, ""); if (!data) throw new Error("أرفق ملف " + kind + ".");
  if (data.length > 14000000) throw new Error("حجم الملف كبير (الحد ١٠ ميغابايت).");
  const id = newId_(), ext = mime === "application/pdf" ? "pdf" : mime === "image/png" ? "png" : "jpg";
  const name = String(c.fileName || "").trim().slice(0, 120) || `${kind} ${title}.${ext}`;
  const file = folder_().createFile(Utilities.newBlob(Utilities.base64Decode(data), mime, `${id} - ${name}`));
  withLock_(() => appendText_(nshSheet_(), [id, kind, String(c.number || "").trim().slice(0, 40), date, title, String(c.source || "").trim().slice(0, 100), aud,
    String(c.note || "").trim().slice(0, 1000), file.getId(), name, mime, nowText_()]));
  return { id };
}
function getCircularFile(code, id) {
  const me = who_(code), r = nshRows_().find(x => x[0] === String(id));
  if (!r || me.isDemo || (!me.isHead && nshTargets_(r[6]).indexOf(me.name) === -1)) throw new Error("النشرة غير موجودة.");
  const blob = DriveApp.getFileById(r[8]).getBlob();
  return { name: r[9], mime: r[10], data: Utilities.base64Encode(blob.getBytes()) };
}
function ackCircular(code, id, sig) {
  const me = who_(code), r = nshRows_().find(x => x[0] === String(id));
  if (!r || me.isDemo || me.isHead || nshTargets_(r[6]).indexOf(me.name) === -1) throw new Error("النشرة غير موجودة.");
  sig = String(sig || ""); if (sig && !validSig_(sig)) throw new Error("التوقيع غير صالح.");
  return withLock_(() => {
    const old = nshAcks_().find(a => a[0] === r[0] && a[1] === me.name);
    if (old) return { at: old[2] };
    const at = nowText_();
    appendText_(nshAckSheet_(), [r[0], me.name, at, sig]);
    if (sig) saveMySig_(me.name, sig, at);
    return { at };
  });
}
/** لرئيس الشعبة: كشف الاطلاع بالتواقيع */
function circularAcks(code, id) {
  head_(code);
  const r = nshRows_().find(x => x[0] === String(id)); if (!r) throw new Error("النشرة غير موجودة.");
  const acks = nshAcks_().filter(a => a[0] === r[0]);
  return nshTargets_(r[6]).map(t => { const a = acks.find(x => x[1] === t); return { teacher: t, at: a ? a[2] : "", sig: a ? a[3] : "" }; });
}
function deleteCircular(code, id) {
  head_(code);
  return withLock_(() => {
    const sh = nshSheet_(), row = findRow_(sh, 1, String(id)); if (row < 0) throw new Error("النشرة غير موجودة.");
    const fid = sh.getRange(row, 9).getDisplayValue();
    try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {}
    sh.deleteRow(row);
    const ak = nshAckSheet_(), v = ak.getDataRange().getDisplayValues();
    for (let i = v.length - 1; i >= 1; i--) if (v[i][0] === String(id)) ak.deleteRow(i + 1);
    return true;
  });
}


// ======================= إنجازات القسم (التقارير الشهرية) =======================
// يُعدّها Claude أول كل شهر عن الشهر السابق: تقرير لقسم الحديث الشريف وتقرير لقسم القرآن الكريم.
// الأعمدة: 0 المعرف | 1 الشهر (yyyy-MM) | 2 القسم (hd/qr) | 3 العنوان | 4 التقرير | 5 الإنجازات المشمولة (معرفات مفصولة بفاصلة) | 6 وقت الإعداد | 7 آخر تعديل
// في نص التقرير: سطر «[[صور:معرف الإنجاز]]» يضع صور ذلك الإنجاز في مكانه.
const MRP_HEAD = ["المعرف", "الشهر", "القسم", "العنوان", "التقرير", "الإنجازات المشمولة", "وقت الإعداد", "آخر تعديل"];
function mrpSheet_() { return sheet_(SHEET_MRP, MRP_HEAD, true); }
function listMonthlyReports(code) {
  head_(code);
  const ach = {}; rows_(SHEET_ACH).slice(1).filter(r => r[0]).forEach(r => ach[r[0]] = achObj_(r, false));
  return mrpSheet_().getDataRange().getDisplayValues().slice(1).filter(r => r[0]).map(r => ({ id: r[0], month: r[1], dept: r[2], title: r[3], report: r[4],
    items: (r[5] || "").split(",").map(x => x.trim()).filter(x => ach[x]).map(x => ({ id: x, title: ach[x].title, teacher: ach[x].teacher, date: ach[x].date, type: ach[x].type, kind: ach[x].kind, photos: ach[x].photos })),
    createdAt: r[6], updatedAt: r[7] })).sort((a, b) => (b.month + b.dept).localeCompare(a.month + a.dept));
}
function saveMonthlyReport(code, id, text) {
  head_(code);
  return withLock_(() => {
    const sh = mrpSheet_(), row = findRow_(sh, 1, String(id)); if (row < 0) throw new Error("التقرير غير موجود.");
    const at = nowText_();
    sh.getRange(row, 5).setValue(String(text || "").slice(0, 45000)); sh.getRange(row, 8).setValue(at);
    return { updatedAt: at };
  });
}
function deleteMonthlyReport(code, id) {
  head_(code);
  return withLock_(() => { const sh = mrpSheet_(), row = findRow_(sh, 1, String(id)); if (row < 0) throw new Error("التقرير غير موجود."); sh.deleteRow(row); return true; });
}


// ======================= الإشعارات: كل تحديث يسويه معلم يصل لرئيس الشعبة =======================
// أعمدة «الإشعارات»: 0 الوقت (yyyy-MM-dd HH:mm:ss) | 1 المعلم | 2 الركن | 3 ماذا فعل | 4 مفتاح الدمج | 5 المرجع (ما يُفتح عند الضغط على الإشعار)
// تُسجَّل الأحداث من doPost بعد نجاح الطلب. حفظ الدرجات والمنهج لنفس الصف خلال ساعة يُدمج في إشعار واحد.
const ACT_HEAD = ["الوقت", "المعلم", "الركن", "ماذا فعل", "مفتاح الدمج", "المرجع"];
function actSheet_() { const sh = sheet_(SHEET_ACTLOG, ACT_HEAD, true); if (sh.getLastColumn() < ACT_HEAD.length) sh.getRange(1, 1, 1, ACT_HEAD.length).setValues([ACT_HEAD]); return sh; }
function stamp_() { return Utilities.formatDate(new Date(), "Asia/Kuwait", "yyyy-MM-dd HH:mm:ss"); }
function arNum_(x) { return String(x).replace(/\d/g, d => "٠١٢٣٤٥٦٧٨٩"[d]); }
function clsName_(id) { const m = String(id || "").match(/^(hd|qr):(\d+)-(\d+)$/); return m ? `${SUBJ[m[1]]} ${arNum_(m[2])}/${arNum_(m[3])}` : String(id || ""); }
function rowTitle_(sheet, id, col) { const r = rows_(sheet).slice(1).find(x => x[0] === String(id)); return r ? r[col] : ""; }
const ACT_LOG = {
  saveEntry: a => ({ sec: "curr", text: `حدّث ما قطع من المنهج: ${clsName_(a[2] + ":" + a[1])}${a[5] ? " — " + a[5] : ""}`, key: "curr|" + a[2] + a[1], ref: a[1] + "__" + a[2] }),
  saveGrades: a => ({ sec: "cls", text: `رصد درجات كشف متابعة المتعلمين: ${clsName_(a[1])} (الفصل الدراسي ${a[2] == 2 ? "الثاني" : "الأول"})`, key: "grd|" + a[1] + "|" + a[2], ref: a[1] + "|" + (a[2] == 2 ? 2 : 1) }),
  saveStudentReport: (a, res) => ({ sec: "cls", ref: "rep|" + (res && res.id || a[1] && a[1].id || ""), text: `${a[1] && a[1].id ? "عدّل" : "أرسل"} تقرير ${SREP_TYPES[a[1] && a[1].type] || "متعلم"}: المتعلم رقم ${arNum_(a[1] && a[1].student || "")} (${clsName_(a[1] && a[1].cls)})` }),
  deleteStudentReport: () => ({ sec: "cls", text: "حذف تقرير متعلم" }),
  submitAchievement: (a, res) => ({ sec: a[1] && a[1].kind === "act" ? "act" : "ach", ref: res && res.id, text: `${a[1] && a[1].kind === "act" ? "رفع نشاطًا في سجل الأنشطة والمسابقات" : "رفع إنجازًا في التنمية المهنية وإنجاز الموظف"}: ${a[1] && a[1].title || ""}` }),
  deleteAchievement: () => ({ sec: "ach", text: "حذف إنجازًا أو نشاطًا مرفوعًا" }),
  editAchievement: (a, res) => ({ sec: res && res.kind === "act" ? "act" : "ach", ref: a[1], text: `عدّل ${res && res.kind === "act" ? "نشاطًا" : "إنجازًا"}: ${a[2] && a[2].title || ""}` }),
  signMinute: a => ({ sec: "min", ref: a[1], text: `اطّلع على المحضر: ${rowTitle_(SHEET_MIN, a[1], 1)}` }),
  signVisit: a => ({ sec: "vis", ref: a[1], text: "اطّلع على بطاقة الزيارة الصفية" }),
  submitPermission: (a, res) => ({ sec: "prm", ref: res && res.id || a[1] && a[1].id, text: `${a[1] && a[1].id ? "عدّل طلب استئذان" : "قدّم طلب استئذان"} (${a[1] && a[1].type || ""}) ليوم ${a[1] && a[1].date || ""}` }),
  deletePermission: () => ({ sec: "prm", text: "حذف طلب استئذان" }),
  saveRemedialPlan: a => ({ sec: "cls", ref: "rep|" + a[1], text: "حدّث الخطة العلاجية لمتعلم متعثر" }),
  requestSwap: a => ({ sec: "res", ref: "swap", text: `طلب تبديل حصة مع أ. ${a[1] && a[1].other || ""} ليوم ${a[1] && a[1].date || ""}` }),
  cancelSwap: () => ({ sec: "res", ref: "swap", text: "ألغى تبديل حصة" }),
  respondTask: (a, res) => { const r = tskRows_().find(x => x[0] === String(a[1])), t = r ? `«${r[1]}: ${r[2]}»` : "مهمة";
    return { sec: "tsk", ref: a[1], text: a[2] === "done" ? `نفّذ المهمة ${t}${res && res.note ? " — " + res.note : ""}` : a[2] === "undo" ? `تراجع عن تنفيذ المهمة ${t}` : `استلم المهمة ${t}` }; },
  tickAlert: (a, res) => { const r = alrRows_().find(x => x[0] === String(a[1])), it = r ? (JSON.parse(r[2] || "[]")[a[2]] || {}).t : "";
    return { sec: "alr", ref: a[1], text: `${a[3] ? "استكمل" : "ألغى استكمال"}: ${it || "نقص"}${res && res.all ? " — واكتملت جميع النواقص ✓" : ""}` }; },
  saveMyPhone: (a, res) => ({ sec: "wa", text: `سجّل رقم الواتساب${res && res.phone ? ": +" + res.phone : ""}`, key: "wa" }),
  approveMyReport: (a, res) => ({ sec: res && res.kind === "act" ? "act" : "ach", ref: a[1], text: `اعتمد التقرير المصاغ: ${res && res.title || ""} (جاهز للطباعة والتوقيع)` }),
  ackCircular: a => { const r = nshRows_().find(x => x[0] === String(a[1])); return { sec: "nsh", ref: a[1], text: `اطّلع على ${r ? r[1] + ": " + r[4] : "نشرة"}` }; }
};
function logAct_(fn, args, result) {
  const me = who_(args[0]); if (me.isHead) return;
  const ev = ACT_LOG[fn](args, result); if (!ev) return;
  const now = stamp_(), key = ev.key ? me.name + "|" + ev.key : "";
  withLock_(() => {
    const sh = actSheet_();
    if (key) {
      const last = sh.getLastRow(), from = Math.max(2, last - 30);
      if (last >= 2) {
        const v = sh.getRange(from, 1, last - from + 1, 6).getDisplayValues();
        for (let i = v.length - 1; i >= 0; i--) if (v[i][4] === key && (new Date(now.replace(" ", "T")) - new Date(v[i][0].replace(" ", "T"))) < 3600000) {
          sh.getRange(from + i, 1, 1, 6).setValues([[now, me.name, ev.sec, ev.text, key, String(ev.ref || "")]]); return;
        }
      }
    }
    appendText_(sh, [now, me.name, ev.sec, ev.text, key, String(ev.ref || "")]);
  });
}
/** لرئيس الشعبة: آخر الإشعارات، وعدد غير المقروء */
function listActivity(code, limit) {
  head_(code); try { ensureReminders_(); } catch (e) {}
  const seen = PropertiesService.getScriptProperties().getProperty("actSeen") || "";
  const v = actSheet_().getDataRange().getDisplayValues().slice(1).filter(r => r[0]);
  const items = v.slice(-Math.min(Number(limit) || 150, 300)).reverse().map(r => ({ at: r[0], teacher: r[1], sec: r[2], text: r[3], ref: r[5] || "", demo: isDemo_(r[1]), unread: r[0] > seen }));
  let secSeen = {}; try { secSeen = JSON.parse(PropertiesService.getScriptProperties().getProperty("actSeenSec") || "{}"); } catch (e) {}
  return { items, unread: v.filter(r => r[0] > seen).length, seen, secSeen };
}
/** اللمبة الزرقاء: فتح الركن يصفّر عدّاد تحديثاته فقط */
const ACT_SECS = ["curr", "cls", "ach", "act", "min", "vis", "prm", "nsh", "res", "tsk", "alr", "wa"];
function markSectionSeen(code, sec) {
  head_(code);
  if (ACT_SECS.indexOf(sec) === -1) throw new Error("ركن غير معروف.");
  const t = stamp_(), props = PropertiesService.getScriptProperties();
  return withLock_(() => {
    let m = {}; try { m = JSON.parse(props.getProperty("actSeenSec") || "{}"); } catch (e) {}
    m[sec] = t; props.setProperty("actSeenSec", JSON.stringify(m));
    return { sec, seen: t };
  });
}
function markActivitySeen(code) { head_(code); const t = stamp_(); PropertiesService.getScriptProperties().setProperty("actSeen", t); return { seen: t }; }


// ======================= المهام والتكاليف =======================
// يكلّف رئيس الشعبة معلمًا أو أكثر بمهمة (نشاط، إذاعة، ورشة…)، فتصل للمعلم في الموقع بلمبة زرقاء،
// يُطبع خطاب التكليف في القسم، ويؤكد المعلم الاستلام ثم يسجّل التنفيذ، وكل ذلك يصل لرئيس الشعبة في الإشعارات.
// «المهام والتكاليف»: 0 المعرف | 1 النوع | 2 العنوان | 3 التفاصيل | 4 التاريخ | 5 الحصة | 6 الوقت | 7 المكان | 8 الفئة المستهدفة
//                     9 آخر موعد للتنفيذ | 10 المكلفون (مفصولون بـ |) | 11 الملف | 12 اسم الملف | 13 نوع الملف | 14 وقت الإنشاء | 15 آخر تعديل
//                     16 التكرار ("" = مرة واحدة، "weekly" = مهمة ثابتة كل أسبوع في يوم التاريخ نفسه)
// المهمة الثابتة: حالة المعلم تُحفظ لكل أسبوع بمفتاح «المعرف@تاريخ الموعد»، وتظهر له لمبة التذكير من اليوم السابق للموعد حتى يومه.
// «حالة التكاليف»:    0 المهمة | 1 المعلم | 2 وقت الاستلام | 3 وقت التنفيذ | 4 ملاحظة المعلم | 5 الحالة
const SHEET_TSK = "المهام والتكاليف";
const SHEET_TSKST = "حالة التكاليف";
const TSK_HEAD = ["المعرف", "النوع", "العنوان", "التفاصيل", "التاريخ", "الحصة", "الوقت", "المكان", "الفئة المستهدفة", "آخر موعد للتنفيذ",
  "المكلفون", "الملف", "اسم الملف", "نوع الملف", "وقت الإنشاء", "آخر تعديل", "التكرار", "السجل المطلوب"];
const TSKST_HEAD = ["المهمة", "المعلم", "وقت الاستلام", "وقت التنفيذ", "ملاحظة المعلم", "الحالة"];
const TSK_TYPES = ["نشاط", "إذاعة مدرسية", "ورشة تدريبية", "درس تطبيقي", "مسابقة", "لجنة", "مناوبة", "إعداد تقرير أو ملف", "تنظيم فعالية", "أخرى"];
const TSK_NEW = "جديدة", TSK_GOT = "تم الاستلام", TSK_DONE = "تم التنفيذ";
function tskSheet_() {
  const sh = sheet_(SHEET_TSK, TSK_HEAD, true);
  if (sh.getLastColumn() < TSK_HEAD.length) sh.getRange(1, 1, 1, TSK_HEAD.length).setValues([TSK_HEAD]);
  return sh;
}
function isoAdd_(iso, n) { const m = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/), d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3] + n)); return d.toISOString().slice(0, 10); }
/** موعد المهمة الثابتة هذا الأسبوع: اليوم إن كان يومها، وإلا أقرب يوم قادم */
function tskOcc_(r) { const t = todayIso_(); if (r[16] !== "weekly") return r[4]; if (t < r[4]) return r[4]; return isoAdd_(t, (dayOf_(r[4]) - dayOf_(t) + 7) % 7); }
function tskKey_(r) { return r[16] === "weekly" ? r[0] + "@" + tskOcc_(r) : r[0]; }
/** التذكير: التكليف العادي قبل موعده بيومين حتى يومه، والمهمة الثابتة في يومها فقط */
function tskRemind_(r, occ) {
  const t = todayIso_();
  if (r[16] === "weekly") return { soon: occ === t, remind: occ === t };
  return { soon: true, remind: r[4] >= t && r[4] <= isoAdd_(t, 2) };
}
/** مرة واحدة يوميًا (عند أول استخدام للموقع في اليوم): يسجّل لرئيس الشعبة إشعار «تم تذكير المعلم» لكل تذكير يبدأ اليوم */
function ensureReminders_() {
  const props = PropertiesService.getScriptProperties(), t = todayIso_();
  if (props.getProperty("remindDay") === t) return;
  withLock_(() => {
    if (props.getProperty("remindDay") === t) return;
    props.setProperty("remindDay", t);
    const sts = tskSts_(), sh = actSheet_(), now = stamp_(), DN = ["الأحد", "الاثنين", "الثلاثاء", "الأربعاء", "الخميس", "الجمعة", "السبت"];
    const dd = iso => DN[dayOf_(iso)] + " " + arNum_(iso.slice(8, 10) + "/" + iso.slice(5, 7));
    tskRows_().forEach(r => {
      const weekly = r[16] === "weekly", occ = tskOcc_(r);
      if (weekly ? occ !== t : r[4] !== isoAdd_(t, 2)) return;
      const key = tskKey_(r);
      r[10].split("|").filter(Boolean).forEach(n => {
        const s = sts.find(x => x[0] === key && x[1] === n);
        if (s && s[5] === TSK_DONE) return;
        const text = weekly ? `⏰ تم تذكيره بمهمة اليوم: ${r[2]}` : `⏰ تم تذكيره: باقي يومين على تكليف «${r[2]}» (${dd(r[4])})`;
        appendText_(sh, [now, n, "tsk", text, "", r[0]]);
      });
    });
  });
}
function tskMine_(key, id) { return key === id || key.indexOf(id + "@") === 0; }
function tskStSheet_() { return sheet_(SHEET_TSKST, TSKST_HEAD, true); }
function tskRows_() { return tskSheet_().getDataRange().getDisplayValues().slice(1).filter(r => r[0]); }
function tskSts_() { return tskStSheet_().getDataRange().getDisplayValues().slice(1).filter(r => r[0]); }
/** من يمكن تكليفه: معلمو الشعبة (بدون الرئيس) والحسابات التجريبية للتجربة */
function tskPool_() { return TEACHERS.filter(t => !t[2]).map(t => t[0]).concat(Object.keys(DEMOS)); }
function tskObj_(r, sts, me) {
  const who = r[10].split("|").filter(Boolean), key = tskKey_(r), occ = tskOcc_(r), weekly = r[16] === "weekly";
  const st = who.map(t => { const s = sts.find(x => x[0] === key && x[1] === t) || [];
    return { teacher: t, demo: isDemo_(t), gotAt: s[2] || "", doneAt: s[3] || "", note: s[4] || "", status: s[5] || TSK_NEW }; });
  const o = { id: r[0], type: r[1], title: r[2], details: r[3], date: r[4], period: r[5], time: r[6], place: r[7], target: r[8], due: r[9],
    fileName: r[12], mime: r[13], hasFile: !!r[11], createdAt: r[14], updatedAt: r[15], teachers: who,
    repeat: weekly ? "weekly" : "", rec: r[17] || "", occ, ...tskRemind_(r, occ) };
  if (me.isHead) o.status = st; else o.mine = st.find(x => x.teacher === me.name) || null;
  return o;
}
function listTasks(code) {
  const me = who_(code); try { ensureReminders_(); } catch (e) {}
  const sts = tskSts_();
  return tskRows_().map(r => tskObj_(r, sts, me)).filter(o => me.isHead ? !me.isDemo : !!o.mine)
    .sort((a, b) => (b.date + b.createdAt).localeCompare(a.date + a.createdAt));
}
function saveTask(code, t) {
  const me = head_(code);
  if (me.isDemo) throw new Error("غير متاح في الحساب التجريبي.");
  const type = TSK_TYPES.indexOf(t.type) !== -1 ? t.type : "أخرى";
  const title = String(t.title || "").trim().slice(0, 200); if (title.length < 3) throw new Error("اكتب عنوان المهمة.");
  const date = String(t.date || "").slice(0, 10); if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error("اختر تاريخ المهمة.");
  const due = String(t.due || "").slice(0, 10); if (due && !/^\d{4}-\d{2}-\d{2}$/.test(due)) throw new Error("آخر موعد غير صحيح.");
  const pool = tskPool_(), who = (Array.isArray(t.teachers) ? t.teachers : []).map(String).filter((x, i, a) => pool.indexOf(x) !== -1 && a.indexOf(x) === i);
  if (!who.length) throw new Error("اختر المعلم المكلَّف.");
  const s = (v, n) => String(v || "").trim().slice(0, n);
  const row = [s(t.id, 20) || newId_(), type, title, s(t.details, 3000), date, s(t.period, 20), s(t.time, 40), s(t.place, 100), s(t.target, 150), t.repeat === "weekly" ? "" : due, who.join("|"), "", "", "", "", "", t.repeat === "weekly" ? "weekly" : "", ACT_SECS.indexOf(t.rec) !== -1 ? t.rec : ""];
  if (row[16] && dayOf_(date) > 4) throw new Error("يوم المهمة الثابتة عطلة (جمعة أو سبت).");
  let file = null;
  if (t.data) {
    const mime = String(t.mime || ""); if (!/^(application\/pdf|image\/(jpeg|png))$/.test(mime)) throw new Error("المرفق لازم يكون PDF أو صورة.");
    const data = String(t.data).replace(/^data:[^,]+,/, ""); if (data.length > 14000000) throw new Error("حجم المرفق كبير (الحد ١٠ ميغابايت).");
    const name = s(t.fileName, 120) || "مرفق";
    file = { id: folder_().createFile(Utilities.newBlob(Utilities.base64Decode(data), mime, `${row[0]} - ${name}`)).getId(), name, mime };
  }
  return withLock_(() => {
    const sh = tskSheet_(), at = nowText_(), idx = t.id ? findRow_(sh, 1, row[0]) : -1;
    if (t.id && idx < 0) throw new Error("المهمة غير موجودة.");
    if (idx > 0) {
      const old = sh.getRange(idx, 1, 1, TSK_HEAD.length).getDisplayValues()[0];
      if (file) { if (old[11]) { try { DriveApp.getFileById(old[11]).setTrashed(true); } catch (e) {} } row[11] = file.id; row[12] = file.name; row[13] = file.mime; }
      else if (t.removeFile) { if (old[11]) { try { DriveApp.getFileById(old[11]).setTrashed(true); } catch (e) {} } }
      else { row[11] = old[11]; row[12] = old[12]; row[13] = old[13]; }
      row[14] = old[14]; row[15] = at;
      writeRow_(sh, idx, row);
      // من أُزيل من التكليف تُحذف حالته، والمهمة المعدّلة تعود «جديدة» عند الباقين حتى يطّلعوا على التعديل
      const st = tskStSheet_(), v = st.getDataRange().getDisplayValues();
      for (let i = v.length - 1; i >= 1; i--) if (tskMine_(v[i][0], row[0])) {
        if (who.indexOf(v[i][1]) === -1) st.deleteRow(i + 1);
        else if (v[i][5] === TSK_GOT) writeRow_(st, i + 1, [v[i][0], v[i][1], "", "", "", TSK_NEW]);
      }
    } else {
      if (file) { row[11] = file.id; row[12] = file.name; row[13] = file.mime; }
      row[14] = at; row[15] = at; appendText_(sh, row);
    }
    return { id: row[0] };
  });
}
function deleteTask(code, id) {
  head_(code);
  return withLock_(() => {
    const sh = tskSheet_(), row = findRow_(sh, 1, String(id)); if (row < 0) throw new Error("المهمة غير موجودة.");
    const fid = sh.getRange(row, 12).getDisplayValue(); if (fid) { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} }
    sh.deleteRow(row);
    const st = tskStSheet_(), v = st.getDataRange().getDisplayValues();
    for (let i = v.length - 1; i >= 1; i--) if (tskMine_(v[i][0], String(id))) st.deleteRow(i + 1);
    return true;
  });
}
function getTaskFile(code, id) {
  const me = who_(code), r = tskRows_().find(x => x[0] === String(id));
  if (!r || !r[11] || (!(me.isHead && !me.isDemo) && r[10].split("|").indexOf(me.name) === -1)) throw new Error("المرفق غير موجود.");
  return { name: r[12], mime: r[13], data: Utilities.base64Encode(DriveApp.getFileById(r[11]).getBlob().getBytes()) };
}
/** المعلم: act = got (استلمت) · done (تم التنفيذ مع ملاحظة) · undo. لا اعتذار في الموقع: يكتبه المعلم يدويًا على خطاب التكليف المطبوع */
function respondTask(code, id, act, note) {
  const me = who_(code);
  if (me.isHead) throw new Error("هذا الإجراء للمعلم المكلَّف.");
  const r = tskRows_().find(x => x[0] === String(id));
  if (!r || r[10].split("|").indexOf(me.name) === -1) throw new Error("المهمة غير موجودة.");
  note = String(note || "").trim().slice(0, 1000);
  return withLock_(() => {
    const st = tskStSheet_(), v = st.getDataRange().getDisplayValues(), at = nowText_();
    const key = tskKey_(r);
    let i = v.findIndex((x, k) => k > 0 && x[0] === key && x[1] === me.name);
    const cur = i > 0 ? v[i] : [key, me.name, "", "", "", TSK_NEW];
    if (act === "got") { if (!cur[2]) cur[2] = at; if (cur[5] === TSK_NEW) cur[5] = TSK_GOT; }
    else if (act === "done") { if (!cur[2]) cur[2] = at; cur[3] = at; cur[4] = note; cur[5] = TSK_DONE; }
    else if (act === "undo") { cur[3] = ""; cur[4] = ""; cur[5] = cur[2] ? TSK_GOT : TSK_NEW; }
    else throw new Error("إجراء غير معروف.");
    if (i > 0) writeRow_(st, i + 1, cur); else appendText_(st, cur);
    return { gotAt: cur[2], doneAt: cur[3], note: cur[4], status: cur[5] };
  });
}

/** مرة واحدة: المهمة الثابتة «مراقبة ومناوبة القسمين» كل ثلاثاء لجميع معلمي الشعبة (لا تتكرر إن وُجدت) */
function addTuesdayDuty() {
  const id = "tue-duty";
  if (tskRows_().some(r => r[0] === id)) return "موجودة مسبقًا";
  const t = todayIso_(), tue = isoAdd_(t, (2 - dayOf_(t) + 7) % 7), at = nowText_();
  const who = TEACHERS.filter(x => !x[2]).map(x => x[0]);
  withLock_(() => appendText_(tskSheet_(), [id, "مناوبة", "مراقبة ومناوبة القسمين", "مهمة ثابتة كل يوم ثلاثاء: مراقبة ومناوبة قسمي الحديث الشريف والقرآن الكريم.",
    tue, "", "", "", "", "", who.join("|"), "", "", "", at, at, "weekly"]));
  return "أُضيفت ليوم " + tue + " لـ " + who.length + " معلمين";
}


// ======================= أرقام واتساب المعلمين =======================
// يكتب كل معلم رقمه في الموقع، ويستخدمه رئيس الشعبة لإرسال التكليف والتنبيه واتساب بضغطة (wa.me) دون أي اشتراك.
const SHEET_WA = "واتساب المعلمين";
function waSheet_() { return sheet_(SHEET_WA, ["المعلم", "رقم الواتساب", "آخر تحديث"], true); }
function phones_() { const out = {}; const sh = sh_(SHEET_WA); if (sh) sh.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[0] && r[1]) out[r[0]] = r[1]; }); return out; }
function cleanPhone_(v) {
  let d = String(v || "").replace(/[٠-٩]/g, x => "٠١٢٣٤٥٦٧٨٩".indexOf(x)).replace(/[^\d]/g, "").replace(/^00/, "");
  if (/^[24569]\d{7}$/.test(d)) d = "965" + d; // رقم كويتي بدون مفتاح الدولة
  if (!/^\d{10,15}$/.test(d)) throw new Error("رقم الواتساب غير صحيح. اكتبه هكذا: 9xxxxxxx");
  return d;
}
function saveMyPhone(code, phone) {
  const me = who_(code), p = cleanPhone_(phone);
  withLock_(() => { const sh = waSheet_(), row = findRow_(sh, 1, me.name), vals = [me.name, p, nowText_()]; if (row > 0) writeRow_(sh, row, vals); else appendText_(sh, vals); });
  return { phone: p };
}
/** لرئيس الشعبة: أرقام المعلمين */
function teacherPhones(code) { head_(code); return phones_(); }
/** لرئيس الشعبة: من سجّل رقمه ومتى */
function phonesList(code) {
  head_(code); const sh = sh_(SHEET_WA), rows = sh ? sh.getDataRange().getDisplayValues().slice(1) : [];
  return tskPool_().map(t => { const r = rows.find(x => x[0] === t); return { teacher: t, demo: isDemo_(t), phone: r ? r[1] : "", at: r ? r[2] : "" }; });
}


// ======================= تنبيهات الاستكمال =======================
// رئيس الشعبة يختار معلمًا أو أكثر وعدة نواقص، فيعلّم المعلم على كل نقص «تم» على حدة.
// «تنبيهات الاستكمال»: 0 المعرف | 1 المعلمون (مفصولون بـ |) | 2 النواقص JSON [{k: الركن أو "", t: النص}] | 3 ملاحظة | 4 آخر موعد | 5 وقت الإنشاء | 6 آخر تعديل
// «حالة الاستكمال»:    0 التنبيه | 1 المعلم | 2 المستكمل JSON {رقم النقص: الوقت} | 3 وقت اكتمال الكل | 4 أول اطلاع
const SHEET_ALR = "تنبيهات الاستكمال", SHEET_ALRST = "حالة الاستكمال";
function alrSheet_() { return sheet_(SHEET_ALR, ["المعرف", "المعلمون", "النواقص", "ملاحظة", "آخر موعد", "وقت الإنشاء", "آخر تعديل"], true); }
function alrStSheet_() { return sheet_(SHEET_ALRST, ["التنبيه", "المعلم", "المستكمل", "اكتمل الكل", "أول اطلاع"], true); }
function alrRows_() { return alrSheet_().getDataRange().getDisplayValues().slice(1).filter(r => r[0]); }
function alrSts_() { return alrStSheet_().getDataRange().getDisplayValues().slice(1).filter(r => r[0]); }
function jsonOr_(s, d) { try { return JSON.parse(s); } catch (e) { return d; } }
function alrObj_(r, sts, me) {
  const who = r[1].split("|").filter(Boolean), items = jsonOr_(r[2], []);
  const st = who.map(t => { const x = sts.find(y => y[0] === r[0] && y[1] === t) || []; const done = jsonOr_(x[2] || "{}", {});
    return { teacher: t, demo: isDemo_(t), done, doneCount: Object.keys(done).length, allAt: x[3] || "" }; });
  const o = { id: r[0], teachers: who, items, note: r[3], due: r[4], createdAt: r[5], updatedAt: r[6] };
  if (me.isHead) o.status = st; else o.mine = st.find(x => x.teacher === me.name) || null;
  return o;
}
function listAlerts(code) {
  const me = who_(code), sts = alrSts_();
  return alrRows_().map(r => alrObj_(r, sts, me)).filter(o => me.isHead ? !me.isDemo : !!o.mine).sort((a, b) => b.createdAt.localeCompare(a.createdAt));
}
function saveAlert(code, a) {
  const me = head_(code); if (me.isDemo) throw new Error("غير متاح في الحساب التجريبي.");
  const pool = tskPool_(), who = (Array.isArray(a.teachers) ? a.teachers : []).map(String).filter((x, i, l) => pool.indexOf(x) !== -1 && l.indexOf(x) === i);
  if (!who.length) throw new Error("اختر المعلم.");
  const items = (Array.isArray(a.items) ? a.items : []).slice(0, 40).map(x => ({ k: ACT_SECS.indexOf(x && x.k) !== -1 ? x.k : "", t: String(x && x.t || "").trim().slice(0, 200) })).filter(x => x.t);
  if (!items.length) throw new Error("اختر نقصًا واحدًا على الأقل.");
  const due = String(a.due || "").slice(0, 10); if (due && !/^\d{4}-\d{2}-\d{2}$/.test(due)) throw new Error("آخر موعد غير صحيح.");
  const note = String(a.note || "").trim().slice(0, 1500);
  return withLock_(() => {
    const sh = alrSheet_(), at = nowText_();
    if (a.id) {
      const row = findRow_(sh, 1, String(a.id)); if (row < 0) throw new Error("التنبيه غير موجود.");
      const old = sh.getRange(row, 1, 1, 7).getDisplayValues()[0];
      writeRow_(sh, row, [old[0], who.join("|"), JSON.stringify(items), note, due, old[5], at]);
      // تعديل النواقص يعيد المتابعة من جديد
      const st = alrStSheet_(), v = st.getDataRange().getDisplayValues();
      for (let i = v.length - 1; i >= 1; i--) if (v[i][0] === old[0]) st.deleteRow(i + 1);
      return { id: old[0] };
    }
    const id = newId_(); appendText_(sh, [id, who.join("|"), JSON.stringify(items), note, due, at, at]); return { id };
  });
}
function deleteAlert(code, id) {
  head_(code);
  return withLock_(() => {
    const sh = alrSheet_(), row = findRow_(sh, 1, String(id)); if (row < 0) throw new Error("التنبيه غير موجود.");
    sh.deleteRow(row);
    const st = alrStSheet_(), v = st.getDataRange().getDisplayValues();
    for (let i = v.length - 1; i >= 1; i--) if (v[i][0] === String(id)) st.deleteRow(i + 1);
    return true;
  });
}
/** المعلم يعلّم نقصًا واحدًا: تم أو لم يتم */
function tickAlert(code, id, idx, done) {
  const me = who_(code); if (me.isHead) throw new Error("هذا للمعلم المنبَّه.");
  const r = alrRows_().find(x => x[0] === String(id));
  if (!r || r[1].split("|").indexOf(me.name) === -1) throw new Error("التنبيه غير موجود.");
  const items = jsonOr_(r[2], []); idx = Math.floor(Number(idx)); if (!(idx >= 0 && idx < items.length)) throw new Error("النقص غير موجود.");
  return withLock_(() => {
    const st = alrStSheet_(), v = st.getDataRange().getDisplayValues(), at = nowText_();
    const i = v.findIndex((x, k) => k > 0 && x[0] === r[0] && x[1] === me.name);
    const cur = i > 0 ? v[i].slice(0, 5) : [r[0], me.name, "{}", "", at];
    const d = jsonOr_(cur[2] || "{}", {});
    if (done) d[idx] = at; else delete d[idx];
    cur[2] = JSON.stringify(d); const all = Object.keys(d).length >= items.length; cur[3] = all ? (cur[3] || at) : "";
    if (i > 0) writeRow_(st, i + 1, cur); else appendText_(st, cur);
    return { done: d, all, allAt: cur[3] };
  });
}
