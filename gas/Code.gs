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
const API = { login, saveEntry, headData, submitAchievement, myAchievements, deleteAchievement, allAchievements, saveReport,
  getPhoto, listMinutes, getMinute, saveMinute, deleteMinute, signMinute, mySignature, saveFileToDrive, demoAccounts, demoAction,
  saveVisit, listVisits, getVisit, signVisit, deleteVisit, reserveData, saveAbsence, deleteAbsence, classesOverview, classSheet, saveGrades,
  listStudentReports, saveStudentReport, seenStudentReport, deleteStudentReport,
  myProfile, saveProfile, listPermissions, getPermission, submitPermission, mailPermissionPdf, decidePermission, deletePermission,
  listCirculars, addCircular, getCircularFile, ackCircular, circularAcks, deleteCircular,
  listMonthlyReports, saveMonthlyReport, deleteMonthlyReport, listActivity, markActivitySeen };

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
  sheet_(SHEET_COV, ["المعرف", "معرف الغياب", "التاريخ", "الحصة", "الفصل", "المعلم الغائب", "المعلم البديل", "الحالة", "تاريخ التسجيل"], true);
  sheet_(SHEET_STU, ["الصف", "م", "اسم الطالب", "ملاحظة"], true);
  gradesSheet_();
  srepSheet_();
  profSheet_();
  permSheet_();
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
  return { name: me.name, classes: me.classes, isHead: me.isHead, isDemo: me.isDemo, entries: mine, all: me.isHead ? all : null };
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
    kind: r[12] === "act" ? "act" : "pd", results: r[13] || "", dept: r[14] === "qr" || r[14] === "hd" ? r[14] : achDept_(r[1]) };
  if (withReport) o.report = r[9];
  return o;
}

/** قسم الإنجاز إن لم يُحدَّد: من مادة المعلم (إن كان يدرّس مادة واحدة)، وإلا الحديث الشريف */
function achDept_(name) { const cl = classesOf_(name) || []; const d = cl.map(c => c.split(":")[0]); return d.length && d.every(x => x === "qr") ? "qr" : "hd"; }
function submitAchievement(code, a) {
  const me = who_(code);
  const title = String(a.title || "").trim().slice(0, 150);
  const desc = String(a.desc || "").trim().slice(0, 3000);
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
  return rows_(SHEET_ACH).slice(1).filter(r => r[0] && r[1] === me.name).map(r => achObj_(r, false)).reverse();
}

function deleteAchievement(code, id) {
  const me = who_(code);
  return withLock_(() => {
    const sh = sh_(SHEET_ACH), row = findRow_(sh, 1, id);
    if (row < 0) throw new Error("النشاط غير موجود.");
    const r = sh.getRange(row, 1, 1, 12).getDisplayValues()[0];
    if (!me.isHead && (r[1] !== me.name || r[8] === ST_READY)) throw new Error("لا يمكن حذف هذا النشاط.");
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
  image = String(image || "");
  if (!validSig_(image)) throw new Error("التوقيع غير صالح، ارسمه مرة أخرى.");
  const r = rows_(SHEET_MIN).slice(1).find(x => x[0] === id);
  if (!r || (me.isDemo !== (id === DEMO_MINUTE))) throw new Error("المحضر غير موجود.");
  if (minObj_(r).attendees.indexOf(me.name) === -1) throw new Error("اسمك ليس ضمن حضور هذا الاجتماع.");
  return withLock_(() => {
    const sh = sh_(SHEET_SIG), vals = sh.getDataRange().getDisplayValues();
    for (let i = 1; i < vals.length; i++) if (vals[i][0] === id && vals[i][1] === me.name) throw new Error("وقّعت على هذا المحضر من قبل.");
    const at = nowText_();
    appendText_(sh, [id, me.name, at, image]);
    saveMySig_(me.name, image, at);
    // عند اكتمال توقيع جميع الحضور يصل إيميل لرئيس الشعبة
    const m = minObj_(r), signed = vals.slice(1).filter(v => v[0] === id).map(v => [v[1], v[2]]).concat([[me.name, at]]);
    if (m.attendees.every(n => signed.some(x => x[0] === n))) {
      notify_(me, `اكتمل التوقيع على محضر: ${m.title}`, `
        <p>وقّع جميع الحضور (${m.attendees.length}) على محضر <b>${esc_(m.title)}</b> بتاريخ ${esc_(m.date)}.</p>
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
      if (old[15]) throw new Error("وقّع المعلم على هذه الزيارة، فلا يمكن تعديلها.");
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
  if (!validSig_(image)) throw new Error("التوقيع غير صالح، ارسمه مرة أخرى.");
  return withLock_(() => {
    const sh = sh_(SHEET_VIS), row = findRow_(sh, 1, id);
    if (row < 0) throw new Error("الزيارة غير موجودة.");
    const r = sh.getRange(row, 1, 1, VIS_COLS).getDisplayValues()[0];
    if (r[1] !== me.name) throw new Error("هذه الزيارة ليست لك.");
    if (r[15]) throw new Error("وقّعت على هذه الزيارة من قبل.");
    const at = nowText_();
    r[14] = image; r[15] = at;
    writeRow_(sh, row, r);
    saveMySig_(me.name, image, at);
    notify_(me, `وقّع أ. ${me.name} على الزيارة الصفية رقم ${r[8]}`, `
      <p>اطّلع <b>أ. ${esc_(me.name)}</b> على الزيارة الصفية رقم ${esc_(r[8])} (${esc_(r[3])}) بتاريخ ${esc_(r[4])}، ووقّع عليها في ${esc_(at)}.</p>
      <p>يمكنك تنزيل تقرير الزيارة موقّعًا من ركن «الزيارات الصفية».</p>`);
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
// الفصل "R" = حصة احتياط في الجدول، "RF" = احتياط فنية (المعلم متاح ومرشّح أول للتغطية)
// يجب أن يطابق TIMETABLE في index.html
const TIMETABLE = {
  "خليفة حمد خليفة القصار": [[3,2,"hd:11-1"],[4,5,"hd:11-1"]],
  "أحمد صادق أحمد أبوسعدة": [[0,1,"hd:11-4"],[0,2,"hd:11-5"],[1,3,"hd:11-3"],[1,6,"hd:11-2"],[2,2,"hd:11-3"],[2,5,"RF"],[2,6,"hd:11-5"],[3,4,"hd:11-2"],[3,6,"R"],[4,4,"RF"],[4,7,"hd:11-4"]],
  "أحمد صلاح زيدان القناوي": [[0,1,"hd:9-1"],[1,2,"hd:9-5"],[1,5,"hd:9-1"],[2,6,"R"],[3,1,"hd:9-2"],[4,4,"hd:9-2"],[4,5,"hd:9-5"]],
  "عماد صادق علي العايق": [[0,5,"hd:12-1"],[1,4,"hd:12-3"],[1,7,"hd:12-1"],[2,3,"hd:10-5"],[2,6,"hd:12-3"],[3,7,"hd:10-3"],[4,1,"hd:10-5"],[4,2,"hd:10-3"]],
  "كرم عبدالحميد أبوزيد أحمد": [[0,2,"hd:12-4"],[0,4,"hd:12-5"],[1,2,"hd:10-1"],[2,5,"hd:12-4"],[2,7,"hd:10-2"],[3,1,"hd:10-1"],[3,4,"hd:10-4"],[3,7,"hd:12-5"],[4,3,"hd:10-2"],[4,6,"hd:10-4"]],
  "منذر خالد مطلق الحربي": [[0,1,"qr:12-2"],[0,3,"qr:12-3"],[0,4,"hd:12-2"],[0,6,"qr:12-1"],[1,2,"qr:12-4"],[1,4,"hd:9-4"],[1,7,"qr:12-5"],[2,2,"qr:12-5"],[2,5,"qr:12-2"],[2,7,"hd:9-3"],[3,3,"hd:12-2"],[3,5,"hd:9-4"],[3,6,"qr:12-4"],[4,2,"qr:12-3"],[4,4,"qr:12-1"],[4,6,"hd:9-3"]],
  "معاذ وليد مبارك الشطي": [[0,4,"qr:9-2"],[0,7,"qr:9-5"],[1,3,"qr:9-4"],[1,6,"qr:9-3"],[2,3,"qr:9-1"],[2,7,"qr:9-2"],[3,4,"qr:9-3"],[3,6,"qr:9-1"],[4,1,"qr:9-4"],[4,2,"qr:9-5"]],
  "يوسف عدنان عبدالكريم الشطي": [[0,4,"qr:10-2"],[0,6,"qr:10-3"],[1,3,"qr:10-3"],[1,5,"qr:10-5"],[2,5,"qr:10-1"],[2,7,"qr:10-4"],[3,4,"qr:10-5"],[3,7,"qr:10-2"],[4,1,"qr:10-1"],[4,3,"qr:10-4"]],
  "عبدالله أحمد عبدالمجد الماجد": [[0,2,"qr:11-2"],[0,4,"qr:11-1"],[1,5,"qr:11-3"],[1,6,"R"],[2,2,"qr:11-5"],[2,3,"qr:11-1"],[2,6,"RF"],[3,1,"qr:11-3"],[3,3,"qr:11-4"],[3,6,"qr:11-5"],[4,5,"qr:11-4"],[4,6,"qr:11-2"]]
};
const COV_DONE = "مغطاة", COV_NONE = "غير مغطاة", COV_PRESENT = "حضر المعلم";

function dayOf_(iso) {
  const m = String(iso || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error("التاريخ غير صحيح.");
  return new Date(+m[1], +m[2] - 1, +m[3], 12).getDay(); // 0 الأحد … 6 السبت
}
function lessonAt_(teacher, day, period) {
  return (TIMETABLE[teacher] || []).find(x => x[0] === day && x[1] === period) || null;
}
function absObj_(r) { return { id: r[0], date: r[1], teacher: r[2], note: r[3], createdAt: r[4], updatedAt: r[5] }; }
function covObj_(r) { return { id: r[0], absenceId: r[1], date: r[2], period: Number(r[3]), classId: r[4], absent: r[5], substitute: r[6], status: r[7] }; }

/** كل بيانات الغياب والاحتياط */
function reserveData(code) {
  head_(code);
  return { absences: rows_(SHEET_ABS).slice(1).filter(r => r[0]).map(absObj_), covers: rows_(SHEET_COV).slice(1).filter(r => r[0]).map(covObj_) };
}

/** تسجيل غياب معلم (أو تعديله) مع توزيع حصصه على البدلاء */
function saveAbsence(code, a) {
  head_(code);
  const teacher = String(a.teacher || ""), date = String(a.date || "").slice(0, 10);
  if (!TIMETABLE[teacher]) throw new Error("اختر المعلم الغائب.");
  const day = dayOf_(date);
  if (day > 4) throw new Error("التاريخ يوم عطلة (جمعة أو سبت).");
  const items = (Array.isArray(a.items) ? a.items : []).slice(0, 7);
  return withLock_(() => {
    const absSh = sh_(SHEET_ABS), covSh = sh_(SHEET_COV);
    const absRows = absSh.getDataRange().getDisplayValues();
    let rowIdx = -1, id = String(a.id || "");
    for (let i = 1; i < absRows.length; i++) {
      if (id ? absRows[i][0] === id : (absRows[i][1] === date && absRows[i][2] === teacher)) { rowIdx = i + 1; id = absRows[i][0]; break; }
    }
    if (a.id && rowIdx < 0) throw new Error("الغياب غير موجود.");
    if (!id) id = newId_();
    const absentToday = absRows.slice(1).filter(r => r[0] && r[1] === date && r[0] !== id).map(r => r[2]);
    if (absentToday.indexOf(teacher) !== -1) throw new Error("هذا المعلم مسجّل غائبًا في هذا اليوم.");
    const covRows = covSh.getDataRange().getDisplayValues();
    const asSub = covRows.slice(1).filter(r => r[0] && r[1] !== id && r[2] === date && r[7] === COV_DONE && r[6] === teacher);
    if (asSub.length) throw new Error("أ. " + teacher + " مسجّل بديلًا في هذا اليوم (الحصة " + asSub.map(r => r[3]).join("، ") + " بدلًا من أ. " + asSub[0][5] + "). غيّر البديل هناك أولًا.");
    const busy = covRows.slice(1).filter(r => r[0] && r[1] !== id && r[2] === date && r[7] === COV_DONE).map(r => r[6] + "|" + r[3]);
    const out = [], seen = {};
    items.forEach(it => {
      const period = Math.floor(Number(it.period));
      const les = lessonAt_(teacher, day, period);
      if (!les || les[2] === "R" || les[2] === "RF") throw new Error("الحصة " + period + " ليست من حصص المعلم في هذا اليوم.");
      if (seen[period]) return; seen[period] = 1;
      let status = [COV_DONE, COV_NONE, COV_PRESENT].indexOf(it.status) !== -1 ? it.status : COV_NONE;
      let sub = String(it.substitute || "");
      if (status === COV_DONE) {
        if (!TIMETABLE[sub] || sub === teacher) throw new Error("اختر المعلم البديل للحصة " + period + ".");
        const l2 = lessonAt_(sub, day, period);
        if (l2 && l2[2] !== "R" && l2[2] !== "RF") throw new Error("أ. " + sub + " عنده حصة في الحصة " + period + ".");
        if (absentToday.indexOf(sub) !== -1) throw new Error("أ. " + sub + " غائب في هذا اليوم.");
        if (busy.indexOf(sub + "|" + period) !== -1) throw new Error("أ. " + sub + " مكلّف باحتياط آخر في الحصة " + period + ".");
      } else sub = "";
      out.push([newId_(), id, date, String(period), les[2], teacher, sub, status, nowText_()]);
    });
    const now = nowText_();
    const absVals = [id, date, teacher, String(a.note || "").slice(0, 300), rowIdx > 0 ? absRows[rowIdx - 1][4] : now, now];
    if (rowIdx > 0) writeRow_(absSh, rowIdx, absVals); else appendText_(absSh, absVals);
    for (let i = covRows.length - 1; i >= 1; i--) if (covRows[i][1] === id) covSh.deleteRow(i + 1);
    out.forEach(r => appendText_(covSh, r));
    return { id };
  });
}

function deleteAbsence(code, id) {
  head_(code);
  return withLock_(() => {
    const absSh = sh_(SHEET_ABS), row = findRow_(absSh, 1, id);
    if (row > 0) absSh.deleteRow(row);
    const covSh = sh_(SHEET_COV), v = covSh.getDataRange().getDisplayValues();
    for (let i = v.length - 1; i >= 1; i--) if (v[i][1] === id) covSh.deleteRow(i + 1);
    return true;
  });
}

// ======================= كشوف متابعة المتعلمين (الدرجات) =======================
// قوائم الطلاب في ورقة «الطلاب» (الصف-الشعبة، م، الاسم، ملاحظة) بلا أرقام مدنية، والدرجات في ورقة «الدرجات».
// توزيع الدرجات حسب خطة التوجيه الفني ٢٠٢٦/٢٠٢٧ (واحد للحديث والقرآن):
// السلوك ٣+٣ · المشاركة ٣+٣ · الواجبات ٣+٣ · المشروع ٣٫٥+٣٫٥ (الفترة الأولى + الفترة الثانية) · اختبار قصير ١٥ = المجموع ٤٠
// قرار رئيس الشعبة: لا يُرصد في الكشف غير هذه البنود (لا امتحان نهائي)
const SHEET_STU = "الطلاب";
const SHEET_GRD = "الدرجات";
const GRD_MAX = [3, 3, 3, 3, 3, 3, 3.5, 3.5, 15];
const GRD_HEAD = ["المفتاح", "القسم", "الصف", "الفصل الدراسي", "م", "اسم المتعلم", "السلوك ق١", "السلوك ق٢", "المشاركة ق١", "المشاركة ق٢",
  "الواجبات ق١", "الواجبات ق٢", "المشروع ق١", "المشروع ق٢", "الاختبار القصير", "المجموع (٤٠)", "المعلم", "آخر تحديث"];
const DEMO_STUDENTS = ["متعلم تجريبي ١", "متعلم تجريبي ٢", "متعلم تجريبي ٣", "متعلم تجريبي ٤", "متعلم تجريبي ٥", "متعلم تجريبي ٦",
  "متعلم تجريبي ٧", "متعلم تجريبي ٨", "متعلم تجريبي ٩", "متعلم تجريبي ١٠", "متعلم تجريبي ١١", "متعلم تجريبي ١٢"];

function teacherOf_(cls) { const t = TEACHERS.find(x => x[1].indexOf(cls) !== -1); return t ? t[0] : ""; }
function canSeeClass_(me, cls) { return me.isHead && !me.isDemo ? !!teacherOf_(cls) : me.classes.indexOf(cls) !== -1; }
function gradesSheet_() { return sheet_(SHEET_GRD, GRD_HEAD, true); }
function studentsOf_(me, cls) {
  if (me.isDemo) return DEMO_STUDENTS.map((name, i) => ({ n: i + 1, name, note: "" }));
  const sec = String(cls).split(":")[1];
  const sh = sh_(SHEET_STU);
  if (!sh) return [];
  return sh.getDataRange().getDisplayValues().slice(1).filter(r => r[0] === sec)
    .map(r => ({ n: Number(r[1]), name: String(r[2]).trim(), note: String(r[3] || "") })).sort((a, b) => a.n - b.n);
}
function sum_(vals) { return Math.round(vals.reduce((a, v) => a + Number(v), 0) * 100) / 100; }
function totals_(v) { return [v.every(x => x !== "") ? String(sum_(v)) : ""]; }

/** نظرة عامة على صفوفي (أو كل الصفوف لرئيس الشعبة): عدد المتعلمين ومدى اكتمال الرصد في كل فصل دراسي */
function classesOverview(code) {
  const me = who_(code);
  const list = me.isHead && !me.isDemo ? TEACHERS.reduce((a, t) => a.concat(t[1]), []) : me.classes;
  const counts = {};
  if (me.isDemo) list.forEach(c => counts[c.split(":")[1]] = DEMO_STUDENTS.length);
  else if (sh_(SHEET_STU)) rows_(SHEET_STU).slice(1).forEach(r => { if (r[0]) counts[r[0]] = (counts[r[0]] || 0) + 1; });
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
  const pre = demoPrefix_(me) + cls + "|" + term + "|";
  const grades = {}; let updatedAt = "";
  if (sh_(SHEET_GRD)) rows_(SHEET_GRD).slice(1).forEach(r => {
    if (String(r[0]).indexOf(pre) !== 0) return;
    grades[String(r[0]).slice(pre.length)] = r.slice(6, 15);
    if (r[17] > updatedAt) updatedAt = r[17];
  });
  return { cls, term, teacher: me.isDemo ? me.name : teacherOf_(cls), students: studentsOf_(me, cls), grades, updatedAt };
}

/** حفظ درجات صف: data = { "اسم المتعلم": [١٠ قيم نصية أو فارغة] } */
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
    const v = (data && data[s.name]) || [];
    const vals = GRD_MAX.map((mx, i) => {
      const x = String(v[i] == null ? "" : v[i]).trim();
      if (x === "") return "";
      if (!/^\d{1,2}(\.\d{1,2})?$/.test(x) || Number(x) > mx) throw new Error(`درجة غير صحيحة للمتعلم ${s.name}: «${x}» (العظمى ${mx}).`);
      return String(Number(x));
    });
    rows.push([pre + s.name, dept, p[1], term === 1 ? "الأول" : "الثاني", String(s.n), s.name].concat(vals, totals_(vals), [teacher, now]));
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
const SREP_HEAD = ["المعرف", "النوع", "المعلم", "الصف", "القسم", "الفصل الدراسي", "اسم المتعلم", "المجموع في الكشف", "البيانات", "تاريخ الإرسال", "اطّلع رئيس الشعبة", "آخر تعديل"];
function srepSheet_() { return sheet_(SHEET_SREP, SREP_HEAD, true); }
function srepObj_(r) {
  let data = {}; try { data = JSON.parse(r[8] || "{}"); } catch (e) {}
  return { id: r[0], type: r[1], teacher: r[2], cls: r[3], term: Number(r[5]) || 1, student: r[6], score: r[7], data, createdAt: r[9], seenAt: r[10], updatedAt: r[11] };
}
function cleanList_(a) { return (Array.isArray(a) ? a : []).slice(0, 20).map(x => String(x).trim().slice(0, 200)).filter(Boolean); }

function listStudentReports(code) {
  const me = who_(code);
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
  if (!studentsOf_(me, cls).some(s => s.name === student)) throw new Error("اختر المتعلم من قائمة الصف.");
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
  notify_(me, `${row > 0 ? "عدّل" : "أرسل"} أ. ${me.name} تقرير ${SREP_TYPES[rep.type]} (${SUBJ[p[0]]}): ${student}`, `
    <p>${row > 0 ? "عدّل" : "أرسل"} <b>أ. ${esc_(me.name)}</b> تقرير <b>${SREP_TYPES[rep.type]}</b> في <b>${SUBJ[p[0]]}</b> عن المتعلم <b>${esc_(student)}</b>.</p>
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
    mode: r[11], substitute: r[12], sentAt: r[14], status: r[15], decidedAt: r[17], headNote: r[18] };
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
  if (!validSig_(p.sig)) throw new Error("وقّع على الطلب قبل الإرسال.");
  const id = newId_(), now = nowText_(), year = date.slice(0, 4);
  const serial = withLock_(() => {
    saveProfile_(me.name, prof);
    const rows = permRows_().filter(r => r[8].slice(0, 4) === year && (me.isDemo ? r[2] === me.name : !isDemo_(r[2])));
    const n = rows.reduce((mx, r) => Math.max(mx, Number(String(r[1]).split("/")[0]) || 0), 0) + 1;
    const s = n + "/" + year;
    appendText_(permSheet_(), [id, s, me.name, type, prof.dept, prof.job, prof.civil, reason, date, p.from, p.to, mode, sub, p.sig, now, PERM_PENDING, "", "", ""]);
    saveMySig_(me.name, p.sig, now);
    return s;
  });
  if (!p.deferMail) permMail_(me, { serial, type, date, from: p.from, to: p.to, reason, mode, sub }, null);
  return { id, serial };
}
function permMail_(me, q, pdf) {
  const toEnd = q.to === "نهاية الدوام";
  notify_(me, `طلب استئذان من أ. ${me.name} (${q.date})`, `
    <p>أرسل <b>أ. ${esc_(me.name)}</b> طلب استئذان${pdf ? "، والنموذج مرفق PDF جاهز للطباعة والتوقيع" : ""}. ويمكن اعتماده أيضًا من الموقع.</p>
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
    if (!(head || (r[2] === me.name && r[15] === PERM_PENDING))) throw new Error("لا يمكن حذف هذا الاستئذان بعد اعتماده.");
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
  if (!validSig_(sig)) throw new Error("وقّع قبل تأكيد الاطلاع.");
  return withLock_(() => {
    const old = nshAcks_().find(a => a[0] === r[0] && a[1] === me.name);
    if (old) return { at: old[2] };
    const at = nowText_();
    appendText_(nshAckSheet_(), [r[0], me.name, at, sig]);
    saveMySig_(me.name, sig, at);
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
// أعمدة «الإشعارات»: 0 الوقت (yyyy-MM-dd HH:mm:ss) | 1 المعلم | 2 الركن | 3 ماذا فعل | 4 مفتاح الدمج
// تُسجَّل الأحداث من doPost بعد نجاح الطلب. حفظ الدرجات والمنهج لنفس الصف خلال ساعة يُدمج في إشعار واحد.
function actSheet_() { return sheet_(SHEET_ACTLOG, ["الوقت", "المعلم", "الركن", "ماذا فعل", "مفتاح الدمج"], true); }
function stamp_() { return Utilities.formatDate(new Date(), "Asia/Kuwait", "yyyy-MM-dd HH:mm:ss"); }
function arNum_(x) { return String(x).replace(/\d/g, d => "٠١٢٣٤٥٦٧٨٩"[d]); }
function clsName_(id) { const m = String(id || "").match(/^(hd|qr):(\d+)-(\d+)$/); return m ? `${SUBJ[m[1]]} ${arNum_(m[2])}/${arNum_(m[3])}` : String(id || ""); }
function rowTitle_(sheet, id, col) { const r = rows_(sheet).slice(1).find(x => x[0] === String(id)); return r ? r[col] : ""; }
const ACT_LOG = {
  saveEntry: a => ({ sec: "curr", text: `حدّث ما قطع من المنهج: ${clsName_(a[2] + ":" + a[1])}${a[5] ? " — " + a[5] : ""}`, key: "curr|" + a[2] + a[1] }),
  saveGrades: a => ({ sec: "cls", text: `رصد درجات كشف متابعة المتعلمين: ${clsName_(a[1])} (الفصل الدراسي ${a[2] == 2 ? "الثاني" : "الأول"})`, key: "grd|" + a[1] + "|" + a[2] }),
  saveStudentReport: a => ({ sec: "cls", text: `${a[1] && a[1].id ? "عدّل" : "أرسل"} تقرير ${SREP_TYPES[a[1] && a[1].type] || "متعلم"}: ${a[1] && a[1].student || ""} (${clsName_(a[1] && a[1].cls)})` }),
  deleteStudentReport: () => ({ sec: "cls", text: "حذف تقرير متعلم" }),
  submitAchievement: a => ({ sec: a[1] && a[1].kind === "act" ? "act" : "ach", text: `${a[1] && a[1].kind === "act" ? "رفع نشاطًا في سجل الأنشطة والمسابقات" : "رفع إنجازًا في التنمية المهنية وإنجاز الموظف"}: ${a[1] && a[1].title || ""}` }),
  deleteAchievement: () => ({ sec: "ach", text: "حذف إنجازًا أو نشاطًا مرفوعًا" }),
  signMinute: a => ({ sec: "min", text: `وقّع على المحضر: ${rowTitle_(SHEET_MIN, a[1], 1)}` }),
  signVisit: () => ({ sec: "vis", text: "اطّلع ووقّع على بطاقة الزيارة الصفية" }),
  submitPermission: a => ({ sec: "prm", text: `قدّم طلب استئذان (${a[1] && a[1].type || ""}) ليوم ${a[1] && a[1].date || ""}` }),
  deletePermission: () => ({ sec: "prm", text: "حذف طلب استئذان" }),
  ackCircular: a => { const r = nshRows_().find(x => x[0] === String(a[1])); return { sec: "nsh", text: `اطّلع ووقّع على ${r ? r[1] + ": " + r[4] : "نشرة"}` }; }
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
        const v = sh.getRange(from, 1, last - from + 1, 5).getDisplayValues();
        for (let i = v.length - 1; i >= 0; i--) if (v[i][4] === key && (new Date(now.replace(" ", "T")) - new Date(v[i][0].replace(" ", "T"))) < 3600000) {
          sh.getRange(from + i, 1, 1, 5).setValues([[now, me.name, ev.sec, ev.text, key]]); return;
        }
      }
    }
    appendText_(sh, [now, me.name, ev.sec, ev.text, key]);
  });
}
/** لرئيس الشعبة: آخر الإشعارات، وعدد غير المقروء */
function listActivity(code, limit) {
  head_(code);
  const seen = PropertiesService.getScriptProperties().getProperty("actSeen") || "";
  const v = actSheet_().getDataRange().getDisplayValues().slice(1).filter(r => r[0]);
  const items = v.slice(-Math.min(Number(limit) || 150, 300)).reverse().map(r => ({ at: r[0], teacher: r[1], sec: r[2], text: r[3], demo: isDemo_(r[1]), unread: r[0] > seen }));
  return { items, unread: v.filter(r => r[0] > seen).length, seen };
}
function markActivitySeen(code) { head_(code); const t = stamp_(); PropertiesService.getScriptProperties().setProperty("actSeen", t); return { seen: t }; }
