/**
 * موقع قسم المواد الشرعية (الحديث الشريف · القرآن الكريم) — معهد عبدالرحمن السميط الديني
 * الأركان: ما قطع من المنهج · الإنجازات · محاضر الاجتماعات · الزيارات الصفية
 * بعد أي تحديث للكود: شغّل setup مرة واحدة، ثم حدّث النشر (Manage deployments ← ✏️ ← New version).
 */

const SHEET_CODES = "الرموز";
const SHEET_DATA = "الإدخالات";
const SHEET_LOG = "السجل";
const SHEET_ACH = "الإنجازات";
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
  saveVisit, listVisits, getVisit, signVisit, deleteVisit, reserveData, saveAbsence, deleteAbsence };

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const f = Object.prototype.hasOwnProperty.call(API, req.fn) ? API[req.fn] : null;
    if (!f) throw new Error("طلب غير معروف.");
    out = { ok: true, result: f.apply(null, Array.isArray(req.args) ? req.args : []) };
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
  sheet_(SHEET_ACH, ["المعرف", "المعلم", "التاريخ", "النوع", "العنوان", "الفئة المستهدفة", "وصف المعلم", "الصور", "الحالة", "التقرير", "تاريخ الرفع", "تاريخ الصياغة"], true);
  sheet_(SHEET_MIN, ["المعرف", "العنوان", "التاريخ", "الوقت", "المكان", "الحضور", "جدول الأعمال", "ما دار في الاجتماع", "التوصيات", "تاريخ الإنشاء"], true);
  sheet_(SHEET_SIG, ["المحضر", "المعلم", "وقت التوقيع", "التوقيع"], true);
  sheet_(SHEET_MYSIG, ["المعلم", "التوقيع", "آخر تحديث"], true);
  sheet_(SHEET_VIS, ["المعرف", "المعلم", "الفصل", "القسم", "التاريخ", "الحصة", "نوع الزيارة", "الموضوع", "رقم الزيارة", "التقديرات",
    "الإيجابيات", "الملاحظات والتوصيات", "توقيع رئيس الشعبة", "تاريخ الإنشاء", "توقيع المعلم", "وقت توقيع المعلم"], true);
  sheet_(SHEET_MAIL, ["الوقت", "إلى", "الموضوع", "النتيجة"], true);
  sheet_(SHEET_ABS, ["المعرف", "التاريخ", "المعلم الغائب", "ملاحظة", "تاريخ التسجيل", "آخر تعديل"], true);
  sheet_(SHEET_COV, ["المعرف", "معرف الغياب", "التاريخ", "الحصة", "الفصل", "المعلم الغائب", "المعلم البديل", "الحالة", "تاريخ التسجيل"], true);
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

function achObj_(r, withReport) {
  const o = { id: r[0], teacher: r[1], date: r[2], type: r[3], title: r[4], target: r[5], desc: r[6],
    photos: r[7] ? r[7].split(",").filter(Boolean) : [], status: r[8] || ST_PENDING, createdAt: r[10], reportedAt: r[11] };
  if (withReport) o.report = r[9];
  return o;
}

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
  const row = [id, me.name, String(a.date || "").slice(0, 10), String(a.type || "").slice(0, 60), title,
    String(a.target || "").slice(0, 150), desc, ids.join(","), ST_PENDING, "", nowText_(), ""];
  withLock_(() => appendText_(sh_(SHEET_ACH), row));
  notify_(me, `رفع أ. ${me.name} إنجازًا جديدًا: ${title}`, `
    <p>رفع <b>أ. ${esc_(me.name)}</b> إنجازًا جديدًا في موقع القسم، وهو بانتظار الصياغة.</p>
    <table cellpadding="6" style="border-collapse:collapse;border:1px solid #d6e2ec">
      <tr><td style="background:#e8f1f8"><b>العنوان</b></td><td>${esc_(title)}</td></tr>
      <tr><td style="background:#e8f1f8"><b>النوع</b></td><td>${esc_(row[3])}</td></tr>
      <tr><td style="background:#e8f1f8"><b>التاريخ</b></td><td>${esc_(row[2])}</td></tr>
      <tr><td style="background:#e8f1f8"><b>الفئة المستهدفة</b></td><td>${esc_(row[5] || "—")}</td></tr>
    </table>
    <p><b>وصف المعلم:</b><br>${esc_(desc).replace(/\n/g, "<br>")}</p>
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
    if (row < 0) throw new Error("الإنجاز غير موجود.");
    const r = sh.getRange(row, 1, 1, 12).getDisplayValues()[0];
    if (!me.isHead && (r[1] !== me.name || r[8] === ST_READY)) throw new Error("لا يمكن حذف هذا الإنجاز.");
    (r[7] || "").split(",").filter(Boolean).forEach(fid => { try { DriveApp.getFileById(fid).setTrashed(true); } catch (e) {} });
    sh.deleteRow(row);
    return true;
  });
}

/** لرئيس الشعبة: كل الإنجازات مع التقارير */
function allAchievements(code) {
  head_(code);
  return rows_(SHEET_ACH).slice(1).filter(r => r[0] && !isDemo_(r[1])).map(r => achObj_(r, true)).reverse();
}

function saveReport(code, id, text) {
  head_(code);
  text = String(text || "").slice(0, 20000);
  return withLock_(() => {
    const sh = sh_(SHEET_ACH), row = findRow_(sh, 1, id);
    if (row < 0) throw new Error("الإنجاز غير موجود.");
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
