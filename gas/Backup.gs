/**
 * نسخة احتياطية أسبوعية لكود الموقع (Code.gs و config.js) في Google Drive.
 * - يجلب أحدث نسخة من المستودع العام في GitHub ويحفظها ملفات نصية داخل مجلد «نسخة yyyy-MM-dd»
 *   في «نسخ احتياطية - موقع القسم» (نفس المجلد الذي تنسخ فيه المهمة المجدولة البيانات والملفات المرفوعة).
 * - لا يحذف ولا يعدّل أي ملف موجود؛ إن وُجد الملف لنفس اليوم لا يكرره.
 * - إن لم يستطع الحساب الوصول لمجلد النسخ (مثلًا لأنه في حساب آخر غير مشارَك) يحفظ في مجلد
 *   «نسخ احتياطية - كود موقع القسم» في Drive هذا الحساب، ويكتب السبب في السجل (Executions).
 *
 * التشغيل لأول مرة: من محرر Apps Script اختر installBackupTrigger ثم Run (ووافق على الصلاحيات).
 * بعدها يعمل تلقائيًا كل سبت بين ٨ و٩ الصبح بتوقيت الكويت (بعد المهمة المجدولة اللي تنشئ مجلد النسخة الساعة ٥:٤٧). وللتجربة الآن: شغّل backupCode.
 */
const BK_ROOT_ID = "1G3MDPYsDvanQHvxF-dr4XgwIYaLu_AEb"; // «نسخ احتياطية - موقع القسم»
const BK_FALLBACK_NAME = "نسخ احتياطية - كود موقع القسم";
const BK_RAW = "https://raw.githubusercontent.com/khalifaq974-web/sumait-dept/main/";
const BK_FILES = [["gas/Code.gs", "Code.gs"], ["config.js", "config.js"]];

function backupCode() {
  const day = Utilities.formatDate(new Date(), "Asia/Kuwait", "yyyy-MM-dd");
  let root;
  try { root = DriveApp.getFolderById(BK_ROOT_ID); root.getName(); }
  catch (e) {
    console.warn("لا يمكن الوصول لمجلد النسخ الاحتياطية، سيُحفظ في مجلد هذا الحساب: " + e);
    const it = DriveApp.getFoldersByName(BK_FALLBACK_NAME);
    root = it.hasNext() ? it.next() : DriveApp.createFolder(BK_FALLBACK_NAME);
  }
  const name = "نسخة " + day;
  const fit = root.getFoldersByName(name);
  const dest = fit.hasNext() ? fit.next() : root.createFolder(name);
  const out = [];
  BK_FILES.forEach(([path, title]) => {
    if (dest.getFilesByName(title).hasNext()) { out.push(title + ": موجود"); return; }
    const res = UrlFetchApp.fetch(BK_RAW + path + "?t=" + Date.now(), { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error("تعذر جلب " + path + " من GitHub (رمز " + res.getResponseCode() + ")");
    const blob = Utilities.newBlob(res.getContent(), "text/plain", title);
    dest.createFile(blob);
    out.push(title + ": " + blob.getBytes().length + " بايت");
  });
  console.log(name + " ← " + out.join(" · "));
  return out;
}

/** يُشغَّل مرة واحدة: ينشئ تشغيلًا أسبوعيًا (كل سبت ٨–٩ صباحًا بتوقيت الكويت) ولا يكرره إن وُجد */
function installBackupTrigger() {
  const has = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === "backupCode");
  if (!has) ScriptApp.newTrigger("backupCode").timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(8).inTimezone("Asia/Kuwait").create();
  return has ? "التشغيل الأسبوعي موجود من قبل" : "تم إنشاء التشغيل الأسبوعي";
}
