/**
 * نقل ملفات الموقع من مجلد الحساب الشخصي إلى مجلد الملفات في حساب الوزارة — يُشغَّل مرة (أو أكثر) من مشروع الوزارة فقط.
 * - ينسخ كل الملفات (مع المجلدات الفرعية) إلى مجلد الموقع (folder_())، ويكتب في وصف كل نسخة «src:<معرف الأصل>»
 *   فلا يتكرر النسخ إذا أعيد التشغيل.
 * - يتوقف قبل ٤٫٥ دقائق ويقول «شغّلها مرة ثانية»؛ وعند اكتمال النسخ يستبدل معرّفات الملفات القديمة في كل خلايا الشيت بمعرّفات النسخ.
 * - لا يحذف ولا يعدّل أي شيء في الحساب الشخصي.
 * - يستثني ملفات «كشف متابعة المتعلمين» القديمة (قبل ٢٦/١٠/٢٠٢٦) لأن فيها أسماء متعلمين (قرار رئيس الشعبة: الكشوف بالأرقام فقط).
 */
const MIG_SRC_FOLDER = "1gArSX174Rax6ZWZCIpotdnlW-10SNCjX";
const MIG_LIMIT_MS = 4.5 * 60 * 1000;
const MIG_SKIP_BEFORE = new Date("2026-10-26T00:00:00+03:00");

function migrateFiles() {
  const t0 = Date.now();
  const src = DriveApp.getFolderById(MIG_SRC_FOLDER);
  const dest = folder_();
  const map = migMap_(dest);              // معرف الأصل ← معرف النسخة
  const st = { copied: 0, existing: 0, skipped: 0, failed: [] };
  const done = migCopy_(src, dest, map, st, t0);
  PropertiesService.getScriptProperties().setProperty("MIG_MAP", JSON.stringify(map));
  let msg = `نُسخ الآن: ${st.copied} · موجود من قبل: ${st.existing} · مستثنى: ${st.skipped} · تعذّر: ${st.failed.length}`;
  if (st.failed.length) msg += "\nتعذّر نسخ: " + st.failed.slice(0, 20).join(" | ");
  if (!done) { msg += "\n⏸ انتهى الوقت — شغّل migrateFiles مرة ثانية ليكمل."; Logger.log(msg); return msg; }
  const rep = migReplaceIds_(map);
  msg += `\n✓ اكتمل النسخ. استُبدلت المعرّفات في ${rep.cells} خلية (${rep.ids} معرّف).`;
  Logger.log(msg);
  return msg;
}

/** يبني الخريطة من أوصاف النسخ الموجودة في مجلد الوزارة (ومجلداته الفرعية) */
function migMap_(folder) {
  const map = {};
  (function walk(f) {
    const it = f.getFiles();
    while (it.hasNext()) { const x = it.next(); const m = String(x.getDescription() || "").match(/^src:([A-Za-z0-9_-]{20,})$/); if (m) map[m[1]] = x.getId(); }
    const fs = f.getFolders();
    while (fs.hasNext()) { const sub = fs.next(); const m = String(sub.getDescription() || "").match(/^src:([A-Za-z0-9_-]{20,})$/); if (m) map[m[1]] = sub.getId(); walk(sub); }
  })(folder);
  return map;
}

function migCopy_(src, dest, map, st, t0) {
  const files = src.getFiles();
  while (files.hasNext()) {
    if (Date.now() - t0 > MIG_LIMIT_MS) return false;
    const f = files.next(), id = f.getId();
    if (map[id]) { st.existing++; continue; }
    if (/^كشف متابعة المتعلمين/.test(f.getName()) && f.getDateCreated() < MIG_SKIP_BEFORE) { st.skipped++; continue; }
    try { const c = f.makeCopy(f.getName(), dest); c.setDescription("src:" + id); map[id] = c.getId(); st.copied++; }
    catch (e) { st.failed.push(f.getName() + ": " + (e && e.message || e)); }
  }
  const subs = src.getFolders();
  while (subs.hasNext()) {
    if (Date.now() - t0 > MIG_LIMIT_MS) return false;
    const s = subs.next(), sid = s.getId();
    let d;
    if (map[sid]) d = DriveApp.getFolderById(map[sid]);
    else { d = dest.createFolder(s.getName()); d.setDescription("src:" + sid); map[sid] = d.getId(); }
    if (!migCopy_(s, d, map, st, t0)) return false;
  }
  return true;
}

/** يستبدل في كل أوراق الشيت أي نص من ٢٥ حرفًا فأكثر (حروف وأرقام و_ و-) إذا كان معرّف ملف قديم منسوخ */
function migReplaceIds_(map) {
  const re = /[A-Za-z0-9_-]{25,}/g;
  let cells = 0; const used = {};
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(sh => {
    const r = sh.getDataRange(), v = r.getValues();
    v.forEach((row, i) => row.forEach((c, j) => {
      if (typeof c !== "string" || c.length < 25) return;
      let changed = false;
      const n = c.replace(re, m => { if (map[m]) { changed = true; used[m] = 1; return map[m]; } return m; });
      if (changed) { sh.getRange(i + 1, j + 1).setValue(n); cells++; }
    }));
  });
  return { cells, ids: Object.keys(used).length };
}

// ======================= مزامنة البيانات من الشيت القديم (الحساب الشخصي) =======================
/**
 * syncDataFromOld — يُشغَّل من مشروع الوزارة فقط، يوم التحويل (أو كلما احتجت تحديث الموقع الجديد).
 * - ينقل كل أوراق الشيت القديم (الإدخالات، الخطط، الزيارات، الغياب، الاستئذان، المحاضر، الرموز…) إلى شيت الموقع الجديد،
 *   فتصير بيانات الموقع الجديد مطابقة للقديم. الأوراق الموجودة في الجديد فقط تبقى كما هي.
 * - قبل أي كتابة يحفظ نسخة كاملة من الشيت الجديد في Drive («نسخة قبل المزامنة - …») ليمكن الرجوع إليها.
 * - لا يعدّل ولا يحذف أي شيء في الشيت القديم (قراءة فقط).
 * - بعده شغّل migrateFiles لنقل الصور والملفات الجديدة واستبدال معرّفاتها في الشيت.
 */
const SYNC_OLD_SHEET_ID = "16shJB2Mng_dguWJmfWNGt4gsmgfFLOepwYuf_jniyNY";

function syncDataFromOld() {
  const old = SpreadsheetApp.openById(SYNC_OLD_SHEET_ID);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss.getId() === SYNC_OLD_SHEET_ID) throw new Error("هذا هو الشيت القديم نفسه — شغّل الدالة من مشروع الوزارة.");
  const stamp = Utilities.formatDate(new Date(), "Asia/Kuwait", "yyyy-MM-dd HH:mm");
  DriveApp.getFileById(ss.getId()).makeCopy("نسخة قبل المزامنة - " + ss.getName() + " - " + stamp);
  const out = [];
  return withLock_(() => {
    old.getSheets().forEach(src => {
      const name = src.getName();
      let dst = ss.getSheetByName(name);
      if (!dst) dst = ss.insertSheet(name);
      const rows = src.getLastRow(), cols = src.getLastColumn();
      dst.clearContents();
      if (rows && cols) {
        if (dst.getMaxRows() < rows) dst.insertRowsAfter(dst.getMaxRows(), rows - dst.getMaxRows());
        if (dst.getMaxColumns() < cols) dst.insertColumnsAfter(dst.getMaxColumns(), cols - dst.getMaxColumns());
        const sr = src.getRange(1, 1, rows, cols), dr = dst.getRange(1, 1, rows, cols);
        dr.setNumberFormats(sr.getNumberFormats());
        dr.setValues(sr.getValues());
      }
      dst.setRightToLeft(src.isRightToLeft());
      if (src.getFrozenRows()) dst.setFrozenRows(src.getFrozenRows());
      out.push(name + ": " + Math.max(rows - 1, 0));
    });
    SpreadsheetApp.flush();
    const msg = "✓ نُقلت البيانات (عدد الصفوف في كل ورقة):\n" + out.join("\n") +
      "\n\nالخطوة التالية: شغّل migrateFiles لنقل الصور والملفات الجديدة.";
    Logger.log(msg);
    return msg;
  });
}
