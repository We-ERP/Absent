/**
 * Seif Dashboard API  -  Google Apps Script (Container-bound على الشيت)
 * الشيتات المطلوبة: "Codes SHR" , "Codes absent" , "STR"
 *
 * الإعداد:
 * 1) Project Settings > Script properties > أضف ADMIN_PASSWORD = كلمة السر بتاعتك
 * 2) Deploy > Manage deployments > Edit > New version  (Execute as: Me , Access: Anyone)
 */
const ALLOWED = ['Codes SHR', 'Codes absent', 'STR'];

function doGet() {
  return json_({ ok: true, msg: 'Seif Dashboard API 2026' });
}

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const pwd = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
    if (!pwd || req.pwd !== pwd) throw new Error('كلمة السر غلط');
    out = handle_(req);
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return json_(out);
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(r) {
  if (r.action === 'login') return { ok: true };
  if (ALLOWED.indexOf(r.sheet) < 0) throw new Error('شيت غير مسموح');
  const sh = SpreadsheetApp.getActive().getSheetByName(r.sheet);
  if (!sh) throw new Error('الشيت مش موجود: ' + r.sheet);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (r.action) {
      case 'read':
        return { ok: true, values: sh.getDataRange().getDisplayValues() };

      case 'setCells': // تعديل صف جزئي
        sh.getRange(r.row, r.c1, 1, r.values.length).setValues([r.values]);
        return { ok: true };

      case 'addRow': { // إضافة في آخر الجدول (البلوك)
        const max = sh.getMaxRows();
        const col = sh.getRange(r.start, r.c1, Math.max(max - r.start + 1, 1), 1).getValues();
        let last = -1;
        col.forEach((v, i) => { if (String(v[0]).trim() !== '') last = i; });
        const row = r.start + last + 1;
        if (row > max) sh.insertRowsAfter(max, row - max);
        sh.getRange(row, r.c1, 1, r.values.length).setValues([r.values]);
        return { ok: true, row: row };
      }

      case 'delRow': // حذف صف من البلوك بس (الجدول التاني ماتتأثرش)
        sh.getRange(r.row, r.c1, 1, r.c2 - r.c1 + 1).deleteCells(SpreadsheetApp.Dimension.ROWS);
        return { ok: true };

      case 'replace': { // رفع شيت بدل البلوك كله
        const width = r.c2 - r.c1 + 1;
        const lastRow = Math.max(sh.getLastRow(), r.start);
        sh.getRange(r.start, r.c1, lastRow - r.start + 1, width).clearContent();
        const need = r.start + r.values.length - 1;
        if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
        if (r.values.length) sh.getRange(r.start, r.c1, r.values.length, width).setValues(r.values);
        return { ok: true, count: r.values.length };
      }
    }
    throw new Error('action مش معروف');
  } finally {
    lock.releaseLock();
  }
}
