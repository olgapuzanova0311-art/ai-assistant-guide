// Google Apps Script: принимает регистрации с сайта и пишет их в таблицу
const SHEET_ID = '1hOFGRX5uW6aREUK3QMP5ss_bqCPfVLqzzxj5J3TnSwA';

function safe(v, n) {
  v = String(v || '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, n);
  return /^[=+\-@]/.test(v) ? "'" + v : v;
}
function norm(v) { return String(v || '').toLowerCase().replace(/[\s()\-+@']/g, ''); }
function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e.postData && e.postData.contents) || '{}');
    const name = safe(d.name, 60), contact = safe(d.contact, 80);
    if (name.length < 2 || contact.length < 3) return out({ ok: false });
    const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
    const last = sh.getLastRow();
    const known = last > 1 ? sh.getRange(2, 3, last - 1, 1).getValues().map(r => norm(r[0])) : [];
    if (known.indexOf(norm(contact)) === -1) {
      sh.appendRow([Utilities.formatDate(new Date(), 'Europe/Moscow', 'dd.MM.yyyy HH:mm'), name, contact, safe(d.src, 120)]);
    }
    return out({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  const sh = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
  return out({ count: Math.max(sh.getLastRow() - 1, 0) });
}
