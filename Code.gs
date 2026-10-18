/**
 * Shared response store for the Asir charity communications assessment.
 * Deploy as a web app that runs as the deploying account and is accessible
 * to anyone. This creates a private Google Sheet owned by the deployer.
 * The public web app accepts survey submissions and exposes aggregate-source
 * records because the companion GitHub Pages dashboard is public.
 */
const SHEET_NAME = 'الاستجابات';
const PROP_KEY = 'ASSESSMENT_SPREADSHEET_ID';

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'health';
  const callback = (e && e.parameter && e.parameter.callback) || '';
  let payload;
  try {
    if (action === 'health') {
      payload = { ok: true, service: 'charity-comms-assessment' };
    } else if (action === 'list') {
      const requiredKey = PropertiesService.getScriptProperties().getProperty('DASHBOARD_READ_KEY');
      const suppliedKey = String((e.parameter && e.parameter.key) || '');
      payload = requiredKey && suppliedKey === requiredKey
        ? { ok: true, records: readRecords_() }
        : { ok: false, error: requiredKey ? 'unauthorized' : 'dashboard_key_not_configured' };
    } else {
      payload = { ok: false, error: 'unknown_action' };
    }
  } catch (err) {
    payload = { ok: false, error: String(err && err.message || err) };
  }

  const json = JSON.stringify(payload).replace(/</g, '\\u003c');
  if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*(\.[A-Za-z_$][0-9A-Za-z_$]*)*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  let result = { ok: false, error: 'invalid_request' };
  try {
    const raw = e && e.parameter && e.parameter.record;
    const record = raw ? JSON.parse(raw) : null;
    if (!record || !record.orgName || !record.city || !record.answers || !record.key) {
      throw new Error('الاستجابة غير مكتملة');
    }
    saveRecord_(record);
    result = { ok: true, id: record.id, updatedAt: record.updatedAt };
  } catch (err) {
    result = { ok: false, error: String(err && err.message || err) };
  }

  // The parent survey page listens for this nonce-bound message from its
  // hidden form iframe; the Google-origin response itself is never rendered.
  const nonce = String((e && e.parameter && e.parameter.nonce) || '').replace(/[^0-9A-Za-z_-]/g, '');
  const message = JSON.stringify({ type: 'assessment-save', nonce: nonce, result: result })
    .replace(/</g, '\\u003c');
  const html = '<!doctype html><meta charset="utf-8"><script>parent.postMessage(' + message + ', "*");</script>';
  return HtmlService.createHtmlOutput(html);
}

function saveRecord_(record) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = getSheet_();
    const rows = sheet.getDataRange().getValues();
    const key = String(record.key);
    const serialized = JSON.stringify(record);
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][1]) === key) {
        const oldTime = new Date(rows[i][4] || 0).getTime();
        const newTime = new Date(record.updatedAt || Date.now()).getTime();
        if (newTime >= oldTime) {
          sheet.getRange(i + 1, 1, 1, 4).setValues([[
            String(record.id), key, String(record.orgName), String(record.city),
          ]]);
          sheet.getRange(i + 1, 5).setValue(String(record.updatedAt || new Date().toISOString()));
          sheet.getRange(i + 1, 6).setValue(serialized);
        }
        return;
      }
    }
    sheet.appendRow([
      String(record.id), key, String(record.orgName), String(record.city),
      String(record.updatedAt || new Date().toISOString()), serialized,
    ]);
  } finally {
    lock.releaseLock();
  }
}

function readRecords_() {
  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  return values.slice(1).map(row => {
    try { return JSON.parse(String(row[5] || '')); } catch (_) { return null; }
  }).filter(Boolean);
}

function getSheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(PROP_KEY);
  let spreadsheet;
  if (id) {
    spreadsheet = SpreadsheetApp.openById(id);
  } else {
    spreadsheet = SpreadsheetApp.create('استطلاع جاهزية العلاقات العامة والإعلام والاتصال المؤسسي - عسير');
    props.setProperty(PROP_KEY, spreadsheet.getId());
  }
  let sheet = spreadsheet.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = spreadsheet.getSheets()[0].setName(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['معرف الاستجابة', 'مفتاح الجمعية', 'اسم الجمعية', 'المحافظة', 'آخر تحديث', 'بيانات الاستجابة']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}
