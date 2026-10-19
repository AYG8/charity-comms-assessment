/**
 * Shared response store for the Asir charity communications assessment.
 * Deploy as a web app that runs as the deploying account and is accessible
 * to anyone. This creates a private Google Sheet owned by the deployer.
 * The public web app accepts survey submissions and exposes aggregate-source
 * records because the companion GitHub Pages dashboard is public.
 */
const SHEET_NAME = 'الاستجابات';
const PROP_KEY = 'ASSESSMENT_SPREADSHEET_ID';
const LOCK_TIMEOUT_MS = 5000;

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

  // The top-level survey page listens for this nonce-bound message; Apps Script
  // HtmlService may wrap the response in a sandboxed iframe.
  const nonce = String((e && e.parameter && e.parameter.nonce) || '').replace(/[^0-9A-Za-z_-]/g, '');
  const message = JSON.stringify({ type: 'assessment-save', nonce: nonce, result: result })
    .replace(/</g, '\\u003c');
  const html = '<!doctype html><meta charset="utf-8"><script>top.postMessage(' + message + ', "*");</script>';
  return HtmlService.createHtmlOutput(html);
}

function saveRecord_(record) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
    throw new Error('الخدمة مشغولة مؤقتًا؛ أعد المحاولة بعد لحظات.');
  }

  try {
    const requestedTime = record.updatedAt ? new Date(record.updatedAt).getTime() : NaN;
    const cleanRecord = {
      ...record,
      id: String(record.id || Utilities.getUuid()),
      key: String(record.key),
      orgName: String(record.orgName),
      city: String(record.city),
      updatedAt: Number.isFinite(requestedTime)
        ? new Date(requestedTime).toISOString()
        : new Date().toISOString()
    };
    const serialized = JSON.stringify(cleanRecord);
    const sheet = getSheet_();
    const lastRow = sheet.getLastRow();

    let existingCell = null;
    if (lastRow > 1) {
      existingCell = sheet
        .getRange(2, 2, lastRow - 1, 1)
        .createTextFinder(cleanRecord.key)
        .matchCase(true)
        .matchEntireCell(true)
        .useRegularExpression(false)
        .findNext();
    }

    const rowValues = [[
      cleanRecord.id,
      cleanRecord.key,
      cleanRecord.orgName,
      cleanRecord.city,
      cleanRecord.updatedAt,
      serialized
    ]];

    if (existingCell) {
      const rowNumber = existingCell.getRow();
      const previousUpdatedAt = sheet.getRange(rowNumber, 5).getValue();
      const oldTime = new Date(previousUpdatedAt || 0).getTime();
      const newTime = new Date(cleanRecord.updatedAt).getTime();

      if (!Number.isFinite(oldTime) || newTime >= oldTime) {
        sheet.getRange(rowNumber, 1, 1, 6).setValues(rowValues);
      }
      return;
    }

    sheet.getRange(lastRow + 1, 1, 1, 6).setValues(rowValues);
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
