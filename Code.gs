/**
 * Backend for the exam app. Bind this script to the Google Sheet that has
 * the "Passcodes" tab (รหัสเข้าสอบ | สถานะ | หมายเหตุ) and a log tab
 * (LOG_SHEET_NAME) for exam results, then deploy as a Web App.
 */

const PASSCODE_SHEET_NAME = 'Passcodes';
const LOG_SHEET_NAME = 'ชีต1';

const STATUS_ACTIVE = 'ใช้งานได้'; // reusable code, never locked
const STATUS_UNUSED = 'ยังไม่ใช้'; // one-time code, still available
const STATUS_USED = 'ใช้แล้ว';     // one-time code, already redeemed

/**
 * Serves the HTML UI when opened directly, or acts as a small JSON API
 * (?action=checkPasscode / ?action=saveResult) for static hosts such as
 * GitHub Pages that embed this app but have no `google.script` bridge.
 */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'checkPasscode') {
    return jsonOutput(checkPasscode(e.parameter.code, e.parameter.name));
  }

  if (action === 'saveResult') {
    return jsonOutput(saveResult(e.parameter));
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบทดสอบความรู้ความสามารถด้านเทคโนโลยีสารสนเทศและคอมพิวเตอร์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Validates a passcode against the Passcodes sheet.
 * - STATUS_ACTIVE codes are reusable and left untouched.
 * - STATUS_UNUSED codes are one-time: on success they flip to STATUS_USED
 *   and the examinee's name is recorded in the หมายเหตุ column.
 * - Anything else (STATUS_USED, unknown code) is rejected.
 */
function checkPasscode(rawCode, examineeName) {
  const code = String(rawCode || '').trim().toUpperCase();
  if (!code) {
    return { valid: false, reason: 'empty' };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(PASSCODE_SHEET_NAME);
    if (!sheet) {
      throw new Error('ไม่พบชีต "' + PASSCODE_SHEET_NAME + '"');
    }

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowCode = String(data[i][0] || '').trim().toUpperCase();
      if (rowCode !== code) continue;

      const status = String(data[i][1] || '').trim();
      const rowNumber = i + 1;

      if (status === STATUS_ACTIVE) {
        return { valid: true };
      }

      if (status === STATUS_UNUSED) {
        sheet.getRange(rowNumber, 2).setValue(STATUS_USED);
        sheet.getRange(rowNumber, 3).setValue(String(examineeName || '').trim());
        return { valid: true };
      }

      // STATUS_USED or any other value => already redeemed / disabled
      return { valid: false, reason: 'used' };
    }

    return { valid: false, reason: 'not_found' };
  } finally {
    lock.releaseLock();
  }
}

/** Appends one exam result row to the log sheet. */
function saveResult(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    throw new Error('ไม่พบชีต "' + LOG_SHEET_NAME + '"');
  }

  sheet.appendRow([
    new Date(),
    payload.name,
    payload.position,
    payload.passcode,
    payload.score + ' / ' + payload.total,
    payload.percentage,
    payload.duration,
    payload.cat1,
    payload.cat2,
    payload.cat3,
    payload.cat4
  ]);

  return { saved: true };
}
