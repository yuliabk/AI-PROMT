// ─── פונקציות עזר כלליות ──────────────────────────────────────────────────────

/**
 * מחזיר את גיליון הלקוחות הראשי.
 * יוצר אותו אם אינו קיים.
 */
function getClientsSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.CLIENTS);
  if (!sheet) sheet = createClientsSheet_(ss);
  return sheet;
}

/**
 * מחזיר את גיליון הלוג.
 */
function getLogSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.LOG);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEETS.LOG);
    sheet.appendRow(['תאריך', 'פעולה', 'לקוח', 'פרטים']);
    sheet.setFrozenRows(1);
    styleHeaderRow_(sheet, 1, 4);
  }
  return sheet;
}

/**
 * כותב שורה ללוג הפעולות.
 */
function writeLog(action, clientName, details) {
  try {
    var logSheet = getLogSheet();
    logSheet.appendRow([
      new Date(),
      action,
      clientName || '',
      details   || ''
    ]);
  } catch (e) {
    // לוג לא חיוני – לא מפסיקים את הביצוע
    console.error('writeLog error:', e.message);
  }
}

/**
 * מייצר מזהה לקוח ייחודי בפורמט CRM-YYYYMMDD-XXXX.
 */
function generateClientId() {
  var date   = Utilities.formatDate(new Date(), CONFIG.timeZone || 'Asia/Jerusalem', 'yyyyMMdd');
  var random = Math.floor(1000 + Math.random() * 9000);
  return 'CRM-' + date + '-' + random;
}

/**
 * מאתר שורת לקוח לפי ID.
 * מחזיר אינדקס שורה (1-based) או -1 אם לא נמצא.
 */
function findClientRowById(clientId) {
  var sheet  = getClientsSheet();
  var data   = sheet.getDataRange().getValues();
  for (var i = DATA_START - 1; i < data.length; i++) {
    if (String(data[i][COL.ID]) === String(clientId)) return i + 1;
  }
  return -1;
}

/**
 * מחזיר את נתוני השורה של לקוח לפי אינדקס (1-based).
 */
function getClientRowData(rowIndex) {
  var sheet = getClientsSheet();
  return sheet.getRange(rowIndex, 1, 1, TOTAL_COLUMNS).getValues()[0];
}

/**
 * מנרמל מספר טלפון לפורמט בינלאומי ישראלי (972...).
 */
function normalizePhone(phone) {
  var p = String(phone).replace(/\D/g, '');
  if (p.startsWith('0')) p = '972' + p.slice(1);
  if (!p.startsWith('972')) p = '972' + p;
  return p;
}

/**
 * בונה קישור WhatsApp Web למספר נתון.
 */
function buildWhatsAppLink(phone) {
  return 'https://wa.me/' + normalizePhone(phone);
}

/**
 * בונה קישור יצירת אירוע מהיר ב-Google Calendar.
 */
function buildCalendarLink(clientName) {
  var title = encodeURIComponent('פגישה עם ' + clientName);
  return 'https://calendar.google.com/calendar/r/eventedit?text=' + title;
}

/**
 * שולח מייל HTML ללקוח.
 */
function sendHtmlEmail(toEmail, subject, htmlBody) {
  if (!toEmail || !isValidEmail_(toEmail)) {
    console.warn('sendHtmlEmail: כתובת מייל לא תקינה –', toEmail);
    return false;
  }
  try {
    GmailApp.sendEmail(toEmail, subject, '', { htmlBody: htmlBody, name: 'יועץ המשכנתאות שלך' });
    return true;
  } catch (e) {
    console.error('sendHtmlEmail error:', e.message);
    return false;
  }
}

/**
 * שולח webhook לכלי אוטומציה חיצוני (Make / n8n).
 */
function triggerWebhook(payload) {
  if (!CONFIG.WEBHOOK_URL) return;
  try {
    UrlFetchApp.fetch(CONFIG.WEBHOOK_URL, {
      method:      'post',
      contentType: 'application/json',
      payload:     JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    console.error('triggerWebhook error:', e.message);
  }
}

// ─── פונקציות פרטיות ─────────────────────────────────────────────────────────

function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email));
}

function styleHeaderRow_(sheet, row, numCols) {
  var range = sheet.getRange(row, 1, 1, numCols);
  range.setBackground('#1a73e8')
       .setFontColor('#ffffff')
       .setFontWeight('bold')
       .setHorizontalAlignment('center');
}

function addDataValidation_(sheet, row, col, values) {
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(row, col).setDataValidation(rule);
}
