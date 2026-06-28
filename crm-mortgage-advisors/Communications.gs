// ─── תהליך 4: תקשורת יזומה וסנכרון יומנים (Micro-Actions) ──────────────────
//
// הקישורים נוצרים אוטומטית בשלב ה-Onboarding (ר' Onboarding.gs).
// קובץ זה מכיל פונקציות עזר ופונקציות שניתן לקרוא מהתפריט.

/**
 * פותח קישור WhatsApp Web ישירות לאיש קשר בשורה הנבחרת.
 * (מייצר URL – המשתמש לוחץ על תא הגיליון)
 */
function openWhatsAppForSelectedClient() {
  var sheet    = getClientsSheet();
  var ui       = SpreadsheetApp.getUi();
  var rowIndex = sheet.getActiveCell().getRow();

  if (rowIndex < DATA_START) { ui.alert('⚠️ בחר שורת לקוח.'); return; }

  var rowData = getClientRowData(rowIndex);
  var phone   = String(rowData[COL.PHONE]).trim();
  var name    = String(rowData[COL.NAME]).trim();

  if (!phone) { ui.alert('⚠️ מספר טלפון חסר ללקוח ' + name + '.'); return; }

  var url = buildWhatsAppLink(phone);
  ui.alert('🔗 קישור WhatsApp ל-' + name + ':\n\n' + url + '\n\nהעתק/י את הקישור ופתח/י בדפדפן.');
}

/**
 * פותח Google Calendar ליצירת פגישה עם הלקוח הנבחר.
 */
function openCalendarForSelectedClient() {
  var sheet    = getClientsSheet();
  var ui       = SpreadsheetApp.getUi();
  var rowIndex = sheet.getActiveCell().getRow();

  if (rowIndex < DATA_START) { ui.alert('⚠️ בחר שורת לקוח.'); return; }

  var rowData = getClientRowData(rowIndex);
  var name    = String(rowData[COL.NAME]).trim();

  if (!name) { ui.alert('⚠️ שם הלקוח חסר.'); return; }

  var url = buildCalendarLink(name);
  ui.alert('🗓️ קישור פגישה עם ' + name + ':\n\n' + url + '\n\nהעתק/י את הקישור ופתח/י בדפדפן.');
}

/**
 * שולח מייל מותאם אישית ללקוח הנבחר (תבנית חופשית).
 */
function sendCustomEmailToSelectedClient() {
  var sheet    = getClientsSheet();
  var ui       = SpreadsheetApp.getUi();
  var rowIndex = sheet.getActiveCell().getRow();

  if (rowIndex < DATA_START) { ui.alert('⚠️ בחר שורת לקוח.'); return; }

  var rowData = getClientRowData(rowIndex);
  var name    = String(rowData[COL.NAME]).trim();
  var email   = String(rowData[COL.EMAIL]).trim();

  if (!email) { ui.alert('⚠️ כתובת מייל חסרה ללקוח ' + name + '.'); return; }

  var subjectResp = ui.prompt('📧 שליחת מייל ל-' + name, 'הכנס נושא ההודעה:', ui.ButtonSet.OK_CANCEL);
  if (subjectResp.getSelectedButton() !== ui.Button.OK) return;

  var bodyResp = ui.prompt('📧 תוכן ההודעה', 'הכנס את תוכן ההודעה:', ui.ButtonSet.OK_CANCEL);
  if (bodyResp.getSelectedButton() !== ui.Button.OK) return;

  var subject = subjectResp.getResponseText();
  var body    = bodyResp.getResponseText();

  if (!subject || !body) { ui.alert('⚠️ נושא ותוכן הם שדות חובה.'); return; }

  var html = buildCustomEmailHtml_(name, body);
  var ok   = sendHtmlEmail(email, subject, html);

  if (ok) {
    writeLog('CUSTOM_EMAIL', name, 'נושא: ' + subject);
    ui.alert('✅ המייל נשלח בהצלחה ל-' + email + '!');
  } else {
    ui.alert('❌ שגיאה בשליחת המייל. בדוק שכתובת המייל תקינה.');
  }
}

function buildCustomEmailHtml_(clientName, bodyText) {
  var escaped = bodyText.replace(/\n/g, '<br>');
  return [
    '<div dir="rtl" style="font-family:Arial,sans-serif;max-width:600px;margin:auto">',
    '  <div style="background:#1a73e8;padding:20px;text-align:center">',
    '    <h2 style="color:#fff;margin:0">הודעה מיועץ המשכנתאות שלך</h2>',
    '  </div>',
    '  <div style="padding:24px">',
    '    <p>שלום <strong>' + clientName + '</strong>,</p>',
    '    <p>' + escaped + '</p>',
    '  </div>',
    '  <div style="background:#f8f9fa;padding:12px;text-align:center;font-size:12px;color:#666">',
    '    מייל זה נשלח ממערכת CRM.',
    '  </div>',
    '</div>'
  ].join('\n');
}
