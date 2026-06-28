// ─── תהליך 1: קליטת לקוח חדש (Onboarding) ───────────────────────────────────

/**
 * טריגר onChange – מופעל בכל שינוי מבנה בגיליון (הוספת שורה).
 */
function onSheetChange(e) {
  if (e.changeType !== 'INSERT_ROW') return;
  processNewClientRows_();
}

/**
 * סורק שורות ללא מזהה לקוח ומבצע את תהליך ה-Onboarding.
 */
function processNewClientRows_() {
  var sheet = getClientsSheet();
  var data  = sheet.getDataRange().getValues();

  for (var i = DATA_START - 1; i < data.length; i++) {
    var row  = data[i];
    var id   = String(row[COL.ID]).trim();
    var name = String(row[COL.NAME]).trim();

    // דלג על שורות ריקות או שכבר עובדו
    if (!name || id.startsWith('CRM-')) continue;

    var rowIndex = i + 1; // 1-based
    onboardNewClient_(sheet, rowIndex, row);
  }
}

/**
 * מבצע את כל שלבי הקליטה עבור שורת לקוח חדשה.
 */
function onboardNewClient_(sheet, rowIndex, rowData) {
  var clientName = String(rowData[COL.NAME]).trim();

  // שלב 2: הנפקת מזהה
  var clientId = generateClientId();
  sheet.getRange(rowIndex, COL.ID + 1).setValue(clientId);

  // תאריך הצטרפות
  if (!rowData[COL.JOIN_DATE]) {
    sheet.getRange(rowIndex, COL.JOIN_DATE + 1).setValue(new Date());
  }

  // סטטוס ברירת מחדל
  if (!rowData[COL.CASE_STATUS]) {
    sheet.getRange(rowIndex, COL.CASE_STATUS + 1).setValue('ליד חדש');
  }
  if (!rowData[COL.DOC_STATUS]) {
    sheet.getRange(rowIndex, COL.DOC_STATUS + 1).setValue('ממתין למסמכים');
  }

  // עדכון timestamp
  sheet.getRange(rowIndex, COL.LAST_UPDATED + 1).setValue(new Date());

  // שלב 4א: קישורי תקשורת
  var phone = String(rowData[COL.PHONE]).trim();
  if (phone) {
    var waLink = buildWhatsAppLink(phone);
    sheet.getRange(rowIndex, COL.WA_LINK + 1)
         .setFormula('=HYPERLINK("' + waLink + '","💬 WhatsApp")');
  }

  var calLink = buildCalendarLink(clientName);
  sheet.getRange(rowIndex, COL.MEETING_LINK + 1)
       .setFormula('=HYPERLINK("' + calLink + '","📅 פגישה")');

  // שלב 4ב: Webhook ל-Make/n8n (WhatsApp ברכה + קישור העלאת מסמכים)
  var email = String(rowData[COL.EMAIL]).trim();
  triggerWebhook({
    event:     'NEW_CLIENT',
    clientId:  clientId,
    name:      clientName,
    phone:     phone,
    email:     email,
    source:    rowData[COL.SOURCE],
    timestamp: new Date().toISOString()
  });

  writeLog('ONBOARDING', clientName, 'לקוח חדש נקלט – ' + clientId);
  SpreadsheetApp.flush();
}

// ─── פתיחת תיקייה ידנית (מתפריט) ────────────────────────────────────────────

/**
 * שלב 3: פתיחת תיקייה ב-Google Drive עבור הלקוח בשורה הנבחרת.
 * הפונקציה מופעלת מתפריט "ניהול מערכת CRM".
 */
function createFolderForSelectedClient() {
  var sheet = getClientsSheet();
  var ui    = SpreadsheetApp.getUi();
  var rowIndex = sheet.getActiveCell().getRow();

  if (rowIndex < DATA_START) {
    ui.alert('⚠️ אנא בחר שורת לקוח (לא שורת כותרת).');
    return;
  }

  var rowData = getClientRowData(rowIndex);
  var id      = String(rowData[COL.ID]).trim();
  var name    = String(rowData[COL.NAME]).trim();

  if (!name) {
    ui.alert('⚠️ שם הלקוח חסר בשורה זו.');
    return;
  }

  // בדוק אם תיקייה כבר קיימת
  var existingLink = String(rowData[COL.FOLDER_LINK]).trim();
  if (existingLink && existingLink.startsWith('http')) {
    ui.alert('ℹ️ תיקייה כבר קיימת ללקוח זה.\n' + existingLink);
    return;
  }

  var folderName  = id + ' – ' + name;
  var rootFolder  = getRootFolder_();

  // בדוק כפילות בדרייב
  var existingFolders = rootFolder.getFoldersByName(folderName);
  var clientFolder;
  if (existingFolders.hasNext()) {
    clientFolder = existingFolders.next();
  } else {
    clientFolder = rootFolder.createFolder(folderName);

    // תתי-תיקיות לסדר
    clientFolder.createFolder('01 – תלושי שכר ועו"ש');
    clientFolder.createFolder('02 – נסח טאבו ומסמכי נכס');
    clientFolder.createFolder('03 – הצעות בנקים');
    clientFolder.createFolder('04 – אישורים וחוזים');
    clientFolder.createFolder('05 – כללי');
  }

  // כתוב קישור תיקייה בחזרה לגיליון
  var folderUrl = clientFolder.getUrl();
  sheet.getRange(rowIndex, COL.FOLDER_LINK + 1)
       .setFormula('=HYPERLINK("' + folderUrl + '","📁 תיקייה")');

  writeLog('FOLDER_CREATED', name, folderUrl);
  ui.alert('✅ תיקייה נפתחה בהצלחה!\n\n' + folderUrl);
}

// ─── מילוי קישורי תקשורת ─────────────────────────────────────────────────────

/**
 * ממלא קישורי WhatsApp ו-Calendar לשורה הנבחרת.
 */
function fillCommunicationLinks() {
  var sheet    = getClientsSheet();
  var ui       = SpreadsheetApp.getUi();
  var rowIndex = sheet.getActiveCell().getRow();

  if (rowIndex < DATA_START) {
    ui.alert('⚠️ אנא בחר שורת לקוח.');
    return;
  }

  var rowData = getClientRowData(rowIndex);
  var name    = String(rowData[COL.NAME]).trim();
  var phone   = String(rowData[COL.PHONE]).trim();

  if (!name) { ui.alert('⚠️ שם הלקוח חסר.'); return; }

  if (phone) {
    sheet.getRange(rowIndex, COL.WA_LINK + 1)
         .setFormula('=HYPERLINK("' + buildWhatsAppLink(phone) + '","💬 WhatsApp")');
  }

  sheet.getRange(rowIndex, COL.MEETING_LINK + 1)
       .setFormula('=HYPERLINK("' + buildCalendarLink(name) + '","📅 פגישה")');

  SpreadsheetApp.flush();
  ui.alert('✅ הקישורים עודכנו בשורה ' + rowIndex + '.');
}
