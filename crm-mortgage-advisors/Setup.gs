// ─── התקנה ואתחול מערכת ──────────────────────────────────────────────────────

/**
 * נקודת הכניסה הראשית – מוסיף את תפריט "ניהול מערכת CRM" לגיליון.
 * מופעל אוטומטית בכל פתיחת הגיליון.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚙️ ניהול מערכת CRM')
    .addItem('🚀 התקן מערכת (ריצה ראשונה)', 'runFirstTimeSetup')
    .addSeparator()
    .addItem('📁 פתח תיקייה ללקוח נבחר', 'createFolderForSelectedClient')
    .addItem('📋 העתק קישורי תקשורת לשורה', 'fillCommunicationLinks')
    .addSeparator()
    .addItem('📊 רענן דשבורד', 'refreshDashboard')
    .addItem('🔔 הפעל תזכורות מסמכים עכשיו', 'runDocumentReminders')
    .addSeparator()
    .addItem('⏰ הגדר טריגרים אוטומטיים', 'installTriggers')
    .addItem('🗑️ הסר טריגרים', 'removeTriggers')
    .addToUi();
}

/**
 * התקנה מלאה של המערכת – ריצה ראשונה בלבד.
 */
function runFirstTimeSetup() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.alert(
    'התקנת מערכת CRM',
    'הפעולה תיצור את גיליון הלקוחות, גיליון הדשבורד ולוג הפעולות.\nהאם להמשיך?',
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();

  createClientsSheet_(ss);
  createDashboardSheet_(ss);
  getLogSheet();
  ensureRootFolder_();
  installTriggers();

  ui.alert('✅ המערכת הותקנה בהצלחה!\n\nניתן להתחיל להזין לקוחות בגיליון "לקוחות".');
  writeLog('SETUP', '', 'התקנה ראשונית הושלמה');
}

// ─── יצירת גיליון לקוחות ─────────────────────────────────────────────────────

function createClientsSheet_(ss) {
  var existing = ss.getSheetByName(CONFIG.SHEETS.CLIENTS);
  if (existing) return existing;

  var sheet = ss.insertSheet(CONFIG.SHEETS.CLIENTS, 0);

  // כותרות
  var headers = [
    'מזהה לקוח', 'שם מלא', 'טלפון', 'מייל', 'מקור הגעה',
    'תאריך הצטרפות', 'בנק מטפל', 'סטטוס תיק', 'סטטוס מסמכים',
    'תאריך עדכון אחרון', 'קישור תיקייה', 'קישור פגישה', 'קישור WhatsApp', 'הערות'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // עיצוב כותרת
  styleHeaderRow_(sheet, 1, headers.length);

  // רוחב עמודות
  sheet.setColumnWidths(1,  1, 160); // מזהה
  sheet.setColumnWidths(2,  1, 160); // שם
  sheet.setColumnWidths(3,  1, 120); // טלפון
  sheet.setColumnWidths(4,  1, 200); // מייל
  sheet.setColumnWidths(5,  1, 140); // מקור
  sheet.setColumnWidths(6,  1, 130); // תאריך הצטרפות
  sheet.setColumnWidths(7,  1, 130); // בנק
  sheet.setColumnWidths(8,  1, 180); // סטטוס תיק
  sheet.setColumnWidths(9,  1, 160); // סטטוס מסמכים
  sheet.setColumnWidths(10, 1, 160); // תאריך עדכון
  sheet.setColumnWidths(11, 1, 60);  // תיקייה
  sheet.setColumnWidths(12, 1, 60);  // פגישה
  sheet.setColumnWidths(13, 1, 60);  // WhatsApp
  sheet.setColumnWidths(14, 1, 220); // הערות

  // הקפאת שורת כותרת
  sheet.setFrozenRows(1);

  // כיוון RTL
  sheet.setRightToLeft(true);

  // ולידציה לעמודת סטטוס תיק (עמודה H=8)
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.CASE_STATUSES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(DATA_START, COL.CASE_STATUS + 1, 500, 1).setDataValidation(statusRule);

  // ולידציה לסטטוס מסמכים (עמודה I=9)
  var docRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.DOC_STATUSES, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(DATA_START, COL.DOC_STATUS + 1, 500, 1).setDataValidation(docRule);

  // עיצוב מותנה – סטטוס תיק
  applyStatusConditionalFormatting_(sheet);

  return sheet;
}

function applyStatusConditionalFormatting_(sheet) {
  var statusColors = {
    'ליד חדש':                  { bg: '#e8f0fe', text: '#1a73e8' },
    'בתהליך איסוף מסמכים':     { bg: '#fff3e0', text: '#e65100' },
    'הוגש לאישור':              { bg: '#fce8b2', text: '#594300' },
    'התקבל אישור עקרוני':       { bg: '#e6f4ea', text: '#137333' },
    'הוגש לאישור סופי':         { bg: '#d2e3fc', text: '#1967d2' },
    'אושר סופי':                { bg: '#b7e1cd', text: '#0d652d' },
    'הועבר לנוטריון':           { bg: '#fde7f3', text: '#a50e0e' },
    'נחתם - עסקה סגורה':        { bg: '#37474f', text: '#ffffff' }
  };

  var rules = [];
  var statusCol = COL.CASE_STATUS + 1; // 1-based

  Object.keys(statusColors).forEach(function(status) {
    var c = statusColors[status];
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(status)
        .setBackground(c.bg)
        .setFontColor(c.text)
        .setRanges([sheet.getRange(DATA_START, statusCol, 500, 1)])
        .build()
    );
  });

  // עיצוב מותנה לסטטוס מסמכים
  var docCol = COL.DOC_STATUS + 1;
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('תקין מלא')
      .setBackground('#e6f4ea').setFontColor('#137333')
      .setRanges([sheet.getRange(DATA_START, docCol, 500, 1)])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('חלקי')
      .setBackground('#fff3e0').setFontColor('#e65100')
      .setRanges([sheet.getRange(DATA_START, docCol, 500, 1)])
      .build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('ממתין למסמכים')
      .setBackground('#fce8b2').setFontColor('#594300')
      .setRanges([sheet.getRange(DATA_START, docCol, 500, 1)])
      .build()
  );

  sheet.setConditionalFormatRules(rules);
}

// ─── תיקיית שורש ────────────────────────────────────────────────────────────

function ensureRootFolder_() {
  var folders = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  if (!folders.hasNext()) {
    DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);
  }
}

function getRootFolder_() {
  ensureRootFolder_();
  return DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME).next();
}

// ─── טריגרים ─────────────────────────────────────────────────────────────────

/**
 * מתקין את כל הטריגרים הנדרשים למערכת.
 */
function installTriggers() {
  removeTriggers(); // נקה ישנים תחילה

  // onChange – מזהה שורות חדשות ומפעיל לוגיקת onboarding
  ScriptApp.newTrigger('onSheetChange')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onChange()
    .create();

  // onEdit – מעדכן timestamp ושולח עדכוני סטטוס
  ScriptApp.newTrigger('onStatusEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();

  // CRON יומי בשעה 09:00 – בדיקת תזכורות מסמכים
  ScriptApp.newTrigger('runDocumentReminders')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();

  // CRON שבועי ראשון בבוקר – רענון דשבורד
  ScriptApp.newTrigger('refreshDashboard')
    .timeBased()
    .everyWeeks(1)
    .onWeekDay(ScriptApp.WeekDay.SUNDAY)
    .atHour(8)
    .create();

  SpreadsheetApp.getUi().alert('✅ הטריגרים הוגדרו בהצלחה!');
  writeLog('TRIGGERS', '', 'טריגרים הותקנו');
}

/**
 * מסיר את כל הטריגרים של הפרויקט.
 */
function removeTriggers() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
}
