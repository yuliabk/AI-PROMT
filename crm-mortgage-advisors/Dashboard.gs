// ─── תהליך 5: בקרת ניהול שבועית (Dashboard) ─────────────────────────────────

/**
 * מרענן את גיליון הדשבורד עם נתוני זמן אמת.
 * מופעל מהתפריט ו/או אוטומטית כל ראשון בבוקר.
 */
function refreshDashboard() {
  var ss           = SpreadsheetApp.getActiveSpreadsheet();
  var clientsSheet = getClientsSheet();
  var dashSheet    = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);

  if (!dashSheet) dashSheet = createDashboardSheet_(ss);

  var data = clientsSheet.getDataRange().getValues();
  var rows = data.slice(DATA_START - 1).filter(function(r) {
    return String(r[COL.NAME]).trim() !== '';
  });

  var totals         = buildTotalsSection_(rows);
  var byStatus       = buildPivot_(rows, COL.CASE_STATUS);
  var byBank         = buildPivot_(rows, COL.BANK);
  var byDocStatus    = buildPivot_(rows, COL.DOC_STATUS);
  var recentActivity = buildRecentActivity_(rows);

  renderDashboard_(dashSheet, totals, byStatus, byBank, byDocStatus, recentActivity);
  writeLog('DASHBOARD', '', 'דשבורד רוענן – ' + rows.length + ' לקוחות');

  if (SpreadsheetApp.getUi()) {
    try { SpreadsheetApp.getUi().alert('✅ הדשבורד רוענן!'); } catch (e) {}
  }
}

function createDashboardSheet_(ss) {
  var existing = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (existing) return existing;

  var sheet = ss.insertSheet(CONFIG.SHEETS.DASHBOARD);
  sheet.setRightToLeft(true);
  return sheet;
}

// ─── חישוב נתונים ────────────────────────────────────────────────────────────

function buildTotalsSection_(rows) {
  var total  = rows.length;
  var closed = rows.filter(function(r) {
    return String(r[COL.CASE_STATUS]) === 'נחתם - עסקה סגורה';
  }).length;
  var missingDocs = rows.filter(function(r) {
    var st = String(r[COL.CASE_STATUS]);
    return String(r[COL.DOC_STATUS]) !== 'תקין מלא' &&
           st !== 'נחתם - עסקה סגורה';
  }).length;
  var nearSigning = rows.filter(function(r) {
    var st = String(r[COL.CASE_STATUS]);
    return st === 'אושר סופי' || st === 'הועבר לנוטריון';
  }).length;

  return {
    total:       total,
    active:      total - closed,
    closed:      closed,
    missingDocs: missingDocs,
    nearSigning: nearSigning
  };
}

function buildPivot_(rows, colIndex) {
  var counts = {};
  rows.forEach(function(r) {
    var val = String(r[colIndex]).trim() || '(לא מוגדר)';
    counts[val] = (counts[val] || 0) + 1;
  });
  return Object.keys(counts)
    .sort(function(a, b) { return counts[b] - counts[a]; })
    .map(function(k) { return [k, counts[k]]; });
}

function buildRecentActivity_(rows) {
  return rows
    .filter(function(r) { return r[COL.LAST_UPDATED]; })
    .sort(function(a, b) {
      return new Date(b[COL.LAST_UPDATED]) - new Date(a[COL.LAST_UPDATED]);
    })
    .slice(0, 10)
    .map(function(r) {
      return [
        String(r[COL.NAME]),
        String(r[COL.CASE_STATUS]),
        String(r[COL.DOC_STATUS]),
        r[COL.LAST_UPDATED]
      ];
    });
}

// ─── ציור הדשבורד ────────────────────────────────────────────────────────────

function renderDashboard_(sheet, totals, byStatus, byBank, byDocStatus, recentActivity) {
  sheet.clearContents();
  sheet.clearFormats();

  var lastUpdated = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'dd/MM/yyyy HH:mm');

  // כותרת ראשית
  var titleRange = sheet.getRange(1, 1, 1, 6);
  titleRange.merge()
    .setValue('📊 דשבורד ניהול תיקי משכנתאות | עודכן: ' + lastUpdated)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontSize(14)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.setRowHeight(1, 40);

  var curRow = 3;

  // ─── סעיף 1: סיכום מהיר ──────────────────────────────────────────────────
  curRow = writeSectionHeader_(sheet, curRow, '⚡ סיכום מהיר', 6);

  var kpis = [
    { label: 'סה"כ לקוחות', value: totals.total,       bg: '#e8f0fe', text: '#1a73e8' },
    { label: 'תיקים פעילים', value: totals.active,      bg: '#e6f4ea', text: '#137333' },
    { label: 'עסקאות סגורות', value: totals.closed,     bg: '#37474f', text: '#ffffff' },
    { label: 'מסמכים חסרים', value: totals.missingDocs, bg: '#fce8b2', text: '#594300' },
    { label: 'קרובים לחתימה', value: totals.nearSigning, bg: '#fde7f3', text: '#9c27b0' }
  ];

  kpis.forEach(function(kpi, idx) {
    var col = idx + 1;
    sheet.getRange(curRow, col)
         .setValue(kpi.label)
         .setBackground(kpi.bg)
         .setFontColor(kpi.text)
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
    sheet.getRange(curRow + 1, col)
         .setValue(kpi.value)
         .setBackground(kpi.bg)
         .setFontColor(kpi.text)
         .setFontSize(24)
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
  });
  sheet.setRowHeight(curRow + 1, 50);
  curRow += 3;

  // ─── סעיף 2: תיקים לפי סטטוס ───────────────────────────────────────────
  curRow = writeSectionHeader_(sheet, curRow, '📋 תיקים לפי סטטוס', 6);
  curRow = writeTableHeader_(sheet, curRow, ['סטטוס תיק', 'מספר תיקים', 'אחוז מהסך'], 3);
  byStatus.forEach(function(r) {
    var pct = totals.total > 0 ? Math.round(r[1] / totals.total * 100) + '%' : '0%';
    sheet.getRange(curRow, 1).setValue(r[0]);
    sheet.getRange(curRow, 2).setValue(r[1]).setHorizontalAlignment('center');
    sheet.getRange(curRow, 3).setValue(pct).setHorizontalAlignment('center');
    curRow++;
  });
  curRow++;

  // ─── סעיף 3: תיקים לפי בנק ─────────────────────────────────────────────
  curRow = writeSectionHeader_(sheet, curRow, '🏦 תיקים לפי בנק מטפל', 6);
  curRow = writeTableHeader_(sheet, curRow, ['בנק מטפל', 'מספר תיקים'], 2);
  byBank.forEach(function(r) {
    sheet.getRange(curRow, 1).setValue(r[0]);
    sheet.getRange(curRow, 2).setValue(r[1]).setHorizontalAlignment('center');
    curRow++;
  });
  curRow++;

  // ─── סעיף 4: סטטוס מסמכים ─────────────────────────────────────────────
  curRow = writeSectionHeader_(sheet, curRow, '📄 סטטוס מסמכים', 6);
  curRow = writeTableHeader_(sheet, curRow, ['סטטוס מסמכים', 'מספר תיקים'], 2);
  byDocStatus.forEach(function(r) {
    sheet.getRange(curRow, 1).setValue(r[0]);
    sheet.getRange(curRow, 2).setValue(r[1]).setHorizontalAlignment('center');
    curRow++;
  });
  curRow++;

  // ─── סעיף 5: פעילות אחרונה ──────────────────────────────────────────────
  curRow = writeSectionHeader_(sheet, curRow, '🕐 10 הלקוחות האחרונים שעודכנו', 6);
  curRow = writeTableHeader_(sheet, curRow, ['שם לקוח', 'סטטוס תיק', 'סטטוס מסמכים', 'עדכון אחרון'], 4);
  recentActivity.forEach(function(r) {
    sheet.getRange(curRow, 1).setValue(r[0]);
    sheet.getRange(curRow, 2).setValue(r[1]);
    sheet.getRange(curRow, 3).setValue(r[2]);
    var dateVal = r[3] ? Utilities.formatDate(new Date(r[3]), 'Asia/Jerusalem', 'dd/MM/yyyy HH:mm') : '';
    sheet.getRange(curRow, 4).setValue(dateVal).setHorizontalAlignment('center');
    curRow++;
  });

  // הגדר רוחב עמודות דשבורד
  sheet.setColumnWidths(1, 1, 220);
  sheet.setColumnWidths(2, 1, 120);
  sheet.setColumnWidths(3, 1, 120);
  sheet.setColumnWidths(4, 1, 160);

  SpreadsheetApp.flush();
}

function writeSectionHeader_(sheet, row, title, numCols) {
  sheet.getRange(row, 1, 1, numCols).merge()
       .setValue(title)
       .setBackground('#e8f0fe')
       .setFontColor('#1a73e8')
       .setFontWeight('bold')
       .setFontSize(12)
       .setHorizontalAlignment('right');
  sheet.setRowHeight(row, 32);
  return row + 1;
}

function writeTableHeader_(sheet, row, headers, numCols) {
  for (var i = 0; i < headers.length; i++) {
    sheet.getRange(row, i + 1)
         .setValue(headers[i])
         .setBackground('#f1f3f4')
         .setFontWeight('bold')
         .setHorizontalAlignment('center');
  }
  return row + 1;
}
