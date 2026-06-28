// ─── תהליך 2: ניהול ואיסוף מסמכים פיננסיים ─────────────────────────────────

/**
 * פונקציית CRON – מופעלת יומית ובודקת מי צריך תזכורת מסמכים.
 * נשלחת תזכורת רק לתיקים פתוחים שסטטוס מסמכיהם אינו "תקין מלא".
 */
function runDocumentReminders() {
  var today  = new Date();
  var dayOfWeek = today.getDay(); // 0=ראשון, 6=שבת

  // שלח תזכורות רק בימים המוגדרים
  if (CONFIG.REMINDER_DAYS.indexOf(dayOfWeek) === -1) return;

  var sheet  = getClientsSheet();
  var data   = sheet.getDataRange().getValues();
  var sent   = 0;
  var errors = 0;

  for (var i = DATA_START - 1; i < data.length; i++) {
    var row       = data[i];
    var name      = String(row[COL.NAME]).trim();
    var email     = String(row[COL.EMAIL]).trim();
    var caseStatus = String(row[COL.CASE_STATUS]).trim();
    var docStatus  = String(row[COL.DOC_STATUS]).trim();

    if (!name) continue;

    // לא מזכיר לתיקים סגורים
    var closedStatuses = ['נחתם - עסקה סגורה'];
    if (closedStatuses.indexOf(caseStatus) !== -1) continue;

    // רק אם מסמכים חסרים
    if (docStatus === 'תקין מלא') continue;

    var reminderType = docStatus === 'חלקי' ? 'partial' : 'missing';

    if (email) {
      var subject = buildReminderEmailSubject_(reminderType);
      var html    = buildReminderEmailBody_(name, docStatus, caseStatus);
      var ok      = sendHtmlEmail(email, subject, html);
      if (ok) {
        sent++;
        writeLog('REMINDER_SENT', name, docStatus + ' → מייל נשלח');
      } else {
        errors++;
      }
    }

    // Webhook ל-WhatsApp (דרך Make/n8n)
    var phone = String(row[COL.PHONE]).trim();
    if (phone) {
      triggerWebhook({
        event:        'DOCUMENT_REMINDER',
        clientId:     row[COL.ID],
        name:         name,
        phone:        phone,
        email:        email,
        docStatus:    docStatus,
        caseStatus:   caseStatus,
        reminderType: reminderType,
        timestamp:    new Date().toISOString()
      });
    }
  }

  writeLog('REMINDER_RUN', '', 'סה"כ תזכורות: נשלחו=' + sent + ', שגיאות=' + errors);
  console.log('Document reminders run: sent=' + sent + ', errors=' + errors);
}

// ─── בניית מייל תזכורת ───────────────────────────────────────────────────────

function buildReminderEmailSubject_(type) {
  if (type === 'partial') {
    return '📎 תזכורת: השלמת מסמכים לתיק המשכנתא שלך';
  }
  return '⚠️ דרוש פעולה: מסמכים חסרים בתיק המשכנתא שלך';
}

function buildReminderEmailBody_(clientName, docStatus, caseStatus) {
  var isPartial = (docStatus === 'חלקי');

  var headerBg    = isPartial ? '#fff3e0' : '#fce8b2';
  var headerColor = isPartial ? '#e65100' : '#594300';
  var icon        = isPartial ? '📎' : '⚠️';
  var titleText   = isPartial ? 'נדרשת השלמת מסמכים' : 'מסמכים חסרים – דרוש טיפול';
  var bodyText    = isPartial
    ? 'בתיקך נמצאים <strong>מסמכים חלקיים</strong>. כדי לקדם את הגשת הבקשה לבנק, יש להשלים את המסמכים החסרים בהקדם.'
    : 'עדיין <strong>לא התקבלו מסמכים</strong> לתיקך. ללא מסמכים לא נוכל להגיש את הבקשה לבנק.';

  var requiredDocs = [
    '✅ תלושי שכר – 3 חודשים אחרונים',
    '✅ דפי עו"ש – 3 חודשים אחרונים',
    '✅ נסח טאבו של הנכס',
    '✅ תעודות זהות של כל הלווים',
    '✅ אישור יתרת חסכונות / קרנות'
  ];

  return [
    '<div dir="rtl" style="font-family:Arial,sans-serif;max-width:600px;margin:auto;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden">',
    '  <div style="background:' + headerBg + ';padding:24px;text-align:center;border-bottom:3px solid ' + headerColor + '">',
    '    <h1 style="color:' + headerColor + ';margin:0;font-size:20px">' + icon + ' ' + titleText + '</h1>',
    '  </div>',
    '  <div style="padding:24px">',
    '    <p style="font-size:16px">שלום <strong>' + clientName + '</strong>,</p>',
    '    <p>' + bodyText + '</p>',
    '    <p>סטטוס תיק נוכחי: <strong>' + caseStatus + '</strong></p>',
    '    <h3 style="color:#1a73e8">מסמכים נדרשים:</h3>',
    '    <ul style="line-height:1.8">',
    requiredDocs.map(function(d) { return '      <li>' + d + '</li>'; }).join('\n'),
    '    </ul>',
    '    <div style="background:#e8f0fe;border-radius:8px;padding:16px;margin-top:20px;text-align:center">',
    '      <p style="margin:0;font-size:15px"><strong>📤 להעלאת מסמכים – פנה/י לנציג שלך</strong></p>',
    '    </div>',
    '  </div>',
    '  <div style="background:#f8f9fa;padding:16px;text-align:center;font-size:12px;color:#666">',
    '    <p>מייל זה נשלח אוטומטית ממערכת CRM יועץ המשכנתאות שלך.</p>',
    '  </div>',
    '</div>'
  ].join('\n');
}
