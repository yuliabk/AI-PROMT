// ─── תהליך 3: ניהול מחזור חיי עסקת המשכנתא ─────────────────────────────────

/**
 * טריגר onEdit – מופעל בכל עריכה ידנית בגיליון.
 * מטפל בעדכון timestamp ושליחת עדכוני סטטוס ללקוח.
 */
function onStatusEdit(e) {
  if (!e) return;

  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEETS.CLIENTS) return;

  var range = e.range;
  var col   = range.getColumn();    // 1-based
  var row   = range.getRow();       // 1-based

  if (row < DATA_START) return; // התעלם משורת כותרת

  // שלב 2: חותמת זמן בכל עריכה בגיליון הלקוחות
  sheet.getRange(row, COL.LAST_UPDATED + 1).setValue(new Date());

  // שלב 1+3: שינוי בעמודת סטטוס תיק
  if (col === COL.CASE_STATUS + 1) {
    var newStatus = String(e.value || '').trim();
    var oldStatus = String(e.oldValue || '').trim();

    if (newStatus && newStatus !== oldStatus) {
      handleCaseStatusChange_(sheet, row, newStatus, oldStatus);
    }
  }
}

/**
 * מטפל בשינוי סטטוס תיק: שולח מייל ומפעיל webhook.
 */
function handleCaseStatusChange_(sheet, row, newStatus, oldStatus) {
  var rowData   = getClientRowData(row);
  var clientId  = String(rowData[COL.ID]).trim();
  var name      = String(rowData[COL.NAME]).trim();
  var email     = String(rowData[COL.EMAIL]).trim();

  if (!name) return;

  // שלב 3: שליחת מייל עדכון ללקוח
  if (email) {
    var subject  = buildStatusEmailSubject_(newStatus);
    var htmlBody = buildStatusEmailBody_(name, newStatus, oldStatus);
    var sent     = sendHtmlEmail(email, subject, htmlBody);
    if (sent) writeLog('STATUS_EMAIL_SENT', name, newStatus);
  }

  // Webhook למידלוור (Make/n8n)
  triggerWebhook({
    event:     'STATUS_CHANGE',
    clientId:  clientId,
    name:      name,
    email:     email,
    phone:     rowData[COL.PHONE],
    oldStatus: oldStatus,
    newStatus: newStatus,
    timestamp: new Date().toISOString()
  });

  writeLog('STATUS_CHANGE', name, oldStatus + ' → ' + newStatus);
}

// ─── בניית תוכן המייל ────────────────────────────────────────────────────────

var STATUS_META = {
  'ליד חדש': {
    icon: '👋',
    title: 'פתחנו את התיק שלך',
    meaning: 'קיבלנו את פרטיך והתחלנו בעבודה על התיק.',
    next: 'נפנה אליך בקרוב לתיאום שיחת ייעוץ ראשונית וריכוז המסמכים הנדרשים.'
  },
  'בתהליך איסוף מסמכים': {
    icon: '📄',
    title: 'אנחנו ממתינים למסמכים שלך',
    meaning: 'פרטי הבקשה שלך מוכנים. כעת נדרשים המסמכים הפיננסיים להגשה.',
    next: 'אנא השלם/י את העלאת המסמכים הנדרשים (תלושי שכר, עו"ש, נסח טאבו) דרך הקישור שנשלח אליך.'
  },
  'הוגש לאישור': {
    icon: '📬',
    title: 'תיקך הוגש לבנק!',
    meaning: 'שלחנו את בקשת המשכנתא לבנק לבחינה ועיבוד.',
    next: 'הליך הבחינה בבנק אורך בדרך כלל 5–10 ימי עסקים. נעדכן אותך בכל התפתחות.'
  },
  'התקבל אישור עקרוני': {
    icon: '🎉',
    title: 'אישור עקרוני התקבל!',
    meaning: 'הבנק אישר את עמידתך בקריטריונים הבסיסיים לקבלת המשכנתא.',
    next: 'נמשיך בהשלמת האישור הסופי. אנא שמור/י על הרמה הפיננסית הנוכחית עד לסגירת העסקה.'
  },
  'הוגש לאישור סופי': {
    icon: '📝',
    title: 'הגשנו לאישור הסופי',
    meaning: 'כל המסמכים הוגשו לאישור סופי בבנק.',
    next: 'הבנק מסיים את הבדיקות. תהליך זה אורך בדרך כלל 3–7 ימי עסקים.'
  },
  'אושר סופי': {
    icon: '✅',
    title: 'המשכנתא אושרה סופית!',
    meaning: 'הבנק אישר את המשכנתא שלך באופן סופי.',
    next: 'נתאם פגישה לחתימה על המסמכים הסופיים מול הנוטריון.'
  },
  'הועבר לנוטריון': {
    icon: '⚖️',
    title: 'העברנו לנוטריון',
    meaning: 'התיק הועבר לנוטריון לצורך חתימת המסמכים הסופיים.',
    next: 'נצור קשר בהקדם לתיאום מועד נוח לחתימה.'
  },
  'נחתם - עסקה סגורה': {
    icon: '🏡',
    title: 'מזל טוב! העסקה נסגרה',
    meaning: 'החוזה נחתם והמשכנתא נרשמה. העסקה הושלמה בהצלחה.',
    next: 'תודה על הבטחון שנתת לנו. אנחנו כאן לכל שאלה גם בעתיד!'
  }
};

function buildStatusEmailSubject_(status) {
  var meta = STATUS_META[status];
  var icon = meta ? meta.icon : '📌';
  return icon + ' עדכון תיק משכנתא: ' + status;
}

function buildStatusEmailBody_(clientName, newStatus, oldStatus) {
  var meta    = STATUS_META[newStatus] || {};
  var icon    = meta.icon    || '📌';
  var title   = meta.title   || 'עדכון סטטוס תיק';
  var meaning = meta.meaning || 'הסטטוס של תיקך עודכן.';
  var next    = meta.next    || 'נציג שלנו יצור אתך קשר בהקדם.';

  return [
    '<div dir="rtl" style="font-family:Arial,sans-serif;max-width:600px;margin:auto;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden">',
    '  <div style="background:#1a73e8;padding:24px;text-align:center">',
    '    <h1 style="color:#fff;margin:0;font-size:22px">' + icon + ' ' + title + '</h1>',
    '  </div>',
    '  <div style="padding:24px">',
    '    <p style="font-size:16px">שלום <strong>' + clientName + '</strong>,</p>',
    '    <p>רצינו לעדכן אותך שסטטוס תיק המשכנתא שלך עודכן ל:</p>',
    '    <div style="background:#e8f0fe;border-right:4px solid #1a73e8;padding:12px 16px;border-radius:4px;margin:16px 0">',
    '      <strong style="font-size:18px;color:#1a73e8">' + newStatus + '</strong>',
    '    </div>',
    '    <h3 style="color:#1a73e8">מה זה אומר?</h3>',
    '    <p>' + meaning + '</p>',
    '    <h3 style="color:#1a73e8">מה קורה עכשיו?</h3>',
    '    <p>' + next + '</p>',
    '  </div>',
    '  <div style="background:#f8f9fa;padding:16px;text-align:center;font-size:12px;color:#666">',
    '    <p>מייל זה נשלח אוטומטית ממערכת CRM יועץ המשכנתאות שלך.</p>',
    '    <p>לשאלות ניתן לפנות ישירות לכתובת מייל זו.</p>',
    '  </div>',
    '</div>'
  ].join('\n');
}
