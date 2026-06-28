# CRM יועצי משכנתאות – Google Apps Script

מערכת CRM מלאה המבוססת על **Google Sheets + Apps Script** לניהול תיקי לקוחות משכנתא.

---

## קבצי הפרויקט

| קובץ | תיאור |
|------|-------|
| `appsscript.json` | מניפסט הפרויקט והרשאות OAuth |
| `Config.gs` | קבועים, אינדקסי עמודות, רשימות סטטוסים |
| `Utils.gs` | פונקציות עזר משותפות (מייל, webhook, לוג) |
| `Setup.gs` | התקנה, תפריט, טריגרים, יצירת גיליונות |
| `Onboarding.gs` | תהליך 1 – קליטת לקוח חדש |
| `Documents.gs` | תהליך 2 – תזכורות מסמכים |
| `Lifecycle.gs` | תהליך 3 – מחזור חיי עסקה ועדכוני מייל |
| `Communications.gs` | תהליך 4 – WhatsApp / Calendar |
| `Dashboard.gs` | תהליך 5 – דשבורד שבועי |

---

## הגדרת הפרויקט (ריצה ראשונה)

### שלב 1 – יצירת פרויקט Apps Script

1. פתח **Google Sheets** חדש.
2. עבור לתפריט **Extensions → Apps Script**.
3. מחק את הקוד הקיים ב-`Code.gs`.
4. צור קובץ חדש לכל `.gs` בפרויקט זה, העתק תוכן כל קובץ.
5. עדכן את `appsscript.json` (כפתור ⚙️ → Project Settings → Show "appsscript.json").

### שלב 2 – הגדרת Webhook (Make / n8n)

ב-`Config.gs`, מלא את שדה `WEBHOOK_URL`:

```javascript
WEBHOOK_URL: 'https://hook.make.com/YOUR_HOOK_ID',
```

ה-Webhook מקבל אובייקט JSON עם `event` מסוג:
- `NEW_CLIENT` – לקוח חדש נקלט
- `STATUS_CHANGE` – סטטוס תיק השתנה
- `DOCUMENT_REMINDER` – תזכורת מסמכים יומית

### שלב 3 – התקנה

1. שמור את כל הקבצים ב-Apps Script.
2. ב-Google Sheets, לחץ על תפריט **⚙️ ניהול מערכת CRM**.
3. בחר **🚀 התקן מערכת (ריצה ראשונה)**.
4. אשר את ההרשאות המבוקשות.

---

## מבנה גיליון הלקוחות

| עמודה | שם | תיאור |
|-------|----|-------|
| A | מזהה לקוח | נוצר אוטומטית (CRM-YYYYMMDD-XXXX) |
| B | שם מלא | שם הלקוח |
| C | טלפון | מספר ישראלי (050...) |
| D | מייל | לשליחת עדכונים אוטומטיים |
| E | מקור הגעה | איך הגיע הלקוח |
| F | תאריך הצטרפות | נמלא אוטומטית |
| G | בנק מטפל | שם הבנק הרלוונטי |
| H | סטטוס תיק | רשימה נפתחת עם עיצוב מותנה |
| I | סטטוס מסמכים | ממתין / חלקי / תקין מלא |
| J | תאריך עדכון אחרון | נמלא אוטומטית בכל עריכה |
| K | קישור תיקייה | 📁 Google Drive |
| L | קישור פגישה | 📅 Google Calendar |
| M | קישור WhatsApp | 💬 WhatsApp Web |
| N | הערות | טקסט חופשי |

---

## סטטוסי תיק

```
ליד חדש → בתהליך איסוף מסמכים → הוגש לאישור →
התקבל אישור עקרוני → הוגש לאישור סופי →
אושר סופי → הועבר לנוטריון → נחתם - עסקה סגורה
```

שינוי סטטוס שולח **מייל HTML אוטומטי** ללקוח עם הסבר המצב ומה צפוי הלאה.

---

## טריגרים אוטומטיים

| טריגר | תזמון | פעולה |
|--------|-------|-------|
| `onSheetChange` | onChange | זיהוי לקוח חדש + Onboarding |
| `onStatusEdit` | onEdit | עדכון timestamp + מייל סטטוס |
| `runDocumentReminders` | יומי 09:00 | תזכורות מסמכים חסרים |
| `refreshDashboard` | ראשון 08:00 | רענון דשבורד |

---

## אירועי Webhook

### NEW_CLIENT
```json
{
  "event": "NEW_CLIENT",
  "clientId": "CRM-20240601-4827",
  "name": "ישראל ישראלי",
  "phone": "0501234567",
  "email": "israel@example.com",
  "source": "אתר אינטרנט",
  "timestamp": "2024-06-01T09:30:00.000Z"
}
```

### STATUS_CHANGE
```json
{
  "event": "STATUS_CHANGE",
  "clientId": "CRM-20240601-4827",
  "name": "ישראל ישראלי",
  "phone": "0501234567",
  "email": "israel@example.com",
  "oldStatus": "הוגש לאישור",
  "newStatus": "התקבל אישור עקרוני",
  "timestamp": "2024-06-05T14:22:00.000Z"
}
```

### DOCUMENT_REMINDER
```json
{
  "event": "DOCUMENT_REMINDER",
  "clientId": "CRM-20240601-4827",
  "name": "ישראל ישראלי",
  "phone": "0501234567",
  "docStatus": "חלקי",
  "caseStatus": "בתהליך איסוף מסמכים",
  "reminderType": "partial",
  "timestamp": "2024-06-09T09:00:00.000Z"
}
```

---

## ממשק Make / n8n לשליחת WhatsApp

בתרחיש Make/n8n, הגדר Router עם 3 מסלולים לפי `event`:

1. **NEW_CLIENT** → שלח WhatsApp ברכה + קישור העלאת מסמכים
2. **DOCUMENT_REMINDER** → שלח WhatsApp תזכורת מסמכים
3. **STATUS_CHANGE** → שלח WhatsApp עדכון קצר + לינק לאימייל המפורט
