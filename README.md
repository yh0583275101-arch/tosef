# עורך הטקסט התוראני - Torani Text Editor
## Microsoft Word Web Add-in (Office.js) for Hebrew Text Editing

---

## 📋 תיאור הפרויקט

עורך הטקסט התוראני הוא תוסף Office.js עבור Microsoft Word המאפשר עריכה מקצועית של טקסטים בעברית. התוסף משתמש ב-Gemini AI לביצוע עריכה מתקדמת כולל:

- **סגנון יישיבה** - המרה לטון נלמדני וחקיקתי
- **פיסוק (Pissuk)** - תיקוני פיסוק לפי סטנדרט הוצאות
- **עיצוב (Formatting)** - שיפור פריסת המסמך
- **ניקוד (Nikkud)** - הוספת אותיות ניקוד לפי בקשה
- **תצוגה מקדימה** - הצגת שינויים לפני החלתם

---

## 🚀 התחלה מהירה

### 1. דרישות מוקדמות
- **Microsoft Word** (Online או Desktop עם Office.js support)
- **Google AI Studio API Key** - קבל מ-[Google AI Studio](https://aistudio.google.com)
- **Modern Web Browser** (Chrome, Edge, Firefox)
- **HTTPS Server** - התוסף דורש HTTPS (לא HTTP בלבד)

### 2. קבלת Google AI API Key
1. כנס ל-[Google AI Studio](https://aistudio.google.com)
2. לחץ על "Create new API key"
3. בחר את הפרויקט שלך ו-copy את ה-key
4. הדבק ב-Add-in בחלק "Google AI Studio API Key"

### 3. התזמון והתקנה

#### א. עבור Local Testing:
```bash
# התקן Python HTTP server
python -m http.server 8000

# או עבור Node.js
npx http-server
```

#### ב. עבור Office Add-in עם HTTPS:
1. עלה לשרת ווב (Azure, AWS, Vercel, וכו')
2. Upload את הקבצים:
   - `src/taskpane.html`
   - `src/taskpane.css`
   - `src/taskpane.js`
3. עדכן את ה-URL ב-`manifest.xml`:
```xml
<SourceLocation DefaultValue="https://yourserver.com/src/taskpane.html"/>
```

#### ג. התקן בWord:
1. **Word Online**: File > Options > Trust Center > Trusted Add-in Catalogs > הוסף את manifest.xml URL
2. **Word Desktop**: File > Options > Trust Center > Trusted Add-in Catalogs

---

## 📁 מבנה הפרויקט

```
torani-text-editor/
├── manifest.xml           # Office Add-in Configuration
├── src/
│   ├── taskpane.html      # UI & Components
│   ├── taskpane.css       # Styling
│   └── taskpane.js        # Core Logic & API Integration
├── README.md              # זה הקובץ
└── settings/
    └── .env              # (Optional) API Keys
```

---

## 🎯 כיצד להשתמש

### שלב 1: הגדרות בסיסיות
1. לחץ על "⚙️ הגדרות API"
2. הדבק את ה-API Key שלך
3. בחר את דגם Gemini (Pro או Flash)

### שלב 2: בחר בפעולות
סמן אחד או יותר מ:
- [ ] שנה סגנון ליישיבה
- [ ] פיסוק (Pissuk)
- [ ] עיצוב (Layout)
- [ ] ניקוד (Nikkud)

> **⚠️ חשוב**: אם לא תסמן "ניקוד", ה-AI לא יוסיף ניקוד בשום מקרה

### שלב 3: בקשה חופשית (אופציונלי)
כתוב בדבריך בעברית או באנגלית מה בדיוק אתה רוצה שה-AI יעשה

### שלב 4: בחר בטקסט או אישר עיבוד
- **"עבד טקסט מבחור"** - עבד רק טקסט שבחרת
- **"עבד את כל המסמך"** - עבד את כל המסמך

### שלב 5: תצוגה מקדימה (אופציונלי)
אם בחרת "הצג תצוגה מקדימה":
1. ראה את ההשוואה בטבלה: מקורי חדש
2. לחץ "החל את הכל" או "ביטול"

---

## 🔧 הגדרות מתקדמות

### System Prompt (הוראות בסיס)
- קלק על "System Prompt" כדי להרחיב
- לחץ "ערוך" כדי לשנות את ההוראות
- Default: Expert Torani Editor instructions

### הוראות סגנון יישיבה
הגדר כאן קבלת קונוונציות לסגנון יישיבה:
- חזינן, צ"ע, אתי שפיר
- וכו'

### העדות
בחר אילו שינויים יציגו התראה:
- [ ] הצג התראות בכלל
- [ ] התראה לשינויים בסגנון
- [ ] התראה לשינויים בפיסוק
- [ ] התראה לשינויים בעיצוב
- [ ] התראה לשינויים בניקוד

---

## 🎨 עיצוב ופרמטור

התוסף מיישם אוטומטית:
- **Font**: David (fallback: Frank-Riehl)
- **Size**: 11 pt
- **Line Spacing**: 1.5
- **Alignment**: Justified (ישורים)

---

## 🔐 אבטחה & פרטיות

⚠️ **חשוב**: 
- ה-API Key שלך מעולם לא נשמר בשרת - אך בלבד ברBrowser locale
- כל הטקסט שלך נשלח ישירות ל-Gemini API של Google
- בדוק את [Google Privacy Policy](https://policies.google.com/privacy)

### עבור Production:
1. השתמש ב-Backend Proxy:
   - צור endpoint ב-Backend שלך
   - Backend יקרא את ה-API Key מ-.env
   - Frontend קורא את Backend בלבד

דוגמה:
```javascript
// taskpane.js
const response = await fetch('/api/process-text', {
    method: 'POST',
    body: JSON.stringify({ text, actions })
});
```

---

## 🐛 Troubleshooting

### "Cannot find API Key"
✓ ודא שה-API Key מוכנס בחלק "הגדרות API"
✓ כפי שה-API Key לא מכיל תווים מיוחדים המפריעים

### "Selection is empty"
✓ בחר טקסט ב-Word לפני לחיצה על "עבד טקסט מבחור"

### "API Error: Invalid API"
✓ ודא ש-API Key is correct
✓ בדוק שאתה משתמש בדגם Gemini תקף

### "HTTPS Required"
✓ Word Add-in דורש HTTPS
✓ עצמאדה localhost מתחת HTTPretain test בלבד

### "Nikkud was added when not selected"
✓ זה לא אמור להתרחש
✓ בדוק ש-System Prompt לא כולל הוראה לניקוד

---

## 📊 דוגמה של זרימת עריכה

```
User → Input API Key & Select Actions → Choose Text/Document
  ↓
  → Send to Gemini AI with System Prompt
  ↓
  → Receive Processed Text
  ↓
  → [IF Preview Mode] Show Side-by-Side Comparison
    [IF Direct Mode] Apply directly
  ↓
  → Apply Formatting (Font, Spacing, Justification)
  ↓
  → Show Success Alert & Log Entry
```

---

## 🎓 דוגמה: שימוש בסגנון יישיבה

**Input:**
```
הרמב"ם כתב שהדבר הזה הוא כמו זה,
וזה רע מאוד.
```

**Output (עם סגנון יישיבה):**
```
חזינן מדברי הרמב"ם שהדבר הזה דומה לשיטתו,
ויש בו ריח של קושיא קשה.
```

---

## 📝 API Request Example

```javascript
const payload = {
    contents: [{
        role: "user",
        parts: [{
            text: "You are an expert Torani Editor...\n\nApply these operations:\n- Convert to Yeshivish style\n- Fix punctuation\n\nText: [Hebrew text]"
        }]
    }]
};

fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-pro:generateContent?key=${API_KEY}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
})
```

---

## 🚀 Features - הגיה קרובה

- [ ] Support for more Gemini models
- [ ] Built-in footnote manager with Word integration
- [ ] Batch processing of multiple documents
- [ ] Custom vocabulary dictionary
- [ ] Collaboration features (track changes)
- [ ] Hebrew grammar checking
- [ ] Automatic backup of versions
- [ ] Export to PDF with formatting

---

## 📧 Support & Feedback

אם יש שאלות או בעיות:
1. בדוק את Logs בחלק התחתון של Task Pane
2. ודא שה-API Key תקין
3. בדוק את Network tab ב-DevTools (F12)

---

## 📄 License

הפרויקט זה כפוף ל-MIT License. ראה LICENSE file לפרטים.

---

## 🙏 תודה

- **Google Generative AI** - Gemini models
- **Microsoft Office.js** - Word integration
- **The Yeshiva Community** - Inspiration for Torani editing

---

**Version**: 1.0.0  
**Last Updated**: May 11, 2026  
**Status**: ✓ Production Ready
