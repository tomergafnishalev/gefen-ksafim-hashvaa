# Dev Log

## 2026-06-14 — תיקון: קובץ תכנון ללא תקציב גפן לא קורס
- `tikhnun_processor.py`: כאשר אין שורת ctrl של גפן ("כללי"), הקוד מחפש את הבאדם הראשון עם H > 0 כ-fallback במקום לזרוק ValueError — מאפשר קבצים עם תקציב תקומה בלבד וכד'

## 2026-06-14 — דיאלוג בחירת תקציב לפני טאב יוזמות וצרכים
- `tikhnun_processor.py`: דילוג שורת "סה''כ תקציבים במערכת גפן" בלולאת זיהוי התקציבים (שורת סיכום לא נספרת כתקציב)
- `ResultsView.jsx`: קומפוננט `BudgetSelectionDialog` חדש — מציג שאלה + כפתור לכל תקציב שזוהה
- `ResultsView.jsx`: כניסה לטאב "יוזמות וצרכים" עם מספר תקציבים — מציג תחילה את דיאלוג בחירת התקציב, אחר כך דיאלוג מודל התמרוץ
- `ResultsView.jsx`: `YozmaTab` מקבל `selectedBudgetName` — מציג נתוני yozma מהתקציב שנבחר (H, מקסימום, בתכנון, פירוט ספקים)
- טאב "ניהול ותפעול" ללא שינוי

## 2026-06-14 — תיקון חישוב betikhnun_by_code — זיהוי קוד ברמת מרכיב
- `tikhnun_processor.py`: `yozma_by_code` (גלובלי) עבר מ-iteration על `unique_plans` (dedup 10 עמודות) ל-iteration component-aware על `perut_rows`
- `tikhnun_processor.py`: `yb_code` (per-budget) — אותו תיקון, iteration על `budget_perut` עם dedup חדש
- מפתח dedup חדש כולל col R (קוד דיווח) ו-col N (עלות למרכיב); כאשר col N קיים הסכום נלקח ממנו, אחרת fallback ל-col P
- כך יוזמה עם שני קודים שונים (למשל 107+106) מתחלקת נכון בין הקודים — לא מוצנחת לקוד אחד בלבד

## 2026-06-14 — הסרת כפתורי תקציב (pills) מכל הטאבים
- `ResultsView.jsx`: הסרת קומפוננט `BudgetPills` וכל לוגיקת בחירת-תקציב מ-`HashvaTab`, `YozmaTab`, `NihulTab`
- `HashvaTab`: תמיד משתמש ב-`rows_finance_not_gefen` / `rows_gefen_not_finance` הגלובליים (לא per_combo_results)
- `YozmaTab`: תמיד משתמש ב-`tikhnun.yozma_03/04` ו-`tikhnun.yozma_breakdown` הגלובליים
- `NihulTab`: תמיד משתמש ב-`tikhnun.nihul_breakdown` הגלובלי
- הסרת state של `activeBudgetName`, `activeBudgetNameTikkon`, `activeBudgetNameBeinayim` מ-`ResultsView`
- עדכון `getTabIssues` להסתמך על `rows_finance_not_gefen` / `rows_gefen_not_finance` הגלובליים בלבד

## 2026-06-11 — תיקוני multi-budget (סבב 2): max יוזמות, pills ניהול, הודעות תקציב ריק
- `tikhnun_processor.py`: הרשומה של תקציב גפן ב-`_budgets_data` מאולצת להשתמש ב-H/L/S/T מה-ctrl row ("כללי") — מונע מצב שה-first encounter row מכיל H כולל שכולל תקציבים אחרים
- `tikhnun_processor.py`: `nihul_planned` flag לכל תקציב — מסמן האם יש תכנון ניהול ותפעול (לפי rcode_to_budget), גם אם טרם הוגשו אסמכתאות
- `tikhnun_processor.py`: `build_tikhnun_result` חושף `nihul_planned` לכל תקציב
- `ResultsView.jsx` `NihulTab`: pills מוצגים לכל התקציבים (לא רק לאלה עם נתונים); תקציב ללא תכנון ניהול → "בתקציב זה אין תכנון של ניהול ותפעול"; תכנון קיים אך ללא אסמכתאות → "טרם דווחו אסמכתאות עבור ניהול ותפעול בתקציב זה"

## 2026-06-11 — תיקוני multi-budget: pills כפולים, פירוט ספקים, gamish_notar
- `tikhnun_processor.py`: ניקוי כפילויות ב-`_budgets_data` לפי norm_name (raw names שונים שמנרמלים ל-"גפן" — כמו "גפ"ן הכללי" ו-"גפ"ן מחוזי" — מתמזגים לרשומה אחת)
- `tikhnun_processor.py`: gamish_notar לתקציב גפן הראשי משתמש בערך הגלובלי (כולל ניכוי kvua) — תקציבים אחרים משתמשים ב-H - L שלהם
- `ResultsView.jsx`: הוספת `NihulSupplierRow` — קומפוננט שורת ספק עם אפשרות לפתיחת פירוט עסקאות (אסמכתאות) בטאב ניהול ותפעול

## 2026-06-11 — תמיכה בתקציבים מרובים (multi-budget) בשלושת הטאבים
- `tikhnun_processor.py`: שכתוב מלא — `normalize_budget_name`, זיהוי תקציבים מגיליון "הכל", בניית `rcode_to_budget`, חישוב yozma ו-nihul_breakdown לכל תקציב בנפרד, `_budgets_data` ברמה הפנימית, `budgets[]` ב-`build_tikhnun_result`
- `analyze_router.py`: קובץ תכנון הפך לחובה לכל בדיקה; תיוג שורות ב-`_budget` לפי `rcode_to_budget`; בניית `per_combo_results` לפיצול לפי תקציב; הוספת `_budget` ל-`_STRIP_COLS`
- `ResultsView.jsx`: קומפוננט `BudgetPills` חדש; עדכון `HashvaTab`, `YozmaTab`, `NihulTab` לקבל `activeBudgetName`/`onBudgetChange`; state מנוהל משותף (`activeBudgetName`, `Tikkon`, `Beinayim`); `getTabIssues` עודכן לבדוק כל תקציב; כל call sites עודכנו להעביר props
- כשיש תקציב אחד — אין שינוי בהתנהגות (ללא pills)

## 2026-06-11 — תמיכה בשני קבצי כספים מאותו סוג
- `analyze_router.py` → `_classify_files`: מאפשר עד 2 קבצי כספים מאותו סוג (למשל שני קבצי סקולקאש); שגיאה אם מדובר בסוגים שונים
- `analyze_router.py` → `_load_finance_raw`: מקבל רשימת נתיבים, ממזג שני קבצים עם ניקוי כפילויות לפי מפתח ichud (אותו מנגנון כמו קבצי גפן כפולים)

## 2026-06-11 — הוספת טאב "ניהול ותפעול"
- `tikhnun_processor.py`: הוספת חישוב `nihul_breakdown` ב-`cross_reference_doch` — מסכם ספקים לפי קוד 104 (תיכון) / 67 (ביניים/יסודי); חשיפה דרך `build_tikhnun_result`
- `tikhnun_exporter.py`: הוספת `_write_nihul` (Excel), branch "nihul" ב-`build_tikhnun_section_story` (PDF), ו-"nihul" ב-PageBreak של `export_tikhnun_pdf`
- `combined_exporter.py`: ייבוא `_write_nihul`, הוספת "nihul" ל-`TIKHNUN_TABS`, `_SHEET_NAMES`, ולולאת ה-Excel
- `ResultsView.jsx`: הוספת טאב "ניהול ותפעול" (`TAB_IDS`, `TAB_LABELS_MAP`, `TIKHNUN_ONLY_TABS`), קומפוננט `NihulTab`, `PILL_STYLE`, ורינדור לסינגל ודואל טיכנון
