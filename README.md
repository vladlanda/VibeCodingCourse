# Vibe Coding — תכנות בסיוע בינה מלאכותית: עקרונות ותהליכי עבודה

קורס בן 13 מפגשים (90 דק' הרצאה + 45 דק' תרגול) לסטודנטים למסלולי תעשייה וניהול / מערכות מידע, ללא רקע קודם בבינה מלאכותית. הקורס מלמד **workflow** לעבודה עם כלי AI לפיתוח תוכנה — לא כלי ספציפי — ומתבסס כולו על כלים חינמיים.

## תוכן הריפו

| קובץ/תיקייה | תיאור |
|---|---|
| `syllabus.md` | סילבוס מפורט: רציונל, מטרות למידה, כלים, מדיניות הערכה, תוכנית 13 המפגשים |
| `project-brief.md` | מפרט מלא לפרויקט המסכם (סוכן AI לבעיה מתעשייה/מערכות מידע) |
| `syllabus_university_format.docx` | סילבוס בפורמט הרשמי של המוסד (להגשה מנהלתית) |
| `materials/` | חומר ההרצאה המלא של כל מפגש (md): הסברים, ניסוח פורמלי, דוגמאות, שאלות בדיקה ודף נוסחאות |
| `sessions/` | מצגות המפגשים: PDF במפגשים 1-4, pptx במפגשים 5-13 |
| `exercises/` | מחברות Colab לתרגול: `HM_NN.ipynb` לסטודנט ו-`HM_NN_SOLUTION.ipynb` פתרון |

## סטטוס

- [x] סילבוס
- [x] מפרט פרויקט מסכם
- [x] סילבוס בפורמט מוסדי
- [x] מפגשים 1-4: חומר הרצאה, מצגת סופית ומחברת תרגול (במפגש 1 אין מחברת)
- [ ] מפגשים 5-13: יש מצגות ומחברות ראשוניות, שעוד לא עברו את התהליך של מפגשים 1-4

## מפגשים 1-4: חומר הרצאה, מצגות ומחברות

| מפגש | נושא | חומר הרצאה | מצגת | מחברת | פתרון |
|---|---|---|---|---|---|
| 1 | מבוא: מהו Vibe Coding? | `materials/session-01-intro-vibe-coding.md` | `sessions/session-01-intro-vibe-coding.pdf` | — | — |
| 2 | איך מודל שפה "חושב" קוד? | `materials/session-02-how-llm-thinks-code.md` | `sessions/session-02-how-llm-thinks-code.pdf` (+ גרסה מקוצרת) | `exercises/HM_02.ipynb` | `exercises/HM_02_SOLUTION.ipynb` |
| 3 | Prompt Engineering לקוד | `materials/session-03-prompt-engineering.md` | `sessions/session-03-prompt-engineering.pdf` (+ גרסה מקוצרת) | `exercises/HM_03.ipynb` | `exercises/HM_03_SOLUTION.ipynb` |
| 4 | Context Engineering: הקשר לאורך פרויקט שלם | `materials/session-04-context-engineering.md` | `sessions/שיעור 4 Context Engineering.pdf` | `exercises/HM_04.ipynb` | `exercises/HM_04_SOLUTION.ipynb` |

המחברות של מפגשים 2-4 בנויות כמעבדה: הסטודנטים מממשים בעצמם, בלי AI, את המודלים מההרצאה. הבדיקות שבמחברות משתמשות במספרים מההרצאה. אחר כך הסטודנטים מודדים בניסוי את מה שבהרצאה הוצג להמחשה. במפגש 4 המצגת מציגה 191 שקופיות, כולל שלבי החשיפה.

## מפגשים 5-13: מצגות ומחברות Colab (תוספת)

לכל מפגש מצגת עם הערות דובר מלאות (`sessions/`), ומחברת Colab אחת לתרגול בסגנון `HM_02`/`HM_03` (`exercises/`): `HM_NN.ipynb` לסטודנט ו־`HM_NN_SOLUTION.ipynb` פתרון. כל מחברת עצמאית (הנתונים מוטמעים בה) ומיועדת לרוץ במלואה ב־Google Colab.

| מפגש | נושא | מצגת | מחברת | פתרון |
|---|---|---|---|---|
| 5 | בחירת הכלי הנכון למשימה (נוף הכלים, MCP) | `sessions/session-05-tool-landscape-selection.pptx` | `exercises/HM_05.ipynb` | `exercises/HM_05_SOLUTION.ipynb` |
| 6 | תכנות אג׳נטי ו־Git | `sessions/session-06-agentic-coding-git.pptx` | `exercises/HM_06.ipynb` | `exercises/HM_06_SOLUTION.ipynb` |
| 7 | APIs, מסדי נתונים ו־Backend | `sessions/session-07-apis-databases-backend.pptx` | `exercises/HM_07.ipynb` | `exercises/HM_07_SOLUTION.ipynb` |
| 8 | בדיקות ואיכות קוד | `sessions/session-08-testing-code-quality.pptx` | `exercises/HM_08.ipynb` | `exercises/HM_08_SOLUTION.ipynb` |
| 9 | דיבוג קוד שנוצר ב־AI | `sessions/session-09-debugging-ai-code.pptx` | `exercises/HM_09.ipynb` | `exercises/HM_09_SOLUTION.ipynb` |
| 10 | אבטחת מידע | `sessions/session-10-security.pptx` | `exercises/HM_10.ipynb` | `exercises/HM_10_SOLUTION.ipynb` |
| 11 | ארכיטקטורה, Refactoring ו־Code Handoff | `sessions/session-11-architecture-refactoring-handoff.pptx` | `exercises/HM_11.ipynb` | `exercises/HM_11_SOLUTION.ipynb` |
| 12 | Deployment, אתיקה ומשפט | `sessions/session-12-deployment-ethics-law.pptx` | `exercises/HM_12.ipynb` | `exercises/HM_12_SOLUTION.ipynb` |
| 13 | סיכום והצגות | `sessions/session-13-summary-showcase.pptx` | `exercises/HM_13.ipynb` | `exercises/HM_13_SOLUTION.ipynb` |
