# references.md — שיעור 2

## מקורות לנתונים (נבדק יולי 2026)

### קנה מידה של GitHub
- [Key GitHub Statistics in 2026 — Kinsta](https://kinsta.com/blog/github-statistics/)
- [GitHub Statistics 2026: Users, Repos & Copilot Data — Skillademia](https://www.skillademia.com/statistics/github-statistics/)
- נתונים מרכזיים: 180M+ מפתחים רשומים (36M+ הצטרפו בשנה אחת בלבד — קצב הצמיחה המהיר ביותר אי-פעם); 630M+ Repositories; 43.2M Pull Requests נמזגו בחודש; כמעט מיליארד Commits ב-2025; 92% מחברות Fortune 100 משתמשות ב-GitHub Enterprise; Copilot: 26M+ משתמשים, 4.7M+ מנויים בתשלום, כ-140,000 ארגונים.

### בנצ'מארקים לכלי Agentic
- [SWE-Bench vs Terminal-Bench: AI Benchmark Guide for 2026 — DigitalApplied](https://www.digitalapplied.com/blog/swe-bench-terminal-bench-benchmark-guide-2026)
- [Best AI Coding Agents (June 2026): Scored Leaderboard — MorphLLM](https://www.morphllm.com/best-ai-coding-agents-2026)
- [12 AI Coding Agents Compared in 2026 — SSOJet](https://ssojet.com/blog/ai-coding-agents-compared)
- נתונים מרכזיים: Claude Code — 80.8% ב-SWE-bench Verified. Terminal-Bench 2.1: Codex CLI מוביל ב-83.4%, Claude Code בגרסאות שונות בין 78.9%-83.1%. **אזהרת מתודולוגיה חשובה:** יש 5 גרסאות שונות של SWE-bench (Original, Verified, Pro, Multilingual, Live) והשוואה ביניהן בעייתית; יש שונות (variance) של 10-20 נקודות אחוז בין הרצות שונות של אותו מודל בדיוק, בהתאם ל-harness. אף פעם אל תציגו מספר בנצ'מארק בודד כ"אמת מוחלטת".
- לוחות תוצאות חיים (לפתיחה בכיתה): [SWE-bench Verified leaderboard](https://www.swebench.com/verified.html) · [Terminal-Bench 2.1 leaderboard](https://www.tbench.ai/leaderboard/terminal-bench/2.1)
- [METR — Task-Completion Time Horizons of Frontier AI Models](https://metr.org/time-horizons/) — בנצ'מארק שני ומשלים: מודד "אופק זמן" (משך משימה, לפי זמן-השלמה אנושי, שבו למודל יש 50% סיכוי להצליח ללא התערבות). אופק הזמן הכפיל את עצמו כל כ-4 חודשים בשנים האחרונות; Claude 3.7 Sonnet (2025) עמד על כ-50 דקות, מודלים מובילים באמצע 2026 הגיעו לכ-12 שעות. **שימו לב:** זה מדד שונה מ-SWE-bench/Terminal-Bench (נכונות פתרון) — כאן נמדדת התמדה עצמאית.

### נתח שימוש בכלים בפועל
- Stack Overflow Developer Survey 2025 (מוזכר במספר מקורות 2026): ChatGPT — 82% (out-of-the-box assistance), GitHub Copilot — 68%; מבין ה-IDE-ים מבוססי AI: Cursor — 18%, Claude Code — 10%, Windsurf — 5%.
- [Top 100 Developer Productivity Statistics with AI Tools 2026 — index.dev](https://www.index.dev/blog/developer-productivity-statistics-with-ai-tools) — בבחירת "כלי ראשי" (primary tool): Claude Code 28%, Cursor 24%.
- **הערה לשימוש בכיתה:** שני הסקרים מודדים דברים שונים (נתח שימוש כללי מול בחירת כלי ראשי אצל מי שכבר בחר כלי Agentic) — לכן המספרים לא סותרים, אבל כדאי להסביר את ההבדל לסטודנטים כדי לא ליצור בלבול.

### DORA — עלות של עבודה בלי Best Practices
- [DORA: ROI of AI-assisted Software Development report](https://dora.dev/ai/)
- אימוץ AI בלי תהליך מסודר מוביל ל-30-41% יותר חוב טכני; Pull Requests שנכתבו ב-AI נושאים פי 1.7 יותר בעיות מ-PR אנושי.

### מקרה בוחן: Amazon (מרץ 2026)
- [AI code wreaked havoc with Amazon outage — Yahoo Tech](https://tech.yahoo.com/articles/ai-code-wreaked-havoc-amazon-160849695.html)
- [AI code wreaked havoc with Amazon outage — Digital Trends](https://www.digitaltrends.com/computing/ai-code-wreaked-havoc-with-amazon-outage-and-now-the-company-is-making-tight-rules/)
- שני תקלות תוך 3 ימים במרץ 2026: 2/3 — כ-6 שעות השבתה, כ-120,000 הזמנות אבודות; 5/3 — ירידה של 99% בהזמנות בצפון אמריקה (כ-6.3 מיליון הזמנות אבודות). לפי Amazon, הגורם המרכזי לתקלה השנייה היה שינוי קוד שעבר לייצור בלי תהליך אישור/תיעוד פורמלי. **ניואנס חשוב לשימוש בכיתה:** Amazon הבהירה שרק תקלה אחת בכלל נגעה לכלי AI, ושם הסיבה הייתה תקלה תהליכית-אנושית, לא הכלי עצמו — אל תציגו את זה כ"AI הרס את Amazon", אלא כ"דילוג על Review הוא הסיכון, גם בענקיות טכנולוגיה". תגובת Amazon: "איפוס בטיחות קוד" בן 90 יום על כ-335 מערכות, אישור שני אנשים + תיעוד פורמלי לכל שינוי, וחתימת מהנדס בכיר ספציפית לשינויי קוד שנוצרו בעזרת AI.

### כלים חינמיים לתרגול (נבדק יולי 2026)
- [Plans for GitHub Copilot — GitHub Docs](https://docs.github.com/en/copilot/get-started/plans) — Copilot Free כולל Agent Mode (לא רק Autocomplete), אך רק עם בחירת מודל אוטומטית (מודלים חלשים יותר כמו Claude Haiku 3.5 / GPT-4o mini), ומכסה מוגבלת (בסדר גודל של 50 בקשות Chat/Agent בחודש) — מספיק ליצירת קובץ בודד כמו בתרגול.
- [Cursor · Pricing (רשמי)](https://cursor.com/pricing) · [ניתוח נוסף — nxcode.io, יולי 2026](https://www.nxcode.io/resources/news/cursor-ai-pricing-plans-guide-2026) — מסלול Hobby (חינמי) כולל "limited Agent requests" בפועל, ללא כרטיס אשראי.
- [GitHub Student Developer Pack (רשמי)](https://education.github.com/pack): הרשמה חדשה להטבת Copilot Student הושהתה באפריל 2026 עקב עומס על תשתית Agentic — סטודנטים שאומתו קודם ממשיכים לקבל את ההטבה, וחברי Pack חדשים מקבלים כרגע Copilot Free הרגיל בלבד. שאר הטבות ה-Pack (JetBrains, Azure for Students, DigitalOcean וכו') פעילות כרגיל.
- [Stack Overflow Developer Survey 2025 (רשמי)](https://survey.stackoverflow.co/2025/)

**הערת עדכניות:** נתוני אימוץ ובנצ'מארק משתנים חודשית בתחום הזה. יש לוודא סטטוס עדכני לפני כל הוראה בפועל.
