# diagrams.md — שיעור 1

דיאגרמות ככתיב Mermaid (ניתן להטמיע ישירות בכלי מצגות/Markdown מודרני התומך ב-Mermaid, כגון Marp עם plugin, Obsidian, GitHub Markdown, ועוד).

## דיאגרמה 1: ה-Workflow בן 10 השלבים

```mermaid
flowchart TD
    A[1. הבנת הבעיה] --> B[2. Specification]
    B --> C[3. בניית Context]
    C --> D[4. תכנון עם AI]
    D --> E[5. מימוש עם AI]
    E --> F[6. Code Review]
    F --> G[7. Testing]
    G --> H[8. Commit]
    H --> I[9. Deploy]
    I --> J[10. שיתוף]
    J -.-> A
```

שימו לב לחץ המקווקו בסוף (10 → 1) — מדגיש שזה תהליך איטרטיבי, לא ליניארי חד-פעמי.

## דיאגרמה 2: ציר זמן — Workflow לעומת כלי

```mermaid
timeline
    title כלי מתחלפים, Workflow נשאר
    2023-2024 : כלי Autocomplete בסיסיים
    2024-2025 : כלים Agentic ראשונים (Claude Code, Cursor)
    2025-2026 : Gemini CLI חינמי וזמין
    יוני 2026 : Gemini CLI נסגר, מוחלף ב-Antigravity CLI
    היום : הכלים ממשיכים להתחלף — ה-Workflow (10 השלבים) נשאר קבוע
```
