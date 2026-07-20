const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

// ---------- palette ----------
const DARK = "1B1F3B";      // deep indigo
const DARK2 = "2A2F55";     // lighter indigo for cards on dark bg
const TEAL = "00C2A8";      // accent
const TEAL_TINT = "E6FBF7"; // light teal fill
const ICE = "CADCFC";
const TEXT_DARK = "1B1F3B";
const MUTED = "6B7280";
const WARN = "FFB000";
const WARN_TINT = "FFF6E6";
const PH_BG = "EDEDF2";
const PH_BG_DARK = "2A2F55";
const PH_BORDER = "B8B8C8";
const WHITE = "FFFFFF";

const TITLE_FONT = "Arial";
const BODY_FONT = "Calibri";

// ---------- helpers ----------
function addImagePlaceholder(slide, x, y, w, h, label, dark) {
  slide.addShape(pres.ShapeType.rect, {
    x, y, w, h,
    fill: { color: dark ? PH_BG_DARK : PH_BG },
    line: { color: dark ? TEAL : PH_BORDER, width: 1.25, dashType: "dash" },
  });
  slide.addText(label, {
    x, y, w, h, margin: 8,
    align: "center", valign: "middle",
    fontFace: BODY_FONT, fontSize: 12, italic: true,
    color: dark ? TEAL : MUTED, rtlMode: true,
  });
}

function addSlideTitle(slide, title, opts = {}) {
  slide.addText(title, {
    x: 0.6, y: 0.4, w: 12.13, h: 0.9,
    align: "right", fontFace: TITLE_FONT, bold: true,
    fontSize: opts.fontSize || 30, color: opts.color || TEXT_DARK,
    rtlMode: true, margin: 0,
  });
}

function addFooter(slide, n, dark) {
  slide.addText(`Vibe Coding · שיעור 1 · שקף ${n}`, {
    x: 0.6, y: 7.1, w: 12.13, h: 0.3,
    align: "right", fontFace: BODY_FONT, fontSize: 9,
    color: dark ? "8891C2" : "A6A6B0", rtlMode: true, margin: 0,
  });
}

function iconCircle(slide, x, y, d, label, fillColor, textColor) {
  slide.addShape(pres.ShapeType.ellipse, { x, y, w: d, h: d, fill: { color: fillColor }, line: { type: "none" } });
  slide.addText(String(label), {
    x, y, w: d, h: d, align: "center", valign: "middle",
    fontFace: TITLE_FONT, bold: true, fontSize: d > 0.6 ? 20 : 14,
    color: textColor, margin: 0,
  });
}

// =====================================================================
// SLIDE 1 — Title
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };

  s.addText("Vibe Coding – AI-Native Software Development", {
    x: 0.6, y: 0.55, w: 6.2, h: 0.4, align: "right",
    fontFace: BODY_FONT, fontSize: 13, color: TEAL, italic: true, rtlMode: true, margin: 0,
  });

  s.addText("שיעור 1: מבוא ל-Vibe Coding", {
    x: 0.6, y: 1.05, w: 6.2, h: 1.6, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 38, color: WHITE, rtlMode: true, margin: 0,
  });

  s.addText("Introduction to Vibe Coding", {
    x: 0.6, y: 2.55, w: 6.2, h: 0.5, align: "right",
    fontFace: BODY_FONT, italic: true, fontSize: 16, color: ICE, rtlMode: true, margin: 0,
  });

  s.addText("בסוף השיעור תדעו לענות על:", {
    x: 0.6, y: 3.35, w: 6.2, h: 0.4, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 16, color: WHITE, rtlMode: true, margin: 0,
  });

  s.addText(
    [
      { text: "• מהו Vibe Coding, ובמה הוא שונה מ״להקליד פרומפט״?", options: { breakLine: true } },
      { text: "• מה היתרונות, החסרונות והמגבלות של פיתוח מבוסס AI?", options: { breakLine: true } },
      { text: "• מהו ה-Workflow בן 10 השלבים שנלמד לאורך כל הקורס?", options: { breakLine: true } },
      { text: "• למה הקורס מלמד ״תהליך״, ולא ״כלי״?", options: { breakLine: false } },
    ],
    {
      x: 0.6, y: 3.85, w: 6.2, h: 2.6, align: "right",
      fontFace: BODY_FONT, fontSize: 14, color: ICE, rtlMode: true,
      paraSpaceAfter: 10, margin: 0,
    }
  );

  addImagePlaceholder(
    s, 7.2, 0.7, 5.53, 6.1,
    "[תמונה: מסך טרמינל / IDE עם סוכן AI כותב קוד בזמן אמת]",
    true
  );

  s.addText("Vibe Coding Course · Lesson 1 of 13", {
    x: 0.6, y: 6.9, w: 6.2, h: 0.3, align: "right",
    fontFace: BODY_FONT, fontSize: 10, color: "8891C2", rtlMode: true, margin: 0,
  });
}

// =====================================================================
// SLIDE 2 — מהו Vibe Coding
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "מהו Vibe Coding?");

  const rows = [
    { h: "AI ככותב קוד, אתם כמנווטים", d: "מנחים AI לכתוב, לבדוק ולתקן קוד בפועל — לא כותבים כל שורה לבד" },
    { h: "שליטה מלאה נשארת אצלכם", d: "כיוון, איכות והחלטות ארכיטקטוניות — תמיד בידיים שלכם" },
    { h: "זה לא ניחוש", d: "תכנון, בדיקה, הבנה ואחריות מלאה על כל שורת קוד שמגיעה לייצור" },
  ];
  rows.forEach((r, i) => {
    const y = 1.7 + i * 1.5;
    iconCircle(s, 6.1, y, 0.6, i + 1, TEAL, WHITE);
    s.addText(r.h, {
      x: 0.6, y, w: 5.35, h: 0.4, align: "right",
      fontFace: TITLE_FONT, bold: true, fontSize: 18, color: TEXT_DARK, rtlMode: true, margin: 0,
    });
    s.addText(r.d, {
      x: 0.6, y: y + 0.42, w: 5.35, h: 0.9, align: "right",
      fontFace: BODY_FONT, fontSize: 14, color: MUTED, rtlMode: true, margin: 0,
    });
  });

  addImagePlaceholder(
    s, 7.1, 1.7, 5.63, 4.6,
    "[תמונה: ממשק Claude Code בפעולה — מפתח נותן הנחיה, הסוכן כותב קוד]"
  );
  addFooter(s, 2);
}

// =====================================================================
// SLIDE 3 — AI-Native Software Engineering
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "AI-Native Software Engineering");

  s.addShape(pres.ShapeType.roundRect, {
    x: 6.93, y: 1.9, w: 5.6, h: 3.6, rectRadius: 0.12,
    fill: { color: "F2F2F2" }, line: { type: "none" },
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: 1.9, w: 5.6, h: 3.6, rectRadius: 0.12,
    fill: { color: TEAL_TINT }, line: { type: "none" },
  });

  s.addText("AI-Native (הגישה של הקורס)", {
    x: 0.9, y: 2.15, w: 5.0, h: 0.5, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 18, color: TEAL, rtlMode: true, margin: 0,
  });
  s.addText(
    [
      { text: "• AI כשותף עבודה מרכזי", options: { breakLine: true } },
      { text: "• מעורב מהבנת הדרישה ועד פריסה בענן", options: { breakLine: true } },
      { text: "• בכל שלב בתהליך הפיתוח, לא רק בכתיבת קוד", options: { breakLine: false } },
    ],
    { x: 0.9, y: 2.75, w: 5.0, h: 2.5, align: "right", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );

  s.addText("הגישה הישנה", {
    x: 7.23, y: 2.15, w: 5.0, h: 0.5, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 18, color: MUTED, rtlMode: true, margin: 0,
  });
  s.addText(
    [
      { text: "• AI ככלי עזר צדדי", options: { breakLine: true } },
      { text: "• השלמת קוד בסיסית (Autocomplete)", options: { breakLine: true } },
      { text: "• תוסף לתהליך פיתוח קיים — לא משנה אותו", options: { breakLine: false } },
    ],
    { x: 7.23, y: 2.75, w: 5.0, h: 2.5, align: "right", fontFace: BODY_FONT, fontSize: 14, color: MUTED, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );

  s.addText("זה לא תוספת לתהליך הפיתוח הקיים — זהו תהליך פיתוח חדש.", {
    x: 0.6, y: 5.85, w: 11.93, h: 0.7, align: "right",
    fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 17, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addFooter(s, 3);
}

// =====================================================================
// SLIDE 4 — Evolution of Software Development
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "Evolution of Software Development");

  const stages = [
    { label: "קוד מכונה", fill: "F2F2F2", color: TEXT_DARK },
    { label: "שפות עיליות", fill: "F2F2F2", color: TEXT_DARK },
    { label: "IDE +\nAutocomplete", fill: "E9EEF9", color: TEXT_DARK },
    { label: "Git", fill: "DCEFFF", color: TEXT_DARK },
    { label: "AI Agentic", fill: TEAL, color: WHITE },
  ];
  // RTL flow: earliest on the right, newest (AI Agentic) on the left
  const positions = [10.73, 8.18, 5.63, 3.08, 0.53];
  stages.forEach((st, i) => {
    const x = positions[i];
    s.addShape(pres.ShapeType.roundRect, {
      x, y: 2.3, w: 2.2, h: 1.5, rectRadius: 0.1,
      fill: { color: st.fill }, line: { type: "none" },
    });
    s.addText(st.label, {
      x, y: 2.3, w: 2.2, h: 1.5, align: "center", valign: "middle",
      fontFace: TITLE_FONT, bold: true, fontSize: 14, color: st.color, margin: 4,
    });
    if (i < stages.length - 1) {
      s.addText("←", {
        x: x - 0.33, y: 2.3, w: 0.3, h: 1.5, align: "center", valign: "middle",
        fontFace: TITLE_FONT, bold: true, fontSize: 20, color: MUTED, margin: 0,
      });
    }
  });

  s.addText("כל קפיצה שינתה מה נחשב ״עבודה של מפתח״", {
    x: 0.6, y: 4.1, w: 12.13, h: 0.5, align: "right",
    fontFace: BODY_FONT, italic: true, fontSize: 15, color: MUTED, rtlMode: true, margin: 0,
  });

  addImagePlaceholder(
    s, 0.6, 4.8, 12.13, 1.9,
    "[תמונה: רצף ויזואלי לכל עידן — כרטיסי ניקוב, IDE ראשון, ממשק Git, סוכן AI]"
  );
  addFooter(s, 4);
}

// =====================================================================
// SLIDE 5 — יתרונות
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "יתרונות");

  const rows = [
    { h: "מהירות", d: "אפליקציה שלקחה ימים — מתחילים לבנות אותה תוך דקות" },
    { h: "נגישות", d: "אנשים בלי רקע עמוק בתכנות יכולים לבנות דברים אמיתיים" },
    { h: "שחרור מבוילרפלייט", d: "קוד חוזר וסטנדרטי (מסלולים, טפסים, בדיקות בסיסיות) נכתב מהר בהרבה" },
  ];
  rows.forEach((r, i) => {
    const y = 1.7 + i * 1.5;
    iconCircle(s, 6.1, y, 0.6, i + 1, TEAL, WHITE);
    s.addText(r.h, {
      x: 0.6, y, w: 5.35, h: 0.4, align: "right",
      fontFace: TITLE_FONT, bold: true, fontSize: 18, color: TEXT_DARK, rtlMode: true, margin: 0,
    });
    s.addText(r.d, {
      x: 0.6, y: y + 0.42, w: 5.35, h: 0.9, align: "right",
      fontFace: BODY_FONT, fontSize: 14, color: MUTED, rtlMode: true, margin: 0,
    });
  });

  addImagePlaceholder(
    s, 7.1, 1.7, 5.63, 4.6,
    "[תמונה: השוואת ״לפני / אחרי״ — זמן פיתוח שהצטמצם מימים לדקות]"
  );
  addFooter(s, 5);
}

// =====================================================================
// SLIDE 6 — חסרונות ומגבלות
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "חסרונות ומגבלות");

  const cards = [
    { h: "Hallucination", d: "המודל ״ממציא״ פונקציות, ספריות או API-ים שלא קיימים — בביטחון מלא" },
    { h: "חוסר הבנת הקשר רחב", d: "בלי Context מנוהל נכון, המודל ״שוכח״ החלטות קודמות בפרויקט" },
    { h: "קוד שרץ ≠ קוד נכון", d: "קוד יכול להיראות תקין ולרוץ בלי שגיאות — ועדיין להכיל טעות לוגית" },
    { h: "תלות בכלי משתנה", d: "כלי ה-AI של היום עשוי להיעלם או להשתנות מהותית בעוד שנה" },
  ];
  const pos = [
    { x: 6.85, y: 1.7 }, { x: 0.6, y: 1.7 },
    { x: 6.85, y: 3.85 }, { x: 0.6, y: 3.85 },
  ];
  cards.forEach((c, i) => {
    const { x, y } = pos[i];
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w: 5.9, h: 2.0, rectRadius: 0.1,
      fill: { color: WARN_TINT }, line: { type: "none" },
    });
    iconCircle(s, x + 5.1, y + 0.25, 0.5, "!", WARN, WHITE);
    s.addText(c.h, {
      x: x + 0.3, y: y + 0.22, w: 4.6, h: 0.5, align: "right",
      fontFace: TITLE_FONT, bold: true, fontSize: 16, color: TEXT_DARK, rtlMode: true, margin: 0,
    });
    s.addText(c.d, {
      x: x + 0.3, y: y + 0.8, w: 5.3, h: 1.0, align: "right",
      fontFace: BODY_FONT, fontSize: 13, color: MUTED, rtlMode: true, margin: 0,
    });
  });

  s.addText(
    "AI מאיץ עבודה, אבל לא מחליף שיפוט הנדסי. ככל שהפרויקט חשוב יותר — כך הביקורת האנושית קריטית יותר.",
    {
      x: 0.6, y: 6.05, w: 12.13, h: 0.6, align: "right",
      fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 14, color: TEXT_DARK, rtlMode: true, margin: 0,
    }
  );
  addFooter(s, 6);
}

// =====================================================================
// SLIDE 7 — Workflow של מפתח מודרני
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "Workflow של מפתח מודרני");

  const steps = [
    "הבנת\nהבעיה", "Specification", "בניית\nContext", "תכנון\nעם AI", "מימוש\nעם AI",
    "Code\nReview", "Testing", "Commit", "Deploy", "שיתוף",
  ];
  const xs = [0.6, 2.93, 5.26, 7.59, 9.92];
  steps.forEach((label, i) => {
    const row = i < 5 ? 0 : 1;
    const col = i % 5;
    const x = xs[col];
    const y = row === 0 ? 1.7 : 3.55;
    s.addShape(pres.ShapeType.roundRect, {
      x, y, w: 2.15, h: 1.5, rectRadius: 0.1,
      fill: { color: TEAL_TINT }, line: { type: "none" },
    });
    iconCircle(s, x + 0.78, y + 0.15, 0.55, i + 1, TEAL, WHITE);
    s.addText(label, {
      x, y: y + 0.78, w: 2.15, h: 0.65, align: "center", valign: "top",
      fontFace: TITLE_FONT, bold: true, fontSize: 12, color: TEXT_DARK, margin: 2,
    });
  });

  s.addText("↻ תהליך איטרטיבי — לא ליניארי חד-פעמי", {
    x: 0.6, y: 5.25, w: 12.13, h: 0.4, align: "right",
    fontFace: BODY_FONT, italic: true, fontSize: 13, color: MUTED, rtlMode: true, margin: 0,
  });
  s.addText("זהו העיקרון-על של כל הקורס", {
    x: 0.6, y: 5.75, w: 12.13, h: 0.5, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 18, color: TEAL, rtlMode: true, margin: 0,
  });
  addFooter(s, 7);
}

// =====================================================================
// SLIDE 8 — מקרה בוחן: Gemini CLI → Antigravity
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "מקרה בוחן: Gemini CLI → Antigravity");

  const points = [
    { x: 11.2, label: "2023-2024", desc: "כלי Autocomplete\nבסיסיים", up: true },
    { x: 8.35, label: "2024-2025", desc: "כלים Agentic\nראשונים", up: false },
    { x: 5.5, label: "2025-2026", desc: "Gemini CLI\nחינמי וזמין", up: true },
    { x: 2.0, label: "יוני 2026", desc: "Gemini CLI נסגר —\nהוחלף ב-Antigravity", up: false, warn: true },
  ];
  s.addShape(pres.ShapeType.line, {
    x: 1.3, y: 2.5, w: 10.6, h: 0, line: { color: PH_BORDER, width: 2 },
  });
  points.forEach((p) => {
    s.addShape(pres.ShapeType.ellipse, {
      x: p.x, y: 2.35, w: 0.3, h: 0.3,
      fill: { color: p.warn ? WARN : TEAL }, line: { color: WHITE, width: 2 },
    });
    s.addText(p.label, {
      x: p.x - 0.85, y: p.up ? 1.55 : 2.8, w: 2.0, h: 0.35, align: "center",
      fontFace: TITLE_FONT, bold: true, fontSize: 13, color: TEXT_DARK, margin: 0,
    });
    s.addText(p.desc, {
      x: p.x - 0.85, y: p.up ? 1.95 : 3.2, w: 2.0, h: 0.6, align: "center",
      fontFace: BODY_FONT, fontSize: 11, color: MUTED, rtlMode: true, margin: 0,
    });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: 4.05, w: 12.13, h: 1.05, rectRadius: 0.1,
    fill: { color: TEAL_TINT }, line: { type: "none" },
  });
  s.addText(
    "מי שלמד רק ״את הכלי״ — מתחיל כמעט מאפס. מי שלמד את ה-Workflow — עובר לכלי הבא כמעט בלי לאבד זמן.",
    {
      x: 0.9, y: 4.15, w: 11.53, h: 0.85, align: "right", valign: "middle",
      fontFace: TITLE_FONT, bold: true, fontSize: 16, color: TEXT_DARK, rtlMode: true, margin: 0,
    }
  );

  addImagePlaceholder(
    s, 0.6, 5.3, 12.13, 1.55,
    "[תמונה/לוגו: Gemini CLI ו-Google Antigravity CLI, זה לצד זה]"
  );
  addFooter(s, 8);
}

// =====================================================================
// SLIDE 9 — עוד דוגמה: זה לא רק Google
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "עוד דוגמה: זה לא רק Google");

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: 1.8, w: 12.13, h: 2.5, rectRadius: 0.1,
    fill: { color: WARN_TINT }, line: { type: "none" },
  });
  iconCircle(s, 11.3, 2.0, 0.5, "!", WARN, WHITE);
  s.addText(
    [
      { text: "• מרץ 2026: GitHub שינה את מסלול הסטודנטים ל-Copilot Student", options: { breakLine: true } },
      { text: "• הפעלות חדשות מוקפאות זמנית — בלי לוח זמנים לחידוש", options: { breakLine: true } },
      { text: "• גם ״תנאי מסלול חינם״ משתנים בלי אזהרה, לא רק כלים שלמים", options: { breakLine: false } },
    ],
    { x: 0.9, y: 2.1, w: 10.2, h: 2.0, align: "right", fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: 4.6, w: 12.13, h: 1.5, rectRadius: 0.1,
    fill: { color: TEAL_TINT }, line: { type: "none" },
  });
  s.addText("פעילות זוגות", {
    x: 0.9, y: 4.78, w: 11.53, h: 0.4, align: "right",
    fontFace: BODY_FONT, italic: true, fontSize: 13, color: TEAL, rtlMode: true, margin: 0,
  });
  s.addText("תחשבו על כלי טכנולוגי שהשתנה דרמטית בחייכם תוך שנה-שנתיים — שתפו את בן/בת הזוג", {
    x: 0.9, y: 5.1, w: 11.53, h: 0.9, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 16, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addFooter(s, 9);
}

// =====================================================================
// SLIDE 10 — הדגמה חיה
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };

  s.addText("הדגמה חיה", {
    x: 0.6, y: 0.5, w: 12.13, h: 0.6, align: "center",
    fontFace: BODY_FONT, italic: true, fontSize: 16, color: TEAL, margin: 0,
  });
  s.addText("20 דקות. בונים יחד, בזמן אמת.", {
    x: 0.6, y: 1.2, w: 12.13, h: 1.2, align: "center",
    fontFace: TITLE_FONT, bold: true, fontSize: 38, color: WHITE, rtlMode: true, margin: 0,
  });
  s.addText("ראו demo/prompt.md להוראות המלאות", {
    x: 0.6, y: 2.35, w: 12.13, h: 0.5, align: "center",
    fontFace: BODY_FONT, italic: true, fontSize: 15, color: ICE, rtlMode: true, margin: 0,
  });

  addImagePlaceholder(
    s, 3.5, 3.15, 6.33, 3.6,
    "[תמונה: מסך שיתוף / הקלטת מסך של סוכן AI בונה קוד בזמן אמת]",
    true
  );
  addFooter(s, 10, true);
}

// =====================================================================
// SLIDE 10 — עוברים לתרגול
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: TEAL };

  s.addText("עוברים לתרגול", {
    x: 0.6, y: 0.6, w: 12.13, h: 0.6, align: "center",
    fontFace: BODY_FONT, italic: true, fontSize: 16, color: DARK, margin: 0,
  });
  s.addText("45 דקות: בניית Todo App באמצעות AI", {
    x: 0.6, y: 1.3, w: 12.13, h: 1.1, align: "center",
    fontFace: TITLE_FONT, bold: true, fontSize: 32, color: WHITE, rtlMode: true, margin: 0,
  });

  s.addText(
    [
      { text: "• הוספת משימה חדשה", options: { breakLine: true } },
      { text: "• סימון משימה כהושלמה", options: { breakLine: true } },
      { text: "• מחיקת משימה", options: { breakLine: true } },
      { text: "• שמירה עם localStorage", options: { breakLine: false } },
    ],
    {
      x: 0.9, y: 2.9, w: 5.6, h: 2.2, align: "right",
      fontFace: BODY_FONT, fontSize: 16, color: WHITE, rtlMode: true, paraSpaceAfter: 10, margin: 0,
    }
  );

  addImagePlaceholder(
    s, 6.9, 2.9, 5.53, 3.2,
    "[אייקון/תמונה: רשימת Todo מסומנת ✓]",
    true
  );
  s.addText("ראו exercises.md ו-solutions.md לפירוט מלא", {
    x: 0.6, y: 6.6, w: 12.13, h: 0.4, align: "center",
    fontFace: BODY_FONT, italic: true, fontSize: 13, color: DARK, rtlMode: true, margin: 0,
  });
}

// =====================================================================
// SLIDE 11 — סיכום
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };

  s.addText("סיכום", {
    x: 0.6, y: 0.5, w: 12.13, h: 0.8, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 34, color: WHITE, rtlMode: true, margin: 0,
  });

  const takeaways = [
    "Vibe Coding = תהליך מקצועי, לא ניחוש",
    "AI מאיץ אך לא מחליף שיפוט הנדסי",
    "לומדים Workflow — כלים יוחלפו, התהליך נשאר",
  ];
  takeaways.forEach((t, i) => {
    const y = 1.7 + i * 1.05;
    iconCircle(s, 11.3, y, 0.55, i + 1, TEAL, WHITE);
    s.addText(t, {
      x: 0.6, y, w: 10.4, h: 0.7, align: "right", valign: "middle",
      fontFace: TITLE_FONT, bold: true, fontSize: 18, color: WHITE, rtlMode: true, margin: 0,
    });
  });

  s.addShape(pres.ShapeType.roundRect, {
    x: 0.6, y: 5.6, w: 12.13, h: 1.3, rectRadius: 0.1,
    fill: { color: TEAL }, line: { type: "none" },
  });
  s.addText("שבוע הבא", {
    x: 0.9, y: 5.78, w: 11.53, h: 0.4, align: "right",
    fontFace: BODY_FONT, italic: true, fontSize: 13, color: DARK, rtlMode: true, margin: 0,
  });
  s.addText("AI Development Environment — מקימים סביבת עבודה אמיתית", {
    x: 0.9, y: 6.15, w: 11.53, h: 0.6, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 18, color: WHITE, rtlMode: true, margin: 0,
  });
}

// =====================================================================
pres.writeFile({ fileName: "lesson_01_slides.pptx" }).then(() => {
  console.log("done");
});
