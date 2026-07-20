const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

// ---------- palette ----------
const DARK = "1B1F3B";
const DARK2 = "2A2F55";
const TEAL = "00C2A8";
const TEAL_TINT = "E6FBF7";
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

function addSourceCaption(slide, text) {
  slide.addText(text, {
    x: 0.6, y: 6.75, w: 12.13, h: 0.3,
    align: "right", fontFace: BODY_FONT, italic: true, fontSize: 10,
    color: MUTED, rtlMode: true, margin: 0,
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

function statCallout(slide, x, y, w, h, big, small, opts = {}) {
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.1,
    fill: { color: opts.fill || TEAL_TINT }, line: { type: "none" },
  });
  slide.addText(big, {
    x, y: y + 0.15, w, h: h - 0.75, align: "center", valign: "bottom",
    fontFace: TITLE_FONT, bold: true, fontSize: opts.bigSize || 40,
    color: opts.bigColor || TEAL, margin: 0,
  });
  slide.addText(small, {
    x: x + 0.15, y: y + h - 0.65, w: w - 0.3, h: 0.6, align: "center", valign: "top",
    fontFace: BODY_FONT, fontSize: 12.5, color: opts.smallColor || TEXT_DARK, rtlMode: true, margin: 0,
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
      { text: "• מה היתרונות, החסרונות והמגבלות — עם נתונים אמיתיים?", options: { breakLine: true } },
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
// SLIDE 2 — לפני שמתחילים: מה זה בכלל LLM?
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "לפני שמתחילים: מה זה בכלל LLM?");

  s.addText(
    [
      { text: "• LLM (Large Language Model) = תוכנה שאומנה על כמויות עצומות של טקסט וקוד", options: { breakLine: true } },
      { text: "• לא \"מבין\" כמו בן אדם — מנבא Token אחר Token מה סביר לבוא בהמשך", options: { breakLine: true } },
      { text: "• ChatGPT ו-Claude = מודלים כאלה, מוכרים לכם כצ'אטבוטים", options: { breakLine: true } },
      { text: "• Vibe Coding = אותה טכנולוגיה בדיוק, מחוברת ישירות לקבצי הקוד שלכם", options: { breakLine: false, bold: true, color: TEAL } },
    ],
    {
      x: 0.6, y: 1.7, w: 12.13, h: 2.6, align: "right",
      fontFace: BODY_FONT, fontSize: 16, color: TEXT_DARK, rtlMode: true,
      paraSpaceAfter: 14, margin: 0,
    }
  );

  addImagePlaceholder(
    s, 0.6, 4.5, 12.13, 2.2,
    "[תמונה: תרשים פשוט — קלט טקסט → מודל שפה (קופסה) → פלט טקסט/קוד, עם חץ לדוגמה של Token אחר Token]"
  );
  addFooter(s, 2);
}

// =====================================================================
// SLIDE 3 — מהו Vibe Coding
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
  addFooter(s, 3);
}

// =====================================================================
// SLIDE 4 — AI-Native Software Engineering
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

  s.addText("AI-Native (הגישה של הקורס)", { x: 0.9, y: 2.15, w: 5.0, h: 0.5, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 18, color: TEAL, rtlMode: true, margin: 0 });
  s.addText(
    [
      { text: "• AI כשותף עבודה מרכזי", options: { breakLine: true } },
      { text: "• מעורב מהבנת הדרישה ועד פריסה בענן", options: { breakLine: true } },
      { text: "• בכל שלב בתהליך הפיתוח, לא רק בכתיבת קוד", options: { breakLine: false } },
    ],
    { x: 0.9, y: 2.75, w: 5.0, h: 2.5, align: "right", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );

  s.addText("הגישה הישנה", { x: 7.23, y: 2.15, w: 5.0, h: 0.5, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 18, color: MUTED, rtlMode: true, margin: 0 });
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
  addFooter(s, 4);
}

// =====================================================================
// SLIDE 5 — המספרים מאחורי המהפכה
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "המספרים מאחורי המהפכה");

  const stats = [
    { big: "92.6%", small: "מהמפתחים משתמשים בכלי AI לפחות פעם בחודש" },
    { big: "26.9%", small: "מקוד הפרודקשן בתעשייה נכתב היום על ידי AI" },
    { big: "1,000+", small: "Pull Requests בשבוע דרך \"Minions\" — המערכת הפנימית של Stripe" },
    { big: "40%", small: "מהאפליקציות הארגוניות ישלבו סוכני AI אוטונומיים עד סוף 2026 (מ-5% ב-2025)" },
  ];
  const xs = [9.75, 6.83, 3.91, 0.99];
  stats.forEach((st, i) => {
    statCallout(s, xs[i], 1.9, 2.7, 2.9, st.big, st.small);
  });

  s.addText("זה לא היפ שמדיה — זה כבר קרה בפועל. השאלה היא לא \"האם\", אלא \"איך נכון\".", {
    x: 0.6, y: 5.2, w: 12.13, h: 0.7, align: "right",
    fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 16, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addSourceCaption(s, "מקור: index.dev Developer Productivity Statistics 2026 · Forbes 2026 · Deloitte");
  addFooter(s, 5);
}

// =====================================================================
// SLIDE 6 — Evolution of Software Development
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
  const positions = [10.73, 8.18, 5.63, 3.08, 0.53];
  stages.forEach((st, i) => {
    const x = positions[i];
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.3, w: 2.2, h: 1.5, rectRadius: 0.1, fill: { color: st.fill }, line: { type: "none" } });
    s.addText(st.label, { x, y: 2.3, w: 2.2, h: 1.5, align: "center", valign: "middle", fontFace: TITLE_FONT, bold: true, fontSize: 14, color: st.color, margin: 4 });
    if (i < stages.length - 1) {
      s.addText("←", { x: x - 0.33, y: 2.3, w: 0.3, h: 1.5, align: "center", valign: "middle", fontFace: TITLE_FONT, bold: true, fontSize: 20, color: MUTED, margin: 0 });
    }
  });

  s.addText("כל קפיצה שינתה מה נחשב \"עבודה של מפתח\"", {
    x: 0.6, y: 4.1, w: 12.13, h: 0.5, align: "right",
    fontFace: BODY_FONT, italic: true, fontSize: 15, color: MUTED, rtlMode: true, margin: 0,
  });

  addImagePlaceholder(
    s, 0.6, 4.8, 12.13, 1.9,
    "[תמונה: רצף ויזואלי לכל עידן — כרטיסי ניקוב, IDE ראשון, ממשק Git, סוכן AI]"
  );
  addFooter(s, 6);
}

// =====================================================================
// SLIDE 7 — ציר זמן מדויק
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "ציר זמן מדויק — והקצב מואץ");

  const points = [
    { x: 11.3, label: "שנות ה-70", desc: "IDE ראשונים" },
    { x: 8.65, label: "2005 / 2008", desc: "Git / GitHub" },
    { x: 6.0, label: "יוני 2021", desc: "GitHub Copilot\n(Preview)" },
    { x: 3.35, label: "נוב' 2022", desc: "ChatGPT" },
    { x: 0.7, label: "2024-2026", desc: "כלים אג'נטיים\nאמיתיים", warn: true },
  ];
  s.addShape(pres.ShapeType.line, { x: 1.0, y: 2.6, w: 10.9, h: 0, line: { color: PH_BORDER, width: 2 } });
  points.forEach((p, i) => {
    s.addShape(pres.ShapeType.ellipse, {
      x: p.x, y: 2.45, w: 0.3, h: 0.3,
      fill: { color: p.warn ? WARN : TEAL }, line: { color: WHITE, width: 2 },
    });
    const up = i % 2 === 0;
    s.addText(p.label, {
      x: p.x - 0.85, y: up ? 1.65 : 2.9, w: 2.0, h: 0.35, align: "center",
      fontFace: TITLE_FONT, bold: true, fontSize: 13, color: TEXT_DARK, margin: 0,
    });
    s.addText(p.desc, {
      x: p.x - 0.85, y: up ? 2.05 : 3.3, w: 2.0, h: 0.6, align: "center",
      fontFace: BODY_FONT, fontSize: 11, color: MUTED, rtlMode: true, margin: 0,
    });
  });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 4.6, w: 12.13, h: 1.5, rectRadius: 0.1, fill: { color: TEAL_TINT }, line: { type: "none" } });
  s.addText("מ-IDE ל-Git: ~30 שנה. מ-Git ל-Copilot: ~16 שנה. מ-Copilot לאג'נטי: ~3 שנים בלבד.", {
    x: 0.9, y: 4.6, w: 11.53, h: 1.5, align: "right", valign: "middle",
    fontFace: TITLE_FONT, bold: true, fontSize: 18, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addSourceCaption(s, "מקור: Wikipedia (GitHub Copilot) · Chronological History of Version Control Systems");
  addFooter(s, 7);
}

// =====================================================================
// SLIDE 8 — יתרונות
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
    s.addText(r.h, { x: 0.6, y, w: 5.35, h: 0.4, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 18, color: TEXT_DARK, rtlMode: true, margin: 0 });
    s.addText(r.d, { x: 0.6, y: y + 0.42, w: 5.35, h: 0.9, align: "right", fontFace: BODY_FONT, fontSize: 14, color: MUTED, rtlMode: true, margin: 0 });
  });

  addImagePlaceholder(s, 7.1, 1.7, 5.63, 4.6, "[תמונה: השוואת ״לפני / אחרי״ — זמן פיתוח שהצטמצם מימים לדקות]");
  addFooter(s, 8);
}

// =====================================================================
// SLIDE 9 — חסרונות ומגבלות
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
  const pos = [{ x: 6.85, y: 1.7 }, { x: 0.6, y: 1.7 }, { x: 6.85, y: 3.85 }, { x: 0.6, y: 3.85 }];
  cards.forEach((c, i) => {
    const { x, y } = pos[i];
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 5.9, h: 2.0, rectRadius: 0.1, fill: { color: WARN_TINT }, line: { type: "none" } });
    iconCircle(s, x + 5.1, y + 0.25, 0.5, "!", WARN, WHITE);
    s.addText(c.h, { x: x + 0.3, y: y + 0.22, w: 4.6, h: 0.5, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 16, color: TEXT_DARK, rtlMode: true, margin: 0 });
    s.addText(c.d, { x: x + 0.3, y: y + 0.8, w: 5.3, h: 1.0, align: "right", fontFace: BODY_FONT, fontSize: 13, color: MUTED, rtlMode: true, margin: 0 });
  });

  s.addText("AI מאיץ עבודה, אבל לא מחליף שיפוט הנדסי. ככל שהפרויקט חשוב יותר — כך הביקורת האנושית קריטית יותר.", {
    x: 0.6, y: 6.05, w: 12.13, h: 0.6, align: "right",
    fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 14, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addFooter(s, 9);
}

// =====================================================================
// SLIDE 10 — האמון עדיין נמוך
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "האמון עדיין נמוך, ובצדק");

  const stats = [
    { big: "84%", small: "משתמשים/מתכננים להשתמש בכלי AI" },
    { big: "29%", small: "סומכים בפועל על הפלט" },
    { big: "61%", small: "מסכימים: \"קוד AI נראה נכון אבל לא אמין\"" },
    { big: "41%+", small: "יותר באגים בקוד AI-heavy" },
  ];
  const xs = [9.75, 6.83, 3.91, 0.99];
  stats.forEach((st, i) => {
    statCallout(s, xs[i], 1.9, 2.7, 2.9, st.big, st.small, { fill: WARN_TINT, bigColor: WARN });
  });

  s.addText("הפער בין \"משתמשים\" ל\"סומכים\" הוא בדיוק מה שהקורס הזה בא לצמצם.", {
    x: 0.6, y: 5.2, w: 12.13, h: 0.7, align: "right",
    fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 16, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addSourceCaption(s, "מקור: Second Talent Developer Survey 2026");
  addFooter(s, 10);
}

// =====================================================================
// SLIDE 11 — המחקר המטריד METR
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };

  s.addText("המחקר המטריד: METR", {
    x: 0.6, y: 0.5, w: 12.13, h: 0.7, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 28, color: WHITE, rtlMode: true, margin: 0,
  });
  s.addText("מפתחים מנוסים עבדו עם כלי AI על משימות אמיתיות:", {
    x: 0.6, y: 1.3, w: 12.13, h: 0.5, align: "right",
    fontFace: BODY_FONT, fontSize: 15, color: ICE, rtlMode: true, margin: 0,
  });

  s.addShape(pres.ShapeType.roundRect, { x: 6.93, y: 2.1, w: 5.6, h: 3.3, rectRadius: 0.12, fill: { color: DARK2 }, line: { type: "none" } });
  s.addText("19%", { x: 6.93, y: 2.4, w: 5.6, h: 1.5, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 60, color: WARN, margin: 0 });
  s.addText("איטיים יותר בפועל", { x: 6.93, y: 3.9, w: 5.6, h: 0.5, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 15, color: WHITE, rtlMode: true, margin: 0 });
  s.addText("(בפועל — מדידה אובייקטיבית)", { x: 6.93, y: 4.35, w: 5.6, h: 0.4, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 11, color: ICE, rtlMode: true, margin: 0 });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 2.1, w: 5.6, h: 3.3, rectRadius: 0.12, fill: { color: TEAL }, line: { type: "none" } });
  s.addText("20%", { x: 0.6, y: 2.4, w: 5.6, h: 1.5, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 60, color: DARK, margin: 0 });
  s.addText("האמינו שהם מהירים יותר", { x: 0.6, y: 3.9, w: 5.6, h: 0.5, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 15, color: DARK, rtlMode: true, margin: 0 });
  s.addText("(תחושה סובייקטיבית)", { x: 0.6, y: 4.35, w: 5.6, h: 0.4, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 11, color: DARK, rtlMode: true, margin: 0 });

  s.addText("התחושה \"AI מאיץ אותי\" יכולה פשוט לא להיות נכונה — גם אצל מומחים. לכן מודדים, לא מרגישים.", {
    x: 0.6, y: 5.7, w: 12.13, h: 0.8, align: "right",
    fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 16, color: WHITE, rtlMode: true, margin: 0,
  });
  addFooter(s, 11, true);
}

// =====================================================================
// SLIDE 12 — Workflow של מפתח מודרני
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
    s.addShape(pres.ShapeType.roundRect, { x, y, w: 2.15, h: 1.5, rectRadius: 0.1, fill: { color: TEAL_TINT }, line: { type: "none" } });
    iconCircle(s, x + 0.78, y + 0.15, 0.55, i + 1, TEAL, WHITE);
    s.addText(label, { x, y: y + 0.78, w: 2.15, h: 0.65, align: "center", valign: "top", fontFace: TITLE_FONT, bold: true, fontSize: 12, color: TEXT_DARK, margin: 2 });
  });

  s.addText("↻ תהליך איטרטיבי — לא ליניארי חד-פעמי", {
    x: 0.6, y: 5.25, w: 12.13, h: 0.4, align: "right",
    fontFace: BODY_FONT, italic: true, fontSize: 13, color: MUTED, rtlMode: true, margin: 0,
  });
  s.addText("זהו העיקרון-על של כל הקורס", {
    x: 0.6, y: 5.75, w: 12.13, h: 0.5, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 18, color: TEAL, rtlMode: true, margin: 0,
  });
  addFooter(s, 12);
}

// =====================================================================
// SLIDE 13 — גם הנתונים תומכים ב-Workflow
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "גם הנתונים תומכים ב-Workflow");

  s.addShape(pres.ShapeType.roundRect, { x: 6.93, y: 1.9, w: 5.6, h: 2.6, rectRadius: 0.12, fill: { color: WARN_TINT }, line: { type: "none" } });
  s.addText("~10%", { x: 6.93, y: 2.0, w: 5.6, h: 1.1, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 42, color: WARN, margin: 0 });
  s.addText("שיפור פרודוקטיביות — קוד קיים בלי Workflow מסודר (Brownfield)", { x: 7.2, y: 3.1, w: 5.1, h: 1.3, align: "center", fontFace: BODY_FONT, fontSize: 13, color: TEXT_DARK, rtlMode: true, margin: 0 });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 1.9, w: 5.6, h: 2.6, rectRadius: 0.12, fill: { color: TEAL_TINT }, line: { type: "none" } });
  s.addText("35-40%", { x: 0.6, y: 2.0, w: 5.6, h: 1.1, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 42, color: TEAL, margin: 0 });
  s.addText("שיפור פרודוקטיביות — אותם כלים, עם Workflow ברור (Greenfield)", { x: 0.85, y: 3.1, w: 5.1, h: 1.3, align: "center", fontFace: BODY_FONT, fontSize: 13, color: TEXT_DARK, rtlMode: true, margin: 0 });

  s.addText("בלי תהליך מסודר: 30-41% יותר חוב טכני · PR שנכתב ב-AI = פי 1.7 יותר בעיות", {
    x: 0.6, y: 4.75, w: 12.13, h: 0.5, align: "right",
    fontFace: BODY_FONT, fontSize: 14, color: MUTED, rtlMode: true, margin: 0,
  });
  s.addText("ההבדל הוא לא הכלי. ההבדל הוא התהליך.", {
    x: 0.6, y: 5.4, w: 12.13, h: 0.6, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 20, color: TEAL, rtlMode: true, margin: 0,
  });
  addSourceCaption(s, "מקור: DORA — ROI of AI-Assisted Software Development, 2026");
  addFooter(s, 13);
}

// =====================================================================
// SLIDE 14 — מקרה בוחן: Gemini CLI → Antigravity
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
  s.addShape(pres.ShapeType.line, { x: 1.3, y: 2.5, w: 10.6, h: 0, line: { color: PH_BORDER, width: 2 } });
  points.forEach((p) => {
    s.addShape(pres.ShapeType.ellipse, { x: p.x, y: 2.35, w: 0.3, h: 0.3, fill: { color: p.warn ? WARN : TEAL }, line: { color: WHITE, width: 2 } });
    s.addText(p.label, { x: p.x - 0.85, y: p.up ? 1.55 : 2.8, w: 2.0, h: 0.35, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 13, color: TEXT_DARK, margin: 0 });
    s.addText(p.desc, { x: p.x - 0.85, y: p.up ? 1.95 : 3.2, w: 2.0, h: 0.6, align: "center", fontFace: BODY_FONT, fontSize: 11, color: MUTED, rtlMode: true, margin: 0 });
  });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 4.05, w: 12.13, h: 1.05, rectRadius: 0.1, fill: { color: TEAL_TINT }, line: { type: "none" } });
  s.addText("מי שלמד רק ״את הכלי״ — מתחיל כמעט מאפס. מי שלמד את ה-Workflow — עובר לכלי הבא כמעט בלי לאבד זמן.", {
    x: 0.9, y: 4.15, w: 11.53, h: 0.85, align: "right", valign: "middle",
    fontFace: TITLE_FONT, bold: true, fontSize: 16, color: TEXT_DARK, rtlMode: true, margin: 0,
  });

  addImagePlaceholder(s, 0.6, 5.3, 12.13, 1.55, "[תמונה/לוגו: Gemini CLI ו-Google Antigravity CLI, זה לצד זה]");
  addFooter(s, 14);
}

// =====================================================================
// SLIDE 15 — עוד דוגמה: זה לא רק Google
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "עוד דוגמה: זה לא רק Google");

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 1.8, w: 12.13, h: 2.5, rectRadius: 0.1, fill: { color: WARN_TINT }, line: { type: "none" } });
  iconCircle(s, 11.3, 2.0, 0.5, "!", WARN, WHITE);
  s.addText(
    [
      { text: "• מרץ 2026: GitHub שינה את מסלול הסטודנטים ל-Copilot Student", options: { breakLine: true } },
      { text: "• הפעלות חדשות מוקפאות זמנית — בלי לוח זמנים לחידוש", options: { breakLine: true } },
      { text: "• גם ״תנאי מסלול חינם״ משתנים בלי אזהרה, לא רק כלים שלמים", options: { breakLine: false } },
    ],
    { x: 0.9, y: 2.1, w: 10.2, h: 2.0, align: "right", fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 4.6, w: 12.13, h: 1.5, rectRadius: 0.1, fill: { color: TEAL_TINT }, line: { type: "none" } });
  s.addText("פעילות זוגות", { x: 0.9, y: 4.78, w: 11.53, h: 0.4, align: "right", fontFace: BODY_FONT, italic: true, fontSize: 13, color: TEAL, rtlMode: true, margin: 0 });
  s.addText("תחשבו על כלי טכנולוגי שהשתנה דרמטית בחייכם תוך שנה-שנתיים — שתפו את בן/בת הזוג", {
    x: 0.9, y: 5.1, w: 11.53, h: 0.9, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 16, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addFooter(s, 15);
}

// =====================================================================
// SLIDE 16 — זה קורה גם בענק
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "זה קורה גם בענק");

  const cards = [
    { t: "Stripe", d: "מערכת \"Minions\" הפנימית — 1,000+ Pull Requests ממוזגים בשבוע" },
    { t: "TELUS", d: "500,000+ שעות עבודה נחסכו דרך 13,000 פתרונות מבוססי AI" },
    { t: "Zapier", d: "89% אימוץ AI ארגוני — כמעט כל הארגון עובד עם הכלים האלה" },
  ];
  const xs = [8.9, 4.63, 0.36];
  cards.forEach((c, i) => {
    const x = xs[i];
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.9, w: 3.85, h: 2.8, rectRadius: 0.1, fill: { color: i === 0 ? TEAL : TEAL_TINT }, line: { type: "none" } });
    s.addText(c.t, { x: x + 0.2, y: 2.1, w: 3.45, h: 0.6, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 20, color: i === 0 ? WHITE : TEXT_DARK, rtlMode: true, margin: 0 });
    s.addText(c.d, { x: x + 0.2, y: 2.8, w: 3.45, h: 1.7, align: "right", fontFace: BODY_FONT, fontSize: 13.5, color: i === 0 ? WHITE : TEXT_DARK, rtlMode: true, margin: 0 });
  });

  s.addText("זה לא רק סטארטאפים של שני אנשים — זה גם ענקיות טכנולוגיה.", {
    x: 0.6, y: 5.1, w: 12.13, h: 0.6, align: "right",
    fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 16, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addSourceCaption(s, "מקור: Forbes 2026");
  addFooter(s, 16);
}

// =====================================================================
// SLIDE 17 — הדגמה חיה
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };

  s.addText("הדגמה חיה", { x: 0.6, y: 0.5, w: 12.13, h: 0.6, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 16, color: TEAL, margin: 0 });
  s.addText("20 דקות. בונים יחד, בזמן אמת.", { x: 0.6, y: 1.2, w: 12.13, h: 1.2, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 38, color: WHITE, rtlMode: true, margin: 0 });
  s.addText("ראו demo/prompt.md להוראות המלאות", { x: 0.6, y: 2.35, w: 12.13, h: 0.5, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 15, color: ICE, rtlMode: true, margin: 0 });

  addImagePlaceholder(s, 3.5, 3.15, 6.33, 3.6, "[תמונה: מסך שיתוף / הקלטת מסך של סוכן AI בונה קוד בזמן אמת]", true);
  addFooter(s, 17, true);
}

// =====================================================================
// SLIDE 18 — עוברים לתרגול
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: TEAL };

  s.addText("עוברים לתרגול", { x: 0.6, y: 0.6, w: 12.13, h: 0.6, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 16, color: DARK, margin: 0 });
  s.addText("45 דקות: בניית Todo App באמצעות AI", { x: 0.6, y: 1.3, w: 12.13, h: 1.1, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 32, color: WHITE, rtlMode: true, margin: 0 });

  s.addText(
    [
      { text: "• הוספת משימה חדשה", options: { breakLine: true } },
      { text: "• סימון משימה כהושלמה", options: { breakLine: true } },
      { text: "• מחיקת משימה", options: { breakLine: true } },
      { text: "• שמירה עם localStorage", options: { breakLine: false } },
    ],
    { x: 0.9, y: 2.9, w: 5.6, h: 2.2, align: "right", fontFace: BODY_FONT, fontSize: 16, color: WHITE, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );

  addImagePlaceholder(s, 6.9, 2.9, 5.53, 3.2, "[אייקון/תמונה: רשימת Todo מסומנת ✓]", true);
  s.addText("ראו exercises.md ו-solutions.md לפירוט מלא", { x: 0.6, y: 6.6, w: 12.13, h: 0.4, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 13, color: DARK, rtlMode: true, margin: 0 });
}

// =====================================================================
// SLIDE 19 — סיכום
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };

  s.addText("סיכום", { x: 0.6, y: 0.5, w: 12.13, h: 0.8, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 34, color: WHITE, rtlMode: true, margin: 0 });

  const takeaways = [
    "Vibe Coding = תהליך מקצועי, לא ניחוש — והנתונים מוכיחים זאת",
    "AI מאיץ עבודה, אך 71% לא סומכים על הפלט בלי ביקורת",
    "לומדים Workflow — זה ההבדל בין 10% ל-40% שיפור פרודוקטיביות",
  ];
  takeaways.forEach((t, i) => {
    const y = 1.7 + i * 1.05;
    iconCircle(s, 11.3, y, 0.55, i + 1, TEAL, WHITE);
    s.addText(t, { x: 0.6, y, w: 10.4, h: 0.7, align: "right", valign: "middle", fontFace: TITLE_FONT, bold: true, fontSize: 17, color: WHITE, rtlMode: true, margin: 0 });
  });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 5.6, w: 12.13, h: 1.3, rectRadius: 0.1, fill: { color: TEAL }, line: { type: "none" } });
  s.addText("שבוע הבא", { x: 0.9, y: 5.78, w: 11.53, h: 0.4, align: "right", fontFace: BODY_FONT, italic: true, fontSize: 13, color: DARK, rtlMode: true, margin: 0 });
  s.addText("AI Development Environment — מקימים סביבת עבודה אמיתית", { x: 0.9, y: 6.15, w: 11.53, h: 0.6, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 18, color: WHITE, rtlMode: true, margin: 0 });
}

// =====================================================================
// SLIDE 20 — מקורות
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "מקורות");

  s.addText(
    [
      { text: "DORA — ROI of AI-Assisted Software Development (dora.dev/ai), 2026", options: { breakLine: true } },
      { text: "METR — מחקר מבוקר על מהירות מפתחים עם כלי AI, 2026", options: { breakLine: true } },
      { text: "index.dev — Developer Productivity Statistics with AI Tools 2026", options: { breakLine: true } },
      { text: "Second Talent — Developer Survey 2026", options: { breakLine: true } },
      { text: "Forbes — 5 Vibe Coding Use Cases Every Company Can Start Using Today, 2026", options: { breakLine: true } },
      { text: "Deloitte — תחזית אימוץ סוכני AI ארגוניים, 2026", options: { breakLine: true } },
      { text: "Google Developers Blog · Wikipedia (GitHub Copilot)", options: { breakLine: false } },
    ],
    { x: 0.6, y: 1.8, w: 12.13, h: 3.5, align: "right", fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, rtlMode: true, paraSpaceAfter: 12, margin: 0 }
  );

  s.addText("פירוט מלא וקישורים: references.md — כל הנתונים נבדקו ביולי 2026 ועשויים להשתנות", {
    x: 0.6, y: 5.7, w: 12.13, h: 0.6, align: "right",
    fontFace: BODY_FONT, italic: true, fontSize: 13, color: MUTED, rtlMode: true, margin: 0,
  });
  addFooter(s, 20);
}

// =====================================================================
pres.writeFile({ fileName: "lesson_01_slides.pptx" }).then(() => {
  console.log("done");
});
