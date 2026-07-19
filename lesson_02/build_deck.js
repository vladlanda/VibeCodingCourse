const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

// ---------- palette (matches lesson_01 for course brand consistency) ----------
const DARK = "1B1F3B";
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
  slide.addText(`Vibe Coding · שיעור 2 · שקף ${n}`, {
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
  s.addText("שיעור 2: AI Development Environment", {
    x: 0.6, y: 1.05, w: 6.2, h: 1.4, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 34, color: WHITE, rtlMode: true, margin: 0,
  });
  s.addText("בסוף השיעור תדעו לענות על:", {
    x: 0.6, y: 3.0, w: 6.2, h: 0.4, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 16, color: WHITE, rtlMode: true, margin: 0,
  });
  s.addText(
    [
      { text: "• מה התפקיד של VS Code, Git ו-GitHub בזרימת העבודה?", options: { breakLine: true } },
      { text: "• מה ההבדל בין כלי Autocomplete לכלי Agentic?", options: { breakLine: true } },
      { text: "• איך פותחים ומנהלים session עם Claude Code בפועל?", options: { breakLine: true } },
      { text: "• אילו כלי AI חינמיים זמינים כדי להתחיל בלי תקציב?", options: { breakLine: false } },
    ],
    { x: 0.6, y: 3.5, w: 6.2, h: 2.6, align: "right", fontFace: BODY_FONT, fontSize: 14, color: ICE, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );

  addImagePlaceholder(s, 7.2, 0.7, 5.53, 6.1, "[תמונה: שולחן עבודה — VS Code פתוח עם פאנל Claude Code בצד]", true);
  s.addText("Vibe Coding Course · Lesson 2 of 13", {
    x: 0.6, y: 6.9, w: 6.2, h: 0.3, align: "right",
    fontFace: BODY_FONT, fontSize: 10, color: "8891C2", rtlMode: true, margin: 0,
  });
}

// =====================================================================
// SLIDE 2 — VS Code, Git, GitHub
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "VS Code, Git ו-GitHub — יסודות");

  const rows = [
    { h: "VS Code", d: "סביבת הפיתוח שרוב כלי ה-AI האג'נטיים מתחברים אליה" },
    { h: "Git", d: "רשת ביטחון: כל שינוי הוא תמונת מצב שאפשר לחזור אליה" },
    { h: "GitHub", d: "אחסון מרוחק ושיתוף פעולה (יורחב בשבוע 10)" },
  ];
  rows.forEach((r, i) => {
    const y = 1.7 + i * 1.5;
    iconCircle(s, 6.1, y, 0.6, i + 1, TEAL, WHITE);
    s.addText(r.h, { x: 0.6, y, w: 5.35, h: 0.4, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 18, color: TEXT_DARK, rtlMode: true, margin: 0 });
    s.addText(r.d, { x: 0.6, y: y + 0.42, w: 5.35, h: 0.9, align: "right", fontFace: BODY_FONT, fontSize: 14, color: MUTED, rtlMode: true, margin: 0 });
  });

  addImagePlaceholder(s, 7.1, 1.7, 5.63, 3.2, "[תמונה: ממשק VS Code עם פאנל Git פתוח]");

  s.addShape(pres.ShapeType.roundRect, { x: 7.1, y: 5.1, w: 5.63, h: 1.2, rectRadius: 0.1, fill: { color: TEAL_TINT }, line: { type: "none" } });
  s.addText("כלל אצבע: Commit נקי לפני שנותנים ל-AI לגעת בקוד", {
    x: 7.35, y: 5.1, w: 5.13, h: 1.2, align: "right", valign: "middle",
    fontFace: TITLE_FONT, bold: true, fontSize: 14, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addFooter(s, 2);
}

// =====================================================================
// SLIDE 3 — Autocomplete vs Agentic
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "מה הופך כלי ל״אג'נטי״?");

  s.addShape(pres.ShapeType.roundRect, { x: 6.93, y: 1.9, w: 5.6, h: 3.9, rectRadius: 0.12, fill: { color: "F2F2F2" }, line: { type: "none" } });
  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 1.9, w: 5.6, h: 3.9, rectRadius: 0.12, fill: { color: TEAL_TINT }, line: { type: "none" } });

  s.addText("כלי Agentic (Claude Code)", { x: 0.9, y: 2.15, w: 5.0, h: 0.5, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 18, color: TEAL, rtlMode: true, margin: 0 });
  s.addText(
    [
      { text: "• מקבל משימה ברמה גבוהה", options: { breakLine: true } },
      { text: "• קורא קבצים, כותב במספר מקומות", options: { breakLine: true } },
      { text: "• מריץ בדיקות ומתקן את עצמו", options: { breakLine: true } },
      { text: "• שותף לביצוע משימה", options: { breakLine: false } },
    ],
    { x: 0.9, y: 2.75, w: 5.0, h: 2.9, align: "right", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );

  s.addText("Autocomplete בסיסי", { x: 7.23, y: 2.15, w: 5.0, h: 0.5, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 18, color: MUTED, rtlMode: true, margin: 0 });
  s.addText(
    [
      { text: "• משלים שורה/פונקציה בודדת", options: { breakLine: true } },
      { text: "• פועל לפי הקשר מקומי בלבד", options: { breakLine: true } },
      { text: "• אתם מחליטים הכל", options: { breakLine: false } },
    ],
    { x: 7.23, y: 2.75, w: 5.0, h: 2.9, align: "right", fontFace: BODY_FONT, fontSize: 14, color: MUTED, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );

  s.addText("ההבדל: עוזר הקלדה מול שותף לביצוע משימה.", {
    x: 0.6, y: 6.0, w: 12.13, h: 0.6, align: "right",
    fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 17, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addFooter(s, 3);
}

// =====================================================================
// SLIDE 4 — Claude Code workflow
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "Claude Code — איך זה עובד בפועל");

  const steps = ["פתיחת\nSession", "הנחיה\nברורה", "תוכנית /\nפעולה", "בדיקת\nDiff", "צעד\nהבא"];
  const xs = [0.6, 2.93, 5.26, 7.59, 9.92];
  steps.forEach((label, i) => {
    const x = xs[i];
    const fill = i === 3 ? WARN_TINT : TEAL_TINT;
    s.addShape(pres.ShapeType.roundRect, { x, y: 2.2, w: 2.15, h: 1.7, rectRadius: 0.1, fill: { color: fill }, line: { type: "none" } });
    iconCircle(s, x + 0.78, y_step(), 0.55, i + 1, i === 3 ? WARN : TEAL, WHITE);
    s.addText(label, { x, y: 3.15, w: 2.15, h: 0.65, align: "center", valign: "top", fontFace: TITLE_FONT, bold: true, fontSize: 13, color: TEXT_DARK, margin: 2 });
    if (i < steps.length - 1) {
      s.addText("←", { x: x + 2.15, y: 2.2, w: 0.35, h: 1.7, align: "center", valign: "middle", fontFace: TITLE_FONT, bold: true, fontSize: 20, color: MUTED, margin: 0 });
    }
  });
  function y_step() { return 2.35; }

  s.addText("בדיקת ה-Diff היא לא אופציונלית — זה השלב שבו אתם עדיין בשליטה", {
    x: 0.6, y: 4.4, w: 12.13, h: 0.5, align: "right",
    fontFace: BODY_FONT, italic: true, fontSize: 15, color: WARN, rtlMode: true, margin: 0,
  });
  s.addText("זה בדיוק שלבים 4-6 מה-Workflow שלמדנו בשיעור 1", {
    x: 0.6, y: 5.4, w: 12.13, h: 0.5, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 16, color: TEAL, rtlMode: true, margin: 0,
  });
  addFooter(s, 4);
}

// =====================================================================
// SLIDE 5 — כלים חינמיים
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "כלים חינמיים להתחלה");

  const cards = [
    { t: "GitHub Copilot Free", d: "מסלול חינמי רחב, ללא צורך באימות סטודנט", fill: TEAL, tcolor: WHITE },
    { t: "GitHub Student Pack", d: "הרחבה למי שמאמת סטודנט (זמינות משתנה)", fill: TEAL_TINT, tcolor: TEXT_DARK },
    { t: "Cursor", d: "IDE מלא, מסלול חינמי מוגבל (מכסת השלמות)", fill: TEAL_TINT, tcolor: TEXT_DARK },
  ];
  const xs = [8.9, 4.63, 0.36];
  cards.forEach((c, i) => {
    const x = xs[i];
    s.addShape(pres.ShapeType.roundRect, { x, y: 1.8, w: 3.85, h: 2.6, rectRadius: 0.1, fill: { color: c.fill }, line: { type: "none" } });
    s.addText(c.t, { x: x + 0.2, y: 2.0, w: 3.45, h: 0.8, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 16, color: c.tcolor, rtlMode: true, margin: 0 });
    s.addText(c.d, { x: x + 0.2, y: 2.85, w: 3.45, h: 1.4, align: "right", fontFace: BODY_FONT, fontSize: 12.5, color: c.tcolor, rtlMode: true, margin: 0 });
  });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 4.75, w: 12.13, h: 1.1, rectRadius: 0.1, fill: { color: WARN_TINT }, line: { type: "none" } });
  iconCircle(s, 11.3, 4.9, 0.5, "!", WARN, WHITE);
  s.addText("שוק הכלים משתנה מהר — ראו מקרה Gemini CLI → Antigravity משיעור 1", {
    x: 0.9, y: 4.85, w: 10.2, h: 0.8, align: "right", valign: "middle",
    fontFace: TITLE_FONT, bold: true, fontSize: 14, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addFooter(s, 5);
}

// =====================================================================
// SLIDE 6 — Best Practices (1)
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "Best Practices לעבודה עם כלי Agentic (1)");

  const rows = [
    { h: "הנחיות ברורות וספציפיות", d: "“תוסיף התחברות” עמום; “תוסיף התחברות עם Google OAuth, עם redirect לדף הבית” ברור" },
    { h: "צעדים קטנים", d: "משימה גדולה מתפרקת לצעדים — קל יותר לבדוק תוצאה של צעד קטן" },
    { h: "בדיקת כל Diff", d: "לפני אישור שינוי, קוראים אותו. תמיד, בלי יוצא מן הכלל" },
  ];
  rows.forEach((r, i) => {
    const y = 1.7 + i * 1.5;
    iconCircle(s, 11.3, y, 0.6, i + 1, TEAL, WHITE);
    s.addText(r.h, { x: 0.6, y, w: 10.4, h: 0.4, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 18, color: TEXT_DARK, rtlMode: true, margin: 0 });
    s.addText(r.d, { x: 0.6, y: y + 0.42, w: 10.4, h: 0.8, align: "right", fontFace: BODY_FONT, fontSize: 14, color: MUTED, rtlMode: true, margin: 0 });
  });
  addFooter(s, 6);
}

// =====================================================================
// SLIDE 7 — Best Practices (2)
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "Best Practices לעבודה עם כלי Agentic (2)");

  const rows = [
    { h: "Git כרשת ביטחון", d: "Commit לפני כל ניסוי משמעותי — תמיד אפשר לחזור אחורה" },
    { h: "לא לסמוך בעיוורון", d: "אם משהו נראה “יותר מדי חכם” או משתמש בפונקציה לא מוכרת — בודקים שהיא קיימת" },
    { h: "איטרציה, לא ניסיון יחיד", d: "אם התוצאה הראשונה לא מושלמת, מבקשים תיקון ממוקד — לא מתחילים מאפס" },
  ];
  rows.forEach((r, i) => {
    const y = 1.7 + i * 1.5;
    iconCircle(s, 11.3, y, 0.6, i + 4, TEAL, WHITE);
    s.addText(r.h, { x: 0.6, y, w: 10.4, h: 0.4, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 18, color: TEXT_DARK, rtlMode: true, margin: 0 });
    s.addText(r.d, { x: 0.6, y: y + 0.42, w: 10.4, h: 0.8, align: "right", fontFace: BODY_FONT, fontSize: 14, color: MUTED, rtlMode: true, margin: 0 });
  });
  addFooter(s, 7);
}

// =====================================================================
// SLIDE 8 — הדגמה חיה
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };
  s.addText("הדגמה חיה", { x: 0.6, y: 0.5, w: 12.13, h: 0.6, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 16, color: TEAL, margin: 0 });
  s.addText("פותחים session אמיתי על פרויקט קיים", { x: 0.6, y: 1.2, w: 12.13, h: 1.2, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 34, color: WHITE, rtlMode: true, margin: 0 });
  s.addText("ראו demo/prompt.md להוראות המלאות", { x: 0.6, y: 2.35, w: 12.13, h: 0.5, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 15, color: ICE, rtlMode: true, margin: 0 });
  addImagePlaceholder(s, 3.5, 3.15, 6.33, 3.6, "[תמונה: מסך שיתוף — Diff מוצג לפני אישור בכלי ה-AI]", true);
  addFooter(s, 8, true);
}

// =====================================================================
// SLIDE 9 — עוברים לתרגול
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: TEAL };
  s.addText("עוברים לתרגול", { x: 0.6, y: 0.6, w: 12.13, h: 0.6, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 16, color: DARK, margin: 0 });
  s.addText("45 דקות: הקמת סביבת עבודה", { x: 0.6, y: 1.3, w: 12.13, h: 1.1, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 32, color: WHITE, rtlMode: true, margin: 0 });

  s.addText(
    [
      { text: "• VS Code + Git מותקנים", options: { breakLine: true } },
      { text: "• Repository חדש ב-GitHub", options: { breakLine: true } },
      { text: "• כלי AI חינמי מחובר", options: { breakLine: true } },
      { text: "• Commit ראשון נדחף בהצלחה", options: { breakLine: false } },
    ],
    { x: 0.9, y: 2.9, w: 5.6, h: 2.2, align: "right", fontFace: BODY_FONT, fontSize: 16, color: WHITE, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );
  addImagePlaceholder(s, 6.9, 2.9, 5.53, 3.2, "[תמונה: מסך טרמינל אחרי git push מוצלח]", true);
  s.addText("ראו exercises.md ו-solutions.md לפירוט מלא", { x: 0.6, y: 6.6, w: 12.13, h: 0.4, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 13, color: DARK, rtlMode: true, margin: 0 });
}

// =====================================================================
// SLIDE 10 — סיכום
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };
  s.addText("סיכום", { x: 0.6, y: 0.5, w: 12.13, h: 0.8, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 34, color: WHITE, rtlMode: true, margin: 0 });

  const takeaways = [
    "VS Code + Git + GitHub = השולחן שעליו עובדים",
    "Agentic ≠ Autocomplete — זה שותף לביצוע משימה",
    "יש נקודת כניסה חינמית — אין תירוץ לא להתחיל",
  ];
  takeaways.forEach((t, i) => {
    const y = 1.7 + i * 1.05;
    iconCircle(s, 11.3, y, 0.55, i + 1, TEAL, WHITE);
    s.addText(t, { x: 0.6, y, w: 10.4, h: 0.7, align: "right", valign: "middle", fontFace: TITLE_FONT, bold: true, fontSize: 18, color: WHITE, rtlMode: true, margin: 0 });
  });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 5.6, w: 12.13, h: 1.3, rectRadius: 0.1, fill: { color: TEAL }, line: { type: "none" } });
  s.addText("שבוע הבא", { x: 0.9, y: 5.78, w: 11.53, h: 0.4, align: "right", fontFace: BODY_FONT, italic: true, fontSize: 13, color: DARK, rtlMode: true, margin: 0 });
  s.addText("Context Engineering — איך ״מזינים״ לכלי את מה שהוא צריך לדעת", { x: 0.9, y: 6.15, w: 11.53, h: 0.6, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 18, color: WHITE, rtlMode: true, margin: 0 });
}

// =====================================================================
pres.writeFile({ fileName: "lesson_02_slides.pptx" }).then(() => {
  console.log("done");
});
