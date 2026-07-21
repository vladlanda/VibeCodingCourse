const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

// ---------- palette (matches lesson_01 for course brand consistency) ----------
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
  slide.addText(`Vibe Coding · שיעור 2 · שקף ${n}`, {
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
    fontFace: TITLE_FONT, bold: true, fontSize: opts.bigSize || 38,
    color: opts.bigColor || TEAL, margin: 0,
  });
  slide.addText(small, {
    x: x + 0.15, y: y + h - 0.65, w: w - 0.3, h: 0.6, align: "center", valign: "top",
    fontFace: BODY_FONT, fontSize: 12.5, color: opts.smallColor || TEXT_DARK, rtlMode: true, margin: 0,
  });
}

// bullet list helper: title/subtitle content + a right-side image placeholder
function conceptSlide(n, title, bullets, imgLabel, opts = {}) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, title);

  s.addText(
    bullets.map((b, i) => ({ text: "• " + b, options: { rtlMode: true, breakLine: i < bullets.length - 1 } })),
    { x: 0.6, y: 1.9, w: 12.13, h: 3.0, align: "right", fontFace: BODY_FONT, fontSize: opts.fontSize || 17, color: TEXT_DARK, rtlMode: true, paraSpaceAfter: 14, margin: 0 }
  );

  addImagePlaceholder(s, 0.6, 5.0, 12.13, 1.75, imgLabel);
  addFooter(s, n);
  return s;
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
      { text: "• מה זה IDE, Git, GitHub — ולמה צריך כל אחד מהם?", options: { rtlMode: true, breakLine: true } },
      { text: "• מה ההבדל בין כלי Autocomplete לכלי Agentic?", options: { rtlMode: true, breakLine: true } },
      { text: "• איך פותחים ומנהלים session עם Claude Code בפועל?", options: { rtlMode: true, breakLine: true } },
      { text: "• אילו כלי AI חינמיים זמינים כדי להתחיל בלי תקציב?", options: { rtlMode: true, breakLine: false } },
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
// SLIDE 2 — מה זה בכלל IDE?
// =====================================================================
conceptSlide(
  2,
  "מה זה בכלל IDE?",
  [
    "קובץ טקסט רגיל (Notepad) = אין צביעת תחביר, אין השלמה, אין הרצה",
    "IDE (סביבת פיתוח משולבת) = עורך קוד + טרמינל + כלי דיבוג, הכל במקום אחד",
    "VS Code = ה-IDE החינמי והפופולרי ביותר כיום",
    "רוב כלי ה-AI האג'נטיים (כולל Claude Code) מתחברים אליו כתוסף",
  ],
  "[תמונה: השוואה ויזואלית — Notepad עם טקסט פשוט לעומת VS Code עם עורך צבעוני, טרמינל, וסייד-בר]"
);

// =====================================================================
// SLIDE 3 — מה הבעיה ש-Version Control פותר?
// =====================================================================
conceptSlide(
  3,
  "מה הבעיה ש-Version Control פותר?",
  [
    "עבודה.docx ← עבודה_v2.docx ← עבודה_v2_final.docx ← עבודה_v2_final_REALLY.docx",
    "זה בדיוק מה שקורה עם קוד בלי כלי מתאים — וגרוע יותר כשכמה אנשים עובדים יחד",
    "Version Control = מערכת ששומרת \"תמונת מצב\" מדויקת של הפרויקט, בכל נקודת זמן",
  ],
  "[תמונה: שרשרת קבצים מבולבלת עם שמות \"v2_final_REALLY\" לעומת היסטוריית Git מסודרת עם תאריכים]"
);

// =====================================================================
// SLIDE 4 — Git
// =====================================================================
{
  const s = conceptSlide(
    4,
    "Git — הכלי שעושה את זה בפועל",
    [
      "Git = כלי ניהול גרסאות שרץ על המחשב שלכם (חינמי, קוד פתוח)",
      "Repository (\"Repo\") = תיקיית הפרויקט שעליה Git \"שומר עין\"",
      "Commit = תמונת מצב שמורה, עם הודעה שמסבירה מה השתנה",
    ],
    "[תמונה: איור של \"נקודת שמירה\" במשחק מחשב, מקביל למושג Commit ב-Git]"
  );
  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 4.35, w: 12.13, h: 0.55, rectRadius: 0.08, fill: { color: TEAL_TINT }, line: { type: "none" } });
  s.addText("כלל אצבע: Commit נקי לפני שנותנים ל-AI לגעת בקוד", {
    x: 0.8, y: 4.35, w: 11.73, h: 0.55, align: "right", valign: "middle",
    fontFace: TITLE_FONT, bold: true, fontSize: 14, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
}

// =====================================================================
// SLIDE 5 — GitHub
// =====================================================================
conceptSlide(
  5,
  "GitHub — לא אותו דבר כמו Git",
  [
    "Git ≠ GitHub — בלבול נפוץ אצל מתחילים",
    "Git = הכלי שרץ אצלכם במחשב, שומר היסטוריה מקומית",
    "GitHub = שירות ענן (של מיקרוסופט) שמאחסן, מגבה, ומשתף Repositories",
    "אנלוגיה: Git הוא ה-Word, GitHub הוא Google Drive שבו הוא מגובה ומשותף",
  ],
  "[תמונה: דיאגרמה — מחשב מקומי עם Git מחובר בחץ לענן עם לוגו GitHub]"
);

// =====================================================================
// SLIDE 6 — GitHub בקנה מידה
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "GitHub בקנה מידה");

  const stats = [
    { big: "180M+", small: "מפתחים רשומים (36M+ הצטרפו בשנה האחרונה)" },
    { big: "630M+", small: "Repositories" },
    { big: "43.2M", small: "Pull Requests נמזגים בכל חודש" },
    { big: "92%", small: "מחברות Fortune 100 משתמשות ב-GitHub Enterprise" },
  ];
  const xs = [9.75, 6.83, 3.91, 0.99];
  stats.forEach((st, i) => statCallout(s, xs[i], 1.9, 2.7, 2.9, st.big, st.small));

  s.addText("זו לא \"עוד פלטפורמה\" — זו התשתית של התעשייה.", {
    x: 0.6, y: 5.2, w: 12.13, h: 0.7, align: "right",
    fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 16, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addSourceCaption(s, "מקור: Kinsta, Skillademia — GitHub Statistics 2026");
  addFooter(s, 6);
}

// =====================================================================
// SLIDE 7 — מה זה בכלל "כלי AI לכתיבת קוד"?
// =====================================================================
conceptSlide(
  7,
  "מה זה בכלל \"כלי AI לכתיבת קוד\"?",
  [
    "תוכנה שמשלבת מודל שפה (LLM, משיעור 1) בתוך סביבת הפיתוח",
    "כותבים בקשה בשפה טבעית (\"תוסיף כפתור שמוחק משימה\") ← הכלי מתרגם לקוד",
    "יש כמה \"רמות\" של כלים כאלה — מהבסיסית ביותר ועד המתקדמת",
  ],
  "[תמונה: תרשים זרימה — בקשה בשפה טבעית → מודל שפה → קוד שנכתב בפועל בקובץ]"
);

// =====================================================================
// SLIDE 8 — Autocomplete vs Agentic
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
      { text: "• מקבל משימה ברמה גבוהה", options: { rtlMode: true, breakLine: true } },
      { text: "• קורא קבצים, כותב במספר מקומות", options: { rtlMode: true, breakLine: true } },
      { text: "• מריץ בדיקות ומתקן את עצמו", options: { rtlMode: true, breakLine: true } },
      { text: "• שותף לביצוע משימה", options: { rtlMode: true, breakLine: false } },
    ],
    { x: 0.9, y: 2.75, w: 5.0, h: 2.9, align: "right", fontFace: BODY_FONT, fontSize: 14, color: TEXT_DARK, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );

  s.addText("Autocomplete בסיסי", { x: 7.23, y: 2.15, w: 5.0, h: 0.5, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 18, color: MUTED, rtlMode: true, margin: 0 });
  s.addText(
    [
      { text: "• משלים שורה/פונקציה בודדת", options: { rtlMode: true, breakLine: true } },
      { text: "• פועל לפי הקשר מקומי בלבד", options: { rtlMode: true, breakLine: true } },
      { text: "• אתם מחליטים הכל", options: { rtlMode: true, breakLine: false } },
    ],
    { x: 7.23, y: 2.75, w: 5.0, h: 2.9, align: "right", fontFace: BODY_FONT, fontSize: 14, color: MUTED, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );

  s.addText("ההבדל: עוזר הקלדה מול שותף לביצוע משימה.", {
    x: 0.6, y: 6.0, w: 12.13, h: 0.6, align: "right",
    fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 17, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addFooter(s, 8);
}

// =====================================================================
// SLIDE 9 — המרוץ לבנצ'מארק
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "המרוץ לבנצ'מארק");

  statCallout(s, 6.93, 1.9, 5.6, 2.3, "80.8%", "Claude Code ב-SWE-bench Verified", { bigSize: 44 });
  statCallout(s, 0.6, 1.9, 5.6, 2.3, "78.9-83.1%", "Claude Code ב-Terminal-Bench 2.1 (טווח לפי מודל)", { bigSize: 36 });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 4.5, w: 12.13, h: 1.8, rectRadius: 0.1, fill: { color: WARN_TINT }, line: { type: "none" } });
  iconCircle(s, 11.3, 4.7, 0.5, "!", WARN, WHITE);
  s.addText(
    "אזהרת מתודולוגיה: יש 5 גרסאות שונות של SWE-bench, ושונות של 10-20 נקודות בין הרצות זהות. תמיד בדקו מתודולוגיה לפני שסומכים על מספר בנצ'מארק בודד.",
    { x: 0.9, y: 4.65, w: 10.2, h: 1.5, align: "right", valign: "middle", fontFace: TITLE_FONT, bold: true, fontSize: 14, color: TEXT_DARK, rtlMode: true, margin: 0 }
  );
  addSourceCaption(s, "מקור: DigitalApplied, MorphLLM — 2026");
  addFooter(s, 9);
}

// =====================================================================
// SLIDE 10 — Claude Code workflow
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
    iconCircle(s, x + 0.78, 2.35, 0.55, i + 1, i === 3 ? WARN : TEAL, WHITE);
    s.addText(label, { x, y: 3.15, w: 2.15, h: 0.65, align: "center", valign: "top", fontFace: TITLE_FONT, bold: true, fontSize: 13, color: TEXT_DARK, margin: 2, rtlMode: true });
    if (i < steps.length - 1) {
      s.addText("←", { x: x + 2.15, y: 2.2, w: 0.35, h: 1.7, align: "center", valign: "middle", fontFace: TITLE_FONT, bold: true, fontSize: 20, color: MUTED, margin: 0 });
    }
  });

  s.addText("בדיקת ה-Diff היא לא אופציונלית — זה השלב שבו אתם עדיין בשליטה", {
    x: 0.6, y: 4.4, w: 12.13, h: 0.5, align: "right",
    fontFace: BODY_FONT, italic: true, fontSize: 15, color: WARN, rtlMode: true, margin: 0,
  });
  s.addText("זה בדיוק שלבים 4-6 מה-Workflow שלמדנו בשיעור 1", {
    x: 0.6, y: 5.4, w: 12.13, h: 0.5, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 16, color: TEAL, rtlMode: true, margin: 0,
  });
  addFooter(s, 10);
}

// =====================================================================
// SLIDE 11 — אתם כבר עובדים ליד AI
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: TEAL_TINT };
  addSlideTitle(s, "אתם כבר עובדים ליד AI");

  s.addText(
    [
      { text: "• רובכם כבר משתמשים ב-ChatGPT/Claude לכתיבת קוד עזר למחקר", options: { rtlMode: true, breakLine: true } },
      { text: "• מה עשיתם עם הקוד שקיבלתם? העתקתם-הדבקתם לקובץ", options: { rtlMode: true, breakLine: true } },
      { text: "• כלי אג'נטי עושה בדיוק את זה — רק כותב ישירות, וקורא את הפרויקט קודם", options: { rtlMode: true, breakLine: false } },
    ],
    { x: 0.6, y: 1.9, w: 12.13, h: 2.2, align: "right", fontFace: BODY_FONT, fontSize: 17, color: TEXT_DARK, rtlMode: true, paraSpaceAfter: 14, margin: 0 }
  );

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 4.4, w: 12.13, h: 1.3, rectRadius: 0.1, fill: { color: DARK }, line: { type: "none" } });
  s.addText("זה לא כלי חדש לגמרי — זה שדרוג למשהו שאתם כבר עושים", {
    x: 0.9, y: 4.4, w: 11.53, h: 1.3, align: "right", valign: "middle",
    fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 18, color: TEAL, rtlMode: true, margin: 0,
  });
  addFooter(s, 11);
}

// =====================================================================
// SLIDE 12 — מי משתמש במה
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "מי משתמש במה, בפועל");

  const rows = [
    { name: "ChatGPT", general: "82%", primary: "—" },
    { name: "GitHub Copilot", general: "68%", primary: "—" },
    { name: "Cursor", general: "18%", primary: "24%" },
    { name: "Claude Code", general: "10%", primary: "28%" },
  ];
  const headerY = 1.8;
  s.addText("כלי", { x: 8.5, y: headerY, w: 3.0, h: 0.5, align: "right", fontFace: TITLE_FONT, bold: true, fontSize: 14, color: MUTED, rtlMode: true, margin: 0 });
  s.addText("נתח שימוש כללי", { x: 4.5, y: headerY, w: 3.7, h: 0.5, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 14, color: MUTED, rtlMode: true, margin: 0 });
  s.addText("ככלי ראשי (Agentic)", { x: 0.6, y: headerY, w: 3.7, h: 0.5, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 14, color: MUTED, rtlMode: true, margin: 0 });

  rows.forEach((r, i) => {
    const y = 2.4 + i * 0.85;
    const highlight = r.name === "Claude Code";
    if (highlight) {
      s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: y - 0.08, w: 11.73, h: 0.75, rectRadius: 0.08, fill: { color: TEAL_TINT }, line: { type: "none" } });
    }
    s.addText(r.name, { x: 8.5, y, w: 3.0, h: 0.6, align: "right", valign: "middle", fontFace: TITLE_FONT, bold: true, fontSize: 15, color: TEXT_DARK, rtlMode: true, margin: 0 });
    s.addText(r.general, { x: 4.5, y, w: 3.7, h: 0.6, align: "center", valign: "middle", fontFace: TITLE_FONT, fontSize: 15, color: TEXT_DARK, margin: 0, rtlMode: true });
    s.addText(r.primary, { x: 0.6, y, w: 3.7, h: 0.6, align: "center", valign: "middle", fontFace: TITLE_FONT, bold: highlight, fontSize: 15, color: highlight ? TEAL : TEXT_DARK, margin: 0, rtlMode: true });
  });

  s.addText("אין \"כלי אחד נכון\" — יש כלי נגיש, ויש כלי שמוביל בקרב מומחים", {
    x: 0.6, y: 6.05, w: 12.13, h: 0.5, align: "right",
    fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 14, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  addSourceCaption(s, "מקור: Stack Overflow Developer Survey 2025 · index.dev 2026");
  addFooter(s, 12);
}

// =====================================================================
// SLIDE 13 — כלים חינמיים
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
  addFooter(s, 13);
}

// =====================================================================
// SLIDE 14 — Best Practices (1)
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
  addFooter(s, 14);
}

// =====================================================================
// SLIDE 15 — Best Practices (2)
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
  addFooter(s, 15);
}

// =====================================================================
// SLIDE 16 — המחיר של דילוג על הכללים
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };

  s.addText("המחיר של דילוג על הכללים", {
    x: 0.6, y: 0.5, w: 12.13, h: 0.7, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 28, color: WHITE, rtlMode: true, margin: 0,
  });

  s.addShape(pres.ShapeType.roundRect, { x: 6.93, y: 1.7, w: 5.6, h: 2.8, rectRadius: 0.12, fill: { color: DARK2 }, line: { type: "none" } });
  s.addText("30-41%", { x: 6.93, y: 1.85, w: 5.6, h: 1.3, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 44, color: WARN, margin: 0, rtlMode: true });
  s.addText("יותר חוב טכני כשאין תהליך מסודר", { x: 7.2, y: 3.1, w: 5.1, h: 1.2, align: "center", fontFace: BODY_FONT, fontSize: 14, color: WHITE, rtlMode: true, margin: 0 });

  s.addShape(pres.ShapeType.roundRect, { x: 0.6, y: 1.7, w: 5.6, h: 2.8, rectRadius: 0.12, fill: { color: TEAL }, line: { type: "none" } });
  s.addText("1.7x", { x: 0.6, y: 1.85, w: 5.6, h: 1.3, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 44, color: DARK, margin: 0, rtlMode: true });
  s.addText("יותר בעיות ב-PR שנכתב ב-AI לעומת PR אנושי", { x: 0.85, y: 3.1, w: 5.1, h: 1.2, align: "center", fontFace: BODY_FONT, fontSize: 14, color: DARK, rtlMode: true, margin: 0 });

  s.addText("ה-Best Practices האלה הם לא בירוקרטיה — הם ההבדל בין \"AI שעוזר\" ל\"AI שיוצר בעיה גדולה יותר\".", {
    x: 0.6, y: 4.9, w: 12.13, h: 0.9, align: "right",
    fontFace: TITLE_FONT, bold: true, italic: true, fontSize: 17, color: WHITE, rtlMode: true, margin: 0,
  });
  addSourceCaption(s, "מקור: DORA — ROI of AI-Assisted Software Development, 2026");
  addFooter(s, 16, true);
}

// =====================================================================
// SLIDE 17 — הדגמה חיה
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: DARK };
  s.addText("הדגמה חיה", { x: 0.6, y: 0.5, w: 12.13, h: 0.6, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 16, color: TEAL, margin: 0, rtlMode: true });
  s.addText("פותחים session אמיתי על פרויקט קיים", { x: 0.6, y: 1.2, w: 12.13, h: 1.2, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 34, color: WHITE, rtlMode: true, margin: 0 });
  s.addText("ראו demo/prompt.md להוראות המלאות", { x: 0.6, y: 2.35, w: 12.13, h: 0.5, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 15, color: ICE, rtlMode: true, margin: 0 });
  addImagePlaceholder(s, 3.5, 3.15, 6.33, 3.6, "[תמונה: מסך שיתוף — Diff מוצג לפני אישור בכלי ה-AI]", true);
  addFooter(s, 17, true);
}

// =====================================================================
// SLIDE 18 — עוברים לתרגול
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: TEAL };
  s.addText("עוברים לתרגול", { x: 0.6, y: 0.6, w: 12.13, h: 0.6, align: "center", fontFace: BODY_FONT, italic: true, fontSize: 16, color: DARK, margin: 0, rtlMode: true });
  s.addText("45 דקות: הקמת סביבת עבודה", { x: 0.6, y: 1.3, w: 12.13, h: 1.1, align: "center", fontFace: TITLE_FONT, bold: true, fontSize: 32, color: WHITE, rtlMode: true, margin: 0 });

  s.addText(
    [
      { text: "• VS Code + Git מותקנים", options: { rtlMode: true, breakLine: true } },
      { text: "• Repository חדש ב-GitHub", options: { rtlMode: true, breakLine: true } },
      { text: "• כלי AI חינמי מחובר", options: { rtlMode: true, breakLine: true } },
      { text: "• Commit ראשון נדחף בהצלחה", options: { rtlMode: true, breakLine: false } },
    ],
    { x: 0.9, y: 2.9, w: 5.6, h: 2.2, align: "right", fontFace: BODY_FONT, fontSize: 16, color: WHITE, rtlMode: true, paraSpaceAfter: 10, margin: 0 }
  );
  addImagePlaceholder(s, 6.9, 2.9, 5.53, 3.2, "[תמונה: מסך טרמינל אחרי git push מוצלח]", true);
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
// SLIDE 20 — מקורות
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addSlideTitle(s, "מקורות");

  s.addText(
    [
      { text: "Kinsta & Skillademia — GitHub Statistics 2026", options: { rtlMode: true, breakLine: true } },
      { text: "DigitalApplied & MorphLLM — Benchmark Guides 2026 (SWE-bench, Terminal-Bench)", options: { rtlMode: true, breakLine: true } },
      { text: "Stack Overflow Developer Survey 2025", options: { rtlMode: true, breakLine: true } },
      { text: "index.dev — Developer Productivity Statistics 2026", options: { rtlMode: true, breakLine: true } },
      { text: "DORA — ROI of AI-Assisted Software Development, 2026", options: { rtlMode: true, breakLine: false } },
    ],
    { x: 0.6, y: 1.8, w: 12.13, h: 3.0, align: "right", fontFace: BODY_FONT, fontSize: 16, color: TEXT_DARK, rtlMode: true, paraSpaceAfter: 14, margin: 0 }
  );

  s.addText("פירוט מלא וקישורים: references.md — כל הנתונים נבדקו ביולי 2026 ועשויים להשתנות", {
    x: 0.6, y: 5.7, w: 12.13, h: 0.6, align: "right",
    fontFace: BODY_FONT, italic: true, fontSize: 13, color: MUTED, rtlMode: true, margin: 0,
  });
  addFooter(s, 20);
}

// =====================================================================
pres.writeFile({ fileName: "lesson_02_slides.pptx" }).then(() => {
  console.log("done");
});
