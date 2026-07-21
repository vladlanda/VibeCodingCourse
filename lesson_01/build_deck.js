const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

// ---------- palette (מינימלי בכוונה - השראה ממצגת NN Optimization) ----------
const TEXT_DARK = "1B1F3B";
const MUTED = "6B7280";
const TEAL = "00A88A";
const PH_BG = "F6F6F8";
const PH_BORDER = "B8B8C8";
const WHITE = "FFFFFF";

const TITLE_FONT = "Arial";
const BODY_FONT = "Calibri";

const LESSON_LABEL = "שיעור 1";
let SLIDE_NUM = 0;

// ---------- טקסט עשיר: פירוק **בולד** לריצות טקסט ----------
// הערה קריטית: כש-addText מקבל מערך של ריצות טקסט (rich text), pptxgenjs
// קורא rtlMode אך ורק מתוך options של הריצה הראשונה במערך (line[0].options.rtlMode) -
// לא מתוך ה-options הכלליים שמועברים ל-addText עצמו. לכן rtlMode:true חייב
// להיות בתוך base כאן, אחרת כל טקסט שמורכב מכמה ריצות (כל בולט, כל שאלה,
// כל "מה עושים בפועל") ייצא ללא rtl="1" ב-XML ויוצג LTR ב-PowerPoint,
// גם אם ב-LibreOffice/QA זה נראה תקין (מזהה RTL אוטומטית לפי התווים).
function runs(text, base = {}) {
  const merged = { rtlMode: true, ...base };
  const parts = text.split(/(\*\*[^*]+\*\*)/g).filter((p) => p.length > 0);
  return parts.map((p) => {
    if (p.startsWith("**") && p.endsWith("**")) {
      return { text: p.slice(2, -2), options: { ...merged, bold: true } };
    }
    return { text: p, options: { ...merged } };
  });
}

// בונה מערך ריצות טקסט עבור רשימת בולטים (כל בולט = שורה נפרדת, "• " בתחילתה)
function bulletRuns(bullets, base = {}) {
  const out = [];
  bullets.forEach((b, i) => {
    const r = runs(b, base);
    r[0] = { ...r[0], text: "• " + r[0].text };
    r[r.length - 1] = {
      ...r[r.length - 1],
      options: { ...r[r.length - 1].options, breakLine: i < bullets.length - 1 },
    };
    out.push(...r);
  });
  return out;
}

// ---------- בלוקים בסיסיים ----------
function addTitle(slide, text) {
  slide.addText(text, {
    x: 0.6, y: 0.35, w: 12.13, h: 0.75,
    align: "right", fontFace: TITLE_FONT, bold: true,
    fontSize: 28, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
}

function addQuestion(slide, text, y = 1.12) {
  slide.addText(runs(text, { fontFace: BODY_FONT, fontSize: 17, italic: true, color: MUTED }), {
    x: 0.6, y, w: 12.13, h: 0.45,
    align: "right", rtlMode: true, margin: 0,
  });
}

function addBullets(slide, bullets, y = 1.65, h = 2.6, opts = {}) {
  slide.addText(bulletRuns(bullets, { fontFace: BODY_FONT, fontSize: opts.fontSize || 19, color: TEXT_DARK }), {
    x: 0.6, y, w: 12.13, h,
    align: "right", rtlMode: true, paraSpaceAfter: 10, margin: 0,
  });
}

function addBestPractice(slide, text, y) {
  // קו דק מעל, בלי תיבת צבע - הפרדה ויזואלית מינימלית
  slide.addShape(pres.ShapeType.line, {
    x: 0.6, y, w: 12.13, h: 0, line: { color: PH_BORDER, width: 0.75 },
  });
  const withLabel = "✔ מה עושים בפועל: " + text;
  slide.addText(runs(withLabel, { fontFace: TITLE_FONT, fontSize: 17, color: TEAL }), {
    x: 0.6, y: y + 0.1, w: 12.13, h: 0.95,
    align: "right", rtlMode: true, margin: 0,
  });
}

function addImagePlaceholder(slide, x, y, w, h, label) {
  slide.addShape(pres.ShapeType.rect, {
    x, y, w, h,
    fill: { color: PH_BG },
    line: { color: PH_BORDER, width: 1, dashType: "dash" },
  });
  slide.addText(label, {
    x, y, w, h, margin: 8,
    align: "center", valign: "middle",
    fontFace: BODY_FONT, fontSize: 13, italic: true, color: MUTED, rtlMode: true,
  });
}

function addSourceCaption(slide, text) {
  slide.addText("מקור: " + text, {
    x: 0.6, y: 6.75, w: 12.13, h: 0.3,
    align: "right", fontFace: BODY_FONT, italic: true, fontSize: 10, color: MUTED, rtlMode: true, margin: 0,
  });
}

function addFooter(slide) {
  SLIDE_NUM += 1;
  slide.addText(`Vibe Coding · ${LESSON_LABEL} · שקף ${SLIDE_NUM}`, {
    x: 0.6, y: 7.1, w: 12.13, h: 0.3,
    align: "right", fontFace: BODY_FONT, fontSize: 9, color: "A6A6B0", rtlMode: true, margin: 0,
  });
}

// שקף תוכן רגיל: כותרת → שאלה → בולטים → מה עושים בפועל → (תמונה) → (מקור) → פוטר
function contentSlide({ title, question, bullets, bestPractice, image, source }) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addTitle(s, title);
  if (question) addQuestion(s, question);

  const bulletsY = question ? 1.65 : 1.35;
  const bulletsH = bestPractice ? 2.5 : 3.4;
  addBullets(s, bullets, bulletsY, bulletsH);

  let nextY = bulletsY + bulletsH + 0.15;
  if (bestPractice) {
    addBestPractice(s, bestPractice, nextY);
    nextY += 1.0;
  }
  if (image) {
    const imgH = Math.max(1.1, 6.55 - nextY);
    addImagePlaceholder(s, 0.6, nextY, 12.13, imgH, image);
    nextY += imgH + 0.1;
  }
  if (source) addSourceCaption(s, source);
  addFooter(s);
  return s;
}

// שקף-ציון-דרך פשוט (מעבר/דמו/תרגול) - עדיין רקע לבן, בלי צבעים כבדים
function milestoneSlide({ eyebrow, title, sub, bullets, image }) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  s.addText(eyebrow, {
    x: 0.6, y: 0.7, w: 12.13, h: 0.4, align: "center",
    fontFace: BODY_FONT, italic: true, fontSize: 15, color: TEAL, rtlMode: true, margin: 0,
  });
  s.addShape(pres.ShapeType.line, { x: 5.16, y: 1.15, w: 3.0, h: 0, line: { color: TEAL, width: 1.5 } });
  s.addText(title, {
    x: 0.6, y: 1.35, w: 12.13, h: 1.0, align: "center",
    fontFace: TITLE_FONT, bold: true, fontSize: 32, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  if (sub) {
    s.addText(sub, {
      x: 0.6, y: 2.35, w: 12.13, h: 0.5, align: "center",
      fontFace: BODY_FONT, italic: true, fontSize: 15, color: MUTED, rtlMode: true, margin: 0,
    });
  }
  if (bullets) {
    s.addText(bulletRuns(bullets, { fontFace: BODY_FONT, fontSize: 18, color: TEXT_DARK }), {
      x: 2.5, y: 3.1, w: 8.33, h: 2.0, align: "right", rtlMode: true, paraSpaceAfter: 8, margin: 0,
    });
  }
  if (image) addImagePlaceholder(s, 3.5, bullets ? 5.3 : 3.1, 6.33, bullets ? 1.5 : 3.0, image);
  addFooter(s);
  return s;
}

// =====================================================================
// שקף 1 — כותרת
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addText("Vibe Coding – AI-Native Software Development", {
    x: 0.6, y: 0.55, w: 12.13, h: 0.4, align: "center",
    fontFace: BODY_FONT, fontSize: 13, color: TEAL, italic: true, rtlMode: true, margin: 0,
  });
  s.addText("שיעור 1: מבוא ל-Vibe Coding", {
    x: 0.6, y: 1.05, w: 12.13, h: 1.1, align: "center",
    fontFace: TITLE_FONT, bold: true, fontSize: 38, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  s.addText("Introduction to Vibe Coding", {
    x: 0.6, y: 2.1, w: 12.13, h: 0.5, align: "center",
    fontFace: BODY_FONT, italic: true, fontSize: 16, color: MUTED, rtlMode: true, margin: 0,
  });

  s.addText("בסוף השיעור תדעו לענות על:", {
    x: 2.5, y: 3.0, w: 8.33, h: 0.4, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 16, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  s.addText(
    bulletRuns(
      [
        "מהו Vibe Coding, ובמה הוא שונה מ״להקליד פרומפט״?",
        "מה היתרונות, החסרונות והמגבלות — עם נתונים אמיתיים?",
        "מהו ה-Workflow בן 10 השלבים שנלמד לאורך כל הקורס?",
        "למה הקורס מלמד ״תהליך״, ולא ״כלי״?",
      ],
      { fontFace: BODY_FONT, fontSize: 17, color: TEXT_DARK }
    ),
    { x: 2.5, y: 3.5, w: 8.33, h: 2.0, align: "right", rtlMode: true, paraSpaceAfter: 8, margin: 0 }
  );

  addImagePlaceholder(s, 2.5, 5.7, 8.33, 1.15, "[תמונה: מסך טרמינל / IDE עם סוכן AI כותב קוד בזמן אמת]");
  s.addText("Vibe Coding Course · Lesson 1 of 13", {
    x: 0.6, y: 7.1, w: 12.13, h: 0.3, align: "center",
    fontFace: BODY_FONT, fontSize: 10, color: MUTED, rtlMode: true, margin: 0,
  });
  SLIDE_NUM += 1;
}

// =====================================================================
// שקף 2 — לפני שמתחילים: מה זה בכלל LLM?
// =====================================================================
contentSlide({
  title: "לפני שמתחילים: מה זה בכלל LLM?",
  question: "השאלה: למה בכלל צריך להבין את זה כדי ללמוד Vibe Coding?",
  bullets: [
    "**LLM (Large Language Model)** = תוכנה שאומנה על כמויות עצומות של טקסט וקוד",
    "לא \"מבין\" כמו בן אדם — מנבא Token אחר Token מה סביר לבוא בהמשך",
    "ChatGPT ו-Claude = מודלים כאלה, מוכרים לכם כצ'אטבוטים",
  ],
  bestPractice: "Vibe Coding = אותה טכנולוגיה בדיוק, רק מחוברת ישירות לקבצי הקוד שלכם במקום לחלון צ'אט.",
  image: "[תמונה: תרשים פשוט — קלט טקסט → מודל שפה (קופסה) → פלט טקסט/קוד, עם חץ לדוגמה של Token אחר Token]",
});

// =====================================================================
// שקף 3 — מהו Vibe Coding?
// =====================================================================
contentSlide({
  title: "מהו Vibe Coding?",
  question: "השאלה: במה זה שונה מסתם \"לבקש מ-ChatGPT קוד\"?",
  bullets: [
    "מנחים AI לכתוב/לבדוק/לתקן קוד — לא כותבים כל שורה לבד",
    "שומרים שליטה על כיוון, איכות והחלטות ארכיטקטוניות",
    "**לא** = \"מקלידים פרומפט ומקווים לטוב\"",
    "**כן** = תכנון, בדיקה, הבנה, ואחריות מלאה על הקוד",
  ],
  bestPractice: "לפני שמאשרים כל שינוי שה-AI מציע — קוראים אותו, מבינים למה הוא נכון, ורק אז ממשיכים.",
  image: "[תמונה: ממשק Claude Code בפעולה — מפתח נותן הנחיה, הסוכן כותב קוד]",
});

// =====================================================================
// שקף 4 — AI-Native Software Engineering
// =====================================================================
contentSlide({
  title: "AI-Native Software Engineering",
  question: "השאלה: האם AI הוא רק \"עוד כלי עזר\", או משהו אחר לגמרי?",
  bullets: [
    "**AI-Native (הגישה של הקורס):** AI כשותף עבודה מרכזי, מעורב מהבנת הדרישה ועד פריסה בענן",
    "**הגישה הישנה:** AI ככלי עזר צדדי, בעיקר להשלמת קוד בסיסית (Autocomplete)",
    "ההבדל אינו כמותי אלא מהותי — זה לא תוספת לתהליך הקיים, זהו תהליך חדש",
  ],
  bestPractice: "בכל שלב בפרויקט (לא רק בזמן הקלדה) שואלים \"איך AI יכול לעזור כאן\" — בתכנון, בבדיקה, בתיעוד.",
});

// =====================================================================
// שקף 5 — המספרים מאחורי המהפכה
// =====================================================================
contentSlide({
  title: "המספרים מאחורי המהפכה",
  bullets: [
    "**92.6%** מהמפתחים משתמשים בכלי AI לפחות פעם בחודש",
    "**26.9%** מקוד הפרודקשן בתעשייה נכתב היום על ידי AI",
    "**1,000+** בקשות שילוב קוד (Pull Requests) בשבוע — דרך \"Minions\", מערכת AI פנימית שבנתה **Stripe** (חברת תשלומים דיגיטליים עולמית) כדי לבצע חלק ניכר מעבודת הפיתוח שלה אוטומטית",
    "**40%** מהאפליקציות הארגוניות — סוכני AI אוטונומיים עד סוף 2026 (מ-5% ב-2025, Deloitte)",
  ],
  bestPractice: "זו כבר לא \"טכנולוגיה חדשה שכדאי להכיר\" — זו התשתית שרוב התעשייה כבר עובדת איתה.",
  source: "index.dev Developer Productivity Statistics 2026, Forbes 2026, Deloitte",
});

// =====================================================================
// שקף 6 — Evolution of Software Development
// =====================================================================
contentSlide({
  title: "Evolution of Software Development",
  question: "השאלה: למה בכלל ללמוד היסטוריה של כלים בשיעור על AI?",
  bullets: [
    "**התשובה:** כי ההיסטוריה מראה שהקצב שבו כלים מתחלפים הולך ומואץ — וזו בדיוק הסיבה שהקורס מלמד תהליך עבודה (Workflow), לא כלי ספציפי אחד",
    "קוד מכונה → שפות עיליות (כמו Python) → IDE + Autocomplete → Git → **AI Agentic**",
    "כל קפיצה שינתה מה נחשב \"עבודה של מפתח\" — ולא נעצרה. את IDE, Git ו-Autocomplete נסביר לעומק בשיעור הבא",
  ],
  bestPractice: "לומדים את התהליך, לא רק את הכלי הנוכחי — כי הקפיצה הבאה כבר בדרך. בשקף הבא נראה את הקצב הזה עם תאריכים מדויקים.",
  image: "[תמונה: רצף ויזואלי לכל עידן — כרטיסי ניקוב, IDE ראשון, ממשק Git, סוכן AI]",
});

// =====================================================================
// שקף 7 — ציר זמן מדויק
// =====================================================================
contentSlide({
  title: "ציר זמן מדויק — והקצב מואץ",
  bullets: [
    "IDE ראשונים: **שנות ה-70**",
    "Git: **2005** · GitHub: **2008**",
    "GitHub Copilot (Preview): **יוני 2021**",
    "ChatGPT: **נובמבר 2022**",
    "כלים אג'נטיים אמיתיים: **2024-2026**",
  ],
  bestPractice: "מ-IDE ל-Git: ~30 שנה. מ-Git ל-Copilot: ~16 שנה. מ-Copilot לאג'נטי: ~3 שנים בלבד.",
  source: "Wikipedia (GitHub Copilot) · Chronological History of Version Control Systems",
});

// =====================================================================
// שקף 8 — יתרונות
// =====================================================================
contentSlide({
  title: "יתרונות",
  question: "השאלה: אז למה בכלל להשתמש בזה?",
  bullets: [
    "**מהירות:** משימות שלקחו ימים — עכשיו לוקחות דקות",
    "**נגישות:** אפשר להתחיל בלי רקע תכנותי עמוק",
    "**שחרור מבוילרפלייט:** פחות זמן על קוד חוזר וסטנדרטי, יותר על החלטות",
  ],
  bestPractice: "משתמשים ביתרון הזה במקום הנכון — קוד סטנדרטי זה מקום מצוין ל-AI; החלטות ארכיטקטורה נשארות אצלכם.",
  image: "[תמונה: השוואת ״לפני / אחרי״ — זמן פיתוח שהצטמצם מימים לדקות]",
});

// =====================================================================
// שקף 9 — חסרונות ומגבלות
// =====================================================================
contentSlide({
  title: "חסרונות ומגבלות",
  question: "השאלה: אם AI כל כך טוב, למה לא סתם לתת לו לעבוד לבד?",
  bullets: [
    "שוכח הקשר בלי Context מנוהל (נעמיק בשיעור 3)",
    "קוד שרץ ≠ קוד נכון — \"עובד\" ו\"נכון\" הם שני דברים שונים",
    "תלות בכלי שמשתנה מהר (נדבר על זה בעוד כמה שקפים)",
    "והבעיה המרכזית שבה נתמקד עכשיו: **Hallucination**",
  ],
  bestPractice: "כל אחת מהמגבלות האלה היא סיבה לשלב Review אנושי בתהליך — לא סיבה לוותר על AI לגמרי.",
});

// =====================================================================
// שקף 10 — Hallucination
// =====================================================================
contentSlide({
  title: "Hallucination — הבעיה המרכזית",
  question: "השאלה: מה קורה כש-AI \"לא יודע\" משהו?",
  bullets: [
    "בניגוד לבן אדם, מודל שפה כמעט אף פעם לא אומר \"אני לא יודע\"",
    "הוא ממשיך \"לנחש\" את הטקסט הכי סביר — גם אם זה אומר להמציא פונקציה, ספרייה, או API שלא קיימים",
    "זה קורה **בביטחון מלא** — אין שום סימן חזותי שההמצאה שונה מעובדה",
  ],
  bestPractice: "לא מכירים פונקציה/ספרייה שה-AI השתמש בה? בודקים בתיעוד הרשמי לפני שממשיכים.",
  image: "[תמונה: קטע קוד עם קריאה לפונקציה שנראית לגיטימית, וסימן שאלה אדום ליד שמה — \"האם זה קיים באמת?\"]",
});

// =====================================================================
// שקף 11 — האמון עדיין נמוך
// =====================================================================
contentSlide({
  title: "האמון עדיין נמוך, ובצדק",
  bullets: [
    "**84%** משתמשים/מתכננים להשתמש בכלי AI",
    "אבל רק **29%** סומכים על הפלט",
    "**61%** מסכימים: \"קוד AI נראה נכון אבל לא אמין\"",
    "קוד AI-heavy: **41%+** באגים, **7.2%-** יציבות מערכת",
  ],
  bestPractice: "זה בדיוק הפער שראינו בשקף הקודם עם Hallucination — עכשיו במספרים אמיתיים.",
  source: "Second Talent Developer Survey 2026",
});

// =====================================================================
// שקף 12 — המחקר המטריד METR
// =====================================================================
contentSlide({
  title: "המחקר המטריד: METR",
  question: "השאלה: אם מפתחים מרגישים ש-AI מאיץ אותם — זה נכון?",
  bullets: [
    "**METR** (ארגון מחקר עצמאי שבוחן ביצועים אמיתיים של כלי AI) עקב אחרי מפתחים מנוסים שעבדו עם כלי AI על משימות אמיתיות",
    "בפועל: **איטיים ב-19%**",
    "אבל האמינו: **מהירים ב-20%**",
  ],
  bestPractice: "לא סומכים על התחושה \"זה הלך מהר\" — עוקבים אחרי זמן אמיתי (שעון, לא הרגשה).",
  source: "METR — מחקר מבוקר על מהירות מפתחים עם כלי AI, 2026",
});

// =====================================================================
// שקף 13 — Workflow של מפתח מודרני (חלק א׳: שלבים 1-5)
// =====================================================================
contentSlide({
  title: "Workflow של מפתח מודרני (חלק א׳: שלבים 1-5)",
  question: "השאלה: אז איך עובדים נכון עם AI, אם לא סומכים על התחושה (כמו שראינו בשקף הקודם)?",
  bullets: [
    "**1. הבנת הבעיה** — מבינים מה בדיוק צריך, לפני שנוגעים בקוד בכלל",
    "**2. Specification** — כותבים מסמך קצר שמתאר מה בונים ולמה",
    "**3. בניית Context** — אוספים את המידע שה-AI צריך כדי להבין את הפרויקט (קבצים קיימים, מוסכמות, החלטות קודמות)",
    "**4. תכנון עם AI** — מבקשים מה-AI הצעת תוכנית לפני שהוא כותב שורת קוד אחת",
    "**5. מימוש עם AI** — ה-AI כותב את הקוד בפועל, לפי התוכנית שאושרה",
  ],
  bestPractice: "חמשת השלבים הבאים (6-10) עוסקים בבדיקה ובשילוב הקוד בעולם האמיתי — בשקף הבא.",
});

// =====================================================================
// שקף 14 — Workflow של מפתח מודרני (חלק ב׳: שלבים 6-10)
// =====================================================================
contentSlide({
  title: "Workflow של מפתח מודרני (חלק ב׳: שלבים 6-10)",
  bullets: [
    "**6. Code Review** — קוראים ובודקים כל שורה שה-AI כתב, לפני שהיא נכנסת לפרויקט",
    "**7. Testing** — מריצים בדיקות אוטומטיות שמוודאות שהקוד באמת עושה מה שהוא אמור",
    "**8. Commit** — שומרים גרסה מתועדת של השינוי (עם Git — נסביר בשיעור הבא)",
    "**9. Deploy** — מעלים את הגרסה לסביבה שבה משתמשים אמיתיים יכולים להשתמש בה",
    "**10. שיתוף** — מציגים ומתעדים את מה שנבנה, לצוות או ללקוח",
  ],
  bestPractice: "בכל שיעור מהיום נחזור לרשימה הזו (כל 10 השלבים) ונעמיק בשלב אחד או שניים ממנה — זהו העיקרון-על של כל הקורס.",
});

function bulletRunsPlain(lines) {
  const out = [];
  lines.forEach((l, i) => {
    out.push({ text: l, options: { rtlMode: true, breakLine: i < lines.length - 1 } });
  });
  return out;
}

// =====================================================================
// שקף 15 — גם הנתונים תומכים ב-Workflow
// =====================================================================
contentSlide({
  title: "גם הנתונים תומכים ב-Workflow",
  bullets: [
    "**דו\"ח DORA** (מחקר שנתי מוביל בתעשייה על תהליכי פיתוח תוכנה, מבית Google) לשנת 2026:",
    "קוד קיים (Brownfield) בלי תהליך מסודר: **~10%** שיפור פרודוקטיביות",
    "אותם כלים, על פרויקט חדש שנבנה מאפס (Greenfield) עם Workflow ברור: **35-40%** שיפור",
    "בלי תהליך מסודר: **30-41%** יותר חוב טכני, PR עם AI = **פי 1.7** יותר בעיות",
  ],
  bestPractice: "ההבדל הוא לא הכלי. ההבדל הוא התהליך.",
  source: "DORA — ROI of AI-Assisted Software Development, 2026",
});

// =====================================================================
// שקף 16 — מקרה בוחן: Gemini CLI → Antigravity
// =====================================================================
contentSlide({
  title: "מקרה בוחן: Gemini CLI → Antigravity",
  question: "השאלה: מה קורה כשהכלי שלמדתם נעלם?",
  bullets: [
    "Google הריצה כלי CLI חינמי — Gemini CLI",
    "אמצע 2026: נסגר, הוחלף ב-Antigravity CLI (לא חינמי באותו אופן)",
    "מי שלמד רק \"את הכלי\" — מתחיל כמעט מאפס",
    "מי שלמד את ה-Workflow — עובר לכלי הבא כמעט בלי לאבד זמן",
  ],
  bestPractice: "משקיעים בהבנת התהליך של עבודה עם כלי אג'נטי, לא רק ב\"איך לוחצים איפה\" בכלי ספציפי אחד.",
  image: "[תמונה/לוגו: Gemini CLI ו-Google Antigravity CLI, זה לצד זה]",
});

// =====================================================================
// שקף 17 — עוד דוגמה: זה לא רק Google
// =====================================================================
contentSlide({
  title: "עוד דוגמה: זה לא רק Google",
  question: "השאלה: זה קרה גם למישהו אחר, או שזה מקרה בודד?",
  bullets: [
    "מרץ 2026: **GitHub** (הפלטפורמה המרכזית בעולם לאחסון ושיתוף קוד — נסביר עליה לעומק בשיעור הבא) שינה את מסלול הסטודנטים ל-Copilot Student",
    "הפעלות חדשות מוקפאות זמנית — בלי לוח זמנים לחידוש",
    "גם \"תנאי מסלול חינם\" משתנים בלי אזהרה, לא רק כלים שלמים",
    "**פעילות זוגות:** תחשבו על כלי שהשתנה דרמטית בחייכם תוך שנה-שנתיים — שתפו את בן/בת הזוג",
  ],
  bestPractice: "בונים על התהליך, לא על תנאי גישה של כלי מסוים.",
});

// =====================================================================
// שקף 18 — זה קורה גם בענק
// =====================================================================
contentSlide({
  title: "זה קורה גם בענק",
  question: "השאלה: אולי זה רלוונטי רק לחברות ענק/סטארטאפים קטנים?",
  bullets: [
    "**Stripe** (חברת התשלומים הדיגיטליים שראינו קודם): מערכת \"Minions\" — 1,000+ PR ממוזגים בשבוע",
    "**TELUS** (חברת תקשורת קנדית גדולה): 500,000+ שעות נחסכו, 13,000 פתרונות AI",
    "**Zapier** (פלטפורמה פופולרית לחיבור אוטומטי בין אפליקציות): 89% אימוץ AI ארגוני",
  ],
  bestPractice: "זה לא רק סטארטאפים של שני אנשים — זה גם ענקיות טכנולוגיה, באותו Workflow שלמדנו היום.",
  source: "Forbes 2026",
});

// =====================================================================
// שקף 19 — הדגמה חיה
// =====================================================================
milestoneSlide({
  eyebrow: "הדגמה חיה",
  title: "20 דקות. בונים יחד, בזמן אמת.",
  sub: "ראו demo/prompt.md להוראות המלאות, לפי ה-Workflow שלמדנו",
  image: "[תמונה: מסך שיתוף / הקלטת מסך של סוכן AI בונה קוד בזמן אמת]",
});

// =====================================================================
// שקף 20 — עוברים לתרגול
// =====================================================================
milestoneSlide({
  eyebrow: "עוברים לתרגול",
  title: "45 דקות: בניית Todo App באמצעות AI",
  bullets: ["הוספת משימה חדשה", "סימון משימה כהושלמה", "מחיקת משימה", "שמירה עם localStorage"],
  image: "[אייקון/תמונה: רשימת Todo מסומנת ✓]",
});

// =====================================================================
// שקף 21 — סיכום
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addTitle(s, "סיכום");

  const takeaways = [
    "Vibe Coding = תהליך מקצועי, לא ניחוש — והנתונים מוכיחים את זה",
    "AI מאיץ עבודה, אבל 71% לא סומכים על הפלט בלי ביקורת — ולכן תמיד בודקים קוד לא מוכר (Hallucination)",
    "לומדים Workflow — כלים יוחלפו, התהליך נשאר (וזה ההבדל בין 10% ל-40%)",
  ];
  s.addText(
    takeaways.map((t, i) => ({
      text: `${i + 1}. ${t}`,
      options: { rtlMode: true, breakLine: i < takeaways.length - 1 },
    })),
    {
      x: 0.6, y: 1.5, w: 12.13, h: 3.0, align: "right", rtlMode: true,
      fontFace: BODY_FONT, fontSize: 18, color: TEXT_DARK, paraSpaceAfter: 16, margin: 0,
    }
  );

  s.addShape(pres.ShapeType.line, { x: 0.6, y: 5.0, w: 12.13, h: 0, line: { color: PH_BORDER, width: 0.75 } });
  s.addText("שבוע הבא: AI Development Environment — מקימים סביבת עבודה אמיתית", {
    x: 0.6, y: 5.15, w: 12.13, h: 0.6, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 17, color: TEAL, rtlMode: true, margin: 0,
  });
  addFooter(s);
}

// =====================================================================
// שקף 22 — מקורות
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addTitle(s, "מקורות");

  s.addText(
    bulletRunsPlain([
      "DORA — ROI of AI-Assisted Software Development (dora.dev/ai), 2026",
      "METR — מחקר מבוקר על מהירות מפתחים עם כלי AI, 2026",
      "index.dev — Developer Productivity Statistics with AI Tools 2026",
      "Second Talent — Developer Survey 2026",
      "Forbes — 5 Vibe Coding Use Cases Every Company Can Start Using Today, 2026",
      "Deloitte — תחזית אימוץ סוכני AI ארגוניים, 2026",
      "Google Developers Blog · Wikipedia (GitHub Copilot)",
    ]),
    { x: 0.6, y: 1.8, w: 12.13, h: 3.5, align: "right", fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, rtlMode: true, paraSpaceAfter: 12, margin: 0 }
  );

  s.addText("פירוט מלא וקישורים: references.md — כל הנתונים נבדקו ביולי 2026 ועשויים להשתנות", {
    x: 0.6, y: 5.7, w: 12.13, h: 0.6, align: "right",
    fontFace: BODY_FONT, italic: true, fontSize: 13, color: MUTED, rtlMode: true, margin: 0,
  });
  addFooter(s);
}

// =====================================================================
pres.writeFile({ fileName: "lesson_01_slides.pptx" }).then(() => {
  console.log("done");
});
