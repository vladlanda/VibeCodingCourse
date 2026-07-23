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

// בונה מערך ריצות טקסט עבור רשימת בולטים (כל בולט = שורה נפרדת).
//
// הערה קריטית שנייה על RTL: בגרסה קודמת, ה"•" הוטמע כתו טקסט מילולי
// שהודבק לתחילת הריצה הראשונה (r[0].text = "• " + ...). כשהריצה הראשונה
// היא באנגלית/מודגשת (כמו "**LLM (Large Language Model)**"), זה יוצר
// ריצה מעורבת "• LLM (Large Language Model)" - ותווים ניטרליים (בולט,
// סוגריים, "=") שנמצאים בגבול בין טקסט אנגלי לעברי עלולים "להתנייד"
// למקום הלא נכון על ידי מנוע ה-bidi האמיתי של PowerPoint, גם כש-rtl="1"
// מוגדר נכון ברמת הפסקה (הדבר לא נראה ב-LibreOffice, רק בפאוארפוינט
// אמיתי - בדיוק כמו באג ה-rtlMode המקורי). התיקון: להשתמש בתכונת
// ה-bullet המובנית של PPTX (buChar ברמת פסקה) במקום תו "•" בטקסט -
// כך הבולט הוא עיצוב פסקה, לא תוכן טקסט, ותמיד יושב בקצה הנכון בלי
// תלות בכיווניות התוכן שאחריו.
function bulletRuns(bullets, base = {}) {
  const out = [];
  bullets.forEach((b, i) => {
    const r = runs(b, base);
    r[0] = {
      ...r[0],
      options: { ...r[0].options, bullet: { code: "2022", indent: 18 } },
    };
    r[r.length - 1] = {
      ...r[r.length - 1],
      options: { ...r[r.length - 1].options, breakLine: i < bullets.length - 1 },
    };
    out.push(...r);
  });
  return out;
}

// בונה מערך ריצות טקסט עבור כמה פסקאות, עם שורה ריקה מפרידה ביניהן
// (בלי בולטים - לתיבת "טקסט מלא להקלדה", כמו פרומפט)
function paragraphRuns(paragraphs, base = {}) {
  const out = [];
  paragraphs.forEach((p, i) => {
    const r = runs(p, base);
    r[r.length - 1] = {
      ...r[r.length - 1],
      options: { ...r[r.length - 1].options, breakLine: true },
    };
    out.push(...r);
    if (i < paragraphs.length - 1) {
      out.push({ text: "", options: { ...base, rtlMode: true, breakLine: true } });
    }
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

// כיתוב מתחת לאיור מקורי (וקטורי, לא תמונת אינטרנט) - "איור:" ולא "מקור:",
// כי אין כאן ציטוט חיצוני, רק תרשים שנבנה ישירות בקוד.
function addDiagramCaption(slide, text, y) {
  slide.addText("איור: " + text, {
    x: 0.6, y, w: 12.13, h: 0.3,
    align: "right", fontFace: BODY_FONT, italic: true, fontSize: 10, color: MUTED, rtlMode: true, margin: 0,
  });
}

// ---------- דיאגרמת מסך טרמינל/IDE מדומה: חלון כהה עם שורת כותרת
// (3 עיגולי בקרה כמו בכל מערכת הפעלה - LTR בכוונה, זהו מוסכמת ממשק
// אמיתית ולא טקסט RTL), וכמה שורות קוד/פלט שממחישות סוכן AI כותב קוד
// בזמן אמת. איור וקטורי מקורי - לא צילום מסך אמיתי שהורד מהאינטרנט.
function addTerminalMockup(slide, x, y, w, h) {
  const DARK = "1B1F3B";
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: DARK }, line: { color: PH_BORDER, width: 0.5 },
  });
  const barH = Math.min(0.28, h * 0.22);
  const dotColors = ["FF5F56", "FFBD2E", "27C93F"];
  dotColors.forEach((c, i) => {
    slide.addShape(pres.ShapeType.ellipse, {
      x: x + 0.14 + i * 0.22, y: y + barH / 2 - 0.05, w: 0.1, h: 0.1,
      fill: { color: c }, line: { type: "none" },
    });
  });
  slide.addText("vibe-agent — zsh", {
    x: x, y: y, w, h: barH, align: "center", valign: "middle", margin: 0,
    fontFace: BODY_FONT, fontSize: 9, color: "8A8FB0",
  });

  const lines = [
    { t: "$ vibe-agent run \"add dark mode toggle\"", c: "FFFFFF" },
    { t: "> קורא את src/App.tsx...", c: "8A8FB0" },
    { t: "> כותב קומפוננטה חדשה...", c: "5FD9C6" },
    { t: "✓ בוצע — 2 קבצים השתנו ▍", c: "27C93F" },
  ];
  const bodyY = y + barH + 0.08;
  const lineH = (h - barH - 0.12) / lines.length;
  lines.forEach((ln, i) => {
    slide.addText(ln.t, {
      x: x + 0.2, y: bodyY + i * lineH, w: w - 0.4, h: lineH,
      align: /[A-Za-z$>✓]/.test(ln.t[0]) ? "left" : "right",
      valign: "middle", margin: 0, rtlMode: !/^[$]/.test(ln.t),
      fontFace: "Courier New", fontSize: 10, color: ln.c,
    });
  });
}

// ---------- דיאגרמת LLM: שורה עליונה "קלט → מודל שפה → פלט" (3 תיבות,
// RTL: קלט מימין, פלט משמאל), ושורה תחתונה שממחישה חיזוי Token אחר
// Token (שרשרת צ'יפים קטנים עם צ'יפ אחרון מודגש כ"הבא בתור"). איור
// וקטורי מקורי - לא תמונה שהורדה מהאינטרנט.
function addLLMFlowDiagram(slide, x, y, w, h) {
  const stages = ["קלט: טקסט טבעי", "מודל שפה (LLM)", "פלט: טקסט/קוד"];
  const rowH = Math.min(0.55, h * 0.45);
  const gap = 0.25;
  const boxW = (w - gap * (stages.length - 1)) / stages.length;

  stages.forEach((stg, i) => {
    const bx = x + w - boxW - i * (boxW + gap);
    const isModel = i === 1;
    slide.addShape(pres.ShapeType.roundRect, {
      x: bx, y, w: boxW, h: rowH, rectRadius: 0.06,
      fill: { color: isModel ? TEAL : PH_BG },
      line: { color: isModel ? TEAL : PH_BORDER, width: 1.25 },
    });
    slide.addText(stg, {
      x: bx, y, w: boxW, h: rowH, margin: 4,
      align: "center", valign: "middle",
      fontFace: BODY_FONT, bold: isModel, fontSize: 12,
      color: isModel ? WHITE : TEXT_DARK, rtlMode: true,
    });
    if (i < stages.length - 1) {
      const arrowRight = bx - gap + 0.02;
      const arrowLeft = bx - 0.02;
      slide.addShape(pres.ShapeType.line, {
        x: arrowRight, y: y + rowH / 2, w: arrowLeft - arrowRight, h: 0,
        line: { color: PH_BORDER, width: 1.5, beginArrowType: "triangle" },
      });
    }
  });

  // שורת Token: צ'יפים קטנים שמומחשים כנבנים אחד אחרי השני, האחרון "צף"
  const tokRowY = y + rowH + Math.min(0.35, h - rowH - 0.15);
  const tokens = ["The", "cat", "sat", "on", "the", "___"];
  const tokW = 0.68, tokGap = 0.1;
  const totalTokW = tokens.length * tokW + (tokens.length - 1) * tokGap;
  const tokStartX = x + w - totalTokW; // RTL: מתחילים מהימין
  slide.addText("חיזוי Token אחר Token:", {
    x, y: tokRowY - 0.02, w: w - totalTokW - 0.15, h: 0.32,
    align: "right", valign: "middle", margin: 0,
    fontFace: BODY_FONT, italic: true, fontSize: 10, color: MUTED, rtlMode: true,
  });
  tokens.forEach((tk, i) => {
    const tx = tokStartX + i * (tokW + tokGap);
    const isNext = i === tokens.length - 1;
    slide.addShape(pres.ShapeType.rect, {
      x: tx, y: tokRowY, w: tokW, h: 0.32,
      fill: { color: isNext ? TEAL : WHITE },
      line: { color: TEAL, width: 1, dashType: isNext ? "dash" : "solid" },
    });
    slide.addText(tk, {
      x: tx, y: tokRowY, w: tokW, h: 0.32, margin: 0,
      align: "center", valign: "middle",
      fontFace: "Courier New", bold: true, fontSize: 10,
      color: isNext ? WHITE : TEXT_DARK,
    });
  });
}

// ---------- דיאגרמת AI-Native מול הגישה הישנה: שתי שורות של אותם 4 שלבי
// פיתוח (הבנת דרישה → תכנון → כתיבת קוד → בדיקה ופריסה), עם תג "AI" קטן
// מעל כל שלב שבו AI מעורב. בשורה "הגישה הישנה" - תג AI רק מעל "כתיבת
// קוד". בשורה "AI-Native" - תג AI מעל כל ארבעת השלבים. איור וקטורי
// מקורי - לא תמונה שהורדה מהאינטרנט.
function addAINativeDiagram(slide, x, y, w, h) {
  const stages = ["הבנת דרישה", "תכנון", "כתיבת קוד", "בדיקה ופריסה"];
  const labelW = 1.75;
  const boxAreaX = x; // RTL: תיבות השלבים מתחילות מ-x (שמאל) עד labelW לפני הקצה הימני
  const boxAreaW = w - labelW - 0.15;
  const gap = 0.12;
  const boxW = (boxAreaW - gap * (stages.length - 1)) / stages.length;
  const rowH = Math.min(0.42, h * 0.32);
  const rowGap = h - rowH * 2 > 0.3 ? h - rowH * 2 - 0.1 : 0.3;
  const row1Y = y;
  const row2Y = y + rowH + rowGap;

  function boxX(i) {
    // תיבה 0 (הבנת דרישה) הכי ימנית (RTL) - ליד התווית
    return boxAreaX + boxAreaW - boxW - i * (boxW + gap);
  }

  function drawRow(rowY, rowLabel, aiIndexes, aiColor) {
    slide.addText(rowLabel, {
      x: x + w - labelW, y: rowY, w: labelW, h: rowH,
      align: "right", valign: "middle", margin: 0,
      fontFace: TITLE_FONT, bold: true, fontSize: 11, color: TEXT_DARK, rtlMode: true,
    });
    stages.forEach((stg, i) => {
      const bx = boxX(i);
      slide.addShape(pres.ShapeType.rect, {
        x: bx, y: rowY, w: boxW, h: rowH,
        fill: { color: PH_BG }, line: { color: PH_BORDER, width: 1 },
      });
      slide.addText(stg, {
        x: bx, y: rowY, w: boxW, h: rowH, margin: 2,
        align: "center", valign: "middle",
        fontFace: BODY_FONT, fontSize: 10, color: TEXT_DARK, rtlMode: true,
      });
      if (aiIndexes.includes(i)) {
        const badgeD = 0.26;
        slide.addShape(pres.ShapeType.ellipse, {
          x: bx + boxW / 2 - badgeD / 2, y: rowY - badgeD * 0.65, w: badgeD, h: badgeD,
          fill: { color: aiColor }, line: { color: WHITE, width: 1 },
        });
        slide.addText("AI", {
          x: bx + boxW / 2 - badgeD / 2, y: rowY - badgeD * 0.65, w: badgeD, h: badgeD, margin: 0,
          align: "center", valign: "middle",
          fontFace: TITLE_FONT, bold: true, fontSize: 7, color: WHITE,
        });
      }
      if (i < stages.length - 1) {
        slide.addShape(pres.ShapeType.line, {
          x: bx - gap, y: rowY + rowH / 2, w: gap, h: 0,
          line: { color: PH_BORDER, width: 1 },
        });
      }
    });
  }

  drawRow(row1Y, "הגישה הישנה", [2], MUTED); // AI רק מעל "כתיבת קוד" (אינדקס 2)
  drawRow(row2Y, "AI-Native (הקורס)", [0, 1, 2, 3], TEAL); // AI בכל שלב
}

// ---------- דיאגרמת מעגל Workflow: 10 עיגולים ממוספרים (1 מימין ל-RTL,
// יורד ל-10 משמאל), חצים בין כל שניים עוקבים, ומסלול חזרה (מלבני) מ-10
// בחזרה ל-1 שממחיש את "התהליך המעגלי". איור וקטורי מקורי - לא תמונה
// שהורדה מהאינטרנט (אין גישת רשת לכך בסביבה הזו).
function addWorkflowCycleDiagram(slide, x, y, w, h) {
  const n = 10;
  const d = Math.min(0.5, h * 0.42); // קוטר עיגול
  const rowY = y + 0.05;
  const usableW = w - d;
  const step = usableW / (n - 1);

  // מרכזי העיגולים: step 1 בקצה הימני (RTL), step 10 בקצה השמאלי
  const centerX = (i) => x + w - d / 2 - i * step; // i = 0..9 (step i+1)

  for (let i = 0; i < n; i++) {
    const cx = centerX(i);
    slide.addShape(pres.ShapeType.ellipse, {
      x: cx - d / 2, y: rowY, w: d, h: d,
      fill: { color: i === n - 1 ? TEAL : WHITE },
      line: { color: TEAL, width: 1.5 },
    });
    slide.addText(String(i + 1), {
      x: cx - d / 2, y: rowY, w: d, h: d,
      align: "center", valign: "middle", margin: 0,
      fontFace: TITLE_FONT, bold: true, fontSize: 12,
      color: i === n - 1 ? WHITE : TEXT_DARK,
    });
    if (i < n - 1) {
      // חץ מהעיגול הנוכחי (i, מימין) לעיגול הבא (i+1, משמאלו) - ראש החץ
      // חייב לשבת בקצה השמאלי (יעד ה"זרימה"), לכן x = הקצה השמאלי (הקטן
      // יותר), w = חיובי עד הקצה הימני, ו-beginArrowType (החץ בנקודת x,y).
      const leftEdge = centerX(i + 1) + d / 2 + 0.03; // קצה ימני של העיגול הבא (i+1)
      const rightEdge = cx - d / 2 - 0.03; // קצה שמאלי של העיגול הנוכחי (i)
      slide.addShape(pres.ShapeType.line, {
        x: leftEdge, y: rowY + d / 2, w: rightEdge - leftEdge, h: 0,
        line: { color: PH_BORDER, width: 1.25, beginArrowType: "triangle" },
      });
    }
  }

  // מסלול חזרה: מ-10 (קצה שמאלי) יורד, פונה ימינה, ועולה בחזרה לתוך 1
  const leftCx = centerX(n - 1);
  const rightCx = centerX(0);
  const dropY = rowY + d + 0.22;
  slide.addShape(pres.ShapeType.line, {
    x: leftCx, y: rowY + d, w: 0, h: dropY - (rowY + d),
    line: { color: TEAL, width: 1.25, dashType: "dash" },
  });
  slide.addShape(pres.ShapeType.line, {
    x: rightCx, y: dropY, w: leftCx - rightCx, h: 0,
    line: { color: TEAL, width: 1.25, dashType: "dash" },
  });
  slide.addShape(pres.ShapeType.line, {
    x: rightCx, y: rowY + d, w: 0, h: dropY - (rowY + d),
    line: { color: TEAL, width: 1.25, dashType: "dash", beginArrowType: "triangle" },
  });
  slide.addText("חוזרים לשלב 1", {
    x: rightCx - 1.6, y: dropY - 0.02, w: 1.5, h: 0.24,
    align: "center", fontFace: BODY_FONT, italic: true, fontSize: 9, color: TEAL, rtlMode: true, margin: 0,
  });
}

// ---------- עוזר משותף: שורת מעגלים ממוספרים עם חצים (בלי מסלול חזרה) -
// משמש גם לתרשים Workflow חלק א' (שלבים 1-5). זהה בעיקרו ללוגיקת
// addWorkflowCycleDiagram, ללא הלולאה בסוף.
function addWorkflowLinearDiagram(slide, x, y, w, h, stepNumbers) {
  const n = stepNumbers.length;
  const d = Math.min(0.55, h * 0.7);
  const rowY = y + (h - d) / 2;
  const usableW = w - d;
  const step = usableW / (n - 1);
  const centerX = (i) => x + w - d / 2 - i * step;

  for (let i = 0; i < n; i++) {
    const cx = centerX(i);
    const isLast = i === n - 1;
    slide.addShape(pres.ShapeType.ellipse, {
      x: cx - d / 2, y: rowY, w: d, h: d,
      fill: { color: isLast ? TEAL : WHITE },
      line: { color: TEAL, width: 1.5 },
    });
    slide.addText(String(stepNumbers[i]), {
      x: cx - d / 2, y: rowY, w: d, h: d,
      align: "center", valign: "middle", margin: 0,
      fontFace: TITLE_FONT, bold: true, fontSize: 13,
      color: isLast ? WHITE : TEXT_DARK,
    });
    if (i < n - 1) {
      const leftEdge = centerX(i + 1) + d / 2 + 0.03;
      const rightEdge = cx - d / 2 - 0.03;
      slide.addShape(pres.ShapeType.line, {
        x: leftEdge, y: rowY + d / 2, w: rightEdge - leftEdge, h: 0,
        line: { color: PH_BORDER, width: 1.25, beginArrowType: "triangle" },
      });
    }
  }
}

// ---------- דיאגרמת "מנחים → סוכן פועל על קבצים → אתם בודקים ומאשרים":
// 3 תיבות עם מספור, מדגימות את הגדרת Vibe Coding עצמה. איור וקטורי מקורי.
function addVibeCodingFlowDiagram(slide, x, y, w, h) {
  const stages = [
    "1. מנחים בשפה טבעית מה רוצים",
    "2. הסוכן פועל ישירות על קבצי הקוד",
    "3. בודקים, מבינים, ומאשרים",
  ];
  const rowH = Math.min(0.6, h * 0.55);
  const gap = 0.25;
  const boxW = (w - gap * (stages.length - 1)) / stages.length;
  stages.forEach((stg, i) => {
    const bx = x + w - boxW - i * (boxW + gap);
    const isLast = i === stages.length - 1;
    slide.addShape(pres.ShapeType.roundRect, {
      x: bx, y, w: boxW, h: rowH, rectRadius: 0.06,
      fill: { color: isLast ? TEAL : PH_BG },
      line: { color: isLast ? TEAL : PH_BORDER, width: 1.25 },
    });
    slide.addText(stg, {
      x: bx, y, w: boxW, h: rowH, margin: 6,
      align: "center", valign: "middle",
      fontFace: BODY_FONT, bold: isLast, fontSize: 11.5,
      color: isLast ? WHITE : TEXT_DARK, rtlMode: true,
    });
    if (i < stages.length - 1) {
      const arrowRight = bx - gap + 0.02;
      const arrowLeft = bx - 0.02;
      slide.addShape(pres.ShapeType.line, {
        x: arrowRight, y: y + rowH / 2, w: arrowLeft - arrowRight, h: 0,
        line: { color: PH_BORDER, width: 1.5, beginArrowType: "triangle" },
      });
    }
  });
  // חץ עגול קטן מ-3 בחזרה ל-1, ממחיש שזו איטרציה חוזרת (לא רק חד-פעמי)
  slide.addText("↺ חוזרים על זה בכל שינוי חדש", {
    x, y: y + rowH + 0.08, w, h: 0.26,
    align: "center", fontFace: BODY_FONT, italic: true, fontSize: 10, color: MUTED, rtlMode: true, margin: 0,
  });
}

// ---------- דיאגרמת ציר-זמן Evolution: 4 עידנים בתיבות עוקבות עם חצים,
// העידן האחרון (AI Agentic) מודגש. איור וקטורי מקורי.
function addEvolutionDiagram(slide, x, y, w, h) {
  const eras = ["קוד מכונה", "שפות עיליות + IDE", "Git", "AI Agentic"];
  const rowH = Math.min(0.55, h * 0.6);
  const gap = 0.25;
  const boxW = (w - gap * (eras.length - 1)) / eras.length;
  eras.forEach((era, i) => {
    const bx = x + w - boxW - i * (boxW + gap);
    const isLast = i === eras.length - 1;
    slide.addShape(pres.ShapeType.roundRect, {
      x: bx, y, w: boxW, h: rowH, rectRadius: 0.06,
      fill: { color: isLast ? TEAL : PH_BG },
      line: { color: isLast ? TEAL : PH_BORDER, width: 1.25 },
    });
    slide.addText(era, {
      x: bx, y, w: boxW, h: rowH, margin: 4,
      align: "center", valign: "middle",
      fontFace: BODY_FONT, bold: isLast, fontSize: 12,
      color: isLast ? WHITE : TEXT_DARK, rtlMode: true,
    });
    if (i < eras.length - 1) {
      const arrowRight = bx - gap + 0.02;
      const arrowLeft = bx - 0.02;
      slide.addShape(pres.ShapeType.line, {
        x: arrowRight, y: y + rowH / 2, w: arrowLeft - arrowRight, h: 0,
        line: { color: PH_BORDER, width: 1.5, beginArrowType: "triangle" },
      });
    }
  });
  slide.addText("הקצב בין קפיצה לקפיצה הולך ומתקצר (ראו שקף הבא)", {
    x, y: y + rowH + 0.08, w, h: 0.26,
    align: "center", fontFace: BODY_FONT, italic: true, fontSize: 10, color: MUTED, rtlMode: true, margin: 0,
  });
}

// ---------- דיאגרמת "לפני/אחרי": שני פסים אופקיים באורך יחסי - "ימים"
// (פס ארוך, מוחלש) מול "דקות" (פס קצר, טיל). איור וקטורי מקורי.
function addBeforeAfterBarDiagram(slide, x, y, w, h) {
  const labelW = 1.3;
  const barAreaW = w - labelW - 0.2;
  const barH = Math.min(0.45, h * 0.3);
  const rows = [
    { label: "לפני: ימים", frac: 0.92, color: MUTED },
    { label: "היום: דקות", frac: 0.12, color: TEAL },
  ];
  const rowGap = Math.max(0.15, (h - barH * 2) / 3);
  rows.forEach((r, i) => {
    const ry = y + rowGap + i * (barH + rowGap);
    slide.addText(r.label, {
      x: x + w - labelW, y: ry, w: labelW, h: barH,
      align: "right", valign: "middle", margin: 0,
      fontFace: TITLE_FONT, bold: true, fontSize: 12, color: TEXT_DARK, rtlMode: true,
    });
    const barW = barAreaW * r.frac;
    slide.addShape(pres.ShapeType.rect, {
      x: x + w - labelW - 0.15 - barW, y: ry, w: barW, h: barH,
      fill: { color: r.color }, line: { type: "none" },
    });
  });
}

// ---------- דיאגרמת 4 מגבלות: 4 תגי אזהרה קטנים ("!") בשורה, כל אחד עם
// תווית קצרה מתחתיו - תמצות ויזואלי של 4 הבולטים בשקף. איור וקטורי מקורי.
function addLimitationsDiagram(slide, x, y, w, h) {
  const items = ["Context נשכח", "עובד ≠ נכון", "תלות בכלי", "Hallucination"];
  const n = items.length;
  const d = Math.min(0.5, h * 0.45);
  const cellW = w / n;
  items.forEach((label, i) => {
    const cx = x + w - cellW * i - cellW / 2; // RTL: פריט ראשון מימין
    slide.addShape(pres.ShapeType.ellipse, {
      x: cx - d / 2, y, w: d, h: d,
      fill: { color: MUTED }, line: { type: "none" },
    });
    slide.addText("!", {
      x: cx - d / 2, y, w: d, h: d, margin: 0,
      align: "center", valign: "middle",
      fontFace: TITLE_FONT, bold: true, fontSize: 16, color: WHITE,
    });
    slide.addText(label, {
      x: cx - cellW / 2 + 0.05, y: y + d + 0.08, w: cellW - 0.1, h: 0.4,
      align: "center", valign: "top", margin: 0,
      fontFace: BODY_FONT, fontSize: 11, color: TEXT_DARK, rtlMode: true,
    });
  });
}

// ---------- דיאגרמת Hallucination: שורת קוד מדומה (חלון כהה, בסגנון
// addTerminalMockup) עם קריאה לפונקציה בדויה, ותג "?" ליד שם הפונקציה
// שממחיש שהיא לא קיימת באמת - למרות שהקוד "נראה" תקין. איור וקטורי מקורי.
function addHallucinationDiagram(slide, x, y, w, h) {
  const DARK = "1B1F3B";
  slide.addShape(pres.ShapeType.roundRect, {
    x, y, w, h, rectRadius: 0.06,
    fill: { color: DARK }, line: { color: PH_BORDER, width: 0.5 },
  });
  const codeLine = "import { smartSort } from \"utils\";";
  slide.addText(codeLine, {
    x: x + 0.25, y, w: w - 1.1, h, margin: 0,
    align: "left", valign: "middle",
    fontFace: "Courier New", fontSize: 13, color: "FFFFFF",
  });
  const badgeD = Math.min(0.42, h * 0.6);
  slide.addShape(pres.ShapeType.ellipse, {
    x: x + w - badgeD - 0.3, y: y + (h - badgeD) / 2, w: badgeD, h: badgeD,
    fill: { color: "E05A47" }, line: { color: WHITE, width: 1.5 },
  });
  slide.addText("?", {
    x: x + w - badgeD - 0.3, y: y + (h - badgeD) / 2, w: badgeD, h: badgeD, margin: 0,
    align: "center", valign: "middle",
    fontFace: TITLE_FONT, bold: true, fontSize: 16, color: WHITE,
  });
}

// ---------- דיאגרמת "כלי הוחלף": שתי כרטיסיות זו לצד זו - Gemini CLI
// (מימין, RTL="לפני") עם סימון "נסגר", וחץ ל-Antigravity CLI (משמאל,
// "אחרי"). איור וקטורי מקורי - לא לוגו אמיתי שהורד מהאינטרנט.
function addToolReplacementDiagram(slide, x, y, w, h) {
  const cardW = (w - 0.9) / 2;
  const cardH = Math.min(0.85, h * 0.75);
  const rightX = x + w - cardW;
  const leftX = x;
  const cardY = y + (h - cardH) / 2;

  slide.addShape(pres.ShapeType.roundRect, {
    x: rightX, y: cardY, w: cardW, h: cardH, rectRadius: 0.06,
    fill: { color: PH_BG }, line: { color: PH_BORDER, width: 1.25 },
  });
  slide.addText([
    { text: "Gemini CLI", options: { bold: true, fontSize: 13, color: TEXT_DARK, breakLine: true, rtlMode: true } },
    { text: "חינמי — נסגר אמצע 2026", options: { fontSize: 10, color: "C0392B", rtlMode: true } },
  ], {
    x: rightX, y: cardY, w: cardW, h: cardH, margin: 6,
    align: "center", valign: "middle",
  });

  slide.addShape(pres.ShapeType.roundRect, {
    x: leftX, y: cardY, w: cardW, h: cardH, rectRadius: 0.06,
    fill: { color: TEAL }, line: { color: TEAL, width: 1.25 },
  });
  slide.addText([
    { text: "Antigravity CLI", options: { bold: true, fontSize: 13, color: WHITE, breakLine: true, rtlMode: true } },
    { text: "לא חינמי באותו אופן", options: { fontSize: 10, color: "E6FFFA", rtlMode: true } },
  ], {
    x: leftX, y: cardY, w: cardW, h: cardH, margin: 6,
    align: "center", valign: "middle",
  });

  const arrowX1 = rightX - 0.08;
  const arrowX2 = leftX + cardW + 0.08;
  slide.addShape(pres.ShapeType.line, {
    x: arrowX2, y: cardY + cardH / 2, w: arrowX1 - arrowX2, h: 0,
    line: { color: PH_BORDER, width: 1.5, beginArrowType: "triangle" },
  });
}

function addFooter(slide) {
  SLIDE_NUM += 1;
  slide.addText(`Vibe Coding · ${LESSON_LABEL} · שקף ${SLIDE_NUM}`, {
    x: 0.6, y: 7.1, w: 12.13, h: 0.3,
    align: "right", fontFace: BODY_FONT, fontSize: 9, color: "A6A6B0", rtlMode: true, margin: 0,
  });
}

// שקף תוכן רגיל: כותרת → שאלה → בולטים → מה עושים בפועל → (תמונה/איור) → (מקור) → פוטר
// diagram: פונקציה (slide, x, y, w, h) => void שמציירת איור וקטורי מקורי;
// diagramCaption: הטקסט שמופיע תחתיו כ"איור: ..." (לא "מקור:", כי זה לא ציטוט חיצוני)
function contentSlide({ title, question, bullets, bestPractice, image, source, diagram, diagramCaption }) {
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
  if (diagram) {
    const dH = Math.max(0.9, 6.6 - nextY);
    diagram(s, 0.6, nextY, 12.13, dH);
    nextY += dH + 0.05;
    if (diagramCaption) {
      addDiagramCaption(s, diagramCaption, nextY);
      nextY += 0.3;
    }
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

  addTerminalMockup(s, 2.5, 5.65, 8.33, 1.2);
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
    "**LLM (Large Language Model)** = תוכנה שאומנה (למדה מדוגמאות) על כמויות עצומות של טקסט וקוד",
    "כותבים בקשה בשפה טבעית (למשל: \"תכתוב לי פונקציה שממיינת רשימה\") — הטקסט הזה נקרא **Prompt** (פרומפט)",
    "המודל לא \"מבין\" את הפרומפט כמו בן אדם — מנבא **Token** (יחידת טקסט קטנה: מילה או חלק ממילה) אחר Token מה סביר לבוא בהמשך",
    "ChatGPT ו-Claude = מודלים כאלה, מוכרים לכם כצ'אטבוטים",
  ],
  bestPractice: "Vibe Coding = אותה טכנולוגיה בדיוק, רק מחוברת ישירות לקבצי הקוד שלכם במקום לחלון צ'אט.",
  diagram: addLLMFlowDiagram,
  diagramCaption: "זרימת LLM: קלט טקסט → מודל שפה → פלט. למטה: דוגמה לחיזוי Token אחר Token (המודל מנחש את המילה הבאה ברצף לפי מה שסביר).",
});

// =====================================================================
// שקף 3 — מהו Vibe Coding?
// =====================================================================
contentSlide({
  title: "מהו Vibe Coding?",
  question: "השאלה: במה זה שונה מסתם \"לבקש מ-ChatGPT קוד\"?",
  bullets: [
    "השם \"Vibe\" מרמז על תחושת זרימה: מתארים במילים מה רוצים, וה-AI הופך את זה לקוד עובד — לרוב תוך דקות ולא שעות",
    "מנחים AI לכתוב/לבדוק/לתקן קוד — לא כותבים כל שורה לבד, ושומרים שליטה על כיוון, איכות והחלטות ארכיטקטוניות",
    "**לא** = \"מקלידים פרומפט ומקווים לטוב\" · **כן** = תכנון, בדיקה, הבנה, ואחריות מלאה על הקוד",
  ],
  bestPractice: "לפני שמאשרים כל שינוי שה-AI מציע — קוראים אותו, מבינים למה הוא נכון, ורק אז ממשיכים.",
  diagram: addVibeCodingFlowDiagram,
  diagramCaption: "מעגל העבודה של Vibe Coding: מנחים בשפה טבעית → הסוכן פועל על הקבצים → בודקים ומאשרים — וחוזר חלילה בכל שינוי.",
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
  diagram: addAINativeDiagram,
  diagramCaption: "אותם 4 שלבי פיתוח בשתי גישות: בגישה הישנה AI מעורב רק בשלב כתיבת הקוד; בגישת AI-Native (הקורס) AI מעורב בכל שלב.",
});

// =====================================================================
// שקף 5 — המספרים מאחורי המהפכה
// =====================================================================
contentSlide({
  title: "המספרים מאחורי המהפכה",
  bullets: [
    "**92.6%** מהמפתחים משתמשים בכלי AI לפחות פעם בחודש",
    "**26.9%** מקוד הפרודקשן בתעשייה נכתב היום על ידי AI",
    "**1,000+** בקשות שילוב קוד (Pull Requests, מכונות בקיצור **PR**) בשבוע — דרך \"Minions\", מערכת AI פנימית שבנתה **Stripe** (חברת תשלומים דיגיטליים עולמית) כדי לבצע חלק ניכר מעבודת הפיתוח שלה אוטומטית",
    "**40%** מהאפליקציות הארגוניות — סוכני AI אוטונומיים עד סוף 2026 (מ-5% ב-2025, לפי **Deloitte** — חברת ייעוץ עסקי בינלאומית גדולה)",
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
    "קוד מכונה → שפות עיליות (כמו Python) → **IDE** (תוכנה אחת לכתיבה, הרצה ובדיקה של קוד) ו-Autocomplete (השלמה אוטומטית של קוד) → **Git** (מערכת לשמירת היסטוריית שינויים בקוד) → **AI Agentic** (AI שפועל עצמאית על משימה שלמה, לא רק עונה לשאלה)",
    "כל קפיצה שינתה מה נחשב \"עבודה של מפתח\" — ולא נעצרה (נעמיק בכל כלי לאורך הקורס)",
  ],
  bestPractice: "לומדים את התהליך, לא רק את הכלי הנוכחי — כי הקפיצה הבאה כבר בדרך. בשקף הבא נראה את הקצב הזה עם תאריכים מדויקים.",
  diagram: addEvolutionDiagram,
  diagramCaption: "4 עידנים בהתפתחות הכלים: קוד מכונה → שפות עיליות + IDE → Git → AI Agentic (העידן הנוכחי, מודגש).",
});

// =====================================================================
// שקף 7 — ציר זמן מדויק
// =====================================================================
contentSlide({
  title: "ציר זמן מדויק — והקצב מואץ",
  bullets: [
    "IDE (סביבת פיתוח לכתיבת קוד — ראינו בשקף הקודם) ראשונים: **שנות ה-70**",
    "Git (ניהול גרסאות קוד): **2005** · GitHub (פלטפורמה לאחסון ושיתוף קוד, מבוססת Git): **2008**",
    "GitHub Copilot (עוזר AI לכתיבת קוד) בגרסת Preview (תצוגה מקדימה): **יוני 2021**",
    "ChatGPT: **נובמבר 2022**",
    "כלים אג'נטיים אמיתיים (AI שפועל עצמאית על משימה שלמה): **2024-2026**",
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
  diagram: addBeforeAfterBarDiagram,
  diagramCaption: "השוואת אורך יחסי: זמן פיתוח שהצטמצם מ״ימים״ (פס ארוך) ל״דקות״ (פס קצר) עבור אותה משימה.",
});

// =====================================================================
// שקף 9 — חסרונות ומגבלות
// =====================================================================
contentSlide({
  title: "חסרונות ומגבלות",
  question: "השאלה: אם AI כל כך טוב, למה לא סתם לתת לו לעבוד לבד?",
  bullets: [
    "שוכח הקשר בלי **Context** (המידע שה-AI 'זוכר' על הפרויקט) מנוהל (נעמיק בשיעור 3)",
    "קוד שרץ ≠ קוד נכון — \"עובד\" ו\"נכון\" הם שני דברים שונים",
    "תלות בכלי שמשתנה מהר (נדבר על זה בעוד כמה שקפים)",
    "והבעיה המרכזית שבה נתמקד עכשיו: **Hallucination**",
  ],
  bestPractice: "כל אחת מהמגבלות האלה היא סיבה לשלב Review (בדיקה) אנושי בתהליך — לא סיבה לוותר על AI לגמרי.",
  diagram: addLimitationsDiagram,
  diagramCaption: "תמצות ויזואלי של 4 המגבלות שלמעלה: Context נשכח, עובד≠נכון, תלות בכלי, Hallucination.",
});

// =====================================================================
// שקף 10 — Hallucination
// =====================================================================
contentSlide({
  title: "Hallucination — הבעיה המרכזית",
  question: "השאלה: מה קורה כש-AI \"לא יודע\" משהו?",
  bullets: [
    "בניגוד לבן אדם, מודל שפה כמעט אף פעם לא אומר \"אני לא יודע\"",
    "הוא ממשיך \"לנחש\" את הטקסט הכי סביר — גם אם זה אומר להמציא פונקציה, ספרייה, או **API** (ממשק שדרכו תוכנות מדברות זו עם זו) שלא קיימים",
    "זה קורה **בביטחון מלא** — אין שום סימן חזותי שההמצאה שונה מעובדה",
    "**למסלול AI:** זו אותה תופעה שמכירים מכל הקשר אחר של מודלי שפה, לא רק קוד — ההבדל כאן הוא שהפלט השגוי הופך לבאג אמיתי שרץ במערכת",
  ],
  bestPractice: "לא מכירים פונקציה/ספרייה שה-AI השתמש בה? בודקים בתיעוד הרשמי לפני שממשיכים. ככלל אצבע: קוד שנראה מוכר מדי בקלות — זה בדיוק הזמן לעצור ולבדוק.",
  diagram: addHallucinationDiagram,
  diagramCaption: "דוגמה: import של פונקציה בשם smartSort שנראית לגיטימית לגמרי — אך לא קיימת בפועל בספרייה. ה-AI \"המציא\" אותה בביטחון מלא.",
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
    "בפרויקטים שהסתמכו הרבה על AI (AI-heavy), לעומת פרויקטים שהסתמכו עליו פחות: **41%+** יותר באגים, **ירידה של 7.2%** ביציבות המערכת",
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
    "**METR** (ארגון מחקר עצמאי שבוחן ביצועים אמיתיים של כלי AI) השווה מפתחים מנוסים שעבדו עם כלי AI על משימות אמיתיות, מול אותם מפתחים על משימות דומות בלי AI",
    "בפועל: **איטיים ב-19%** (לעומת עבודה בלי כלי AI)",
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
  diagram: (s, x, y, w, h) => addWorkflowLinearDiagram(s, x, y, w, h, [1, 2, 3, 4, 5]),
  diagramCaption: "שלבים 1-5 בזרימה מ-1 עד 5 (ממשיכים לשלבים 6-10 בשקף הבא, כולל מסלול החזרה המלא ל-1).",
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
  bestPractice: "בכל שיעור מהיום נחזור לרשימה הזו ונעמיק בשלב אחד או שניים ממנה. שימו לב: אחרי שלב 10 חוזרים לשלב 1 — זה תהליך מעגלי ואיטרטיבי, לא רצף חד-פעמי שנגמר.",
  diagram: addWorkflowCycleDiagram,
  diagramCaption: "עשרת שלבי ה-Workflow בזרימה מ-1 עד 10, עם מסלול חזרה מודגש (מקווקו) מ-10 בחזרה ל-1 — ממחיש שזה תהליך מעגלי, לא רצף חד-פעמי.",
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
    "**דו\"ח DORA** (מחקר שנתי מוביל בתעשייה על תהליכי פיתוח תוכנה, מבית Google) לשנת 2026, שיפור פרודוקטיביות בהשוואה לעבודה בלי כלי AI כלל:",
    "קוד קיים (Brownfield) בלי תהליך מסודר: **~10%** שיפור",
    "אותם כלים, על פרויקט חדש שנבנה מאפס (Greenfield) עם Workflow ברור: **35-40%** שיפור",
    "בלי תהליך מסודר: **30-41%** יותר **חוב טכני** (עבודה \"מהירה\" עכשיו שדורשת תיקון יקר יותר בהמשך), ו-PR עם AI = **פי 1.7** יותר בעיות מ-PR שכתב בן אדם",
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
    "Google הריצה כלי **CLI** (ממשק שורת-פקודה — עובדים דרך הקלדת פקודות טקסט, בלי עכבר) חינמי — Gemini CLI",
    "אמצע 2026: נסגר, הוחלף ב-Antigravity CLI (לא חינמי באותו אופן)",
    "מי שלמד רק \"את הכלי\" — מתחיל כמעט מאפס",
    "מי שלמד את ה-Workflow — עובר לכלי הבא כמעט בלי לאבד זמן",
  ],
  bestPractice: "משקיעים בהבנת התהליך של עבודה עם כלי אג'נטי — כמו **Claude Code** (כלי AI לכתיבת קוד מבית **Anthropic**, החברה שיצרה את Claude), **Cursor** (עורך קוד עם AI מובנה) או Gemini CLI — לא רק ב\"איך לוחצים איפה\" בכלי ספציפי אחד.",
  diagram: addToolReplacementDiagram,
  diagramCaption: "Gemini CLI (חינמי, נסגר אמצע 2026) הוחלף ב-Antigravity CLI (לא חינמי באותו אופן). איור המחשה - לא לוגו רשמי.",
});

// =====================================================================
// שקף 17 — עוד דוגמה: זה לא רק Google
// =====================================================================
contentSlide({
  title: "עוד דוגמה: זה לא רק Google",
  question: "השאלה: זה קרה גם למישהו אחר, או שזה מקרה בודד?",
  bullets: [
    "מרץ 2026: **GitHub** (הפלטפורמה המרכזית בעולם לאחסון ושיתוף קוד — נסביר עליה לעומק בשיעור הבא) שינה את מסלול הסטודנטים ל-**Copilot Student** (התוכנית שנותנת לסטודנטים גישה חינמית ל-GitHub Copilot)",
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

// עוזר לשקפי "תיבת פרומפט" - כותרת+תת-כותרת ממורכזות, ואז תיבת/תיבות טקסט
// עם הפרומפט/ים המדויקים להקלדה. הכל מוטמע בשקף - בלי הפניה לקובץ חיצוני.
function promptSlide({ eyebrow, title, sub, boxes }) {
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addText(eyebrow, {
    x: 0.6, y: 0.45, w: 12.13, h: 0.4, align: "center",
    fontFace: BODY_FONT, italic: true, fontSize: 15, color: TEAL, rtlMode: true, margin: 0,
  });
  s.addShape(pres.ShapeType.line, { x: 5.16, y: 0.9, w: 3.0, h: 0, line: { color: TEAL, width: 1.5 } });
  s.addText(title, {
    x: 0.6, y: 1.05, w: 12.13, h: 0.6, align: "center",
    fontFace: TITLE_FONT, bold: true, fontSize: 26, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  if (sub) {
    s.addText(sub, {
      x: 0.6, y: 1.65, w: 12.13, h: 0.35, align: "center",
      fontFace: BODY_FONT, italic: true, fontSize: 14, color: MUTED, rtlMode: true, margin: 0,
    });
  }

  let y = sub ? 2.15 : 1.8;
  const totalH = 6.85 - y;
  const gap = 0.2;
  const boxH = (totalH - gap * (boxes.length - 1)) / boxes.length;

  boxes.forEach((box) => {
    s.addText(box.label, {
      x: 0.6, y, w: 12.13, h: 0.3, align: "right",
      fontFace: TITLE_FONT, bold: true, fontSize: 14, color: TEXT_DARK, rtlMode: true, margin: 0,
    });
    const innerY = y + 0.32;
    const innerH = boxH - 0.32;
    s.addShape(pres.ShapeType.rect, {
      x: 1.0, y: innerY, w: 11.33, h: innerH,
      fill: { color: PH_BG },
      line: { color: PH_BORDER, width: 1 },
    });
    s.addText(paragraphRuns(box.lines, { fontFace: "Courier New", fontSize: 14, color: TEXT_DARK }), {
      x: 1.3, y: innerY + 0.15, w: 10.73, h: innerH - 0.3,
      align: "right", valign: "top", rtlMode: true, margin: 0,
    });
    y += boxH + gap;
  });

  addFooter(s);
  return s;
}

// =====================================================================
// שקף 19 — הדגמה חיה (חלק א׳: פרומפט 1)
// =====================================================================
promptSlide({
  eyebrow: "הדגמה חיה (חלק א׳)",
  title: "20 דקות. בונים יחד, בזמן אמת.",
  sub: "פרומפט 1 — ההתחלה, מקלידים מול הכיתה (קובץ HTML=שפת מבנה הדף, CSS=שפת עיצוב, JS/JavaScript=שפת תכנות להתנהגות אינטראקטיבית — שלושתם ביחד = דף אינטרנט עובד):",
  boxes: [
    {
      label: "",
      lines: [
        "בוא נבנה יחד, שלב אחר שלב, אפליקציית Note-Taking פשוטה שרצה בדפדפן — קובץ HTML יחיד עם CSS ו-JS מוטמעים, בלי שרת.",
        "שלב 1: תתחיל רק במבנה הבסיסי — אזור טקסט לכתיבת פתק חדש, וכפתור שמירה. אחרי שאני מאשר, נמשיך לשלב הבא (הצגת רשימת הפתקים שנשמרו).",
      ],
    },
  ],
});

// =====================================================================
// שקף 20 — הדגמה חיה (חלק ב׳: פרומפטים 2-3)
// =====================================================================
promptSlide({
  eyebrow: "הדגמה חיה (חלק ב׳)",
  title: "ממשיכים בזמן אמת",
  sub: "אחרי שהתוצאה של פרומפט 1 עובדת מקלידים בהמשך (פרומפט 2 מבקש שמירה ב-localStorage — זיכרון מקומי בדפדפן שנשמר גם אחרי רענון הדף):",
  boxes: [
    {
      label: "פרומפט 2 (אחרי אישור שלב 1):",
      lines: [
        "מעולה. עכשיו תוסיף: הצגת כל הפתקים השמורים מתחת לטופס, כל אחד בכרטיס נפרד, עם כפתור מחיקה. השתמש ב-localStorage לשמירה בין רענוני דף.",
      ],
    },
    {
      label: "פרומפט 3 (אופציונלי, אם נשאר זמן):",
      lines: [
        "תוסיף חיפוש טקסט חופשי שמסנן את הפתקים המוצגים בזמן אמת תוך כדי הקלדה.",
      ],
    },
  ],
});

// =====================================================================
// שקף 21 — עוברים לתרגול (חלק א׳: המשימה)
// =====================================================================
contentSlide({
  title: "עוברים לתרגול: Todo App (חלק א׳ — המשימה)",
  question: "מה המטרה של 45 הדקות האלה?",
  bullets: [
    "לחוות בפועל את מלוא ה-Workflow שלמדנו היום (בקנה מידה זעיר): מהבנת דרישה, דרך שימוש ב-AI למימוש, ועד בדיקה שהתוצר עובד — בזוגות",
    "המשימה: **Todo App** באמצעות כלי AI (Claude Code / Cursor / GitHub Copilot — לפי מה שמותקן) שתומך ב: הוספת משימה, סימון כהושלמה, מחיקה, והצגת הרשימה",
    "בלי **Backend** (קוד שרץ בצד שרת מרוחק, לא בדפדפן שלכם) — מספיק קובץ HTML/CSS/JS אחד שרץ בדפדפן, עם שמירה ב-localStorage",
  ],
  bestPractice: "אחרי קבלת קוד ראשוני — לא עוצרים. מבקשים שיפור אחד נוסף (מיון, ספירה, עיצוב) כדי לתרגל איטרציה על קוד קיים, לא רק בקשה ראשונה.",
});

// =====================================================================
// שקף 22 — עוברים לתרגול (חלק ב׳: פרומפט פתיחה)
// =====================================================================
promptSlide({
  eyebrow: "עוברים לתרגול (חלק ב׳)",
  title: "פרומפט פתיחה לתרגול",
  sub: "התחילו מכאן, והתאימו לפי הצורך:",
  boxes: [
    {
      label: "",
      lines: [
        "אני רוצה לבנות Todo App פשוט שרץ בדפדפן, קובץ HTML יחיד עם CSS ו-JS מוטמעים (בלי צורך בשרת). דרישות: טופס להוספת משימה חדשה, רשימת משימות עם checkbox לסימון \"הושלם\", כפתור מחיקה לכל משימה, שמירת המצב ב-localStorage כך שהרשימה נשארת אחרי רענון הדף, עיצוב נקי ופשוט.",
        "תכתוב את הקוד המלא בקובץ אחד. תסביר בקצרה מה כל חלק עושה.",
      ],
    },
  ],
});

// =====================================================================
// שקף 23 — עוברים לתרגול (חלק ג׳: מתי סיימתם)
// =====================================================================
contentSlide({
  title: "עוברים לתרגול: מתי סיימתם? (חלק ג׳)",
  question: "השאלה: איך יודעים שהתרגול הושלם בהצלחה?",
  bullets: [
    "קובץ index.html שנפתח בדפדפן ועובד בפועל — לא רק \"קוד שנראה טוב\"",
    "אפשר להוסיף, לסמן כהושלם, ולמחוק משימה — כל השלושה נבדקו בפועל בדפדפן",
    "המשימות נשמרות גם אחרי רענון הדף (localStorage)",
    "כל אחד/ת מבני הזוג יודע/ת להסביר מה הקוד עושה, גם אם AI כתב אותו — זה חלק מהתרגול, לא בונוס",
  ],
  bestPractice: "אם ה-AI \"ממציא\" קוד שלא עובד או משתמש בפונקציה לא קיימת — זה בדיוק Hallucination שדיברנו עליו היום. לא תקלה שלכם: תקנו על ידי בקשה מפורשת מה-AI לתקן, או תיאור מדויק יותר של השגיאה שקיבלתם.",
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
    "AI מאיץ עבודה, אבל כפי שראינו — רק 29% סומכים על הפלט (כלומר 71% לא) — ולכן תמיד בודקים קוד לא מוכר (Hallucination)",
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
  s.addText(
    runs(
      "**שבוע הבא:** AI Development Environment — מקימים סביבת עבודה אמיתית",
      { fontFace: TITLE_FONT, bold: true, fontSize: 17, color: TEAL }
    ),
    { x: 0.6, y: 5.15, w: 12.13, h: 0.4, align: "right", rtlMode: true, margin: 0 }
  );
  s.addText(
    "בהמשך הקורס: Context נכון (שבוע 3) · עבודה עם קוד קיים (שבוע 4) · Agents (שבוע 5) · תכנון לפני קוד (שבוע 6) · עד פריסה מלאה בענן ואבטחה (שבועות 11-12)",
    {
      x: 0.6, y: 5.6, w: 12.13, h: 0.6, align: "right",
      fontFace: BODY_FONT, italic: true, fontSize: 13, color: MUTED, rtlMode: true, margin: 0,
    }
  );
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
