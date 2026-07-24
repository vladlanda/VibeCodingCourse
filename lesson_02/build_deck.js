const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5

// ---------- palette (זהה לשיעור 1 לעקביות מותג הקורס) ----------
const TEXT_DARK = "1B1F3B";
const MUTED = "6B7280";
const TEAL = "00A88A";
const PH_BG = "F6F6F8";
const PH_BORDER = "B8B8C8";
const WHITE = "FFFFFF";

const TITLE_FONT = "Arial";
const BODY_FONT = "Calibri";

const LESSON_LABEL = "שיעור 2";
let SLIDE_NUM = 0;

// ---------- טקסט עשיר: פירוק **בולד** לריצות טקסט ----------
// הערה קריטית (זהה לשיעור 1): rtlMode חייב להיות בתוך options של כל ריצה
// וריצה, לא רק ב-options הכלליים של addText - אחרת pptxgenjs לא מסמן
// rtl="1" ב-XML והמצגת תיפתח LTR בפועל ב-PowerPoint אמיתי.
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

// בונה מערך ריצות טקסט עבור רשימת בולטים - בולט מובנה (buChar) ולא תו "•"
// מילולי, כדי למנוע באג מיקום בולט/סוגריים כשריצה ראשונה מתחילה באנגלית
// מודגשת (מתועד ב-lesson_01/build_deck.js ובפרומפט המרכזי).
function bulletRuns(bullets, base = {}) {
  const out = [];
  const bulletOpt = { code: "2022", indent: 18 };
  bullets.forEach((b, i) => {
    const r = runs(b, base);
    // הערה קריטית: כשבולט מכיל טקסט מודגש (**...**) בתוך המשפט, runs() מפרק
    // אותו למספר ריצות (runs). pptxgenjs פולט <a:pPr> נפרד לכל ריצה שמקבלת
    // options שונים - וכל <a:pPr> כזה שאין לו bullet מפורש מקבל <a:buNone/>
    // אוטומטית. ה-<a:buNone/> הזה "מנצח" את ה-<a:buChar/> הקודם ב-render
    // בפועל (למשל ב-LibreOffice), והבולט נעלם! לכן חובה להצמיד את אותו
    // אובייקט bullet לכל ריצה בפסקה, לא רק לריצה הראשונה - כדי שכל ה-pPr
    // שנוצרים יהיו זהים ועקביים.
    r.forEach((run) => {
      run.options = { ...run.options, bullet: bulletOpt };
    });
    r[r.length - 1] = {
      ...r[r.length - 1],
      options: { ...r[r.length - 1].options, breakLine: i < bullets.length - 1 },
    };
    out.push(...r);
  });
  return out;
}

// בונה מערך ריצות טקסט עבור כמה פסקאות עם שורה ריקה מפרידה (בלי בולטים) -
// לתיבת "טקסט מלא להקלדה", כמו פרומפט
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

function addBullets(slide, bullets, y = 1.65, h = 3.6, opts = {}) {
  slide.addText(bulletRuns(bullets, { fontFace: BODY_FONT, fontSize: opts.fontSize || 19, color: TEXT_DARK }), {
    x: 0.6, y, w: 12.13, h,
    align: "right", rtlMode: true, paraSpaceAfter: 10, margin: 0,
  });
}

// הערה קריטית: מיקום ה"מקור:" בתחתית לא יכול להיות קבוע (y=6.75 קשיח) -
// כי טקסט "מה עושים בפועל" משתנה באורך בין שקפים, ופעמים רבות עוטף לשתי
// שורות. עם מיקום קבוע, שקף עם bestPractice ארוך (שתי שורות) "מקור:"
// נדרס/מתנגש איתו בפועל ב-render. לכן addBestPractice מעריך כמה שורות
// הטקסט ידרוש (לפי אורך התו ורוחב התיבה), ומחזיר את מיקום ה-Y בפועל
// שבו הטקסט מסתיים - כדי ש-addSourceCaption תמיד תמוקם מתחתיו, לא על גביו.
function estimateLines(text, charsPerLine) {
  return Math.max(1, Math.ceil(text.length / charsPerLine));
}

function addBestPractice(slide, text, y) {
  slide.addShape(pres.ShapeType.line, {
    x: 0.6, y, w: 12.13, h: 0, line: { color: PH_BORDER, width: 0.75 },
  });
  const withLabel = "✔ מה עושים בפועל: " + text;
  const lines = estimateLines(withLabel, 108);
  const textH = Math.max(0.5, lines * 0.34);
  slide.addText(runs(withLabel, { fontFace: TITLE_FONT, fontSize: 17, color: TEAL }), {
    x: 0.6, y: y + 0.1, w: 12.13, h: textH,
    align: "right", rtlMode: true, margin: 0,
  });
  return y + 0.1 + textH;
}

function addSourceCaption(slide, text, y = 6.75) {
  slide.addText("מקור: " + text, {
    x: 0.6, y, w: 12.13, h: 0.3,
    align: "right", fontFace: BODY_FONT, italic: true, fontSize: 10, color: MUTED, rtlMode: true, margin: 0,
  });
}

// כיתוב מתחת לאיור מקורי (וקטורי, לא תמונת אינטרנט) - "איור:" ולא "מקור:",
// כי אין כאן ציטוט חיצוני, רק תרשים שנבנה ישירות בקוד (זהה לשיטת שיעור 1).
function addDiagramCaption(slide, text, y) {
  slide.addText("איור: " + text, {
    x: 0.6, y, w: 12.13, h: 0.3,
    align: "right", fontFace: BODY_FONT, italic: true, fontSize: 10, color: MUTED, rtlMode: true, margin: 0,
  });
}

// ---------- דיאגרמת מסך טרמינל מדומה לשקף הכותרת (זהה בעיצובה לשיעור 1,
// עם פקודות רלוונטיות לשיעור 2). איור וקטורי מקורי - לא צילום מסך.
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
  slide.addText("my-vibe-coding-env — zsh", {
    x, y, w, h: barH, align: "center", valign: "middle", margin: 0,
    fontFace: BODY_FONT, fontSize: 9, color: "8A8FB0",
  });
  const lines = [
    { t: "$ git clone github.com/you/my-vibe-coding-env", c: "FFFFFF" },
    { t: "$ code .", c: "FFFFFF" },
    { t: "> מבקשים מכלי ה-AI ליצור hello.py...", c: "5FD9C6" },
    { t: "✓ נוצר, נבדק, ונדחף ל-GitHub ▍", c: "27C93F" },
  ];
  const bodyY = y + barH + 0.08;
  const lineH = (h - barH - 0.12) / lines.length;
  lines.forEach((ln, i) => {
    slide.addText(ln.t, {
      x: x + 0.2, y: bodyY + i * lineH, w: w - 0.4, h: lineH,
      align: /^[$>✓]/.test(ln.t) ? "left" : "right",
      valign: "middle", margin: 0, rtlMode: !/^[$]/.test(ln.t),
      fontFace: "Courier New", fontSize: 10, color: ln.c,
    });
  });
}

// ---------- דיאגרמת השוואה: Notepad מול IDE. שתי כרטיסיות זו לצד זו -
// ימין (RTL="ראשון") = Notepad עם שורה בודדת של טקסט פשוט; שמאל = IDE
// עם 3 "פסים" קטנים (עורך/טרמינל/דיבאגר) שממחישים ריבוי-פאנלים.
// איור וקטורי מקורי.
function addIDEComparisonDiagram(slide, x, y, w, h) {
  const cardW = (w - 0.4) / 2;
  const cardH = h;
  const rightX = x + w - cardW; // Notepad (RTL: ראשון/ימני)
  const leftX = x; // IDE

  slide.addShape(pres.ShapeType.roundRect, {
    x: rightX, y, w: cardW, h: cardH, rectRadius: 0.06,
    fill: { color: PH_BG }, line: { color: PH_BORDER, width: 1.25 },
  });
  slide.addText("Notepad", {
    x: rightX, y: y + 0.08, w: cardW, h: 0.3, align: "center",
    fontFace: TITLE_FONT, bold: true, fontSize: 12, color: MUTED, margin: 0,
  });
  slide.addShape(pres.ShapeType.rect, {
    x: rightX + 0.3, y: y + 0.5, w: cardW - 0.6, h: h - 0.65,
    fill: { color: WHITE }, line: { color: PH_BORDER, width: 0.75 },
  });
  slide.addText("טקסט רגיל, בלי צבע ובלי עזרה", {
    x: rightX + 0.3, y: y + 0.5, w: cardW - 0.6, h: h - 0.65, margin: 4,
    align: "center", valign: "middle",
    fontFace: "Courier New", fontSize: 9, color: MUTED, rtlMode: true,
  });

  slide.addShape(pres.ShapeType.roundRect, {
    x: leftX, y, w: cardW, h: cardH, rectRadius: 0.06,
    fill: { color: "1B1F3B" }, line: { color: TEAL, width: 1.25 },
  });
  slide.addText("VS Code (IDE)", {
    x: leftX, y: y + 0.08, w: cardW, h: 0.3, align: "center",
    fontFace: TITLE_FONT, bold: true, fontSize: 12, color: WHITE, margin: 0,
  });
  const panels = [
    { label: "עורך קוד צבעוני", frac: 0.55 },
    { label: "טרמינל מובנה", frac: 0.25 },
    { label: "איתור באגים", frac: 0.2 },
  ];
  let py = y + 0.48;
  const totalH = h - 0.6;
  panels.forEach((p) => {
    const ph = totalH * p.frac - 0.05;
    slide.addShape(pres.ShapeType.rect, {
      x: leftX + 0.25, y: py, w: cardW - 0.5, h: ph,
      fill: { color: "2A2F55" }, line: { color: TEAL, width: 0.5 },
    });
    slide.addText(p.label, {
      x: leftX + 0.25, y: py, w: cardW - 0.5, h: ph, margin: 2,
      align: "center", valign: "middle",
      fontFace: BODY_FONT, fontSize: 9, color: "CADCFC", rtlMode: true,
    });
    py += ph + 0.06;
  });
}

// ---------- דיאגרמת "כאוס גרסאות" מול Git מסודר: שורה עליונה = שרשרת
// 4 "תגיות קובץ" מתרחבות (v1→v2→v2_final→v2_final_REALLY) בצבע מוחלש;
// שורה תחתונה = שרשרת 4 נקודות Commit מסודרות בטיל, עם תאריך קטן מתחת.
// איור וקטורי מקורי.
function addVersionChaosDiagram(slide, x, y, w, h) {
  const rowH = Math.min(0.42, h * 0.4);
  const gap = 0.18;
  const n = 4;
  const cellW = (w - gap * (n - 1)) / n;
  const row1Y = y;
  const row2Y = y + h - rowH;

  const messyLabels = ["עבודה.docx", "עבודה_v2.docx", "v2_final.docx", "v2_final_REALLY"];
  messyLabels.forEach((label, i) => {
    const bx = x + w - cellW - i * (cellW + gap);
    slide.addShape(pres.ShapeType.rect, {
      x: bx, y: row1Y, w: cellW, h: rowH,
      fill: { color: PH_BG }, line: { color: MUTED, width: 1, dashType: "dash" },
    });
    slide.addText(label, {
      x: bx, y: row1Y, w: cellW, h: rowH, margin: 2,
      align: "center", valign: "middle",
      fontFace: "Courier New", fontSize: 8, color: MUTED,
    });
    if (i < n - 1) {
      const leftEdge = bx - gap + 0.02;
      slide.addShape(pres.ShapeType.line, {
        x: leftEdge, y: row1Y + rowH / 2, w: gap - 0.04, h: 0,
        line: { color: MUTED, width: 1, beginArrowType: "triangle" },
      });
    }
  });

  const commitD = Math.min(0.3, rowH * 0.8);
  for (let i = 0; i < n; i++) {
    const cx = x + w - cellW / 2 - i * (cellW + gap);
    slide.addShape(pres.ShapeType.ellipse, {
      x: cx - commitD / 2, y: row2Y + (rowH - commitD) / 2, w: commitD, h: commitD,
      fill: { color: TEAL }, line: { type: "none" },
    });
    if (i < n - 1) {
      const leftEdge = cx - (cellW + gap) + commitD / 2 + 0.03;
      const rightEdge = cx - commitD / 2 - 0.03;
      slide.addShape(pres.ShapeType.line, {
        x: leftEdge, y: row2Y + rowH / 2, w: rightEdge - leftEdge, h: 0,
        line: { color: TEAL, width: 1.25, beginArrowType: "triangle" },
      });
    }
  }
  slide.addText("Commit נקי, מסודר וניתן-לחזרה", {
    x: x, y: row2Y + rowH + 0.02, w, h: 0.22, align: "center",
    fontFace: BODY_FONT, italic: true, fontSize: 9, color: TEAL, rtlMode: true, margin: 0,
  });
}

// ---------- דיאגרמת מחשב-מקומי מול ענן: כרטיסייה ימנית (RTL) = "המחשב
// שלכם (Git)", כרטיסייה שמאלית = "GitHub (בענן)", עם חץ דו-כיווני
// (Push/Pull) ביניהן. איור וקטורי מקורי.
function addLocalCloudDiagram(slide, x, y, w, h) {
  const cardW = (w - 1.1) / 2;
  const cardH = Math.min(0.9, h * 0.8);
  const rightX = x + w - cardW;
  const leftX = x;
  const cardY = y + (h - cardH) / 2;

  slide.addShape(pres.ShapeType.roundRect, {
    x: rightX, y: cardY, w: cardW, h: cardH, rectRadius: 0.06,
    fill: { color: PH_BG }, line: { color: PH_BORDER, width: 1.25 },
  });
  slide.addText([
    { text: "המחשב שלכם", options: { bold: true, fontSize: 13, color: TEXT_DARK, breakLine: true, rtlMode: true } },
    { text: "Git — היסטוריה מקומית", options: { fontSize: 10, color: MUTED, rtlMode: true } },
  ], {
    x: rightX, y: cardY, w: cardW, h: cardH, margin: 6,
    align: "center", valign: "middle",
  });

  slide.addShape(pres.ShapeType.roundRect, {
    x: leftX, y: cardY, w: cardW, h: cardH, rectRadius: 0.06,
    fill: { color: TEAL }, line: { color: TEAL, width: 1.25 },
  });
  slide.addText([
    { text: "GitHub", options: { bold: true, fontSize: 13, color: WHITE, breakLine: true, rtlMode: true } },
    { text: "בענן — מגובה ומשותף", options: { fontSize: 10, color: "E6FFFA", rtlMode: true } },
  ], {
    x: leftX, y: cardY, w: cardW, h: cardH, margin: 6,
    align: "center", valign: "middle",
  });

  const arrowX1 = rightX - 0.1;
  const arrowX2 = leftX + cardW + 0.1;
  const arrowYTop = cardY + cardH * 0.35;
  const arrowYBottom = cardY + cardH * 0.65;
  slide.addShape(pres.ShapeType.line, {
    x: arrowX2, y: arrowYTop, w: arrowX1 - arrowX2, h: 0,
    line: { color: PH_BORDER, width: 1.25, beginArrowType: "triangle" },
  });
  slide.addShape(pres.ShapeType.line, {
    x: arrowX2, y: arrowYBottom, w: arrowX1 - arrowX2, h: 0,
    line: { color: PH_BORDER, width: 1.25, endArrowType: "triangle" },
  });
  slide.addText("Push / Pull", {
    x: leftX + cardW + 0.05, y: cardY + cardH / 2 - 0.13, w: 1.0, h: 0.26,
    align: "center", fontFace: BODY_FONT, italic: true, fontSize: 9, color: MUTED, margin: 0,
  });
}

// ---------- דיאגרמת זרימת כלי AI: 3 תיבות "בקשה בשפה טבעית → כלי AI →
// קוד בקובץ" (זהה במבנה ל-addLLMFlowDiagram משיעור 1). איור וקטורי מקורי.
function addAIToolFlowDiagram(slide, x, y, w, h) {
  const stages = ["בקשה בשפה טבעית", "כלי AI", "קוד בקובץ שלכם"];
  const rowH = Math.min(0.55, h * 0.7);
  const gap = 0.25;
  const boxW = (w - gap * (stages.length - 1)) / stages.length;
  const rowY = y + (h - rowH) / 2;

  stages.forEach((stg, i) => {
    const bx = x + w - boxW - i * (boxW + gap);
    const isModel = i === 1;
    slide.addShape(pres.ShapeType.roundRect, {
      x: bx, y: rowY, w: boxW, h: rowH, rectRadius: 0.06,
      fill: { color: isModel ? TEAL : PH_BG },
      line: { color: isModel ? TEAL : PH_BORDER, width: 1.25 },
    });
    slide.addText(stg, {
      x: bx, y: rowY, w: boxW, h: rowH, margin: 4,
      align: "center", valign: "middle",
      fontFace: BODY_FONT, bold: isModel, fontSize: 12,
      color: isModel ? WHITE : TEXT_DARK, rtlMode: true,
    });
    if (i < stages.length - 1) {
      const arrowRight = bx - gap + 0.02;
      const arrowLeft = bx - 0.02;
      slide.addShape(pres.ShapeType.line, {
        x: arrowRight, y: rowY + rowH / 2, w: arrowLeft - arrowRight, h: 0,
        line: { color: PH_BORDER, width: 1.5, beginArrowType: "triangle" },
      });
    }
  });
}

// ---------- דיאגרמת Autocomplete מול Agentic: שתי שורות של אותם 4
// שלבי-פעולה (קורא קבצים / כותב קוד / מריץ בדיקות / מתקן את עצמו), עם
// תג "✓" מעל כל שלב שבו הסוג הזה של כלי בפועל פועל. זהה במבנה ל-
// addAINativeDiagram משיעור 1. איור וקטורי מקורי.
function addAgenticComparisonDiagram(slide, x, y, w, h) {
  const stages = ["קורא קבצים", "כותב קוד", "מריץ בדיקות", "מתקן את עצמו"];
  const labelW = 1.65;
  const boxAreaW = w - labelW - 0.15;
  const gap = 0.12;
  const boxW = (boxAreaW - gap * (stages.length - 1)) / stages.length;
  const rowH = Math.min(0.42, h * 0.32);
  const rowGap = h - rowH * 2 > 0.3 ? h - rowH * 2 - 0.1 : 0.3;
  const row1Y = y;
  const row2Y = y + rowH + rowGap;

  function boxX(i) {
    return x + boxAreaW - boxW - i * (boxW + gap);
  }

  function drawRow(rowY, rowLabel, okIndexes, okColor) {
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
      if (okIndexes.includes(i)) {
        const badgeD = 0.24;
        slide.addShape(pres.ShapeType.ellipse, {
          x: bx + boxW / 2 - badgeD / 2, y: rowY - badgeD * 0.65, w: badgeD, h: badgeD,
          fill: { color: okColor }, line: { color: WHITE, width: 1 },
        });
        slide.addText("✓", {
          x: bx + boxW / 2 - badgeD / 2, y: rowY - badgeD * 0.65, w: badgeD, h: badgeD, margin: 0,
          align: "center", valign: "middle",
          fontFace: TITLE_FONT, bold: true, fontSize: 9, color: WHITE,
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

  drawRow(row1Y, "Autocomplete", [1], MUTED);
  drawRow(row2Y, "Agentic", [0, 1, 2, 3], TEAL);
}

// ---------- עוזר משותף: שורת מעגלים ממוספרים עם חצים (זהה ל-
// addWorkflowLinearDiagram משיעור 1) - משמש לשקף Claude Code 5 השלבים.
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

// ---------- דיאגרמת "צמיחת אופק זמן" (METR): 3 פסים אנכיים בגובה עולה
// (לא לפי קנה-מידה מדויק - המרווח בין השניות לשעות עצום מכדי לצייר
// ליניארית - אלא המחשה ויזואלית של "כל פעם גדול משמעותית יותר"), כל אחד
// עם תווית מודל+זמן מתחתיו. איור וקטורי מקורי.
function addGrowthBarsDiagram(slide, x, y, w, h) {
  const bars = [
    { label: "Claude 3.7 Sonnet\n~50 דקות", frac: 0.28, color: PH_BORDER },
    { label: "מודלים מובילים\nאמצע 2026: ~12 שעות", frac: 1.0, color: TEAL },
  ];
  const n = bars.length;
  const gap = 0.6;
  const barW = 0.9;
  const totalW = n * barW + (n - 1) * gap;
  const startX = x + (w - totalW) / 2;
  const maxBarH = h - 0.5;

  bars.forEach((b, i) => {
    const bx = startX + i * (barW + gap);
    const barH = Math.max(0.3, maxBarH * b.frac);
    const barY = y + (maxBarH - barH);
    slide.addShape(pres.ShapeType.rect, {
      x: bx, y: barY, w: barW, h: barH,
      fill: { color: b.color }, line: { type: "none" },
    });
    slide.addText(b.label, {
      x: bx - 0.35, y: y + maxBarH + 0.05, w: barW + 0.7, h: 0.45,
      align: "center", valign: "top", margin: 0,
      fontFace: BODY_FONT, fontSize: 9, color: TEXT_DARK, rtlMode: true,
    });
  });

  // חץ עולה קטן משמאל לימין (RTL: מהמוקדם/ימין למאוחר/שמאל) שממחיש מגמה
  slide.addShape(pres.ShapeType.line, {
    x: startX + totalW + 0.15, y: y + maxBarH * 0.15, w: 0, h: maxBarH * 0.7,
    line: { color: MUTED, width: 1, dashType: "dash", beginArrowType: "triangle" },
    rotate: 0,
  });
}

function addFooter(slide) {
  SLIDE_NUM += 1;
  slide.addText(`Vibe Coding · ${LESSON_LABEL} · שקף ${SLIDE_NUM}`, {
    x: 0.6, y: 7.1, w: 12.13, h: 0.3,
    align: "right", fontFace: BODY_FONT, fontSize: 9, color: "A6A6B0", rtlMode: true, margin: 0,
  });
}

// שקף תוכן רגיל: כותרת → שאלה → בולטים → (איור) → מה עושים בפועל → (מקור) → פוטר
// diagram: פונקציה (slide, x, y, w, h) => void שמציירת איור וקטורי מקורי;
// diagramCaption: הטקסט שמופיע תחתיו כ"איור: ..." (לא "מקור:", כי זה לא ציטוט חיצוני)
function contentSlide({ title, question, bullets, bestPractice, source, diagram, diagramCaption }) {
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addTitle(s, title);
  if (question) addQuestion(s, question);

  const bulletsY = question ? 1.65 : 1.35;
  const maxBulletsH = bestPractice ? 4.3 : 5.2;

  // הערה קריטית: תיבת הבולטים ב-pptxgenjs ממורכזת אנכית (anchor="ctr")
  // בברירת המחדל. אם מקצים לה פחות גובה משהתוכן בפועל צריך (כדי לפנות
  // מקום לאיור), הטקסט לא נחתך - הוא "גולש" בשני הכיוונים סביב מרכז
  // התיבה, ולפעמים בולע ממש לתוך השאלה שמעליה! לכן חובה קודם להעריך כמה
  // גובה הבולטים דורשים בפועל (לפי אורך תו), ולא לצמצם את bulletsH
  // מתחת לזה - גם אם זה אומר לצמצם את האיור, או לוותר עליו, במקום.
  const bulletCharsPerLine = 100;
  const bulletLineH = 0.335; // גובה שורה בפועל בגופן 19, כולל paraSpaceAfter יחסי
  let neededBulletsH = 0;
  bullets.forEach((b) => {
    neededBulletsH += estimateLines(b, bulletCharsPerLine) * bulletLineH + 0.1;
  });
  neededBulletsH = Math.min(neededBulletsH, maxBulletsH);

  // מגבלת הגובה הפנוי לפני הפוטר (7.1): מעריכים מראש כמה מקום bestPractice
  // ו-source ידרשו בפועל, ומחשבים כמה נשאר לאיור (אם יש) - ולא להיפך.
  let reservedFixedBottom = 0.15;
  let bpLines = 1;
  if (bestPractice) {
    bpLines = estimateLines("✔ מה עושים בפועל: " + bestPractice, 108);
    reservedFixedBottom += 0.1 + bpLines * 0.34 + 0.12;
  }
  if (source) reservedFixedBottom += 0.32;

  // ללא איור: כמו קודם, ממלאים את כל הגובה הפנוי (עד maxBulletsH).
  // עם איור: הבולטים מקבלים רק את מה שהם צריכים בפועל (neededBulletsH),
  // לא יותר - כדי שיישאר מקום אמיתי לאיור, בלי לגרום לגלישה.
  const bulletsH = diagram
    ? neededBulletsH
    : Math.min(maxBulletsH, 7.1 - bulletsY - reservedFixedBottom);
  addBullets(s, bullets, bulletsY, bulletsH);

  // מה שנשאר עבור האיור: הגובה האידיאלי (1.65) אם יש מקום, אחרת מצטמצם
  // עד למינימום סביר (1.0) - ואם גם זה לא מספיק, מוותרים על האיור בשקף
  // הזה במקום לגרום לחפיפה (עדיף שקף בלי איור מאשר שקף שבור).
  const availableForDiagram = 7.1 - (bulletsY + bulletsH + 0.15) - reservedFixedBottom;
  const diagramCaptionH = diagram && diagramCaption ? 0.28 : 0;
  let diagramH = 0;
  let showDiagram = false;
  if (diagram) {
    if (availableForDiagram >= 1.0 + diagramCaptionH + 0.1) {
      diagramH = Math.min(1.65, availableForDiagram - diagramCaptionH - 0.1);
      showDiagram = true;
    } else {
      console.warn(`[דילוג על איור - אין מקום] שקף: ${title}`);
    }
  }

  let nextY = bulletsY + bulletsH + 0.15;
  if (showDiagram) {
    diagram(s, 0.6, nextY, 12.13, diagramH);
    nextY += diagramH + 0.1;
    if (diagramCaption) {
      addDiagramCaption(s, diagramCaption, nextY);
      nextY += 0.28;
    }
  }
  if (bestPractice) {
    nextY = addBestPractice(s, bestPractice, nextY) + 0.12;
  }
  if (source) addSourceCaption(s, source, Math.min(nextY, 6.85));
  addFooter(s);
  return s;
}

// שקף פרומפט: eyebrow/title/sub במרכז, ואז תיבה אחת או יותר עם טקסט
// מוטמע במלואו (בפונט Courier New) - לפרומפטים מדויקים לתרגול/הדגמה.
function promptSlide({ eyebrow, title, sub, boxes }) {
  const s = pres.addSlide();
  s.background = { color: WHITE };

  s.addText(eyebrow, {
    x: 0.6, y: 0.55, w: 12.13, h: 0.4, align: "center",
    fontFace: BODY_FONT, italic: true, fontSize: 15, color: TEAL, rtlMode: true, margin: 0,
  });
  s.addShape(pres.ShapeType.line, { x: 5.16, y: 1.0, w: 3.0, h: 0, line: { color: TEAL, width: 1.5 } });
  s.addText(title, {
    x: 0.6, y: 1.15, w: 12.13, h: 0.75, align: "center",
    fontFace: TITLE_FONT, bold: true, fontSize: 26, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  if (sub) {
    s.addText(sub, {
      x: 0.6, y: 1.9, w: 12.13, h: 0.75, align: "center",
      fontFace: BODY_FONT, italic: true, fontSize: 14, color: MUTED, rtlMode: true, margin: 0,
    });
  }

  const startY = 2.85;
  const availH = 6.6 - startY;
  const gap = 0.25;
  const boxH = (availH - gap * (boxes.length - 1)) / boxes.length;

  boxes.forEach((box, i) => {
    const y = startY + i * (boxH + gap);
    if (box.label) {
      s.addText(box.label, {
        x: 0.6, y: y - 0.02, w: 12.13, h: 0.3, align: "right",
        fontFace: TITLE_FONT, bold: true, fontSize: 13, color: TEAL, rtlMode: true, margin: 0,
      });
    }
    const boxY = box.label ? y + 0.3 : y;
    const boxHAdj = box.label ? boxH - 0.3 : boxH;
    s.addShape(pres.ShapeType.roundRect, {
      x: 0.6, y: boxY, w: 12.13, h: boxHAdj, rectRadius: 0.06,
      fill: { color: PH_BG }, line: { color: PH_BORDER, width: 1 },
    });
    s.addText(paragraphRuns(box.lines, { fontFace: "Courier New", fontSize: 13, color: TEXT_DARK }), {
      x: 0.85, y: boxY + 0.15, w: 11.63, h: boxHAdj - 0.3,
      align: "right", rtlMode: true, margin: 0,
    });
  });

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
  s.addText("שיעור 2: AI Development Environment", {
    x: 0.6, y: 1.05, w: 12.13, h: 1.1, align: "center",
    fontFace: TITLE_FONT, bold: true, fontSize: 36, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  s.addText("Introduction to AI Development Environment", {
    x: 0.6, y: 2.1, w: 12.13, h: 0.5, align: "center",
    fontFace: BODY_FONT, italic: true, fontSize: 16, color: MUTED, rtlMode: true, margin: 0,
  });

  s.addText("בסוף השיעור תדעו לענות על:", {
    x: 2.5, y: 3.1, w: 8.33, h: 0.4, align: "right",
    fontFace: TITLE_FONT, bold: true, fontSize: 16, color: TEXT_DARK, rtlMode: true, margin: 0,
  });
  s.addText(
    bulletRuns(
      [
        "מה זה IDE, Git, ו-GitHub — ולמה כל אחד מהם קריטי לעבודה עם AI?",
        "מה ההבדל בין כלי Autocomplete לכלי Agentic?",
        "איך נראה תהליך עבודה טיפוסי מול כלי AI אג'נטי כמו Claude Code?",
        "אילו כלים חינמיים זמינים כדי להתחיל בלי תקציב?",
      ],
      { fontFace: BODY_FONT, fontSize: 17, color: TEXT_DARK }
    ),
    { x: 2.5, y: 3.6, w: 8.33, h: 2.4, align: "right", rtlMode: true, paraSpaceAfter: 8, margin: 0 }
  );

  s.addText("Vibe Coding Course · Lesson 2 of 13", {
    x: 0.6, y: 7.1, w: 12.13, h: 0.3, align: "center",
    fontFace: BODY_FONT, fontSize: 10, color: MUTED, rtlMode: true, margin: 0,
  });
  SLIDE_NUM += 1;
}

// =====================================================================
// שקף 2 — מה זה בכלל IDE?
// =====================================================================
contentSlide({
  title: "מה זה בכלל IDE?",
  question: "השאלה: איך בכלל כותבים קוד, טכנית?",
  bullets: [
    "אפשר טכנית לפתוח קובץ טקסט רגיל (כמו Notepad) ולהקליד בו — זה עובד, אבל בלי שום עזרה: אין צביעת תחביר (סימון צבעוני של מילות קוד), אין השלמה אוטומטית, ואין דרך להריץ קוד בלי לעבור לתוכנה אחרת לגמרי",
    "**IDE** (ראשי תיבות: Integrated Development Environment — סביבת פיתוח משולבת) = תוכנה שמאחדת במקום אחד עורך טקסט חכם, טרמינל מובנה (שורת פקודה בתוך התוכנה עצמה), וכלים לאיתור באגים",
    "**VS Code** (ראשי תיבות של Visual Studio Code) = ה-IDE החינמי והפופולרי ביותר כיום",
    "רוב כלי ה-AI האג'נטיים (כולל Claude Code) מתחברים ל-VS Code כתוסף (Extension), או פועלים לצידו מתוך הטרמינל המובנה שלו",
  ],
  bestPractice: "כל התרגול היום (ובכל שיעור מכאן והלאה) יתבצע בתוך VS Code — זה \"השולחן\" שעליו כל שאר הכלים יושבים.",
  diagram: addIDEComparisonDiagram,
  diagramCaption: "Notepad — טקסט פשוט בלי עזרה, מול VS Code — עורך צבעוני + טרמינל מובנה + איתור באגים, הכל בחלון אחד.",
});

// =====================================================================
// שקף 3 — מה הבעיה ש-Version Control פותר?
// =====================================================================
contentSlide({
  title: "מה הבעיה ש-Version Control פותר?",
  question: "השאלה: איך שומרים גרסאות של קובץ בלי להתבלבל?",
  bullets: [
    "דמיינו שאתם כותבים עבודה ושומרים: `עבודה.docx`, `עבודה_v2.docx`, `עבודה_v2_final.docx`, `עבודה_v2_final_REALLY.docx` — זה בדיוק מה שקורה עם קוד בלי כלי מתאים, ומחמיר עוד יותר כשכמה אנשים עובדים על אותו פרויקט",
    "**Version Control** (ניהול גרסאות) = מערכת ששומרת \"תמונת מצב\" (Snapshot — עותק מדויק של כל הפרויקט ברגע נתון) בכל נקודת זמן שתבחרו, עם אפשרות לחזור אליה, להשוות בין גרסאות, ולעבוד יחד עם אחרים בלי לדרוך אחד על השני",
  ],
  bestPractice: "Version Control הוא לא \"נחמד שיהיה\" — הוא הבסיס שעליו כל שאר השיעור בנוי, כי בלעדיו אי אפשר לעבוד בבטחון עם AI (בשקף הבא נראה למה).",
  diagram: addVersionChaosDiagram,
  diagramCaption: "שרשרת שמות קבצים מבולבלת (למעלה, מוחלש) מול שרשרת Commit-ים מסודרת עם מסלול חזרה (למטה, בטיל).",
});

// =====================================================================
// שקף 4 — Git
// =====================================================================
contentSlide({
  title: "Git — הכלי שעושה את זה בפועל",
  question: "השאלה: איך בפועל מיישמים Version Control?",
  bullets: [
    "**Git** = כלי ניהול הגרסאות הנפוץ ביותר, רץ על המחשב שלכם בעצמו (בחינם, קוד פתוח — כל אחד יכול לראות ולשנות את הקוד שלו)",
    "**Repository** (בקיצור **Repo**) = תיקיית הפרויקט שעליה Git \"שומר עין\" — כל קובץ בתוכה עוקב אחר ההיסטוריה שלו",
    "**Commit** = \"תמונת מצב\" שמורה אחת, עם הודעה שמסבירה מה השתנה (למשל: \"הוספתי כפתור התחברות\") — בדיוק כמו נקודת שמירה במשחק מחשב, תמיד אפשר לחזור אליה",
    "**למסלול AI:** זו רשת הביטחון שלכם כש-AI כותב קוד — אם הוא הורס משהו, Git מאפשר לחזור לתמונת המצב הקודמת תוך שניות",
  ],
  bestPractice: "לפני שנותנים ל-AI לגעת בקוד, מוודאים שיש Commit נקי מאחור. זה לא בירוקרטיה — זה הביטוח שהופך ניסוי-וטעייה עם AI לבטוח.",
});

// =====================================================================
// שקף 5 — GitHub
// =====================================================================
contentSlide({
  title: "GitHub — לא אותו דבר כמו Git",
  question: "השאלה: אם Git כבר שומר הכל אצלי במחשב, למה צריך עוד שירות?",
  bullets: [
    "בלבול נפוץ אצל מתחילים: **Git ≠ GitHub**",
    "Git = הכלי שרץ אצלכם במחשב, שומר היסטוריה מקומית בלבד",
    "**GitHub** = שירות אחסון בענן (של Microsoft) ששומר עותק מגובה של ה-Repository באינטרנט, מאפשר לשתף אותו עם אחרים, ולשתף פעולה על אותו קוד (נרחיב בשיעור 10)",
    "אנלוגיה: Git הוא כמו קובץ Word בודד במחשב שלכם; GitHub הוא כמו Google Drive — המקום שבו הוא מגובה ומשותף",
    "**פעילות:** מי כבר השתמש ב-Git, ולו רק כדי לשכפל Repository של מישהו אחר? איך שיתפתם קוד לפני שהכרתם GitHub (וואטסאפ? מייל?)",
  ],
  bestPractice: "בתרגול היום תיצרו Repository משלכם ב-GitHub — הבית הקבוע של הפרויקט לאורך הסמסטר.",
  diagram: addLocalCloudDiagram,
  diagramCaption: "Git רץ על המחשב שלכם (היסטוריה מקומית); GitHub הוא העותק בענן, מגובה ומשותף — Push מעלה, Pull מוריד.",
});

// =====================================================================
// שקף 6 — GitHub בקנה מידה
// =====================================================================
contentSlide({
  title: "GitHub בקנה מידה",
  bullets: [
    "**180M+** מפתחים רשומים (36M+ הצטרפו בשנה האחרונה בלבד — קצב הצמיחה המהיר ביותר בהיסטוריית הפלטפורמה)",
    "**630M+** Repositories",
    "**43.2M** בקשות שילוב קוד (Pull Requests) נמזגות בכל חודש",
    "**92%** מחברות **Fortune 100** (מדד 100 החברות הגדולות בעולם לפי הכנסות) משתמשות ב-GitHub Enterprise (הגרסה הארגונית בתשלום)",
  ],
  bestPractice: "זו לא \"עוד פלטפורמה\" — זו התשתית שרוב תעשיית התוכנה בנויה עליה.",
  source: "Kinsta & Skillademia — GitHub Statistics 2026",
});

// =====================================================================
// שקף 7 — מה זה בכלל "כלי AI לכתיבת קוד"?
// =====================================================================
contentSlide({
  title: "מה זה בכלל \"כלי AI לכתיבת קוד\"?",
  question: "השאלה: מה בדיוק קורה מאחורי הקלעים כשכלי AI \"כותב קוד\"?",
  bullets: [
    "כלי AI לכתיבת קוד = תוכנה שמשלבת מודל שפה (**LLM**, כפי שלמדנו בשיעור 1) בתוך סביבת הפיתוח שלכם",
    "כותבים בקשה בשפה טבעית (\"תוסיף כפתור שמוחק משימה\") — הכלי מתרגם אותה לקוד בפועל בקבצים שלכם",
    "יש כמה \"רמות\" של כלים כאלה, מהבסיסית ביותר ועד המתקדמת — וזה בדיוק ההבדל בין Autocomplete ל-Agentic שנפרט בשקף הבא",
  ],
  bestPractice: "בשיעור הקודם ראיתם הדגמה של זה בפעולה — עכשיו נגדיר את זה במדויק.",
  diagram: addAIToolFlowDiagram,
  diagramCaption: "בקשה בשפה טבעית → כלי AI → קוד שנכתב בפועל בקובץ שלכם.",
});

// =====================================================================
// שקף 8 — מה הופך כלי ל"אג'נטי"?
// =====================================================================
contentSlide({
  title: "מה הופך כלי ל\"אג'נטי\"?",
  question: "השאלה: מה ההבדל בין \"עוזר הקלדה\" לבין \"שותף לביצוע משימה\"?",
  bullets: [
    "**Autocomplete בסיסי** (הדורות הראשונים של כלי AI לקוד) — משלים שורה או פונקציה בודדת לפי הקשר מקומי; אתם עדיין מקבלים כל החלטה",
    "**כלי Agentic** (כמו Claude Code) — מקבל משימה ברמה גבוהה (\"תוסיף התחברות עם Google לאפליקציה\"), ומבצע בעצמו סדרת פעולות: קורא קבצים רלוונטיים, כותב קוד במספר מקומות, מריץ בדיקות, ולעיתים גם מתקן את עצמו אם משהו נכשל",
    "**למסלול AI:** דומה למעבר מ-Feature Engineering ידני ל-Deep Learning — גם שם עוברים מ\"מגדיר כל שלב\" ל\"מגדיר מטרה, המערכת מבצעת\", אבל אתם עדיין בודקים את התוצאה (כמו Evaluation ב-ML)",
  ],
  bestPractice: "הקורס הזה עוסק בסוג השני (Agentic) — \"עוזר הקלדה\" זה לא המטרה כאן.",
  diagram: addAgenticComparisonDiagram,
  diagramCaption: "אותם 4 שלבי-פעולה: Autocomplete נוגע רק בכתיבת קוד; Agentic פועל בכל השלבים - קורא, כותב, בודק, ומתקן את עצמו.",
});

// =====================================================================
// שקף 8.5 — פעילות בזוגות: Autocomplete או Agentic
// =====================================================================
contentSlide({
  title: "פעילות בזוגות: Autocomplete או Agentic?",
  bullets: [
    "**המשימה (2-3 דקות בזוגות):** לכל אחת מהמשימות הבאות, קבעו יחד — מספיק Autocomplete בסיסי, או צריך כלי Agentic מלא? נמקו.",
    "**1.** להשלים שם משתנה שכבר התחלתם להקליד",
    "**2.** להוסיף מסך \"התחברות\" שלם לאפליקציה, כולל שמירה במסד נתונים",
    "**3.** לתקן שגיאת הזחה (Indentation) בשורה בודדת",
    "**4.** לעדכן פונקציה אחת בעשרה קבצים שונים בפרויקט, בעקביות",
  ],
  bestPractice: "התשובה: 1+3 = Autocomplete מספיק (שינוי נקודתי); 2+4 = דורש Agentic (משתרע על כמה קבצים, דורש החלטות). תרחיש 4 הכי פחות חד-משמעי — זה בדיוק העניין.",
});

// =====================================================================
// שקף 9 — איך מודדים "כמה טוב" כלי אג'נטי
// =====================================================================
contentSlide({
  title: "איך מודדים \"כמה טוב\" כלי אג'נטי?",
  question: "השאלה: איך משווים בין כלי AI שונים בצורה אובייקטיבית?",
  bullets: [
    "הדרך המקצועית היא בנצ'מארקים (מבחני השוואה סטנדרטיים): **SWE-bench** (בעיות תכנות אמיתיות שנאספו מ-GitHub) ו-**Terminal-Bench** (משימות טרמינל/שורת-פקודה מורכבות)",
    "**אזהרה מקצועית:** יש היום חמש גרסאות שונות של SWE-bench (Original, Verified, Pro, Multilingual, Live), וההשוואה ביניהן לא תמיד הוגנת",
    "אפילו על אותם משקלי מודל בדיוק, יש שונות (Variance) של 10-20 נקודות אחוז, בהתאם ל-**Harness** (התשתית הטכנית שמריצה את הבדיקה בפועל)",
  ],
  bestPractice: "תמיד בודקים את המתודולוגיה מאחורי מספר בנצ'מארק לפני שסומכים עליו — בדיוק כמו שבודקים מתודולוגיה לפני שסומכים על הערכת מודל בכל הקשר אחר שאתם מכירים.",
});

// =====================================================================
// שקף 10 — המרוץ לבנצ'מארק: המספרים
// =====================================================================
contentSlide({
  title: "המרוץ לבנצ'מארק: המספרים",
  bullets: [
    "**80.8%** — התוצאה של Claude Code ב-SWE-bench Verified (הגרסה ה\"מנוקה\" של הבנצ'מארק)",
    "**78.9%-83.1%** — טווח התוצאות של Claude Code ב-Terminal-Bench 2.1, תלוי בגרסת המודל",
    "כפי שראינו בשקף הקודם — המספרים האלה חשובים, אבל **רק** ביחד עם אזהרת המתודולוגיה, לא בלעדיה",
  ],
  bestPractice: "מספר בנצ'מארק בודד הוא נקודת התחלה לבדיקה, לא מסקנה סופית.",
  source: "DigitalApplied, MorphLLM — 2026",
});

// =====================================================================
// שקף 10.5 — לוחות תוצאות חיים
// =====================================================================
contentSlide({
  title: "לוחות תוצאות חיים — לפתיחה בכיתה",
  question: "השאלה: איך רואים את זה לא כטקסט על שקף, אלא כמשהו אמיתי ומתעדכן?",
  bullets: [
    "לוח SWE-bench Verified (התוצאות המדויקות שראינו בשקף הקודם): **swebench.com/verified.html**",
    "לוח Terminal-Bench 2.1: **tbench.ai/leaderboard/terminal-bench/2.1**",
    "שני הלוחות מתעדכנים באופן שוטף — המספר שראיתם בשקף הקודם עשוי כבר להיות לא הכי עדכני ברגע שאתם קוראים את זה",
  ],
  bestPractice: "המרצה פותח את שני הלינקים בדפדפן מול הכיתה. **פעילות (הרימו יד):** לפני שפותחים — מי בטוח שהמודל שהיה מוביל בזמן הכנת השקף עדיין מוביל היום? זו ההדגמה הכי טובה לאזהרת המתודולוגיה מהשקף הקודם: המספרים משתנים ממש מתחת לרגליים.",
});

// =====================================================================
// שקף 10.6 — בנצ'מארק שני: METR Time Horizon
// =====================================================================
contentSlide({
  title: "בנצ'מארק שני: לא רק \"עובד/לא עובד\"",
  question: "השאלה: SWE-bench בודק אם הפתרון נכון — אבל איך בודקים כמה זמן כלי יכול להתמיד במשימה בלי עזרה?",
  bullets: [
    "**METR** (ארגון מחקר שמפרסם מדדי יכולת עצמאית של סוכני AI) מודד **\"אופק זמן\" (Time Horizon)**: משך המשימה (לפי זמן שלוקח למומחה אנושי) שבו למודל יש 50% סיכוי להצליח, ללא התערבות",
    "זה שונה מ-SWE-bench: שם בודקים \"האם הפתרון הספציפי הזה נכון\", כאן בודקים \"כמה זמן הכלי יכול להתמיד לבד לפני שהוא נתקע או טועה\"",
    "**הנתון:** אופק הזמן הכפיל את עצמו בערך כל 4 חודשים בשנים האחרונות; במודלים המובילים באמצע 2026 האופק הגיע לכ-12 שעות (לשם השוואה: Claude 3.7 Sonnet, מ-2025, עמד על כ-50 דקות בלבד)",
  ],
  bestPractice: "שני סוגי הבנצ'מארק משלימים זה את זה — אחד אומר \"האם זה נכון\", השני אומר \"כמה הוא יכול להתמיד לבד\". כשמעריכים כלי AI (בקורס הזה או בעתיד המקצועי שלכם), שווה לשאול את שתי השאלות, לא רק אחת.",
  diagram: addGrowthBarsDiagram,
  diagramCaption: "המחשה (לא לפי קנה-מידה): אופק הזמן גדל משמעותית בין דורות מודלים - מדקות בודדות לשעות ארוכות.",
  source: "metr.org/time-horizons",
});

// =====================================================================
// שקף 11 — Claude Code: איך זה עובד בפועל
// =====================================================================
contentSlide({
  title: "Claude Code — איך זה עובד בפועל",
  question: "השאלה: אז איך בפועל נראית עבודה מול כלי אג'נטי, צעד-צעד?",
  bullets: [
    "**1. פתיחת Session** (חיבור פעיל אחד מול כלי ה-AI, מתחילתו ועד סופו) — מהטרמינל, או כתוסף בתוך VS Code",
    "**2. הנחיה ברורה וספציפית** — אומרים בדיוק מה רוצים",
    "**3. הכלי מציג תוכנית או מתחיל לפעול** — קורא קבצים, כותב קוד",
    "**4. בודקים כל Diff** (תצוגה שמראה בדיוק אילו שורות קוד השתנו, נוספו, או נמחקו) **לפני שמאשרים**",
    "**5. ממשיכים לצעד הבא**, או חוזרים לשלב 2 אם צריך תיקון",
  ],
  bestPractice: "זה בדיוק החיבור לשלבים 4-6 מה-Workflow שלמדנו בשיעור 1 (תכנון עם AI → מימוש עם AI → Code Review) — היום אתם עושים את זה בפועל, לא רק רואים הדגמה.",
  diagram: (s, x, y, w, h) => addWorkflowLinearDiagram(s, x, y, w, h, [1, 2, 3, 4, 5]),
  diagramCaption: "5 השלבים בזרימה מ-1 עד 5 (כזכור מהבולט למעלה - אם שלב 4 דורש תיקון, חוזרים לשלב 2).",
});

// =====================================================================
// שקף 12 — אתם כבר עובדים ליד AI
// =====================================================================
contentSlide({
  title: "אתם כבר עובדים ליד AI",
  question: "השאלה: זה נשמע מסובך — האם אתם באמת מתחילים מאפס?",
  bullets: [
    "**פעילות:** מי מכם כבר השתמש ב-ChatGPT או Claude כדי לכתוב סקריפט לניקוי דאטה, או פונקציית עזר למודל? מה עשיתם עם הקוד שקיבלתם בסוף?",
    "ברוב הכיתות: התשובה היא העתק-הדבק לקובץ בעצמכם",
    "כלי אג'נטי עושה בדיוק את זה בשבילכם — רק כותב ישירות לקבצים, וגם קורא את שאר הפרויקט קודם כדי להבין הקשר",
  ],
  bestPractice: "זה לא כלי חדש לגמרי — זה שדרוג למשהו שאתם כבר עושים, רק עם פחות חיכוך והרבה יותר הקשר.",
});

// =====================================================================
// שקף 13 — מי משתמש במה, בפועל
// =====================================================================
contentSlide({
  title: "מי משתמש במה, בפועל",
  bullets: [
    "**ChatGPT** — 82% נתח שימוש כללי לעזרה בקוד (הכלי הכי נפוץ)",
    "**GitHub Copilot** — 68% נתח שימוש כללי",
    "מבין ה-IDE-ים הממוקדי-AI: **Cursor** (עורך קוד עם AI מובנה) — 18% נתח שימוש כללי, אך 24% כ\"כלי ראשי\" בקרב מי שכבר עובד אג'נטית",
    "**Claude Code** — 10% נתח שימוש כללי, אך **28%** כ\"כלי ראשי\" בקרב מי שכבר עובד אג'נטית — כלומר מי שבחר להתמקצע בעבודה אג'נטית מלאה נוטה לבחור בו דווקא",
    "(ChatGPT ו-Copilot לא נמדדו בקטגוריית \"כלי ראשי אג'נטי\" כי הם משמשים בעיקר כעזר כללי, לא כסביבת עבודה אג'נטית מלאה)",
  ],
  bestPractice: "אין \"כלי אחד נכון\" — הקורס מלמד עם Claude Code כי הוא מייצג היטב את העבודה האג'נטית המלאה, אבל ה-Workflow שתלמדו עובד בכל אחד מהכלים האלה.",
  source: "Stack Overflow Developer Survey (סקר שנתי בקרב מפתחים, מבית אתר Stack Overflow) 2025 — survey.stackoverflow.co/2025 · index.dev 2026",
});

// =====================================================================
// שקף 14 — כלים חינמיים להתחלה
// =====================================================================
contentSlide({
  title: "כלים חינמיים להתחלה",
  bullets: [
    "**GitHub Copilot Free** — מסלול חינמי רחב, ללא צורך באימות סטודנט, זמין ישירות ב-VS Code — נקודת הכניסה הפשוטה ביותר. כולל גם מצב Agent (לא רק השלמות), עם בחירת מודל אוטומטית ומכסה חודשית (docs.github.com/.../plans)",
    "**GitHub Student Developer Pack** — חבילת הטבות לסטודנטים מאומתים; הרשמה חדשה להטבת **Copilot Student** (מסלול מורחב יותר מ-Copilot Free הרגיל) מושהית מאפריל 2026 בשל עומס על תשתית Agentic — מי שכבר אומת קודם ממשיך לקבל אותה, וחברי Pack חדשים מקבלים כרגע את Copilot Free הרגיל (education.github.com/pack)",
    "**Cursor** — IDE מלא מבוסס AI, עם מסלול חינמי מוגבל (Agent requests מוגבל + מכסת השלמות בסיסית, בלי כרטיס אשראי) — cursor.com/pricing",
  ],
  bestPractice: "שוק הכלים משתנה מהר (בשיעור 1 ראינו את Gemini CLI החינמי נסגר תוך פחות משנה) — ולכן ההתקנה וההרשמה בפועל לכלי הנבחר קורות בתרגול המיידי, לא כאן. תמיד כדאי לבדוק סטטוס עדכני קרוב למועד השימוש.",
});

// =====================================================================
// שקף 15 — Best Practices (חלק א')
// =====================================================================
contentSlide({
  title: "Best Practices לעבודה עם כלי Agentic (חלק א׳)",
  bullets: [
    "**1. הנחיות ברורות וספציפיות** — \"תוסיף התחברות\" עמום מדי; \"תוסיף התחברות עם Google **OAuth** (פרוטוקול התחברות מאובטח דרך חשבון קיים), עם Redirect (הפניה אוטומטית) לדף הבית אחרי הצלחה\" ברור וממוקד",
    "**2. צעדים קטנים** — משימה גדולה מתפרקת לצעדים; קל יותר לבדוק תוצאה של צעד קטן מאשר של שינוי ענק",
    "**3. בדיקת כל Diff** — לפני אישור שינוי, קוראים אותו. תמיד, בלי יוצא מן הכלל",
  ],
  bestPractice: "שלושת הכללים האלה (וגם השלושה בשקף הבא) הם מה שמבדיל שימוש מקצועי מ\"ניסוי מקרי\".",
});

// =====================================================================
// שקף 16 — Best Practices (חלק ב')
// =====================================================================
contentSlide({
  title: "Best Practices לעבודה עם כלי Agentic (חלק ב׳)",
  bullets: [
    "**4. Git כרשת ביטחון** — Commit לפני כל ניסוי משמעותי",
    "**5. לא לסמוך בעיוורון** — אם משהו נראה \"יותר מדי חכם\", או משתמש בפונקציה שלא הכרתם — בודקים שהיא אכן קיימת (זוכרים את Hallucination משיעור 1?)",
    "**6. איטרציה, לא ניסיון יחיד** — אם התוצאה הראשונה לא מושלמת, מבקשים תיקון ממוקד, לא מתחילים מאפס",
  ],
  bestPractice: "בשקף הבא נראה שיש לזה גם מחיר מדיד, לא רק \"עצה טובה\".",
});

// =====================================================================
// שקף 16.5 — תרגול בזוגות: לתקן הנחיה מעורפלת
// =====================================================================
contentSlide({
  title: "תרגול בזוגות: לתקן הנחיה מעורפלת",
  bullets: [
    "**המשימה (3 דקות בזוגות):** אחד/אחת בזוג כותב/ת בכוונה הנחיה מעורפלת לכלי AI (למשל: \"תשפר את הקוד\", \"תתקן את הבאג\", \"תוסיף אבטחה\"). השני/ה כותב/ת אותה מחדש לפי כלל 1 (הנחיות ברורות וספציפיות) — כמו הדוגמה מהשקף הקודם.",
  ],
  bestPractice: "אחרי 3 דקות, 2-3 זוגות מקריאים בקול את ההנחיה ה\"לפני\" וה\"אחרי\". מטרת התרגיל: להרגיש בעצמכם למה כלי AI \"מנחש\" כשההנחיה עמומה — לא רק לשמוע את זה כעצה.",
});

// =====================================================================
// שקף 17 — המחיר של דילוג על הכללים
// =====================================================================
contentSlide({
  title: "המחיר של דילוג על הכללים",
  bullets: [
    "דו\"ח **DORA** (מחקר שנתי מוביל על תהליכי פיתוח תוכנה, מבית Google — הכרנו בשיעור 1) מצא: אימוץ AI בלי תהליך עבודה מסודר (בדיוק הכללים משני השקפים הקודמים) מוביל ל-**30-41%** יותר **חוב טכני** (עבודה \"מהירה\" עכשיו שדורשת תיקון יקר יותר בהמשך), לעומת עבודה עם תהליך מסודר",
    "בקשות שילוב קוד (Pull Requests) שנכתבו על ידי AI נושאות **פי 1.7** יותר בעיות מ-PR שכתב בן אדם",
  ],
  bestPractice: "ה-Best Practices משני השקפים הקודמים הם לא בירוקרטיה — הם ההבדל בין \"AI שעוזר\" ל\"AI שיוצר בעיה גדולה יותר ממה שפתר\".",
  source: "DORA — ROI of AI-Assisted Software Development, 2026",
});

// =====================================================================
// שקף 17.5 — מקרה בוחן: Amazon
// =====================================================================
contentSlide({
  title: "מקרה בוחן: כשמדלגים על Review בפועל",
  question: "השאלה: זה נשמע תיאורטי — האם יש דוגמה אמיתית וגדולה למחיר הזה?",
  bullets: [
    "במרץ 2026, אתר המסחר של **Amazon** עבר שני תקלות חמורות תוך שלושה ימים: ב-2 במרץ קריסה של כ-6 שעות (כ-1.6 מיליון שגיאות אתר, כ-120,000 הזמנות אבודות), וב-5 במרץ קריסה חמורה עוד יותר — ירידה של 99% בהזמנות בצפון אמריקה ביחס לקצב הרגיל, כ-6.3 מיליון הזמנות אבודות",
    "לפי דיווחי Amazon עצמה, הגורם המרכזי לתקלה השנייה היה **שינוי קוד שעבר לייצור בלי תהליך אישור ותיעוד פורמלי** — בדיוק סוג הדילוג-על-Review שדיברנו עליו. Amazon הבהירה שרק תקלה אחת בכלל נגעה לכלי AI, ושם הסיבה הייתה תקלה אנושית בתהליך",
    "**התגובה של Amazon:** \"איפוס בטיחות קוד\" בן 90 יום על כ-335 מערכות קריטיות — כל שינוי דורש עכשיו אישור שני אנשים, תיעוד פורמלי, ובדיקות אוטומטיות מחמירות יותר; שינויי קוד שנוצרו בעזרת AI דורשים ספציפית גם חתימה של מהנדס בכיר",
  ],
  bestPractice: "גם ענקית כמו Amazon נדרשה לחזק בדיוק את הכללים שלמדתם היום — לא כי AI \"רע\", אלא כי דילוג על Review הוא הסיכון האמיתי, עם או בלי AI.",
  source: "Yahoo Tech, Digital Trends — מרץ 2026",
});

// =====================================================================
// שקף 18 — הדגמה חיה (חלק א׳)
// =====================================================================
contentSlide({
  title: "הדגמה חיה (חלק א׳): Autocomplete מול Agentic",
  bullets: [
    "20-25 דקות. פותחים session אמיתי מול קוד קיים — פרויקט ה-Todo App שבנינו יחד בשיעור 1",
    "**שלב 1 — Autocomplete (לא אג'נטי):** המרצה מקליד בעצמו בתוך `index.html`, למשל תוך כדי הוספת פונקציה חדשה, ונותן לכלי להשלים שורה אחת בלבד. הדגש: \"זה עוזר, אבל אני עדיין מחליט כל דבר.\"",
    "**שלב 2 — שימוש אג'נטי אמיתי:** פותחים session אג'נטי מלא (Claude Code / Cursor Agent Mode) על אותה תיקייה, ומקלידים את הפרומפט שבשקף הבא",
  ],
  bestPractice: "הניגוד בין שני השלבים הוא הנקודה הפדגוגית המרכזית — לתת לכיתה לראות בעיניים את ההבדל, לא רק לשמוע עליו.",
});

// =====================================================================
// שקף 19 — הדגמה חיה (חלק ב׳)
// =====================================================================
promptSlide({
  eyebrow: "הדגמה חיה (חלק ב׳)",
  title: "פרומפט אג'נטי מלא",
  sub: "אחרי הדגמת Autocomplete בשלב 1, מקלידים את זה מול הכיתה:",
  boxes: [
    {
      label: "",
      lines: [
        "תוסיף לאפליקציית ה-Todo הזו שדה \"עדיפות\" לכל משימה (נמוכה/בינונית/גבוהה),\nעם צבע רקע שונה לכל רמת עדיפות ברשימה. תשמור על אותו סגנון קוד שכבר קיים\nבקובץ.",
      ],
    },
    {
      label: "לפני האישור, עוצרים ובודקים מול הכיתה:",
      lines: [
        "עוצרים לפני אישור השינוי, ובודקים ביחד את ה-Diff שהתקבל: אילו שורות נוספו, אילו השתנו. שני דברים לשים לב אליהם: (א) האם הכלי \"קרא\" את הקוד הקיים לפני שהוסיף (שמר על מוסכמות שמות קיימות)? (ב) האם הכלי הציע לשנות משהו לא קשור (Scope Creep — \"זחילת היקף\", הרחבה לא מתוכננת של המשימה)? אם כן — זו הזדמנות להראות איך דוחים חלק מהצעה ומאשרים רק את מה שרלוונטי.",
      ],
    },
  ],
});

// =====================================================================
// שקף 20 — עוברים לתרגול (חלק א׳: המשימה)
// =====================================================================
contentSlide({
  title: "עוברים לתרגול: הקמת סביבה (חלק א׳ — המשימה)",
  question: "השאלה: מה המטרה של 45 הדקות האלה?",
  bullets: [
    "להקים סביבת עבודה אישית שתלווה אתכם לאורך כל הסמסטר, ולבצע בה פעולה אמיתית ראשונה עם כלי AI",
    "המשימה: מתקינים VS Code + Git, יוצרים Repository חדש בשם `my-vibe-coding-env` ב-GitHub, משכפלים אותו למחשב (`git clone`), ומפעילים כלי AI חינמי (Copilot Free או שקול)",
    "אז מבקשים מה-AI ליצור קובץ `hello.py` (או `hello.js`) שמדפיס הודעת פתיחה אישית — זו הפעולה האג'נטית הראשונה שלכם בקורס",
  ],
  bestPractice: "אם משהו בהתקנה לא עובד — זה בדיוק סוג הבעיה ש-AI טוב בפתרון שלה. מעתיקים את הודעת השגיאה המדויקת ושואלים את כלי ה-AI \"מה השגיאה הזו אומרת ואיך פותרים אותה?\" — זה תרגול לגיטימי, לא \"רמאות\".",
});

// =====================================================================
// שקף 21 — עוברים לתרגול (חלק ב׳: פרומפט פתיחה)
// =====================================================================
promptSlide({
  eyebrow: "עוברים לתרגול (חלק ב׳)",
  title: "פרומפט פתיחה לתרגול",
  sub: "התחילו מכאן, והתאימו את השם שלכם:",
  boxes: [
    {
      label: "",
      lines: [
        "תיצור קובץ hello.py שמדפיס \"שלום מ-Vibe Coding, שיעור 2!\" ואת השם שלי\n[השם שלכם]. תוסיף גם הדפסה של התאריך הנוכחי.",
      ],
    },
  ],
});

// =====================================================================
// שקף 22 — עוברים לתרגול (חלק ג׳: מתי סיימתם)
// =====================================================================
contentSlide({
  title: "עוברים לתרגול: מתי סיימתם? (חלק ג׳)",
  question: "השאלה: איך יודעים שהתרגול הושלם בהצלחה?",
  bullets: [
    "VS Code מותקן ופתוח על תיקיית הפרויקט",
    "Git מותקן, וה-Repository `my-vibe-coding-env` קיים ב-GitHub",
    "כלי AI חינמי מותקן ופעיל (Copilot Free או שקול)",
    "קובץ `hello.py`/`hello.js` נוצר בעזרת AI, נבדק, ורץ בהצלחה",
    "בוצע Commit אחד לפחות, ונדחף (**Push** — העלאה מהמחשב המקומי לשרת המרוחק) ל-GitHub בהצלחה",
  ],
  bestPractice: "הבעיה הכי נפוצה בתרגול הזה היא אימות/הרשאות מול GitHub (HTTPS Token או SSH key — שתי שיטות שונות להוכיח שאתם באמת אתם). אם נתקעתם כאן — זה בדיוק המקום לבקש עזרה מהמרצה או מ-AI, זה לא קשור ליכולת התכנות שלכם.",
});

// =====================================================================
// שקף 23 — סיכום
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addTitle(s, "סיכום");

  s.addText(
    bulletRuns(
      [
        "VS Code + Git + GitHub = השולחן שעליו עובדים לאורך כל הקורס",
        "Agentic ≠ Autocomplete — זה שותף לביצוע משימה, לא רק עוזר הקלדה (וזוכרים: תמיד בודקים Diff לפני אישור)",
        "יש נקודת כניסה חינמית לכל הכלים שלמדנו היום — אין תירוץ לא להתחיל",
      ],
      { fontFace: BODY_FONT, fontSize: 19, color: TEXT_DARK }
    ),
    { x: 0.6, y: 1.5, w: 12.13, h: 2.2, align: "right", rtlMode: true, paraSpaceAfter: 12, margin: 0 }
  );

  s.addShape(pres.ShapeType.line, { x: 0.6, y: 4.0, w: 12.13, h: 0, line: { color: PH_BORDER, width: 0.75 } });

  s.addText(
    [
      { text: "שבוע הבא: ", options: { bold: true, rtlMode: true } },
      { text: "Context Engineering — איך \"מזינים\" לכלי את מה שהוא צריך לדעת על הפרויקט שלכם", options: { rtlMode: true, breakLine: true } },
    ],
    { x: 0.6, y: 4.2, w: 12.13, h: 0.6, align: "right", fontFace: TITLE_FONT, fontSize: 17, color: TEXT_DARK, rtlMode: true, margin: 0 }
  );

  s.addText(
    "בהמשך הקורס: עבודה עם קוד קיים (שבוע 4) · Agents (שבוע 5) · תכנון לפני קוד (שבוע 6) · עד פריסה מלאה בענן ואבטחה (שבועות 11-12)",
    { x: 0.6, y: 4.9, w: 12.13, h: 0.5, align: "right", fontFace: BODY_FONT, italic: true, fontSize: 13, color: MUTED, rtlMode: true, margin: 0 }
  );

  addFooter(s);
}

// =====================================================================
// שקף 24 — מקורות
// =====================================================================
{
  const s = pres.addSlide();
  s.background = { color: WHITE };
  addTitle(s, "מקורות");

  s.addText(
    "Kinsta & Skillademia — GitHub Statistics 2026 · DigitalApplied & MorphLLM — Benchmark Guides 2026 · swebench.com · tbench.ai · metr.org/time-horizons · Stack Overflow Developer Survey 2025 (survey.stackoverflow.co/2025) · index.dev 2026 · DORA — ROI of AI-Assisted Software Development, 2026 · cursor.com/pricing · docs.github.com/copilot/plans · education.github.com/pack · Yahoo Tech / Digital Trends (מקרה Amazon, מרץ 2026)",
    { x: 0.6, y: 1.5, w: 12.13, h: 2.6, align: "right", fontFace: BODY_FONT, fontSize: 15, color: TEXT_DARK, rtlMode: true, margin: 0 }
  );

  s.addText(
    "פירוט מלא: references.md — כל הנתונים נבדקו ביולי 2026 ועשויים להשתנות",
    { x: 0.6, y: 4.3, w: 12.13, h: 0.4, align: "right", fontFace: BODY_FONT, italic: true, fontSize: 12, color: MUTED, rtlMode: true, margin: 0 }
  );

  addFooter(s);
}

// =====================================================================
pres.writeFile({ fileName: "lesson_02_slides.pptx" }).then(() => {
  console.log("done");
});
