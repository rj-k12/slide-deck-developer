const pptxgen = require("pptxgenjs");
const fs = require("fs");

const lessonPath = process.argv[2] || "lesson.json";
const outPath = process.argv[3] || "output.pptx";
const lesson = JSON.parse(fs.readFileSync(lessonPath, "utf-8"));

// Same verified palette as generate_deck.js / generate_deck_literature_response.js.
const NAVY_INK = "0E0142";
const PURPLE = "3928AA";
const CORAL = "F19A65";
const TITLE_PURPLE = "382DB0";
const CREAM_YELLOW = "FFF3CB";
const PEACH = "FFDBC5";
const MUTED = "6B6478";
const TABLE_BORDER_BODY = "FFAB7B";
const TP_BORDER_ORANGE = "FFAB7B";
const TP_BORDER_CLOSING = "DCD6EE";
const WHITE = "FFFFFF";
const BODY = "33322E";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
const PAGE_W = 13.3, PAGE_H = 7.5;

const HEADER_GRADIENT = require("path").join(__dirname, "assets", "header_gradient.png");
function addHeaderGradient(slide) {
  const scale = PAGE_W / 10;
  slide.addImage({ path: HEADER_GRADIENT, x: 0, y: 0, w: 10 * scale, h: 1.2578 * scale });
}

function unitHeader() {
  return `${lesson.grade}, Knowledge Unit ${lesson.unit_number} | ${lesson.unit_title}`;
}
function footer(slide, pageNum, onDark) {
  const c = onDark ? "C9BEEB" : MUTED;
  slide.addText(unitHeader(), { x: 0.5, y: PAGE_H - 0.4, w: PAGE_W - 1.5, h: 0.3, fontFace: "Arial", fontSize: 12, color: c, align: "left", margin: 0 });
  slide.addText(String(pageNum), { x: PAGE_W - 0.9, y: PAGE_H - 0.4, w: 0.4, h: 0.3, fontFace: "Arial", fontSize: 12, color: c, align: "right", margin: 0 });
}
function slideTitle(slide, title, onDark) {
  slide.addText(title, { x: 0.55, y: 0.35, w: PAGE_W - 1.2, h: 0.6, fontFace: "Arial", fontSize: 36, bold: true, color: onDark ? WHITE : TITLE_PURPLE, margin: 0 });
}
function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}
function boldPageNumbers(text) {
  const re = /\bpages?\s+\d+(?:[\u2013-]\d+)?\b/gi;
  const parts = [];
  let lastIndex = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > lastIndex) parts.push({ text: text.slice(lastIndex, m.index) });
    parts.push({ text: m[0], options: { bold: true } });
    lastIndex = m.index + m[0].length;
  }
  if (lastIndex < text.length) parts.push({ text: text.slice(lastIndex) });
  return parts.length ? parts : [{ text }];
}

const VOCAB_ICON_DIR = process.argv[5] || require("path").join(__dirname, "vocab_icons");
function iconPathFor(word) {
  const safe = word.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  const p = require("path").join(VOCAB_ICON_DIR, `${safe}.png`);
  return fs.existsSync(p) ? p : null;
}
function vocabBlock(slide, words, x, y, w, onDark) {
  const colW = (w - 0.3) / 2;
  const cardBg = onDark ? WHITE : PEACH;
  words.forEach((item, i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const bx = x + col * (colW + 0.3), by = y + row * 2.85;
    slide.addShape("roundRect", { x: bx, y: by, w: colW, h: 2.7, rectRadius: 0.20, fill: { color: cardBg }, line: onDark ? { color: CORAL, width: 1 } : { type: "none" } });
    const iconPath = iconPathFor(item.word);
    if (iconPath) {
      slide.addImage({ path: iconPath, x: bx + 0.2, y: by + 0.2, w: 0.55, h: 0.55 });
    }
    slide.addText(item.word, { x: bx + 0.9, y: by + 0.22, w: colW - 1.05, h: 0.5, fontFace: "Arial", fontSize: 24, bold: true, color: NAVY_INK, margin: 0 });
    slide.addText(item.definition, { x: bx + 0.2, y: by + 1.0, w: colW - 0.4, h: 1.55, fontFace: "Arial", fontSize: 21.5, color: BODY, margin: 0, valign: "top" });
  });
}

// ===== Generic chart (Reading family's fixed 2-column shape) =====
function addChart(slide, chart, x, y, w, h) {
  const nRows = (chart.rows && chart.rows.length) || 4;
  const headerBorder = { type: "solid", color: PURPLE, pt: 1 };
  const bodyBorder = { type: "solid", color: TABLE_BORDER_BODY, pt: 1 };
  const headerRow = chart.columns.map(c => ({
    text: c, options: { bold: true, color: WHITE, fill: { color: PURPLE }, align: "left", valign: "middle", fontFace: "Arial", fontSize: 14, border: headerBorder }
  }));
  const bodyRows = [];
  for (let r = 0; r < nRows; r++) {
    const zebra = r % 2 === 1;
    const cells = chart.columns.map((_, ci) => {
      const text = (chart.rows && chart.rows[r]) ? chart.rows[r][ci] : "";
      return { text, options: { color: BODY, fontFace: "Arial", fontSize: 13, valign: "top", fill: { color: zebra ? PEACH : WHITE }, border: bodyBorder } };
    });
    bodyRows.push(cells);
  }
  slide.addTable([headerRow, ...bodyRows], {
    x, y, w, h,
    fontFace: "Arial", autoPage: false,
    rowH: [0.45, ...bodyRows.map(() => (h - 0.45) / nRows)],
  });
}

// ===== Generic planner-sections renderer =====
// NEW for this generator. Language and Discourse Deep Dive both show
// reference charts and structured content that don't fit the Reading
// family's single fixed 2-column "chart" shape above -- e.g. a 3-column
// suffix reference chart with two examples per row, or a chart where
// each cell is itself a bulleted list ("Strategies for Listening and
// Building On"). Prototyped and visually verified (rendered with
// LibreOffice, not just XML-checked) against three real examples from
// RT_Gr4U2_KNO_full_unit_Teacher_Guide.pdf before being wired in here.
//
// Schema: an array of { type: "text"|"list"|"table", label?, ... } --
// see extract_lesson.py's planner_sections field for the full shape.
//
// Known open question, NOT resolved here: whether adjacent short
// sections (e.g. two "list" sections back to back) should render
// side-by-side rather than stacked -- there's no real Language/Discourse
// Deep Dive template to confirm this against, so everything stacks
// top-to-bottom for now. Revisit if a real template surfaces.
function estimateTableRowHeight(row, colWidths) {
  // Per-row height driven by whichever cell needs the most vertical
  // space -- a cell can be a plain string (wraps to N lines) or an array
  // of strings (each item is its own bulleted line, each of which can
  // ALSO wrap). The original flat 0.5in/row estimate visibly collided
  // two stacked tables in testing -- a cell with 3-4 wrapped bullet
  // items needs much more than 0.5in.
  let maxLines = 1;
  row.forEach((cell, ci) => {
    const w = colWidths[ci];
    const charsPerLine = w * 11;
    if (Array.isArray(cell)) {
      let lines = 0;
      cell.forEach(item => { lines += Math.max(1, Math.ceil(item.length / charsPerLine)); });
      maxLines = Math.max(maxLines, lines);
    } else {
      const lines = Math.max(1, Math.ceil((cell || "").length / charsPerLine));
      maxLines = Math.max(maxLines, lines);
    }
  });
  return Math.max(0.4, maxLines * 0.22 + 0.15);
}
function estimatePlannerSectionHeight(section, w) {
  if (section.type === "text") {
    const charsPerLine = w * 11;
    const lines = Math.ceil(section.content.length / charsPerLine);
    return 0.35 + lines * 0.28 + 0.15;
  }
  if (section.type === "list") {
    const charsPerLine = (w - 0.3) * 11;
    let lines = 0;
    section.items.forEach(item => { lines += Math.max(1, Math.ceil(item.length / charsPerLine)); });
    return 0.35 + lines * 0.26 + 0.15;
  }
  if (section.type === "table") {
    const colWidths = section.columns.map(() => w / section.columns.length);
    const rowHeights = section.rows.map(row => estimateTableRowHeight(row, colWidths));
    return 0.35 + 0.4 + rowHeights.reduce((a, b) => a + b, 0) + 0.15;
  }
  return 0.5;
}
function renderPlannerSections(slide, sections, x, y, w) {
  let cursorY = y;
  (sections || []).forEach(section => {
    const sectionH = estimatePlannerSectionHeight(section, w);
    if (section.label) {
      slide.addText(section.label, { x, y: cursorY, w, h: 0.3, fontFace: "Arial", fontSize: 16, bold: true, color: PURPLE, margin: 0 });
      cursorY += 0.35;
    }
    if (section.type === "text") {
      slide.addText(section.content, { x, y: cursorY, w, h: 1, fontFace: "Arial", fontSize: 14, color: BODY, margin: 0, valign: "top" });
    } else if (section.type === "list") {
      const bulletText = section.items.map(item => ({ text: item, options: { bullet: { code: "2022" }, breakLine: true } }));
      slide.addText(bulletText, { x, y: cursorY, w, h: 1.5, fontFace: "Arial", fontSize: 13, color: BODY, margin: 0, valign: "top", paraSpaceAfter: 4 });
    } else if (section.type === "table") {
      const nCols = section.columns.length;
      const colWidths = section.columns.map(() => w / nCols);
      const headerRow = section.columns.map(c => ({
        text: c, options: { bold: true, color: WHITE, fill: { color: PURPLE }, align: "center", valign: "middle", fontFace: "Arial", fontSize: 12, border: { type: "solid", color: PURPLE, pt: 1 } }
      }));
      const bodyRows = section.rows.map((row, ri) => {
        const zebra = ri % 2 === 1;
        return row.map(cell => {
          const cellText = Array.isArray(cell)
            ? cell.map(item => ({ text: item, options: { bullet: { code: "2022" }, breakLine: true } }))
            : cell;
          return { text: cellText, options: { color: BODY, fontFace: "Arial", fontSize: 11, valign: "top", fill: { color: zebra ? PEACH : WHITE }, border: { type: "solid", color: TABLE_BORDER_BODY, pt: 1 } } };
        });
      });
      // Real per-row heights (see estimateTableRowHeight), not a uniform
      // average -- rows with denser bulleted cells get more room than
      // rows with a single short line.
      const rowHeights = section.rows.map(row => estimateTableRowHeight(row, colWidths));
      slide.addTable([headerRow, ...bodyRows], {
        x, y: cursorY, w, fontFace: "Arial", autoPage: false,
        rowH: [0.4, ...rowHeights],
        colW: colWidths,
      });
    }
    cursorY += sectionH;
  });
  return cursorY - y;
}

// Splits a planner_sections array into groups that each fit within
// availableH, so a slide with two dense reference tables (as seen in
// Lesson 11b: Discourse Deep Dive's real content) doesn't silently
// overflow into the footer. Mirrors this codebase's existing convention
// for handling overflow, e.g. vocabBlock's chunk-of-4-per-slide-with-
// "(continued)" pattern, rather than trying to shrink content to force
// a fit -- shrinking risks unreadably small text; splitting doesn't.
function paginatePlannerSections(sections, w, availableH) {
  const groups = [];
  let current = [], currentH = 0;
  (sections || []).forEach(section => {
    const labelH = section.label ? 0.35 : 0;
    const sectionH = estimatePlannerSectionHeight(section, w) + labelH;
    if (current.length && currentH + sectionH > availableH) {
      groups.push(current);
      current = [];
      currentH = 0;
    }
    current.push(section);
    currentH += sectionH;
  });
  if (current.length) groups.push(current);
  return groups.length ? groups : [[]];
}

// ===== Slide: Cover (optional) =====
const coverImagePath = process.argv[4];
if (coverImagePath && fs.existsSync(coverImagePath)) {
  const s = pres.addSlide();
  s.addImage({ path: coverImagePath, x: 0, y: 0, w: PAGE_W, h: PAGE_H });
}

// ===== Slide: Lesson title =====
{
  const s = pres.addSlide();
  s.background = { color: PURPLE };
  s.addShape("roundRect", { x: 0.55, y: 0.6, w: 1.9, h: 0.55, rectRadius: 0.28, fill: { color: CORAL }, line: { type: "none" } });
  s.addText(`LESSON ${lesson.lesson_number}`, { x: 0.55, y: 0.6, w: 1.9, h: 0.55, fontFace: "Arial", fontSize: 16, bold: true, color: NAVY_INK, align: "center", valign: "middle", margin: 0 });
  s.addText(`Lesson ${lesson.lesson_number}: ${lesson.lesson_type}`, { x: 0.55, y: 1.3, w: 12.2, h: 1.8, fontFace: "Arial", fontSize: 50, bold: true, color: WHITE, margin: 0, valign: "top" });
  s.addText(`Grade ${lesson.grade.replace('Grade ','')}, Knowledge Unit ${lesson.unit_number}: ${lesson.unit_title}`, { x: 0.55, y: 3.35, w: 12.2, h: 1.1, fontFace: "Arial", fontSize: 25.5, color: "E5EFF9", margin: 0, valign: "top" });
  if (lesson.core_text) {
    const byLine = lesson.author ? `${lesson.core_text} by ${lesson.author}` : lesson.core_text;
    s.addText(`${byLine}${lesson.pages ? "   |   Pages " + lesson.pages : ""}`, { x: 0.55, y: 4.65, w: 12.2, h: 0.6, fontFace: "Arial", fontSize: 21.5, italic: true, color: "C9BEEB", margin: 0, valign: "top" });
  }
  s.addText("Copyright \u00a9 2026 Lavinia Group. All Rights Reserved. RedThread is a trademark of K12 Coalition.", { x: 0.55, y: PAGE_H - 0.5, w: 11.5, h: 0.3, fontFace: "Arial", fontSize: 10, color: "9A8FD1", margin: 0 });
}

// ===== Slide: Teaching Point (+ Essential Question, if present) =====
// Real structural difference from the Reading/Literature Response
// family: Language lessons frequently have NO Essential Question at all
// (confirmed directly in Lesson 7a: Language, which opens with only a
// Teaching Point). When essential_question is empty, this renders a
// single full-height Teaching Point box instead of the usual stacked
// EQ-on-top layout -- not a smaller version of the same layout with an
// empty gap where the EQ would go.
{
  const s = pres.addSlide();
  const hasEQ = !!(lesson.essential_question && lesson.essential_question.trim());

  if (hasEQ) {
    addHeaderGradient(s);
    s.addShape("rect", { x: 0.55, y: 0.55, w: 0.06, h: 1.35, fill: { color: PURPLE }, line: { type: "none" } });
    s.addShape("roundRect", { x: 0.61, y: 0.55, w: PAGE_W - 1.2, h: 1.35, rectRadius: 0.139, fill: { color: CREAM_YELLOW }, line: { type: "none" } });
    s.addText("Essential Question", { x: 0.85, y: 0.7, w: PAGE_W - 1.6, h: 0.3, fontFace: "Arial", fontSize: 18.5, bold: true, color: PURPLE, margin: 0 });
    s.addText(lesson.essential_question, { x: 0.85, y: 1.05, w: PAGE_W - 1.6, h: 0.8, fontFace: "Arial", fontSize: 24, bold: true, color: NAVY_INK, margin: 0, valign: "top" });
  }

  const tpTop = hasEQ ? 2.15 : 0.75;
  const tpBottom = PAGE_H - 0.8;
  s.addShape("rect", { x: 0.55, y: tpTop, w: PAGE_W - 1.1, h: 0.05, fill: { color: CORAL }, line: { type: "none" } });
  s.addShape("roundRect", { x: 0.55, y: tpTop + 0.05, w: PAGE_W - 1.1, h: tpBottom - tpTop - 0.05, rectRadius: hasEQ ? 0.113 : 0.139, fill: { color: WHITE }, line: { color: hasEQ ? TP_BORDER_CLOSING : TP_BORDER_ORANGE, width: hasEQ ? 1 : 1.5 } });
  s.addText("Teaching Point", { x: 0.75, y: tpTop + 0.25, w: PAGE_W - 1.5, h: 0.35, fontFace: "Arial", fontSize: 22, bold: true, color: CORAL, margin: 0 });
  if (lesson.teaching_point) {
    s.addText(lesson.teaching_point, { x: 0.75, y: tpTop + 0.65, w: PAGE_W - 1.5, h: tpBottom - tpTop - 0.85, fontFace: "Arial", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });
  }
  footer(s, 2, false);
}

// ===== Slides: Engage / Launch sections =====
let pageNum = 3;
(lesson.sections || []).forEach(section => {
  const hasContent = (section.vocabulary && section.vocabulary.length) ||
    (section.chart && section.chart.rows && section.chart.rows.length) ||
    (section.planner_sections && section.planner_sections.length) ||
    section.read_directions;
  if (!hasContent) return;

  if (section.vocabulary && section.vocabulary.length) {
    chunk(section.vocabulary, 4).forEach((words, i) => {
      const s = pres.addSlide();
      if (section.section_name === "Engage") addHeaderGradient(s);
      slideTitle(s, `${section.section_name} Vocabulary${i > 0 ? " (continued)" : ""}`, false);
      vocabBlock(s, words, 0.55, 1.25, PAGE_W - 1.1, false);
      footer(s, pageNum++, false);
    });
  }
  if (section.chart && section.chart.rows && section.chart.rows.length) {
    const s = pres.addSlide();
    slideTitle(s, `${section.section_name} Chart`, false);
    if (section.read_directions) {
      s.addText(boldPageNumbers(section.read_directions), { x: 0.55, y: 1.1, w: PAGE_W - 1.1, h: 0.6, fontFace: "Arial", fontSize: 16, color: BODY, margin: 0, valign: "top" });
    }
    addChart(s, section.chart, 0.55, section.read_directions ? 1.8 : 1.2, PAGE_W - 1.1, PAGE_H - (section.read_directions ? 2.3 : 1.7));
    footer(s, pageNum++, false);
  }
  if (section.planner_sections && section.planner_sections.length) {
    const availableH = PAGE_H - 1.1 - 0.6; // top margin + footer clearance
    const groups = paginatePlannerSections(section.planner_sections, PAGE_W - 1.1, availableH);
    groups.forEach((group, gi) => {
      const s = pres.addSlide();
      const titleSuffix = section.chart || section.vocabulary ? " Reference" : "";
      slideTitle(s, `${section.section_name}${titleSuffix}${gi > 0 ? " (continued)" : ""}`, false);
      renderPlannerSections(s, group, 0.55, 1.1, PAGE_W - 1.1);
      footer(s, pageNum++, false);
    });
  }
  if (section.read_directions && !(section.chart && section.chart.rows && section.chart.rows.length)) {
    const s = pres.addSlide();
    slideTitle(s, `${section.section_name} Read Directions`, false);
    s.addText(boldPageNumbers(section.read_directions), { x: 0.55, y: 1.5, w: PAGE_W - 1.1, h: 1, fontFace: "Arial", fontSize: 21.5, color: BODY, margin: 0, valign: "top" });
    footer(s, pageNum++, false);
  }
});

// ===== Slide: Application (Language) =====
if (lesson.application && (lesson.application.read_directions || (lesson.application.vocabulary || []).length)) {
  const app = lesson.application;
  if (app.vocabulary && app.vocabulary.length) {
    chunk(app.vocabulary, 4).forEach((words, i) => {
      const s = pres.addSlide();
      slideTitle(s, `Application${i > 0 ? " (continued)" : ""}`, false);
      if (i === 0 && app.read_directions) {
        s.addText(boldPageNumbers(app.read_directions), { x: 0.55, y: 1.1, w: PAGE_W - 1.1, h: 0.5, fontFace: "Arial", fontSize: 16, color: BODY, margin: 0, valign: "top" });
      }
      vocabBlock(s, words, 0.55, app.read_directions && i === 0 ? 1.75 : 1.25, PAGE_W - 1.1, false);
      footer(s, pageNum++, false);
    });
  } else if (app.read_directions) {
    const s = pres.addSlide();
    slideTitle(s, "Application", false);
    s.addText(boldPageNumbers(app.read_directions), { x: 0.55, y: 1.5, w: PAGE_W - 1.1, h: 1, fontFace: "Arial", fontSize: 21.5, color: BODY, margin: 0, valign: "top" });
    footer(s, pageNum++, false);
  }
  if (app.chart && app.chart.rows && app.chart.rows.length) {
    const s = pres.addSlide();
    slideTitle(s, "Application Chart", false);
    addChart(s, app.chart, 0.55, 1.2, PAGE_W - 1.1, PAGE_H - 1.7);
    footer(s, pageNum++, false);
  }
  if (app.planner_sections && app.planner_sections.length) {
    const availableH = PAGE_H - 1.1 - 0.6;
    const groups = paginatePlannerSections(app.planner_sections, PAGE_W - 1.1, availableH);
    groups.forEach((group, gi) => {
      const s = pres.addSlide();
      slideTitle(s, `Application Reference${gi > 0 ? " (continued)" : ""}`, false);
      renderPlannerSections(s, group, 0.55, 1.1, PAGE_W - 1.1);
      footer(s, pageNum++, false);
    });
  }
}

// ===== Slide: Evidence Collection (Discourse Deep Dive) =====
if (lesson.evidence_collection && (lesson.evidence_collection.prompts || []).length) {
  const s = pres.addSlide();
  slideTitle(s, "Evidence Collection", false);
  let y = 1.2;
  if (lesson.evidence_collection.intro_text) {
    s.addText(lesson.evidence_collection.intro_text, { x: 0.55, y, w: PAGE_W - 1.1, h: 0.6, fontFace: "Arial", fontSize: 16, color: BODY, margin: 0, valign: "top" });
    y += 0.8;
  }
  const bulletText = lesson.evidence_collection.prompts.map(p => ({ text: p, options: { bullet: { code: "2022" }, breakLine: true } }));
  s.addText(bulletText, { x: 0.55, y, w: PAGE_W - 1.1, h: PAGE_H - y - 0.6, fontFace: "Arial", fontSize: 21.5, color: NAVY_INK, margin: 0, valign: "top", paraSpaceAfter: 10 });
  footer(s, pageNum++, false);
}

// ===== Slide: Whole-Class Discourse (singular or, for Discourse Deep
// Dive, multiple numbered prompts) =====
if (lesson.whole_class_discourse_prompt) {
  const s = pres.addSlide();
  slideTitle(s, "Whole-Class Discourse", false);
  s.addShape("roundRect", { x: 0.55, y: 1.5, w: PAGE_W - 1.1, h: 2.2, rectRadius: 0.167, fill: { color: "FDEEF2" }, line: { color: "ED6A91", width: 1.5, dashType: "dash" } });
  s.addText("Prompt", { x: 0.85, y: 1.65, w: PAGE_W - 1.7, h: 0.35, fontFace: "Arial", fontSize: 24, bold: true, color: PURPLE, margin: 0 });
  s.addText(lesson.whole_class_discourse_prompt, { x: 0.85, y: 2.05, w: PAGE_W - 1.7, h: 1.5, fontFace: "Arial", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });
  footer(s, pageNum++, false);
} else if (lesson.whole_class_discourse_prompts && lesson.whole_class_discourse_prompts.length) {
  const s = pres.addSlide();
  slideTitle(s, "Whole-Class Discourse", false);
  let y = 1.3;
  const boxH = Math.min(1.7, (PAGE_H - 1.8) / lesson.whole_class_discourse_prompts.length);
  lesson.whole_class_discourse_prompts.forEach((p, i) => {
    s.addShape("roundRect", { x: 0.55, y, w: PAGE_W - 1.1, h: boxH - 0.15, rectRadius: 0.1, fill: { color: "FDEEF2" }, line: { color: "ED6A91", width: 1.5, dashType: "dash" } });
    s.addText(`Prompt #${i + 1}`, { x: 0.75, y: y + 0.1, w: PAGE_W - 1.5, h: 0.3, fontFace: "Arial", fontSize: 16, bold: true, color: PURPLE, margin: 0 });
    s.addText(p, { x: 0.75, y: y + 0.42, w: PAGE_W - 1.5, h: boxH - 0.6, fontFace: "Arial", fontSize: 16, color: NAVY_INK, margin: 0, valign: "top" });
    y += boxH;
  });
  footer(s, pageNum++, false);
}

// ===== Slide: Closing =====
{
  const s = pres.addSlide();
  slideTitle(s, "Closing", false);
  const hasEQ = !!(lesson.essential_question && lesson.essential_question.trim());
  let tpTop = 1.15;

  if (hasEQ) {
    s.addShape("rect", { x: 0.55, y: 1.15, w: 0.06, h: 1.35, fill: { color: PURPLE }, line: { type: "none" } });
    s.addShape("roundRect", { x: 0.61, y: 1.15, w: PAGE_W - 1.2, h: 1.35, rectRadius: 0.139, fill: { color: CREAM_YELLOW }, line: { type: "none" } });
    s.addText("Essential Question", { x: 0.85, y: 1.3, w: PAGE_W - 1.6, h: 0.3, fontFace: "Arial", fontSize: 22, bold: true, color: PURPLE, margin: 0 });
    s.addText(lesson.essential_question, { x: 0.85, y: 1.65, w: PAGE_W - 1.6, h: 0.75, fontFace: "Arial", fontSize: 24, bold: true, color: NAVY_INK, margin: 0, valign: "top" });
    tpTop = 2.75;
  }

  s.addShape("rect", { x: 0.55, y: tpTop, w: PAGE_W - 1.1, h: 0.05, fill: { color: CORAL }, line: { type: "none" } });
  s.addShape("roundRect", { x: 0.55, y: tpTop + 0.05, w: PAGE_W - 1.1, h: PAGE_H - tpTop - 0.85, rectRadius: 0.113, fill: { color: WHITE }, line: { color: TP_BORDER_CLOSING, width: 1 } });
  s.addText("Teaching Point", { x: 0.75, y: tpTop + 0.25, w: PAGE_W - 1.5, h: 0.35, fontFace: "Arial", fontSize: 22, bold: true, color: CORAL, margin: 0 });
  s.addText(lesson.teaching_point, { x: 0.75, y: tpTop + 0.65, w: PAGE_W - 1.5, h: PAGE_H - tpTop - 1.5, fontFace: "Arial", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });

  footer(s, pageNum++, false);
}

// ===== Slides: Lesson Vocabulary Review =====
if (lesson.lesson_vocabulary_review && lesson.lesson_vocabulary_review.length) {
  chunk(lesson.lesson_vocabulary_review, 4).forEach((words, i) => {
    const s = pres.addSlide();
    s.background = { color: PURPLE };
    slideTitle(s, `Lesson Vocabulary Review${i > 0 ? " (continued)" : ""}`, true);
    vocabBlock(s, words, 0.55, 1.25, PAGE_W - 1.1, true);
    footer(s, pageNum++, true);
  });
}

pres.writeFile({ fileName: outPath }).then(() => {
  console.log(`Wrote ${outPath}`);
});
