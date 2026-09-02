// generate_deck_unit_launch.js -- Unit Launch lesson type.
//
// No dedicated visual .pptx template exists for Unit Launch (only Reading
// and Close Reading have one). This is built against the written
// "Knowledge Slide Template Outlines" spec doc's "Unit Launch Lesson (1)"
// section instead, reusing Reading's verified colors/chrome/helper
// functions wholesale (same approach generate_deck_language.js already
// takes for Language/Discourse Deep Dive -- "Same verified palette as
// generate_deck.js"). Slide-by-slide structure per that outline:
//   1. Cover + title
//   2. EQ/TP/Language Goal (blank here -- this lesson doesn't state one)
//   3-9. Engage/Launch/Background Knowledge: vocab + activity/chart slides
//   10. Shared Analysis
//   11. Quick Write
//   12. Inquiry: Unit Essential Questions
//   13. Discourse Clubs
//   14-15. Key Artifacts of Learning: Project / Writing
//   16. Closing
//   17. Lesson Vocabulary Review
// Verified against the real Lesson 1 ("Unit Launch") content in
// RT_Gr4U2_KNO_full_unit_Teacher_Guide.pdf (pages 55-65).
const pptxgen = require("pptxgenjs");
const fs = require("fs");

const lessonPath = process.argv[2] || "lesson_1_REAL.json";
const outPath = process.argv[3] || "output.pptx";
const lesson = JSON.parse(fs.readFileSync(lessonPath, "utf-8"));

// ===== Same verified palette as generate_deck.js =====
const NAVY_INK = "0E0142";
const PURPLE = "3928AA";
const CORAL = "F19A65";
const TITLE_PURPLE = "382DB0";
const CREAM_YELLOW = "FFF3CB";
const PEACH = "FFDBC5";
const MUTED = "6B6478";
const DETAIL_DESC_COLOR = "110045";
const TABLE_BORDER_BODY = "FFAB7B";
const TP_BORDER_ORANGE = "FFAB7B";
const TP_BORDER_CLOSING = "DCD6EE";
const LG_BORDER_LIGHT_PURPLE = "B8AEE0";
const DETAIL_BOX_BORDER = "FDEEF2";
const PINK_BORDER = "ED6A91";
const PINK_PALE = "FDEEF2";
const WHITE = "FFFFFF";
const BODY = "33322E";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
const PAGE_W = 13.3, PAGE_H = 7.5;
const TEMPLATE_SCALE = 4 / 3; // 13.3in canvas over the template's 10in -- see generate_deck.js

const HEADER_GRADIENT = require("path").join(__dirname, "assets", "header_gradient.png");
function addHeaderGradient(slide) {
  const scale = PAGE_W / 10;
  slide.addImage({ path: HEADER_GRADIENT, x: 0, y: 0, w: 10 * scale, h: 1.2578 * scale });
}
const DECORATIVE_BG = require("path").join(__dirname, "assets", "decorative_bg.png");
function addDecorativeBg(slide) {
  const scale = PAGE_W / 10;
  slide.addImage({ path: DECORATIVE_BG, x: 5.99 * scale, y: 0, w: 4.01 * scale, h: 3.46 * scale });
}

function unitHeader() {
  return `${lesson.grade}, Knowledge Unit ${lesson.unit_number} | ${lesson.unit_title}`;
}
function footer(slide, pageNum, onDark) {
  const c = onDark ? "C9BEEB" : MUTED;
  slide.addText(unitHeader(), { x: 0.5, y: PAGE_H - 0.4, w: PAGE_W - 1.5, h: 0.3, fontFace: "Arial", fontSize: 12, italic: true, color: c, align: "left", margin: 0 });
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
function parseInlineMarkup(text) {
  if (!text) return [{ text: text || "" }];
  // Same auto-underline safety net as generate_deck.js -- extraction
  // isn't reliably tagging the core text title with <u> tags, so detect
  // literal occurrences here rather than relying on it.
  if (lesson.core_text && text.includes(lesson.core_text) && !text.includes(`<u>${lesson.core_text}</u>`)) {
    text = text.split(lesson.core_text).join(`<u>${lesson.core_text}</u>`);
  }
  const parts = [];
  const underlineRe = /<u>(.*?)<\/u>/g;
  let lastIndex = 0, m;
  while ((m = underlineRe.exec(text)) !== null) {
    if (m.index > lastIndex) parts.push({ text: text.slice(lastIndex, m.index) });
    parts.push({ text: m[1], options: { underline: true } });
    lastIndex = m.index + m[0].length;
  }
  if (lastIndex < text.length) parts.push({ text: text.slice(lastIndex) });
  const pageNumRe = /\bpages?\s+\d+(?:[\u2013-]\d+)?\b/gi;
  const expanded = [];
  parts.forEach(part => {
    if (part.options && part.options.underline) { expanded.push(part); return; }
    let li = 0, mm, found = false;
    pageNumRe.lastIndex = 0;
    while ((mm = pageNumRe.exec(part.text)) !== null) {
      found = true;
      if (mm.index > li) expanded.push({ text: part.text.slice(li, mm.index) });
      expanded.push({ text: mm[0], options: { bold: true } });
      li = mm.index + mm[0].length;
    }
    if (found) { if (li < part.text.length) expanded.push({ text: part.text.slice(li) }); }
    else expanded.push(part);
  });
  return expanded.length ? expanded : [{ text }];
}

const VOCAB_ICON_DIR = process.argv[5] || require("path").join(__dirname, "vocab_icons");
function iconPathFor(word) {
  const safe = word.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  const p = require("path").join(VOCAB_ICON_DIR, `${safe}.png`);
  return fs.existsSync(p) ? p : null;
}
function estimateDefinitionLines(definition, colW) {
  const charsPerLine = (colW - 0.4) * 5.4;
  return Math.max(1, Math.ceil(definition.length / charsPerLine));
}
function chunkVocabForHeight(words, colW, maxHeight) {
  const groups = [];
  let current = [], heightSoFar = 0, rowsInCurrent = 0;
  for (let i = 0; i < words.length; i += 2) {
    const rowWords = words.slice(i, i + 2);
    const rowH = Math.max(2.7, ...rowWords.map(item => 1.0 + estimateDefinitionLines(item.definition, colW) * 0.34 + 0.25));
    const gap = rowsInCurrent > 0 ? 0.15 : 0;
    if (rowsInCurrent > 0 && (rowsInCurrent >= 2 || heightSoFar + gap + rowH > maxHeight)) {
      groups.push(current);
      current = []; heightSoFar = 0; rowsInCurrent = 0;
    }
    current = current.concat(rowWords);
    heightSoFar += (rowsInCurrent > 0 ? 0.15 : 0) + rowH;
    rowsInCurrent += 1;
  }
  if (current.length) groups.push(current);
  return groups;
}
function vocabBlock(slide, words, x, y, w, onDark) {
  const colW = (w - 0.3) / 2;
  const cardBg = onDark ? WHITE : PEACH;
  const rows = chunk(words, 2);
  let cursorY = y;
  rows.forEach(rowWords => {
    const rowH = Math.max(2.7, ...rowWords.map(item => 1.0 + estimateDefinitionLines(item.definition, colW) * 0.34 + 0.25));
    rowWords.forEach((item, ci) => {
      const bx = x + ci * (colW + 0.3), by = cursorY;
      slide.addShape("roundRect", { x: bx, y: by, w: colW, h: rowH, rectRadius: 0.20, fill: { color: cardBg }, line: onDark ? { color: CORAL, width: 1 } : { type: "none" } });
      const iconPath = iconPathFor(item.word);
      if (iconPath) slide.addImage({ path: iconPath, x: bx + 0.2, y: by + 0.2, w: 0.55, h: 0.55 });
      slide.addText(parseInlineMarkup(item.word), { x: bx + 0.9, y: by + 0.22, w: colW - 1.05, h: 0.5, fontFace: "Arial", fontSize: 24, bold: true, color: DETAIL_DESC_COLOR, margin: 0 });
      slide.addText(parseInlineMarkup(item.definition), { x: bx + 0.2, y: by + 1.0, w: colW - 0.4, h: rowH - 1.15, fontFace: "Arial", fontSize: 21.5, color: DETAIL_DESC_COLOR, margin: 0, valign: "top" });
    });
    cursorY += rowH + 0.15;
  });
}

function estimateDescLines(text, w) {
  const charsPerLine = w * 7.0;
  return Math.max(1, Math.ceil((text || "").length / charsPerLine));
}
function detailPromptSlide(s, opts) {
  const { cardColor, accentColor, iconPath, cardTitle, description, promptLabel, promptText, boxBorderColor } = opts;
  const cardX = 0.521 * TEMPLATE_SCALE, cardY = 1.458 * TEMPLATE_SCALE, cardW = 2.292 * TEMPLATE_SCALE;
  const textW = cardW - 0.417 * TEMPLATE_SCALE;
  const descLines = estimateDescLines(description, textW);
  const descTextOffset = 1.7 * TEMPLATE_SCALE;
  const cardH = Math.max(2.917 * TEMPLATE_SCALE, descTextOffset + descLines * 0.32 + 0.15);
  s.addShape("roundRect", { x: cardX, y: cardY, w: cardW, h: cardH, rectRadius: 0.167 * TEMPLATE_SCALE, fill: { color: cardColor }, line: { type: "none" } });
  if (iconPath) {
    const iconW = (opts.iconW || 0.65) * TEMPLATE_SCALE, iconH = (opts.iconH || 0.65) * TEMPLATE_SCALE;
    s.addImage({ path: iconPath, x: cardX + 0.208 * TEMPLATE_SCALE, y: cardY + 0.270 * TEMPLATE_SCALE, w: iconW, h: iconH });
  }
  const textX = cardX + 0.208 * TEMPLATE_SCALE;
  s.addText(cardTitle, { x: textX, y: cardY + 1.150 * TEMPLATE_SCALE, w: textW, h: 0.5 * TEMPLATE_SCALE, fontFace: "Arial", fontSize: 18.5, bold: true, color: PURPLE, margin: 0, valign: "top" });
  s.addText(description, { x: textX, y: cardY + 1.150 * TEMPLATE_SCALE + 0.55 * TEMPLATE_SCALE, w: textW, h: cardH - (1.150 * TEMPLATE_SCALE + 0.55 * TEMPLATE_SCALE), fontFace: "Arial", fontSize: 18.5, color: DETAIL_DESC_COLOR, margin: 0, valign: "top" });
  const boxX = 3.437 * TEMPLATE_SCALE, boxW = 5.729 * TEMPLATE_SCALE;
  s.addShape("roundRect", { x: boxX - 0.312 * TEMPLATE_SCALE, y: cardY, w: 0.08 * TEMPLATE_SCALE, h: cardH, rectRadius: 0.04 * TEMPLATE_SCALE, fill: { color: accentColor }, line: { type: "none" } });
  s.addShape("roundRect", { x: boxX - 0.312 * TEMPLATE_SCALE + 0.08 * TEMPLATE_SCALE, y: cardY, w: boxW + 0.312 * TEMPLATE_SCALE - 0.08 * TEMPLATE_SCALE, h: cardH, rectRadius: 0.167 * TEMPLATE_SCALE, fill: { color: WHITE }, line: { color: boxBorderColor || DETAIL_BOX_BORDER, width: 1 } });
  s.addText(promptLabel, { x: boxX, y: cardY + 0.15 * TEMPLATE_SCALE, w: boxW, h: 0.3 * TEMPLATE_SCALE, fontFace: "Arial", fontSize: 24, bold: true, color: accentColor, margin: 0 });
  s.addText(parseInlineMarkup(promptText), { x: boxX, y: cardY + 0.15 * TEMPLATE_SCALE + 0.45 * TEMPLATE_SCALE, w: boxW, h: cardH - (0.15 * TEMPLATE_SCALE + 0.45 * TEMPLATE_SCALE), fontFace: "Arial", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });
}

let pageNum = 2;

// ===== Slide: Cover =====
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
    s.addText(parseInlineMarkup(`${byLine}${lesson.pages ? "   |   Pages " + lesson.pages : ""}`), { x: 0.55, y: 4.65, w: 12.2, h: 0.6, fontFace: "Arial", fontSize: 21.5, italic: true, color: "C9BEEB", margin: 0, valign: "top" });
  }
  s.addText("Copyright \u00a9 2026 Lavinia Group. All Rights Reserved. RedThread is a trademark of K12 Coalition.", { x: 0.55, y: PAGE_H - 0.5, w: 11.5, h: 0.3, fontFace: "Arial", fontSize: 10, color: "9A8FD1", margin: 0 });
}

// ===== Slide: EQ / TP / LG -- blank per spec for this lesson (Unit Launch
// doesn't state a single lesson-level EQ/TP the way Reading does; its
// Essential Questions get introduced later, in the dedicated Inquiry
// slide) =====
{
  const s = pres.addSlide();
  addHeaderGradient(s);
  const colGap = 0.2;
  const colW = (PAGE_W - 1.1 - colGap) / 2;
  const rowTop = 2.15, rowBottom = PAGE_H - 0.8;
  const rowH = rowBottom - rowTop;
  s.addShape("roundRect", { x: 0.55, y: 0.55, w: 0.06, h: 1.35, rectRadius: 0.03, fill: { color: PURPLE }, line: { type: "none" } });
  s.addShape("roundRect", { x: 0.61, y: 0.55, w: PAGE_W - 1.2, h: 1.35, rectRadius: 0.139, fill: { color: CREAM_YELLOW }, line: { type: "none" } });
  s.addText("Essential Question", { x: 0.85, y: 0.7, w: PAGE_W - 1.6, h: 0.3, fontFace: "Arial", fontSize: 18.5, bold: true, color: PURPLE, margin: 0 });
  if (lesson.essential_question) {
    s.addText(parseInlineMarkup(lesson.essential_question), { x: 0.85, y: 1.05, w: PAGE_W - 1.6, h: 0.8, fontFace: "Arial", fontSize: 24, bold: true, color: NAVY_INK, margin: 0, valign: "top" });
  } else {
    s.addText("Introduced in the Inquiry section of this lesson (see the unit's full set of Essential Questions).", { x: 0.85, y: 1.05, w: PAGE_W - 1.6, h: 0.8, fontFace: "Arial", fontSize: 16, italic: true, color: MUTED, margin: 0, valign: "top" });
  }
  s.addShape("roundRect", { x: 0.55, y: rowTop, w: colW, h: 0.05, rectRadius: 0.025, fill: { color: CORAL }, line: { type: "none" } });
  s.addShape("roundRect", { x: 0.55, y: rowTop + 0.05, w: colW, h: rowH - 0.05, rectRadius: 0.139, fill: { color: WHITE }, line: { color: TP_BORDER_ORANGE, width: 1.5 } });
  s.addText("Teaching Point", { x: 0.75, y: rowTop + 0.25, w: colW - 0.4, h: 0.3, fontFace: "Arial", fontSize: 18.5, bold: true, color: CORAL, margin: 0 });
  if (lesson.teaching_point) {
    s.addText(parseInlineMarkup(lesson.teaching_point), { x: 0.75, y: rowTop + 0.6, w: colW - 0.4, h: rowH - 0.8, fontFace: "Arial", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });
  } else {
    s.addText("Not specified in this lesson's source material.", { x: 0.75, y: rowTop + 0.6, w: colW - 0.4, h: rowH - 0.8, fontFace: "Arial", fontSize: 16, italic: true, color: MUTED, margin: 0, valign: "top" });
  }
  const lgX = 0.55 + colW + colGap;
  s.addShape("roundRect", { x: lgX, y: rowTop, w: colW, h: 0.05, rectRadius: 0.025, fill: { color: PURPLE }, line: { type: "none" } });
  s.addShape("roundRect", { x: lgX, y: rowTop + 0.05, w: colW, h: rowH - 0.05, rectRadius: 0.139, fill: { color: "F4F2FC" }, line: { color: LG_BORDER_LIGHT_PURPLE, width: 1, dashType: "dash" } });
  s.addText("Language Goal", { x: lgX + 0.2, y: rowTop + 0.25, w: colW - 0.4, h: 0.3, fontFace: "Arial", fontSize: 18.5, bold: true, color: PURPLE, margin: 0 });
  if (lesson.language_goal) {
    s.addText(parseInlineMarkup(lesson.language_goal), { x: lgX + 0.2, y: rowTop + 0.6, w: colW - 0.4, h: rowH - 0.8, fontFace: "Arial", fontSize: 24, italic: true, color: NAVY_INK, margin: 0, valign: "top" });
  } else {
    s.addText("Not specified in this lesson's source material.", { x: lgX + 0.2, y: rowTop + 0.6, w: colW - 0.4, h: rowH - 0.8, fontFace: "Arial", fontSize: 16, italic: true, color: MUTED, margin: 0, valign: "top" });
  }
  footer(s, pageNum++, false);
}

// ===== Slides: Engage / Launch / Background Knowledge sections =====
(lesson.sections || []).forEach(section => {
  // Activity text (Unit Launch-specific: a live classroom activity with no
  // chart or read_directions, e.g. the Region 1/Region 2 simulation) --
  // rendered as a plain directions-style slide, same visual treatment as
  // a Reading section's read_directions slide.
  if (section.activity_text) {
    const s = pres.addSlide();
    addHeaderGradient(s);
    slideTitle(s, `${section.section_name} Activity`, false);
    s.addText(parseInlineMarkup(section.activity_text), { x: 0.55, y: 1.15, w: PAGE_W - 1.1, h: PAGE_H - 1.75, fontFace: "Arial", fontSize: 16, color: DETAIL_DESC_COLOR, margin: 0, valign: "top" });
    footer(s, pageNum++, false);
  }
  if (section.vocabulary && section.vocabulary.length) {
    const vocabColW = ((PAGE_W - 1.1) - 0.3) / 2;
    chunkVocabForHeight(section.vocabulary, vocabColW, 5.65).forEach((words, i) => {
      const s = pres.addSlide();
      addHeaderGradient(s);
      slideTitle(s, `${section.section_name} Vocabulary${i > 0 ? " (continued)" : ""}`, false);
      vocabBlock(s, words, 0.55, 1.25, PAGE_W - 1.1, false);
      footer(s, pageNum++, false);
    });
  }
  if (section.chart) {
    const s = pres.addSlide();
    addHeaderGradient(s);
    slideTitle(s, `${section.section_name} Chart`, false);
    let chartY = 1.15;
    if (section.read_directions) {
      s.addText(parseInlineMarkup(section.read_directions), { x: 0.55, y: chartY, w: PAGE_W - 1.1, h: 0.7, fontFace: "Arial", fontSize: 21.5, color: DETAIL_DESC_COLOR, margin: 0, valign: "top" });
      chartY += 0.75;
    }
    footer(s, pageNum++, false);
  }
});

// ===== Slide: Shared Analysis (Teaching Point + directions + vocabulary,
// same left-column layout as Reading's Independent Reading/Shared
// Analysis slide) =====
if (lesson.shared_analysis) {
  const block = lesson.shared_analysis;
  const s = pres.addSlide();
  addHeaderGradient(s);
  slideTitle(s, "Shared Analysis", false);
  const leftW = 6.6;
  let y = 1.1;
  if (lesson.teaching_point) {
    s.addText("TEACHING POINT", { x: 0.55, y, w: leftW, h: 0.35, fontFace: "Arial", fontSize: 20, bold: true, color: CORAL, margin: 0 });
    s.addText(parseInlineMarkup(lesson.teaching_point), { x: 0.55, y: y + 0.4, w: leftW, h: 1.15, fontFace: "Arial", fontSize: 16, color: NAVY_INK, margin: 0, valign: "top" });
    y += 1.65;
  }
  if (block.read_directions) {
    s.addShape("roundRect", { x: 0.55, y, w: leftW, h: 1.0, rectRadius: 0.06, fill: { color: "FBE8AB" }, line: { type: "none" } });
    s.addText(parseInlineMarkup(block.read_directions), { x: 0.75, y: y + 0.1, w: leftW - 0.4, h: 0.8, fontFace: "Arial", fontSize: 16, color: NAVY_INK, margin: 0, valign: "top" });
    y += 1.15;
  }
  if (block.vocabulary && block.vocabulary.length) {
    s.addText("IMPORTANT VOCABULARY", { x: 0.55, y, w: leftW, h: 0.35, fontFace: "Arial", fontSize: 20, bold: true, color: CORAL, margin: 0 });
    y += 0.42;
    block.vocabulary.forEach(v => {
      s.addShape("ellipse", { x: 0.6, y: y + 0.1, w: 0.11, h: 0.11, fill: { color: "9343F6" }, line: { type: "none" } });
      s.addText([{ text: `${v.word}: `, options: { bold: true, color: NAVY_INK } }, { text: v.definition, options: { color: NAVY_INK } }],
        { x: 0.85, y, w: leftW - 0.3, h: 0.65, fontFace: "Arial", fontSize: 16, margin: 0, valign: "top" });
      const itemLines = Math.max(1, Math.ceil((`${v.word}: ${v.definition}`).length / ((leftW - 0.3) * 9.6)));
      y += Math.max(0.66, 0.18 + itemLines * 0.24);
    });
  }
  footer(s, pageNum++, false);
}

// ===== Slide: Quick Write =====
if (lesson.quick_write_prompt) {
  const s = pres.addSlide();
  addHeaderGradient(s);
  addDecorativeBg(s);
  slideTitle(s, "Quick Write", false);
  detailPromptSlide(s, {
    cardColor: "FFAB7B", accentColor: PINK_BORDER,
    iconPath: require("path").join(__dirname, "assets", "detail_icons", "quick_write.png"),
    iconW: 1.005, iconH: 0.600,
    cardTitle: "Quick Write",
    description: "Take a few minutes to write down your initial thoughts and response to the prompt.",
    promptLabel: "Prompt", promptText: lesson.quick_write_prompt,
  });
  footer(s, pageNum++, false);
}

// ===== Slide: Inquiry: Unit Essential Questions -- Unit Launch-specific,
// no equivalent in Reading. Two labeled groups (content-based,
// reading/writing-based), matching how the lesson's own "Inquiry" section
// presents them. =====
if (lesson.inquiry_essential_questions) {
  const eqs = lesson.inquiry_essential_questions;
  const s = pres.addSlide();
  addHeaderGradient(s);
  slideTitle(s, "Inquiry: Unit Essential Questions", false);
  let y = 1.15;
  const renderGroup = (label, questions) => {
    if (!questions || !questions.length) return;
    s.addText(label, { x: 0.55, y, w: PAGE_W - 1.1, h: 0.35, fontFace: "Arial", fontSize: 18.5, bold: true, color: CORAL, margin: 0 });
    y += 0.45;
    questions.forEach((q, i) => {
      s.addText(`${i + 1}. ${q}`, { x: 0.75, y, w: PAGE_W - 1.5, h: 0.5, fontFace: "Arial", fontSize: 17, color: NAVY_INK, margin: 0, valign: "top" });
      y += 0.52;
    });
    y += 0.2;
  };
  renderGroup("CONTENT-BASED ESSENTIAL QUESTIONS", eqs.content_based);
  renderGroup("READING- AND WRITING-BASED ESSENTIAL QUESTIONS", eqs.reading_writing_based);
  footer(s, pageNum++, false);
}

// ===== Slide: Discourse Clubs =====
if (lesson.discourse_club_prompt) {
  const s = pres.addSlide();
  addHeaderGradient(s);
  addDecorativeBg(s);
  slideTitle(s, "Discourse Clubs", false);
  detailPromptSlide(s, {
    cardColor: PINK_PALE, accentColor: CORAL,
    boxBorderColor: TP_BORDER_ORANGE,
    iconPath: require("path").join(__dirname, "assets", "detail_icons", "discourse_clubs.png"),
    iconW: 0.733, iconH: 0.707,
    cardTitle: "Discourse Clubs",
    description: "Choose the Essential Question that most interests you today, and discuss it with your peers.",
    promptLabel: "Prompt", promptText: lesson.discourse_club_prompt,
  });
  footer(s, pageNum++, false);
}

// ===== Slides: Key Artifacts of Learning: Project / Writing -- both
// flagged resource_unavailable per lesson_1_REAL.json (the actual project
// and writing prompts aren't reprinted in Lesson 1 itself; they reference
// the unit's Student Outcomes section, which is rubrics/benchmarks and
// mentor examples, not the prompts) =====
if (lesson.key_artifacts_resource_unavailable) {
  ["Project", "Writing"].forEach(kind => {
    const s = pres.addSlide();
    addHeaderGradient(s);
    slideTitle(s, `Key Artifacts of Learning: ${kind}`, false);
    s.addShape("roundRect", { x: 0.55, y: 1.5, w: PAGE_W - 1.1, h: 1.4, rectRadius: 0.08, fill: { color: "FEF3C7" }, line: { color: "D97706", width: 1 } });
    s.addText([{ text: "\u26a0 Needs manual follow-up: " }, ...parseInlineMarkup(lesson.key_artifacts_resource_unavailable)], { x: 0.8, y: 1.65, w: PAGE_W - 1.6, h: 1.1, fontFace: "Arial", fontSize: 16, italic: true, color: "92400E", margin: 0, valign: "top" });
    footer(s, pageNum++, false);
  });
}

// ===== Slide: Closing =====
{
  const s = pres.addSlide();
  addHeaderGradient(s);
  slideTitle(s, "Closing", false);
  s.addShape("roundRect", { x: 0.55, y: 1.15, w: 0.06, h: 1.35, rectRadius: 0.03, fill: { color: PURPLE }, line: { type: "none" } });
  s.addShape("roundRect", { x: 0.61, y: 1.15, w: PAGE_W - 1.2, h: 1.35, rectRadius: 0.139, fill: { color: CREAM_YELLOW }, line: { type: "none" } });
  s.addText("Essential Question", { x: 0.85, y: 1.3, w: PAGE_W - 1.6, h: 0.3, fontFace: "Arial", fontSize: 22, bold: true, color: PURPLE, margin: 0 });
  if (lesson.essential_question) {
    s.addText(parseInlineMarkup(lesson.essential_question), { x: 0.85, y: 1.65, w: PAGE_W - 1.6, h: 0.75, fontFace: "Arial", fontSize: 24, bold: true, color: NAVY_INK, margin: 0, valign: "top" });
  } else {
    s.addText("Not specified in this lesson's source material.", { x: 0.85, y: 1.65, w: PAGE_W - 1.6, h: 0.75, fontFace: "Arial", fontSize: 16, italic: true, color: MUTED, margin: 0, valign: "top" });
  }
  s.addShape("roundRect", { x: 0.55, y: 2.75, w: PAGE_W - 1.1, h: 0.05, rectRadius: 0.025, fill: { color: CORAL }, line: { type: "none" } });
  s.addShape("roundRect", { x: 0.55, y: 2.8, w: PAGE_W - 1.1, h: 2.4, rectRadius: 0.113, fill: { color: WHITE }, line: { color: TP_BORDER_CLOSING, width: 1 } });
  s.addText("Teaching Point", { x: 0.75, y: 3.0, w: PAGE_W - 1.5, h: 0.35, fontFace: "Arial", fontSize: 22, bold: true, color: CORAL, margin: 0 });
  if (lesson.teaching_point) {
    s.addText(parseInlineMarkup(lesson.teaching_point), { x: 0.75, y: 3.4, w: PAGE_W - 1.5, h: 1.7, fontFace: "Arial", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });
  } else {
    s.addText("Not specified in this lesson's source material.", { x: 0.75, y: 3.4, w: PAGE_W - 1.5, h: 1.7, fontFace: "Arial", fontSize: 16, italic: true, color: MUTED, margin: 0, valign: "top" });
  }
  footer(s, pageNum++, false);
}

// ===== Slides: Lesson Vocabulary Review =====
if (lesson.lesson_vocabulary_review && lesson.lesson_vocabulary_review.length) {
  const reviewColW = ((PAGE_W - 1.1) - 0.3) / 2;
  chunkVocabForHeight(lesson.lesson_vocabulary_review, reviewColW, 5.65).forEach((words, i) => {
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
