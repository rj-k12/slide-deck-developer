const pptxgen = require("pptxgenjs");
const fs = require("fs");

const lessonPath = process.argv[2] || "lesson_6b_REAL.json";
const outPath = process.argv[3] || "output.pptx";
const lesson = JSON.parse(fs.readFileSync(lessonPath, "utf-8"));

// Same verified palette as generate_deck.js, pulled directly from
// KNO_Reading_TEMPLATE_8_7.pptx's XML.
const NAVY_INK = "0E0142";
const PURPLE = "3928AA";
const BLUE_PALE = "E5EFF9";
const CORAL = "F19A65";
const TITLE_PURPLE = "382DB0";
const CREAM_YELLOW = "FFF3CB";
const PEACH = "FFDBC5";
const MUTED = "6B6478";
const VIOLET = "9343F6";
const TABLE_BORDER = "DCD6EE";
const PINK_BORDER = "ED6A91";
const PINK_PALE = "FDEEF2";
const AI_ICON_BLUE = "5B9BD5";
const WHITE = "FFFFFF";
const BODY = "33322E";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
const PAGE_W = 13.3, PAGE_H = 7.5;

function unitHeader() {
  return `${lesson.grade}, Knowledge Unit ${lesson.unit_number} | ${lesson.unit_title}`;
}
function footer(slide, pageNum, onDark) {
  const c = onDark ? "C9BEEB" : MUTED;
  slide.addText(unitHeader(), { x: 0.5, y: PAGE_H - 0.4, w: PAGE_W - 1.5, h: 0.3, fontFace: "Figtree", fontSize: 12, italic: true, color: c, align: "left", margin: 0 });
  slide.addText(String(pageNum), { x: PAGE_W - 0.9, y: PAGE_H - 0.4, w: 0.4, h: 0.3, fontFace: "Figtree", fontSize: 12, color: c, align: "right", margin: 0 });
}
function slideTitle(slide, title, onDark) {
  slide.addText(title, { x: 0.55, y: 0.35, w: PAGE_W - 1.2, h: 0.6, fontFace: "Figtree", fontSize: 36, bold: true, color: onDark ? WHITE : TITLE_PURPLE, margin: 0 });
}
function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
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
    slide.addShape("roundRect", { x: bx, y: by, w: colW, h: 2.7, rectRadius: 0.06, fill: { color: cardBg }, line: onDark ? { color: CORAL, width: 1 } : { type: "none" } });
    const iconPath = iconPathFor(item.word);
    if (iconPath) {
      slide.addImage({ path: iconPath, x: bx + 0.2, y: by + 0.2, w: 0.55, h: 0.55 });
    }
    slide.addText(item.word, { x: bx + 0.9, y: by + 0.22, w: colW - 1.05, h: 0.5, fontFace: "Figtree", fontSize: 24, bold: true, color: NAVY_INK, margin: 0 });
    slide.addText(item.definition, { x: bx + 0.2, y: by + 1.0, w: colW - 0.4, h: 1.55, fontFace: "Figtree", fontSize: 21.5, color: BODY, margin: 0, valign: "top" });
  });
}
function promptBox(slide, prompt, x, y, w, h) {
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.06, fill: { color: PINK_PALE }, line: { color: PINK_BORDER, width: 1.5, dashType: "dash" } });
  slide.addText([{ text: "Prompt: ", options: { bold: true, color: PURPLE } }, { text: prompt, options: { color: NAVY_INK } }],
    { x: x + 0.3, y: y + 0.15, w: w - 0.6, h: h - 0.3, fontFace: "Figtree", fontSize: 24, margin: 0, valign: "top" });
}

// ===== Slide: Cover (page 1 of the real template) =====
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
  s.addText(`LESSON ${lesson.lesson_number}`, { x: 0.55, y: 0.6, w: 1.9, h: 0.55, fontFace: "Figtree", fontSize: 16, bold: true, color: NAVY_INK, align: "center", valign: "middle", margin: 0 });
  s.addText(`Lesson ${lesson.lesson_number}: ${lesson.lesson_type}`, { x: 0.55, y: 1.3, w: 12.2, h: 1.8, fontFace: "Figtree", fontSize: 50, bold: true, color: WHITE, margin: 0, valign: "top" });
  s.addText(`Grade ${lesson.grade.replace('Grade ','')}, Knowledge Unit ${lesson.unit_number}: ${lesson.unit_title}`, { x: 0.55, y: 3.35, w: 12.2, h: 1.1, fontFace: "Figtree", fontSize: 25.5, color: "E5EFF9", margin: 0, valign: "top" });
  if (lesson.core_text) {
    const byLine = lesson.author ? `${lesson.core_text}  by ${lesson.author}` : lesson.core_text;
    s.addText(byLine, { x: 0.55, y: 4.65, w: 12.2, h: 0.6, fontFace: "Figtree", fontSize: 21.5, italic: true, color: "C9BEEB", margin: 0, valign: "top" });
  }
  s.addText("Copyright \u00a9 2026 Lavinia Group. All Rights Reserved. RedThread is a trademark of K12 Coalition.", { x: 0.55, y: PAGE_H - 0.5, w: 11.5, h: 0.3, fontFace: "Figtree", fontSize: 10, color: "9A8FD1", margin: 0 });
}

// ===== Slide: EQ / TP / LG -- blank, real 3-box structure. EQ full-width
// on top, Teaching Point and Language Goal SIDE BY SIDE below (confirmed
// against the template's own shape positions -- the earlier stacked
// version left significant dead space at the bottom of the slide) =====
{
  const s = pres.addSlide();
  const colGap = 0.2;
  const colW = (PAGE_W - 1.1 - colGap) / 2;
  const rowTop = 2.15, rowBottom = PAGE_H - 0.8;
  const rowH = rowBottom - rowTop;

  s.addShape("rect", { x: 0.55, y: 0.55, w: 0.06, h: 1.35, fill: { color: PURPLE }, line: { type: "none" } });
  s.addShape("rect", { x: 0.61, y: 0.55, w: PAGE_W - 1.2, h: 1.35, fill: { color: CREAM_YELLOW }, line: { type: "none" } });
  s.addText("ESSENTIAL QUESTION", { x: 0.85, y: 0.7, w: PAGE_W - 1.6, h: 0.3, fontFace: "Figtree", fontSize: 18.5, bold: true, color: PURPLE, margin: 0 });
  if (lesson.essential_question) {
    s.addText(lesson.essential_question, { x: 0.85, y: 1.05, w: PAGE_W - 1.6, h: 0.8, fontFace: "Figtree", fontSize: 24, bold: true, color: NAVY_INK, margin: 0, valign: "top" });
  }

  s.addShape("rect", { x: 0.55, y: rowTop, w: colW, h: 0.05, fill: { color: CORAL }, line: { type: "none" } });
  s.addShape("rect", { x: 0.55, y: rowTop + 0.05, w: colW, h: rowH - 0.05, fill: { color: WHITE }, line: { color: BLUE_PALE, width: 1 } });
  s.addText("TEACHING POINT", { x: 0.75, y: rowTop + 0.25, w: colW - 0.4, h: 0.3, fontFace: "Figtree", fontSize: 18.5, bold: true, color: CORAL, margin: 0 });
  if (lesson.teaching_point) {
    s.addText(lesson.teaching_point, { x: 0.75, y: rowTop + 0.6, w: colW - 0.4, h: rowH - 0.8, fontFace: "Figtree", fontSize: 24, color: BODY, margin: 0, valign: "top" });
  }

  const lgX = 0.55 + colW + colGap;
  s.addShape("rect", { x: lgX, y: rowTop, w: colW, h: 0.05, fill: { color: PURPLE }, line: { type: "none" } });
  s.addShape("rect", { x: lgX, y: rowTop + 0.05, w: colW, h: rowH - 0.05, fill: { color: "F4F2FC" }, line: { color: PURPLE, width: 1, dashType: "dash" } });
  s.addText("LANGUAGE GOAL", { x: lgX + 0.2, y: rowTop + 0.25, w: colW - 0.4, h: 0.3, fontFace: "Figtree", fontSize: 18.5, bold: true, color: PURPLE, margin: 0 });
  if (lesson.language_goal) {
    s.addText(lesson.language_goal, { x: lgX + 0.2, y: rowTop + 0.6, w: colW - 0.4, h: rowH - 0.8, fontFace: "Figtree", fontSize: 24, italic: true, color: NAVY_INK, margin: 0, valign: "top" });
  } else {
    s.addText("Not specified in this lesson's source material.", { x: lgX + 0.2, y: rowTop + 0.6, w: colW - 0.4, h: rowH - 0.8, fontFace: "Figtree", fontSize: 16, italic: true, color: MUTED, margin: 0, valign: "top" });
  }

  footer(s, 2, false);
}

// ===== Slides: Engage / Launch sections =====
const DECORATIVE_BG = require("path").join(__dirname, "assets", "decorative_bg.png");
function addDecorativeBg(slide) {
  const scale = PAGE_W / 10;
  slide.addImage({ path: DECORATIVE_BG, x: 5.99 * scale, y: 0, w: 4.01 * scale, h: 3.46 * scale });
}

let pageNum = 3;
(lesson.sections || []).forEach(section => {
  if (section.vocabulary && section.vocabulary.length) {
    chunk(section.vocabulary, 4).forEach((words, i) => {
      const s = pres.addSlide();
      if (section.section_name === "Launch") addDecorativeBg(s);
      slideTitle(s, `${section.section_name} Vocabulary${i > 0 ? " (continued)" : ""}`, false);
      vocabBlock(s, words, 0.55, 1.25, PAGE_W - 1.1, false);
      footer(s, pageNum++, false);
    });
  }
  if (section.mentor_prompt) {
    const s = pres.addSlide();
    slideTitle(s, `${section.section_name} Mentor Prompt`, false);
    promptBox(s, section.mentor_prompt, 0.55, 1.5, PAGE_W - 1.1, 2.2);
    footer(s, pageNum++, false);
  }
  if (section.claims_to_evaluate && section.claims_to_evaluate.length) {
    const s = pres.addSlide();
    slideTitle(s, `${section.section_name} Claims to Evaluate`, false);
    let y = 1.4;
    section.claims_to_evaluate.forEach(claim => {
      s.addShape("roundRect", { x: 0.55, y, w: PAGE_W - 1.1, h: 1.3, rectRadius: 0.06, fill: { color: PEACH }, line: { type: "none" } });
      s.addText(claim, { x: 0.8, y: y + 0.12, w: PAGE_W - 1.6, h: 1.05, fontFace: "Figtree", fontSize: 21.5, color: NAVY_INK, margin: 0, valign: "top" });
      y += 1.5;
    });
    footer(s, pageNum++, false);
  }
  if (section.resource_unavailable) {
    const s = pres.addSlide();
    slideTitle(s, `${section.section_name} Resource`, false);
    s.addShape("roundRect", { x: 0.55, y: 1.5, w: PAGE_W - 1.1, h: 1.1, rectRadius: 0.08, fill: { color: "FEF3C7" }, line: { color: "D97706", width: 1 } });
    s.addText(`\u26a0 Needs manual follow-up: ${section.resource_unavailable}`, { x: 0.8, y: 1.65, w: PAGE_W - 1.6, h: 0.8, fontFace: "Figtree", fontSize: 16, italic: true, color: "92400E", margin: 0, valign: "top" });
    footer(s, pageNum++, false);
  }
});

// ===== Slide: Literature Response =====
if (lesson.literature_response_prompt) {
  const s = pres.addSlide();
  slideTitle(s, "Literature Response", false);
  s.addText("TEACHING POINT", { x: 0.55, y: 1.3, w: PAGE_W - 1.1, h: 0.35, fontFace: "Figtree", fontSize: 22, bold: true, color: CORAL, margin: 0 });
  s.addText(lesson.teaching_point, { x: 0.55, y: 1.7, w: PAGE_W - 1.1, h: 1.1, fontFace: "Figtree", fontSize: 21.5, color: BODY, margin: 0, valign: "top" });
  promptBox(s, lesson.literature_response_prompt, 0.55, 3.0, PAGE_W - 1.1, 1.4);
  footer(s, pageNum++, false);
}

// ===== Slide: Writers' Circle & Revise =====
if (lesson.writers_circle) {
  const s = pres.addSlide();
  slideTitle(s, "Writers' Circle & Revise", false);
  s.addText("TEACHING POINT", { x: 0.55, y: 1.3, w: PAGE_W - 1.1, h: 0.35, fontFace: "Figtree", fontSize: 22, bold: true, color: CORAL, margin: 0 });
  s.addText(lesson.teaching_point, { x: 0.55, y: 1.7, w: PAGE_W - 1.1, h: 1.1, fontFace: "Figtree", fontSize: 21.5, color: BODY, margin: 0, valign: "top" });
  if (lesson.writers_circle.focus_points && lesson.writers_circle.focus_points.length) {
    s.addText("FOCUS FOR FEEDBACK", { x: 0.55, y: 3.0, w: PAGE_W - 1.1, h: 0.35, fontFace: "Figtree", fontSize: 22, bold: true, color: CORAL, margin: 0 });
    let y = 3.45;
    lesson.writers_circle.focus_points.forEach(point => {
      s.addShape("ellipse", { x: 0.6, y: y + 0.1, w: 0.11, h: 0.11, fill: { color: VIOLET }, line: { type: "none" } });
      s.addText(point, { x: 0.85, y, w: PAGE_W - 1.4, h: 0.85, fontFace: "Figtree", fontSize: 21.5, color: BODY, margin: 0, valign: "top" });
      y += 0.9;
    });
  }
  footer(s, pageNum++, false);
}

// ===== Slide: Closing =====
{
  const s = pres.addSlide();
  slideTitle(s, "Closing", false);
  s.addShape("rect", { x: 0.55, y: 1.15, w: 0.06, h: 1.35, fill: { color: PURPLE }, line: { type: "none" } });
  s.addShape("rect", { x: 0.61, y: 1.15, w: PAGE_W - 1.2, h: 1.35, fill: { color: CREAM_YELLOW }, line: { type: "none" } });
  s.addText("ESSENTIAL QUESTION", { x: 0.85, y: 1.3, w: PAGE_W - 1.6, h: 0.3, fontFace: "Figtree", fontSize: 22, bold: true, color: PURPLE, margin: 0 });
  s.addText(lesson.essential_question, { x: 0.85, y: 1.65, w: PAGE_W - 1.6, h: 0.75, fontFace: "Figtree", fontSize: 24, bold: true, color: NAVY_INK, margin: 0, valign: "top" });

  s.addShape("rect", { x: 0.55, y: 2.75, w: PAGE_W - 1.1, h: 0.05, fill: { color: CORAL }, line: { type: "none" } });
  s.addShape("rect", { x: 0.55, y: 2.8, w: PAGE_W - 1.1, h: 2.4, fill: { color: WHITE }, line: { color: BLUE_PALE, width: 1 } });
  s.addText("TEACHING POINT", { x: 0.75, y: 3.0, w: PAGE_W - 1.5, h: 0.35, fontFace: "Figtree", fontSize: 22, bold: true, color: CORAL, margin: 0 });
  s.addText(lesson.teaching_point, { x: 0.75, y: 3.4, w: PAGE_W - 1.5, h: 1.7, fontFace: "Figtree", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });

  footer(s, pageNum++, false);
}

// ===== Slides: Lesson Vocabulary Review -- purple bookend =====
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
