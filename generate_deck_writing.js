// generate_deck_writing.js -- Writing lesson type.
//
// Same situation as Unit Launch/Project: no dedicated visual .pptx
// template exists for Writing, so this is built against the written
// "Knowledge Slide Template Outlines" spec's "Writing" section, reusing
// Reading's verified colors/chrome/helper functions wholesale. Structure
// per that outline:
//   1. Cover + title
//   2. EQ/TP/Language Goal
//   3-6. Engage/Launch: vocab, charts, prompt/mentor/criteria list
//   7. Independent Writing (Teaching Point, prompt/criteria, planner)
//   8. Writers' Circle & Revise
//   9. Closing
//   10. Lesson Vocabulary Review
// Verified against the real Lesson 18b ("Writing") content in
// RT_Gr4U2_KNO_full_unit_Teacher_Guide.pdf (physical pages 393-399).
// This lesson also reprints a full Mentor "Realistic Narrative Writing
// Planner" -- a structured, multi-section reference example (message +
// two bulleted lists + a 3-column bulleted table) -- the first real test
// of the schema's planner_sections concept (previously only used by
// generate_deck_language.js) for a mentor/student-planner comparison,
// which the outline calls out as Project/Writing-specific content
// ("mentor draft comparisons").
const pptxgen = require("pptxgenjs");
const fs = require("fs");

const lessonPath = process.argv[2] || "lesson_18b_REAL.json";
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
const WHITE = "FFFFFF";
const BODY = "33322E";
const VIOLET = "9343F6";

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

// ===== Mentor / student planner renderer -- new for this lesson type.
// Handles the 4 section shapes this real planner actually has (text,
// list, list, table-with-bulleted-cells); each rendered compactly since
// all four need to fit on one slide, matching how the source itself
// prints this as a single dense page. =====
function renderPlanner(slide, planner, x, y, w, h) {
  let cy = y;
  const rowH = h / (planner.sections.length <= 3 ? 3 : 3.6); // table section gets more room than a fixed 1/n split
  planner.sections.forEach(section => {
    if (section.type === "text") {
      slide.addShape("roundRect", { x, y: cy, w, h: 0.75, rectRadius: 0.05, fill: { color: "FBE8AB" }, line: { type: "none" } });
      slide.addText(section.label, { x: x + 0.15, y: cy + 0.05, w: w - 0.3, h: 0.25, fontFace: "Arial", fontSize: 13, bold: true, color: CORAL, margin: 0 });
      slide.addText(section.content || "[to be completed]", { x: x + 0.15, y: cy + 0.3, w: w - 0.3, h: 0.4, fontFace: "Arial", fontSize: 13, italic: !section.content, color: section.content ? NAVY_INK : MUTED, margin: 0, valign: "top" });
      cy += 0.85;
    } else if (section.type === "list") {
      // Two "list" sections render side by side (Protagonist / Change
      // Protagonist Faces in the real planner) -- handled by the caller
      // passing a half-width w for these two calls instead of looping
      // them here, so this branch just renders whatever width it's given.
      slide.addText(section.label, { x, y: cy, w, h: 0.25, fontFace: "Arial", fontSize: 13, bold: true, color: CORAL, margin: 0 });
      let iy = cy + 0.28;
      if (section.items && section.items.length) {
        section.items.forEach(item => {
          slide.addText(`\u2022 ${item}`, { x, y: iy, w, h: 0.28, fontFace: "Arial", fontSize: 12, color: NAVY_INK, margin: 0, valign: "top" });
          iy += 0.26;
        });
      } else {
        slide.addText("[to be completed]", { x, y: iy, w, h: 0.28, fontFace: "Arial", fontSize: 12, italic: true, color: MUTED, margin: 0 });
      }
    } else if (section.type === "table") {
      slide.addText(section.label, { x, y: cy, w, h: 0.25, fontFace: "Arial", fontSize: 13, bold: true, color: CORAL, margin: 0 });
      cy += 0.3;
      const tableH = y + h - cy;
      const headerRow = section.columns.map(c => ({
        text: c, options: { bold: true, color: WHITE, fill: { color: PURPLE }, align: "left", valign: "middle", fontFace: "Arial", fontSize: 11, border: { type: "solid", color: PURPLE, pt: 1 } }
      }));
      const bodyCells = section.columns.map((_, ci) => {
        const cellItems = (section.rows && section.rows[0] && section.rows[0][ci]) || [];
        const text = cellItems.length
          ? cellItems.map(item => ({ text: item, options: { bullet: { code: "2022" }, breakLine: true } }))
          : "[to be completed]";
        return { text, options: { color: cellItems.length ? BODY : MUTED, italic: !cellItems.length, fontFace: "Arial", fontSize: 10, valign: "top", fill: { color: WHITE }, border: { type: "solid", color: TABLE_BORDER_BODY, pt: 1 } } };
      });
      slide.addTable([headerRow, bodyCells], { x, y: cy, w, h: tableH, fontFace: "Arial", autoPage: false, rowH: [0.35, tableH - 0.35] });
      cy += tableH;
    }
  });
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

// ===== Slide: EQ / TP / LG =====
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
    s.addText("Not specified in this lesson's source material.", { x: 0.85, y: 1.05, w: PAGE_W - 1.6, h: 0.8, fontFace: "Arial", fontSize: 16, italic: true, color: MUTED, margin: 0, valign: "top" });
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

// ===== Slides: Engage / Launch sections =====
(lesson.sections || []).forEach(section => {
  if (section.read_directions) {
    const s = pres.addSlide();
    addHeaderGradient(s);
    slideTitle(s, `${section.section_name} Read Directions`, false);
    s.addText(parseInlineMarkup(section.read_directions), { x: 0.55, y: 1.5, w: PAGE_W - 1.1, h: 1, fontFace: "Arial", fontSize: 21.5, color: DETAIL_DESC_COLOR, margin: 0, valign: "top" });
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
  if (section.resource_unavailable) {
    const s = pres.addSlide();
    addHeaderGradient(s);
    slideTitle(s, `${section.section_name} Resource`, false);
    s.addShape("roundRect", { x: 0.55, y: 1.5, w: PAGE_W - 1.1, h: 1.4, rectRadius: 0.08, fill: { color: "FEF3C7" }, line: { color: "D97706", width: 1 } });
    s.addText([{ text: "\u26a0 Needs manual follow-up: " }, ...parseInlineMarkup(section.resource_unavailable)], { x: 0.8, y: 1.65, w: PAGE_W - 1.6, h: 1.1, fontFace: "Arial", fontSize: 16, italic: true, color: "92400E", margin: 0, valign: "top" });
    footer(s, pageNum++, false);
  }
});

// ===== Slide: Mentor Planner -- reference example, fully filled in
// (this lesson's own real content, not invented) =====
if (lesson.mentor_planner) {
  const s = pres.addSlide();
  addHeaderGradient(s);
  slideTitle(s, "Mentor Planner", false);
  const p = lesson.mentor_planner;
  const textSection = p.sections.find(sec => sec.type === "text");
  const listSections = p.sections.filter(sec => sec.type === "list");
  const tableSection = p.sections.find(sec => sec.type === "table");
  let y = 1.15;
  if (textSection) {
    renderPlanner(s, { sections: [textSection] }, 0.55, y, PAGE_W - 1.1, 0.85);
    y += 0.95;
  }
  if (listSections.length) {
    const halfW = (PAGE_W - 1.1 - 0.3) / 2;
    listSections.forEach((sec, i) => {
      renderPlanner(s, { sections: [sec] }, 0.55 + i * (halfW + 0.3), y, halfW, 1.0);
    });
    y += 1.15;
  }
  if (tableSection) {
    renderPlanner(s, { sections: [tableSection] }, 0.55, y, PAGE_W - 1.1, PAGE_H - y - 0.6);
  }
  footer(s, pageNum++, false);
}

// ===== Slide: Independent Writing =====
if (lesson.independent_writing) {
  const iw = lesson.independent_writing;
  const s = pres.addSlide();
  addHeaderGradient(s);
  slideTitle(s, "Independent Writing", false);
  let y = 1.15;
  if (lesson.teaching_point) {
    s.addText("TEACHING POINT", { x: 0.55, y, w: PAGE_W - 1.1, h: 0.3, fontFace: "Arial", fontSize: 16, bold: true, color: CORAL, margin: 0 });
    y += 0.35;
    s.addText(parseInlineMarkup(lesson.teaching_point), { x: 0.55, y, w: PAGE_W - 1.1, h: 0.7, fontFace: "Arial", fontSize: 13, color: NAVY_INK, margin: 0, valign: "top" });
    y += 0.75;
  }
  if (iw.directions) {
    s.addText(parseInlineMarkup(iw.directions), { x: 0.55, y, w: PAGE_W - 1.1, h: 0.65, fontFace: "Arial", fontSize: 13, color: DETAIL_DESC_COLOR, margin: 0, valign: "top" });
    y += 0.7;
  }
  if (iw.planner) {
    const p = iw.planner;
    s.addText(p.title, { x: 0.55, y, w: PAGE_W - 1.1, h: 0.3, fontFace: "Arial", fontSize: 16, bold: true, color: CORAL, margin: 0 });
    y += 0.35;
    const textSection = p.sections.find(sec => sec.type === "text");
    const listSections = p.sections.filter(sec => sec.type === "list");
    const tableSection = p.sections.find(sec => sec.type === "table");
    if (textSection) {
      renderPlanner(s, { sections: [textSection] }, 0.55, y, PAGE_W - 1.1, 0.7);
      y += 0.8;
    }
    if (listSections.length) {
      const halfW = (PAGE_W - 1.1 - 0.3) / 2;
      listSections.forEach((sec, i) => {
        renderPlanner(s, { sections: [sec] }, 0.55 + i * (halfW + 0.3), y, halfW, 0.7);
      });
      y += 0.8;
    }
    if (tableSection) {
      renderPlanner(s, { sections: [tableSection] }, 0.55, y, PAGE_W - 1.1, PAGE_H - y - 0.6);
    }
  }
  footer(s, pageNum++, false);
}

// ===== Slide: Writers' Circle =====
if (lesson.writers_circle && lesson.writers_circle.focus_points && lesson.writers_circle.focus_points.length) {
  const s = pres.addSlide();
  addHeaderGradient(s);
  slideTitle(s, "Writers' Circle", false);
  let y = 1.25;
  s.addText("FOCUS FOR FEEDBACK", { x: 0.55, y, w: PAGE_W - 1.1, h: 0.35, fontFace: "Arial", fontSize: 20, bold: true, color: CORAL, margin: 0 });
  y += 0.5;
  lesson.writers_circle.focus_points.forEach(point => {
    s.addShape("ellipse", { x: 0.6, y: y + 0.1, w: 0.11, h: 0.11, fill: { color: VIOLET }, line: { type: "none" } });
    s.addText(parseInlineMarkup(point), { x: 0.85, y, w: PAGE_W - 1.4, h: 0.7, fontFace: "Arial", fontSize: 17, color: NAVY_INK, margin: 0, valign: "top" });
    y += 0.75;
  });
  footer(s, pageNum++, false);
}

// ===== Slide: Revise -- only rendered when this lesson actually has
// student-facing directions text for it; Lesson 18b's own Revise section
// has nothing beyond a teacher-only instruction ("Direct students to
// revise their writing based on the discussion"), which is exactly the
// kind of facilitation note excluded elsewhere in this pipeline -- so
// per the outline's own "remove any placeholder slides that do not have
// content" instruction, this lesson gets no Revise slide. =====
if (lesson.revise_directions) {
  const s = pres.addSlide();
  addHeaderGradient(s);
  slideTitle(s, "Revise", false);
  s.addText(parseInlineMarkup(lesson.revise_directions), { x: 0.55, y: 1.25, w: PAGE_W - 1.1, h: 1.3, fontFace: "Arial", fontSize: 18, color: DETAIL_DESC_COLOR, margin: 0, valign: "top" });
  footer(s, pageNum++, false);
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
