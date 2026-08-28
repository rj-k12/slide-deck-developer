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
// Missing from this file (present in generate_deck.js) -- needed for the
// vocab word/definition color fix below, confirmed 110045 against the
// template's Engage/Launch Vocabulary XML.
const DETAIL_DESC_COLOR = "110045";
// Same confirmed values as generate_deck.js -- see that file for the
// template-measurement notes.
const TP_BORDER_ORANGE = "FFAB7B";
const TP_BORDER_CLOSING = "DCD6EE";
const LG_BORDER_LIGHT_PURPLE = "B8AEE0"; // NOT independently confirmed -- see generate_deck.js
const PINK_BORDER = "ED6A91";
const PINK_PALE = "FDEEF2";
const AI_ICON_BLUE = "5B9BD5";
const WHITE = "FFFFFF";
const BODY = "33322E";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
const PAGE_W = 13.3, PAGE_H = 7.5;

// Header gradient band (white -> pale yellow). Same asset and reasoning
// as generate_deck.js's addHeaderGradient -- this file explicitly shares
// the same verified palette/template source (see the comment at the top
// of this file), so the same header treatment applies. Declared here,
// before first use below, since HEADER_GRADIENT is a top-level const and
// isn't hoisted the way a function declaration is.
//
// Only wired up for the EQ/TP/LG slide and "Engage"-named vocabulary
// sections below, mirroring the two rules confirmed directly against
// KNO_Reading_TEMPLATE_8_7.pptx's XML in generate_deck.js. This file has
// no chart-type section (no direct equivalent of Reading's "Launch
// Chart"), so no gradient rule was added for one -- extend this if a
// literature-response-specific template turns out to need it.
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
  // Confirmed directly from the template's raw XML: the unit-header text
  // on the left is italic (i="1"), the page number on the right is not
  // (i="0") -- Ashley's comment ("shouldn't be italicized") led to
  // removing italic from here entirely; RJ's decision was to trust the
  // template's actual formatting over that comment, so it's back, but
  // ONLY on the left text, matching the template's real asymmetry.
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

// Catherine confirmed (Google Slides comment thread on this lesson) that
// underlined text in the source PDF -- most commonly book/text titles
// like "Finding Langston" -- should be preserved when it shows up in
// extracted content. extract_lesson.py wraps underlined spans in
// <u>...</u> within whatever field they appear in; this turns a string
// containing that markup into a run array pptxgenjs can actually render
// with real underline formatting, since addText only applies formatting
// per-run, not via inline tags.
//
// Also folds in bold-page-number detection (e.g. "pages 46-50") in the
// same pass, rather than as a separate parser -- a single string can
// plausibly need both ("read pages 46-50 of <u>Finding Langston</u>"),
// and running two independent regex passes over the same text would
// require re-splitting an already-split run array, which is more
// fragile than handling both in one pass.
function parseInlineMarkup(text) {
  if (!text) return [{ text: text || "" }];
  // Ashley flagged "Finding Langston" appearing un-underlined (title
  // slide byline, prompt text) even though extract_lesson.py is supposed
  // to wrap underlined source spans in <u>...</u> -- checked
  // lesson_6b_REAL.json directly and it has no <u> tags at all, so
  // extraction didn't tag the title this time. Since core_text IS the
  // book/text title by definition, auto-underline every literal,
  // not-already-tagged occurrence of it here as a safety net. Ported
  // from the same fix in generate_deck.js.
  if (typeof lesson !== "undefined" && lesson.core_text && text.includes(lesson.core_text) && !text.includes(`<u>${lesson.core_text}</u>`)) {
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
    if (part.options && part.options.underline) {
      expanded.push(part);
      return;
    }
    let li = 0, mm, found = false;
    pageNumRe.lastIndex = 0;
    while ((mm = pageNumRe.exec(part.text)) !== null) {
      found = true;
      if (mm.index > li) expanded.push({ text: part.text.slice(li, mm.index) });
      expanded.push({ text: mm[0], options: { bold: true } });
      li = mm.index + mm[0].length;
    }
    if (found) {
      if (li < part.text.length) expanded.push({ text: part.text.slice(li) });
    } else {
      expanded.push(part);
    }
  });
  return expanded.length ? expanded : [{ text }];
}
const VOCAB_ICON_DIR = process.argv[5] || require("path").join(__dirname, "vocab_icons");
function iconPathFor(word) {
  const safe = word.toLowerCase().replace(/[^a-z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  const p = require("path").join(VOCAB_ICON_DIR, `${safe}.png`);
  return fs.existsSync(p) ? p : null;
}
// Estimates how many wrapped lines a definition needs at vocabBlock's
// fontSize 21.5 in a card of width colW, so card/row height can be sized
// to actually fit the text -- PowerPoint text boxes don't clip
// overflowing text, they just spill past the box (and the card behind
// it). Found via a real long definition in a Discourse Deep Dive lesson
// ("process": needed 5 wrapped lines, but the old fixed 1.55in box only
// fit ~4 -- the 5th line visibly overlapped the card's rounded border).
function estimateDefinitionLines(definition, colW) {
  const charsPerLine = (colW - 0.4) * 5.4; // empirically tuned for Arial 21.5pt
  return Math.max(1, Math.ceil(definition.length / charsPerLine));
}
function vocabBlock(slide, words, x, y, w, onDark) {
  const colW = (w - 0.3) / 2;
  const cardBg = onDark ? WHITE : PEACH;
  const rows = chunk(words, 2);
  let cursorY = y;
  rows.forEach(rowWords => {
    // Row height = whatever the tallest card in this row actually needs,
    // floored at the original 2.7in default so short definitions keep
    // the established look. Radius confirmed from template (Engage
    // Vocabulary slide) -- same value used in generate_deck.js's
    // vocabBlock.
    const rowH = Math.max(2.7, ...rowWords.map(item => 1.0 + estimateDefinitionLines(item.definition, colW) * 0.34 + 0.25));
    rowWords.forEach((item, ci) => {
      const bx = x + ci * (colW + 0.3), by = cursorY;
      slide.addShape("roundRect", { x: bx, y: by, w: colW, h: rowH, rectRadius: 0.20, fill: { color: cardBg }, line: onDark ? { color: CORAL, width: 1 } : { type: "none" } });
      const iconPath = iconPathFor(item.word);
      if (iconPath) {
        slide.addImage({ path: iconPath, x: bx + 0.2, y: by + 0.2, w: 0.55, h: 0.55 });
      }
      // Ashley: "template - font color is hex #110045" (on the Engage
      // Vocabulary "compelling" word/definition) -- confirmed against the
      // template's Engage/Launch Vocabulary XML directly: both [word] and
      // [definition] runs are 110045, not NAVY_INK/BODY. Ported from the
      // same fix in generate_deck.js's vocabBlock.
      slide.addText(parseInlineMarkup(item.word), { x: bx + 0.9, y: by + 0.22, w: colW - 1.05, h: 0.5, fontFace: "Arial", fontSize: 24, bold: true, color: DETAIL_DESC_COLOR, margin: 0 });
      slide.addText(parseInlineMarkup(item.definition), { x: bx + 0.2, y: by + 1.0, w: colW - 0.4, h: rowH - 1.15, fontFace: "Arial", fontSize: 21.5, color: DETAIL_DESC_COLOR, margin: 0, valign: "top" });
    });
    cursorY += rowH + 0.15;
  });
}
function promptBox(slide, prompt, x, y, w, h) {
  slide.addShape("roundRect", { x, y, w, h, rectRadius: 0.06, fill: { color: PINK_PALE }, line: { color: PINK_BORDER, width: 1.5, dashType: "dash" } });
  slide.addText("Prompt", { x: x + 0.3, y: y + 0.15, w: w - 0.6, h: 0.35, fontFace: "Arial", fontSize: 24, bold: true, color: PURPLE, margin: 0 });
  slide.addText(parseInlineMarkup(prompt), { x: x + 0.3, y: y + 0.55, w: w - 0.6, h: h - 0.7, fontFace: "Arial", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });
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
  s.addText(`LESSON ${lesson.lesson_number}`, { x: 0.55, y: 0.6, w: 1.9, h: 0.55, fontFace: "Arial", fontSize: 16, bold: true, color: NAVY_INK, align: "center", valign: "middle", margin: 0 });
  s.addText(`Lesson ${lesson.lesson_number}: ${lesson.lesson_type}`, { x: 0.55, y: 1.3, w: 12.2, h: 1.8, fontFace: "Arial", fontSize: 50, bold: true, color: WHITE, margin: 0, valign: "top" });
  s.addText(`Grade ${lesson.grade.replace('Grade ','')}, Knowledge Unit ${lesson.unit_number}: ${lesson.unit_title}`, { x: 0.55, y: 3.35, w: 12.2, h: 1.1, fontFace: "Arial", fontSize: 25.5, color: "E5EFF9", margin: 0, valign: "top" });
  if (lesson.core_text) {
    const byLine = lesson.author ? `${lesson.core_text} by ${lesson.author}` : lesson.core_text;
    s.addText(parseInlineMarkup(`${byLine}${lesson.pages ? "   |   Pages " + lesson.pages : ""}`), { x: 0.55, y: 4.65, w: 12.2, h: 0.6, fontFace: "Arial", fontSize: 21.5, italic: true, color: "C9BEEB", margin: 0, valign: "top" });
  }
  s.addText("Copyright \u00a9 2026 Lavinia Group. All Rights Reserved. RedThread is a trademark of K12 Coalition.", { x: 0.55, y: PAGE_H - 0.5, w: 11.5, h: 0.3, fontFace: "Arial", fontSize: 10, color: "9A8FD1", margin: 0 });
}

// ===== Slide: EQ / TP / LG -- blank, real 3-box structure. EQ full-width
// on top, Teaching Point and Language Goal SIDE BY SIDE below (confirmed
// against the template's own shape positions -- the earlier stacked
// version left significant dead space at the bottom of the slide) =====
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
  }

  s.addShape("roundRect", { x: 0.55, y: rowTop, w: colW, h: 0.05, rectRadius: 0.025, fill: { color: CORAL }, line: { type: "none" } });
  s.addShape("roundRect", { x: 0.55, y: rowTop + 0.05, w: colW, h: rowH - 0.05, rectRadius: 0.139, fill: { color: WHITE }, line: { color: TP_BORDER_ORANGE, width: 1.5 } });
  s.addText("Teaching Point", { x: 0.75, y: rowTop + 0.25, w: colW - 0.4, h: 0.3, fontFace: "Arial", fontSize: 18.5, bold: true, color: CORAL, margin: 0 });
  if (lesson.teaching_point) {
    // Ashley's #110045 comment here (on "Writers study a prompt...")
    // checked against the template's EQ/TP/LG slide (slide3.xml)
    // directly: the Teaching Point body run is actually 0E0142
    // (NAVY_INK), not 110045. BODY was wrong either way; using the
    // measured template value. Same discrepancy as generate_deck.js --
    // worth flagging to Ashley that the real template doesn't match what
    // she expected here.
    s.addText(parseInlineMarkup(lesson.teaching_point), { x: 0.75, y: rowTop + 0.6, w: colW - 0.4, h: rowH - 0.8, fontFace: "Arial", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });
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

  footer(s, 2, false);
}

// ===== Slides: Engage / Launch sections =====
const DECORATIVE_BG = require("path").join(__dirname, "assets", "decorative_bg.png");
function addDecorativeBg(slide) {
  const scale = PAGE_W / 10;
  slide.addImage({ path: DECORATIVE_BG, x: 5.99 * scale, y: 0, w: 4.01 * scale, h: 3.46 * scale });
}

// Catherine's explicit rule (Slack, this lesson's Launch section):
// "vocabulary for a section should always be the last slide of that
// section" -- regardless of where that vocab word happens to be defined
// inline in the source document's own script (Ashley found "ensure",
// "evaluate", and "evolve" each introduced at three different, scattered
// points in the real lesson text). extract_lesson.py's extraction order
// will tend to mirror that same scattered document order for every
// future lesson too, so this needs to be a rendering rule enforced here,
// not a one-off reorder of this lesson's own JSON.
//
// Groups lesson.sections by section_name (in first-appearance order),
// and within each group, renders every entry's NON-vocabulary content
// first (preserving each entry's own original relative order), then
// every entry's vocabulary content last -- handles both real shapes:
// separate section objects sharing one section_name (this lesson's
// actual case), and a single object that happens to carry vocabulary
// alongside other fields together.
function renderGroupedSections(sections, renderNonVocab, renderVocab) {
  const names = [];
  (sections || []).forEach(s => { if (!names.includes(s.section_name)) names.push(s.section_name); });
  names.forEach(name => {
    const group = (sections || []).filter(s => s.section_name === name);
    group.forEach(section => renderNonVocab(section));
    group.forEach(section => renderVocab(section));
  });
}

let pageNum = 3;
renderGroupedSections(
  lesson.sections,
  section => {
    if (section.mentor_prompt) {
      const s = pres.addSlide();
      addHeaderGradient(s);
      slideTitle(s, `${section.section_name} Mentor Prompt`, false);
      promptBox(s, section.mentor_prompt, 0.55, 1.5, PAGE_W - 1.1, 2.2);
      footer(s, pageNum++, false);
    }
    if (section.claims_to_evaluate && section.claims_to_evaluate.length) {
      const s = pres.addSlide();
      addHeaderGradient(s);
      slideTitle(s, `${section.section_name} Claims to Evaluate`, false);
      let y = 1.4;
      section.claims_to_evaluate.forEach(claim => {
        s.addShape("roundRect", { x: 0.55, y, w: PAGE_W - 1.1, h: 1.3, rectRadius: 0.06, fill: { color: PEACH }, line: { type: "none" } });
        s.addText(parseInlineMarkup(claim), { x: 0.8, y: y + 0.12, w: PAGE_W - 1.6, h: 1.05, fontFace: "Arial", fontSize: 21.5, color: NAVY_INK, margin: 0, valign: "top" });
        y += 1.5;
      });
      footer(s, pageNum++, false);
    }
    if (section.resource_unavailable) {
      const s = pres.addSlide();
      addHeaderGradient(s);
      slideTitle(s, `${section.section_name} Resource`, false);
      s.addShape("roundRect", { x: 0.55, y: 1.5, w: PAGE_W - 1.1, h: 1.1, rectRadius: 0.08, fill: { color: "FEF3C7" }, line: { color: "D97706", width: 1 } });
      s.addText([{ text: "\u26a0 Needs manual follow-up: " }, ...parseInlineMarkup(section.resource_unavailable)], { x: 0.8, y: 1.65, w: PAGE_W - 1.6, h: 0.8, fontFace: "Arial", fontSize: 16, italic: true, color: "92400E", margin: 0, valign: "top" });
      footer(s, pageNum++, false);
    }
  },
  section => {
    if (section.vocabulary && section.vocabulary.length) {
      chunk(section.vocabulary, 4).forEach((words, i) => {
        const s = pres.addSlide();
        // RJ: gradient was missing here and elsewhere in this file --
        // same fix as generate_deck.js. Rendered all 18 real template
        // slides to PNG and sampled pixel colors directly (grep for
        // gradFill in each slide's own XML missed most of these because
        // the gradient is inherited from the slide LAYOUT, not embedded
        // per-slide). Ground truth: every content slide has it except
        // the cover, the purple Lesson-title slide, and the dark-purple
        // Lesson Vocabulary Review slides -- applying unconditionally
        // here and at every other light-background slide below, not
        // just "Engage" as before.
        addHeaderGradient(s);
        if (section.section_name === "Launch") addDecorativeBg(s);
        slideTitle(s, `${section.section_name} Vocabulary${i > 0 ? " (continued)" : ""}`, false);
        vocabBlock(s, words, 0.55, 1.25, PAGE_W - 1.1, false);
        footer(s, pageNum++, false);
      });
    }
  }
);

// ===== Slide: Literature Response =====
if (lesson.literature_response_prompt) {
  const s = pres.addSlide();
  addHeaderGradient(s);
  slideTitle(s, "Literature Response", false);
  s.addText("Teaching Point", { x: 0.55, y: 1.3, w: PAGE_W - 1.1, h: 0.35, fontFace: "Arial", fontSize: 22, bold: true, color: CORAL, margin: 0 });
  s.addText(parseInlineMarkup(lesson.teaching_point), { x: 0.55, y: 1.7, w: PAGE_W - 1.1, h: 1.1, fontFace: "Arial", fontSize: 21.5, color: BODY, margin: 0, valign: "top" });
  promptBox(s, lesson.literature_response_prompt, 0.55, 3.0, PAGE_W - 1.1, 1.4);
  footer(s, pageNum++, false);
}

// ===== Slide: Writers' Circle & Revise -- heading + Teaching Point only.
// RJ (relaying a newer Ashley comment on 6b slide 9, not visible in this
// doc's own comment threads when checked) : add this slide back, but
// only the heading and Teaching Point -- the "Focus for Feedback"
// bulleted list stays cut. That list was the actual problem in Ashley's
// original review comment (teacher-facing facilitation guidance, not
// content for the student-facing deck) -- the heading + Teaching Point
// recap alone doesn't have that issue, so only that half comes back.
// lesson.writers_circle.focus_points still isn't rendered here.
if (lesson.writers_circle) {
  const s = pres.addSlide();
  addHeaderGradient(s);
  slideTitle(s, "Writers' Circle & Revise", false);
  s.addText("Teaching Point", { x: 0.55, y: 1.3, w: PAGE_W - 1.1, h: 0.35, fontFace: "Arial", fontSize: 22, bold: true, color: CORAL, margin: 0 });
  s.addText(parseInlineMarkup(lesson.teaching_point), { x: 0.55, y: 1.7, w: PAGE_W - 1.1, h: 1.1, fontFace: "Arial", fontSize: 21.5, color: NAVY_INK, margin: 0, valign: "top" });
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
  s.addText(parseInlineMarkup(lesson.essential_question), { x: 0.85, y: 1.65, w: PAGE_W - 1.6, h: 0.75, fontFace: "Arial", fontSize: 24, bold: true, color: NAVY_INK, margin: 0, valign: "top" });

  s.addShape("roundRect", { x: 0.55, y: 2.75, w: PAGE_W - 1.1, h: 0.05, rectRadius: 0.025, fill: { color: CORAL }, line: { type: "none" } });
  s.addShape("roundRect", { x: 0.55, y: 2.8, w: PAGE_W - 1.1, h: 2.4, rectRadius: 0.113, fill: { color: WHITE }, line: { color: TP_BORDER_CLOSING, width: 1 } });
  s.addText("Teaching Point", { x: 0.75, y: 3.0, w: PAGE_W - 1.5, h: 0.35, fontFace: "Arial", fontSize: 22, bold: true, color: CORAL, margin: 0 });
  s.addText(parseInlineMarkup(lesson.teaching_point), { x: 0.75, y: 3.4, w: PAGE_W - 1.5, h: 1.7, fontFace: "Arial", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });

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
