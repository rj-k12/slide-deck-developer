const pptxgen = require("pptxgenjs");
const fs = require("fs");

const lessonPath = process.argv[2] || "lesson_6a_REAL.json";
const outPath = process.argv[3] || "output.pptx";
const lesson = JSON.parse(fs.readFileSync(lessonPath, "utf-8"));

// Verified directly from KNO_Reading_TEMPLATE_8_7.pptx's raw XML -- not
// approximated from a rendered image.
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
const TAN_BG = "FBE8AB";
// Confirmed directly from the template's raw XML (chart on its Launch
// Chart slide): header row border matches the header's own PURPLE fill,
// body rows use this orange -- NOT the uniform TABLE_BORDER gray this
// generator used everywhere before.
const TABLE_BORDER_BODY = "FFAB7B";
// Confirmed directly from the template: the Teaching Point box's solid
// outline on the EQ/TP/LG slide.
const TP_BORDER_ORANGE = "FFAB7B";
// Confirmed directly from the template's Closing slide specifically --
// its Teaching Point box (full-width there, not split into columns) uses
// a different, lighter border than the EQ/TP/LG slide's orange one.
const TP_BORDER_CLOSING = "DCD6EE";
// NOT independently confirmed against the template XML -- no border was
// found on the Language Goal box in the one template file available.
// Ashley's comment asked for "a lighter purple" outline; this is a
// reasonable tint of PURPLE, not a measured value. Worth confirming with
// her directly before treating this one as settled.
const LG_BORDER_LIGHT_PURPLE = "B8AEE0";
// Confirmed directly from the template's Discourse Clubs and Quick Write
// detail slides -- both use this same pale pink for the white prompt
// box's outline (the accent-bar color is what actually varies per
// slide type, and that was already correct in this file before this
// change).
const DETAIL_BOX_BORDER = "FDEEF2";
const PINK_BORDER = "ED6A91";
const PINK_PALE = "FDEEF2";
const AI_ICON_BLUE = "5B9BD5";
const WHITE = "FFFFFF";
const BODY = "33322E";

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE";
const PAGE_W = 13.3, PAGE_H = 7.5;

// Header gradient band (white -> pale yellow). Declared here, before the
// EQ/TP/LG slide below (its first use), rather than down near
// addDecorativeBg -- HEADER_GRADIENT is a top-level const, and const
// declarations aren't hoisted the way function declarations are, so
// calling addHeaderGradient() before this line had executed would throw
// at runtime ("Cannot access 'HEADER_GRADIENT' before initialization").
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

// Individual detail slide (Discourse Clubs / Quick Write / Whole-Class
// Discourse each get one): left icon card + right white box with a solid
// accent-color left bar. Matches the real template exactly -- verified
// directly against its own shape positions and colors, not approximated.
// This is a DIFFERENT visual pattern from the pink dashed combined intro
// slide (Quick Write & Discourse Clubs together), which is correct as-is.
function detailPromptSlide(s, opts) {
  const { cardColor, accentColor, iconPath, cardTitle, description, promptLabel, promptText, boxBorderColor } = opts;
  const cardX = 0.7, cardY = 1.9, cardW = 3.0, cardH = 3.85;
  // Radius confirmed from template (Discourse Clubs / Quick Write detail
  // slides): adj=5454 on the icon card -> ~0.167in at this file's scale.
  // Was 0.08, visibly too subtle.
  s.addShape("roundRect", { x: cardX, y: cardY, w: cardW, h: cardH, rectRadius: 0.167, fill: { color: cardColor }, line: { type: "none" } });
  if (iconPath) {
    s.addImage({ path: iconPath, x: cardX + 0.25, y: cardY + 0.3, w: 0.65, h: 0.65 });
  }
  s.addText(cardTitle, { x: cardX + 0.25, y: cardY + 1.15, w: cardW - 0.5, h: 0.65, fontFace: "Arial", fontSize: 21.5, bold: true, color: NAVY_INK, margin: 0, valign: "top" });
  s.addText(description, { x: cardX + 0.25, y: cardY + 1.85, w: cardW - 0.5, h: cardH - 2.05, fontFace: "Arial", fontSize: 17.5, color: MUTED, margin: 0, valign: "top" });

  const boxX = cardX + cardW + 0.45, boxW = PAGE_W - boxX - 0.55;
  s.addShape("roundRect", { x: boxX, y: cardY, w: 0.08, h: cardH, rectRadius: 0.04, fill: { color: accentColor }, line: { type: "none" } });
  // Radius confirmed from template: adj=4285 on the white prompt box ->
  // ~0.167in, essentially the same absolute radius as the card above.
  // Border color: defaults to DETAIL_BOX_BORDER (pale pink), confirmed
  // directly from the template on both the Discourse Clubs and Quick
  // Write slides -- but Ashley's own comment on the Discourse Clubs
  // slide specifically says "light orange outline," which contradicts
  // that measurement. Going with her stated color for that one call site
  // (see below) rather than the measured value, since she's looking at
  // the real rendered deck and may be catching something this static
  // template file doesn't reflect.
  s.addShape("roundRect", { x: boxX + 0.08, y: cardY, w: boxW - 0.08, h: cardH, rectRadius: 0.167, fill: { color: WHITE }, line: { color: boxBorderColor || DETAIL_BOX_BORDER, width: 1 } });
  s.addText(promptLabel, { x: boxX + 0.4, y: cardY + 0.35, w: boxW - 0.8, h: 0.3, fontFace: "Arial", fontSize: 18.5, bold: true, color: accentColor, margin: 0 });
  s.addText(parseInlineMarkup(promptText), { x: boxX + 0.4, y: cardY + 0.75, w: boxW - 0.8, h: cardH - 1.1, fontFace: "Arial", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });
}
function chunk(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

// Ashley's comment: "bold pg numbers" on read-directions text (e.g.
// "Read pages 46-50 of Finding Langston..."). extract_lesson.py doesn't
// give page numbers as a separate field within read_directions -- it's
// one plain string -- so this finds a "page(s) N" or "page(s) N-N"
// pattern within the string and bolds just that substring, returning a
// run array for addText rather than a plain string.
// Catherine confirmed (Google Slides comment thread on Lesson 6b) that
// underlined text in the source PDF -- most commonly book/text titles
// like "Finding Langston" -- should be preserved when it shows up in
// extracted content. extract_lesson.py wraps underlined spans in
// <u>...</u> within whatever field they appear in; this turns a string
// containing that markup into a run array pptxgenjs can actually render
// with real underline formatting. Originally built for
// generate_deck_literature_response.js; ported here unchanged since the
// same underline-preservation need applies to Reading/Close Reading
// content (Reading lessons reference "Finding Langston" constantly).
//
// Supersedes the old boldPageNumbers() -- folds bold-page-number
// detection into the same single pass, since a string can plausibly
// need both ("read pages 46-50 of <u>Finding Langston</u>").
function parseInlineMarkup(text) {
  if (!text) return [{ text: text || "" }];
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

// Vocabulary card grid, matching the template's 2-col layout with the
// "AI ICON matching definition" oval placeholder next to each word.
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
    // Row height = whatever the tallest card in this row actually needs
    // (both cards in a row must match height), floored at the original
    // 2.7in default so short definitions keep the established look.
    // Radius confirmed from template (Engage Vocabulary slide): adj=4000
    // on a w=4.31/h=3.80 card at this file's scale -> ~0.20in. Was 0.06,
    // visibly too subtle compared to the template.
    const rowH = Math.max(2.7, ...rowWords.map(item => 1.0 + estimateDefinitionLines(item.definition, colW) * 0.34 + 0.25));
    rowWords.forEach((item, ci) => {
      const bx = x + ci * (colW + 0.3), by = cursorY;
      slide.addShape("roundRect", { x: bx, y: by, w: colW, h: rowH, rectRadius: 0.20, fill: { color: cardBg }, line: onDark ? { color: CORAL, width: 1 } : { type: "none" } });
      const iconPath = iconPathFor(item.word);
      if (iconPath) {
        slide.addImage({ path: iconPath, x: bx + 0.2, y: by + 0.2, w: 0.55, h: 0.55 });
      }
      slide.addText(parseInlineMarkup(item.word), { x: bx + 0.9, y: by + 0.22, w: colW - 1.05, h: 0.5, fontFace: "Arial", fontSize: 24, bold: true, color: NAVY_INK, margin: 0 });
      slide.addText(parseInlineMarkup(item.definition), { x: bx + 0.2, y: by + 1.0, w: colW - 0.4, h: rowH - 1.15, fontFace: "Arial", fontSize: 21.5, color: BODY, margin: 0, valign: "top" });
    });
    cursorY += rowH + 0.15;
  });
}

// Chart table: purple header row, alternating pale-blue/white zebra rows,
// matching the template exactly (used both for reference charts, filled
// in with real content, and blank student-fillable charts).
function addChart(slide, chart, x, y, w, h) {
  const nRows = (chart.rows && chart.rows.length) || 4;
  // Border colors confirmed directly from the template's Launch Chart
  // slide: the header row's border matches its own PURPLE fill, and body
  // rows use TABLE_BORDER_BODY (orange) -- not the single uniform gray
  // (TABLE_BORDER) this function used for every row before.
  const headerBorder = { type: "solid", color: PURPLE, pt: 1 };
  const bodyBorder = { type: "solid", color: TABLE_BORDER_BODY, pt: 1 };
  const headerRow = chart.columns.map(c => ({
    text: c, options: { bold: true, color: WHITE, fill: { color: PURPLE }, align: "left", valign: "middle", fontFace: "Arial", fontSize: 14, border: headerBorder }
  }));
  const bodyRows = [];
  for (let r = 0; r < nRows; r++) {
    const zebra = r % 2 === 1;
    const cells = chart.columns.map((_, ci) => {
      const cell = (chart.rows && chart.rows[r]) ? chart.rows[r][ci] : "";
      // A cell can be a plain string, or an array of strings if the
      // source lesson's own chart content is bulleted (Ashley's review
      // question: "does the bulleted formatting need to carry over from
      // the lesson?" -- yes, when it's genuinely there). Matches the
      // same array-cell pattern already proven in
      // generate_deck_language.js's planner table renderer.
      const text = Array.isArray(cell)
        ? cell.map(item => ({ text: item, options: { bullet: { code: "2022" }, breakLine: true } }))
        : cell;
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

// ===== Slide: Cover (page 1 of the real template -- RedThread wordmark +
// wave graphic, with this lesson's Grade/Unit substituted into the
// template's own text run before rendering, so the graphic is pixel-exact
// to the source rather than hand-recreated) =====
const coverImagePath = process.argv[4];
if (coverImagePath && fs.existsSync(coverImagePath)) {
  const s = pres.addSlide();
  s.addImage({ path: coverImagePath, x: 0, y: 0, w: PAGE_W, h: PAGE_H });
}

// ===== Slide: Lesson title (page 2 equivalent) =====
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

// ===== Slide: EQ / TP / LG -- blank per spec. Matches the real template's
// actual layout: EQ full-width on top, then Teaching Point and Language
// Goal SIDE BY SIDE below (not stacked -- confirmed directly against the
// template's own shape positions, since the earlier stacked version left
// a lot of dead space at the bottom of the slide) =====
{
  const s = pres.addSlide();
  addHeaderGradient(s);
  const colGap = 0.2;
  const colW = (PAGE_W - 1.1 - colGap) / 2;
  const rowTop = 2.15, rowBottom = PAGE_H - 0.8;
  const rowH = rowBottom - rowTop;

  // Essential Question box (full width). Corner radius confirmed from
  // template: adj=8333 on an h=1.25in box at this file's 10"-reference
  // scale -> radius = 0.08333 * 1.25 * (PAGE_W/10) = ~0.139in. No border
  // in the template -- curved corners only.
  s.addShape("roundRect", { x: 0.55, y: 0.55, w: 0.06, h: 1.35, rectRadius: 0.03, fill: { color: PURPLE }, line: { type: "none" } });
  s.addShape("roundRect", { x: 0.61, y: 0.55, w: PAGE_W - 1.2, h: 1.35, rectRadius: 0.139, fill: { color: CREAM_YELLOW }, line: { type: "none" } });
  s.addText("Essential Question", { x: 0.85, y: 0.7, w: PAGE_W - 1.6, h: 0.3, fontFace: "Arial", fontSize: 18.5, bold: true, color: PURPLE, margin: 0 });
  if (lesson.essential_question) {
    s.addText(parseInlineMarkup(lesson.essential_question), { x: 0.85, y: 1.05, w: PAGE_W - 1.6, h: 0.8, fontFace: "Arial", fontSize: 24, bold: true, color: NAVY_INK, margin: 0, valign: "top" });
  }

  // Teaching Point box (left column). Corner radius confirmed from
  // template: adj=3703 on this box's actual dimensions -> ~0.139in at
  // this file's scale (works out to the same absolute radius as the EQ
  // box above, despite the different adj% -- the template appears to
  // target one consistent absolute corner radius across box sizes).
  // Border color FFAB7B (orange) confirmed directly from the template --
  // this was BLUE_PALE before, which didn't match anything in the source.
  s.addShape("roundRect", { x: 0.55, y: rowTop, w: colW, h: 0.05, rectRadius: 0.025, fill: { color: CORAL }, line: { type: "none" } });
  s.addShape("roundRect", { x: 0.55, y: rowTop + 0.05, w: colW, h: rowH - 0.05, rectRadius: 0.139, fill: { color: WHITE }, line: { color: TP_BORDER_ORANGE, width: 1.5 } });
  s.addText("Teaching Point", { x: 0.75, y: rowTop + 0.25, w: colW - 0.4, h: 0.3, fontFace: "Arial", fontSize: 18.5, bold: true, color: CORAL, margin: 0 });
  if (lesson.teaching_point) {
    s.addText(parseInlineMarkup(lesson.teaching_point), { x: 0.75, y: rowTop + 0.6, w: colW - 0.4, h: rowH - 0.8, fontFace: "Arial", fontSize: 24, color: BODY, margin: 0, valign: "top" });
  }

  // Language Goal box (right column). Corner radius same 0.139in as
  // above. Border color is NOT independently confirmed -- see
  // LG_BORDER_LIGHT_PURPLE's definition above.
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

// ===== Slides: Engage / Launch sections (variable) =====
const DECORATIVE_BG = require("path").join(__dirname, "assets", "decorative_bg.png");
function addDecorativeBg(slide) {
  // Scaled from the template's own position (left=5.99", top=0, w=4.01",
  // h=3.46" on its 10"-wide canvas) to this 13.3"-wide canvas -- confirmed
  // this specific decoration only appears on Launch Vocabulary and the
  // three detail-prompt slides in the real template, not on every slide.
  const scale = PAGE_W / 10;
  slide.addImage({ path: DECORATIVE_BG, x: 5.99 * scale, y: 0, w: 4.01 * scale, h: 3.46 * scale });
}

// Header gradient band (white -> pale yellow), pptxgenjs's addShape only
// supports 'none'/'solid' fills -- no gradient option exists in its API
// (checked directly against node_modules/pptxgenjs/types/index.d.ts,
// ShapeFillProps.type is typed as 'none' | 'solid' only) -- so this is a
// pre-rendered PNG asset, same pattern as DECORATIVE_BG above, rather
// than an addShape gradient fill that would silently render with no
// visible fill at all.
//
// Position/size scaled from the template's own values (left=0, top=0,
// w=10", h=1.2578" on its 10"-wide canvas) using the same `PAGE_W / 10`
// convention as addDecorativeBg above, for consistency with the rest of
// this file's template-to-canvas translation.
//
// Confirmed directly against the template's raw XML on exactly 4 slide
// types: the Essential Question/Teaching Point/Language Goal slide,
// Engage Vocabulary, Engage Resource, and Launch Chart -- NOT on Launch
// Vocabulary or any other slide type. The template's shape has rot=180degrees,
// which flips the gradient direction -- rendered result is pale yellow
// at the very top of the slide, fading to white by the bottom edge of
// the 1.26"-tall band. header_gradient.png is pre-rendered in that final
// (post-rotation) direction already, so no rotation is needed here.
// Catherine's explicit rule (Slack): "vocabulary for a section should
// always be the last slide of that section" -- see
// generate_deck_literature_response.js's identical fix for the full
// reasoning (same rule, same underlying problem: extract_lesson.py's
// extraction order tends to mirror the source document's own scattered
// inline definition order, not the order slides should render in).
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
    if (section.chart) {
      const s = pres.addSlide();
      if (section.section_name === "Launch") addHeaderGradient(s);
      slideTitle(s, `${section.section_name} Chart`, false);
      let y = 1.15;
      if (section.read_directions) {
        s.addText(parseInlineMarkup(section.read_directions), { x: 0.55, y, w: PAGE_W - 1.1, h: 0.4, fontFace: "Arial", fontSize: 21.5, color: BODY, margin: 0, valign: "top" });
        y += 0.5;
      }
      addChart(s, section.chart, 0.55, y, PAGE_W - 1.1, PAGE_H - y - 0.6);
      footer(s, pageNum++, false);
    } else if (section.read_directions) {
      const s = pres.addSlide();
      slideTitle(s, `${section.section_name} Read Directions`, false);
      s.addText(parseInlineMarkup(section.read_directions), { x: 0.55, y: 1.5, w: PAGE_W - 1.1, h: 1, fontFace: "Arial", fontSize: 21.5, color: BODY, margin: 0, valign: "top" });
      footer(s, pageNum++, false);
    }
    // Was a genuine gap, not a gradient-placement question: this
    // generator never rendered resource_unavailable at all, even though
    // extract_lesson.py's schema already supports it generically per
    // section (matching generate_deck_literature_response.js's existing
    // "Launch Resource" pattern exactly). The real template's slide5.xml
    // ("Engage | Resource") is the only confirmed gradient+resource
    // combination across all 18 template slides -- gradient is applied
    // here only when section_name is "Engage" for that reason; other
    // section names still render this slide (the field itself isn't
    // Engage-specific), just without the gradient, since that combination
    // was never confirmed against the template.
    if (section.resource_unavailable) {
      const s = pres.addSlide();
      if (section.section_name === "Engage") addHeaderGradient(s);
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
        if (section.section_name === "Engage") addHeaderGradient(s);
        if (section.section_name === "Launch") addDecorativeBg(s);
        slideTitle(s, `${section.section_name} Vocabulary${i > 0 ? " (continued)" : ""}`, false);
        vocabBlock(s, words, 0.55, 1.25, PAGE_W - 1.1, false);
        footer(s, pageNum++, false);
      });
    }
  }
);

// ===== Slide: Independent Reading OR Shared Analysis (two-column layout:
// text on the left, chart on the right, matching the template exactly) =====
const ir = lesson.independent_reading, sa = lesson.shared_analysis;
if (ir || sa) {
  const block = ir || sa;
  const title = ir ? "Independent Reading" : "Shared Analysis";
  const s = pres.addSlide();
  slideTitle(s, title, false);
  const leftW = 6.6;
  let y = 1.1;
  s.addText("Teaching Point", { x: 0.55, y, w: leftW, h: 0.35, fontFace: "Arial", fontSize: 20, bold: true, color: CORAL, margin: 0 });
  s.addText(parseInlineMarkup(lesson.teaching_point), { x: 0.55, y: y + 0.4, w: leftW, h: 1.15, fontFace: "Arial", fontSize: 16, color: BODY, margin: 0, valign: "top" });
  y += 1.65;
  if (block.read_directions) {
    s.addShape("roundRect", { x: 0.55, y, w: leftW, h: 1.0, rectRadius: 0.06, fill: { color: TAN_BG }, line: { type: "none" } });
    s.addText(parseInlineMarkup(block.read_directions), { x: 0.75, y: y + 0.1, w: leftW - 0.4, h: 0.8, fontFace: "Arial", fontSize: 16, color: NAVY_INK, margin: 0, valign: "top" });
    y += 1.15;
  }
  if (block.vocabulary && block.vocabulary.length) {
    s.addText("IMPORTANT VOCABULARY", { x: 0.55, y, w: leftW, h: 0.35, fontFace: "Arial", fontSize: 20, bold: true, color: CORAL, margin: 0 });
    y += 0.42;
    block.vocabulary.forEach(v => {
      s.addShape("ellipse", { x: 0.6, y: y + 0.1, w: 0.11, h: 0.11, fill: { color: VIOLET }, line: { type: "none" } });
      s.addText([{ text: `${v.word}: `, options: { bold: true, color: NAVY_INK } }, { text: v.definition, options: { color: BODY } }],
        { x: 0.85, y, w: leftW - 0.3, h: 0.65, fontFace: "Arial", fontSize: 16, margin: 0, valign: "top" });
      y += 0.66;
    });
  }
  if (block.chart) {
    addChart(s, block.chart, 7.4, 1.1, PAGE_W - 7.95, PAGE_H - 1.7);
  }
  footer(s, pageNum++, false);
}

// ===== Slide: Quick Write & Discourse Clubs -- pale pink dashed boxes =====
if (lesson.quick_write_prompt || lesson.discourse_club_prompt) {
  const s = pres.addSlide();
  const hasBoth = lesson.quick_write_prompt && lesson.discourse_club_prompt;
  slideTitle(s, hasBoth ? "Quick Write & Discourse Clubs" : (lesson.quick_write_prompt ? "Quick Write" : "Discourse Clubs"), false);
  const boxes = [];
  if (lesson.quick_write_prompt) boxes.push({ label: "QUICK WRITE", prompt: lesson.quick_write_prompt });
  if (lesson.discourse_club_prompt) boxes.push({ label: "DISCOURSE CLUBS", prompt: lesson.discourse_club_prompt });
  const boxW = boxes.length === 2 ? (PAGE_W - 1.3) / 2 : PAGE_W - 1.1;
  boxes.forEach((b, i) => {
    const bx = 0.55 + i * (boxW + 0.2);
    s.addText(b.label, { x: bx, y: 1.2, w: boxW, h: 0.35, fontFace: "Arial", fontSize: 22, bold: true, color: CORAL, margin: 0 });
    s.addShape("roundRect", { x: bx, y: 1.6, w: boxW, h: 2.6, rectRadius: 0.06, fill: { color: PINK_PALE }, line: { color: PINK_BORDER, width: 1.5, dashType: "dash" } });
    s.addText("Prompt", { x: bx + 0.3, y: 2.0, w: boxW - 0.6, h: 0.35, fontFace: "Arial", fontSize: 24, bold: true, color: PURPLE, margin: 0 });
    s.addText(parseInlineMarkup(b.prompt), { x: bx + 0.3, y: 2.4, w: boxW - 0.6, h: 1.2, fontFace: "Arial", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });
  });
  footer(s, pageNum++, false);
}

// ===== Slides: individual Discourse Clubs / Quick Write detail slides --
// the template shows these in addition to (not instead of) the combined
// intro slide above: an overview first, then each one elaborated with an
// icon card and its own prompt box. Previously missing entirely. =====
const DETAIL_ICON_DIR = require("path").join(__dirname, "assets", "detail_icons");
if (lesson.discourse_club_prompt) {
  const s = pres.addSlide();
  addDecorativeBg(s);
  slideTitle(s, "Discourse Clubs", false);
  detailPromptSlide(s, {
    cardColor: PINK_PALE, accentColor: CORAL,
    boxBorderColor: TP_BORDER_ORANGE, // Ashley's comment explicitly says "light orange outline" for this box
    iconPath: require("path").join(DETAIL_ICON_DIR, "discourse_clubs.png"),
    cardTitle: "Discourse Clubs",
    description: "Collaborate with your club to discuss and deepen your understanding of the text.",
    promptLabel: "Prompt", promptText: lesson.discourse_club_prompt,
  });
  footer(s, pageNum++, false);
}
if (lesson.quick_write_prompt) {
  const s = pres.addSlide();
  addDecorativeBg(s);
  slideTitle(s, "Quick Write", false);
  detailPromptSlide(s, {
    cardColor: "FFAB7B", accentColor: PINK_BORDER,
    iconPath: require("path").join(DETAIL_ICON_DIR, "quick_write.png"),
    cardTitle: "Quick Write",
    description: "Take a few minutes to write down your initial thoughts and response to the prompt.",
    promptLabel: "Prompt", promptText: lesson.quick_write_prompt,
  });
  footer(s, pageNum++, false);
}

// ===== Slide: Whole-Class Discourse -- icon card + purple accent bar,
// matching the template exactly (this was previously using the wrong
// pink-dashed style, which only belongs on the combined intro slide) =====
if (lesson.whole_class_discourse_prompt) {
  const s = pres.addSlide();
  addDecorativeBg(s);
  slideTitle(s, "Whole-Class Discourse", false);
  detailPromptSlide(s, {
    cardColor: "ECEAF9", accentColor: PURPLE,
    iconPath: require("path").join(__dirname, "assets", "detail_icons", "whole_class_discourse.png"),
    cardTitle: "Whole-Class Discourse",
    description: "Share your ideas, listen to your classmates, and build on each other's thinking.",
    promptLabel: "Prompt", promptText: lesson.whole_class_discourse_prompt,
  });
  footer(s, pageNum++, false);
}

// ===== Slide: Closing -- EQ box (blue, left bar) + TP box (white, coral
// top bar), no vocabulary here =====
{
  const s = pres.addSlide();
  slideTitle(s, "Closing", false);
  s.addShape("roundRect", { x: 0.55, y: 1.15, w: 0.06, h: 1.35, rectRadius: 0.03, fill: { color: PURPLE }, line: { type: "none" } });
  s.addShape("roundRect", { x: 0.61, y: 1.15, w: PAGE_W - 1.2, h: 1.35, rectRadius: 0.139, fill: { color: CREAM_YELLOW }, line: { type: "none" } });
  s.addText("Essential Question", { x: 0.85, y: 1.3, w: PAGE_W - 1.6, h: 0.3, fontFace: "Arial", fontSize: 22, bold: true, color: PURPLE, margin: 0 });
  s.addText(parseInlineMarkup(lesson.essential_question), { x: 0.85, y: 1.65, w: PAGE_W - 1.6, h: 0.75, fontFace: "Arial", fontSize: 24, bold: true, color: NAVY_INK, margin: 0, valign: "top" });

  // Corner radius and border color confirmed from the template's Closing
  // slide specifically -- this box is full-width here (no Language Goal
  // column), and uses a lighter border (DCD6EE) than the split-column
  // version on the EQ/TP/LG slide (which uses orange).
  s.addShape("roundRect", { x: 0.55, y: 2.75, w: PAGE_W - 1.1, h: 0.05, rectRadius: 0.025, fill: { color: CORAL }, line: { type: "none" } });
  s.addShape("roundRect", { x: 0.55, y: 2.8, w: PAGE_W - 1.1, h: 2.4, rectRadius: 0.113, fill: { color: WHITE }, line: { color: TP_BORDER_CLOSING, width: 1 } });
  s.addText("Teaching Point", { x: 0.75, y: 3.0, w: PAGE_W - 1.5, h: 0.35, fontFace: "Arial", fontSize: 22, bold: true, color: CORAL, margin: 0 });
  s.addText(parseInlineMarkup(lesson.teaching_point), { x: 0.75, y: 3.4, w: PAGE_W - 1.5, h: 1.7, fontFace: "Arial", fontSize: 24, color: NAVY_INK, margin: 0, valign: "top" });

  footer(s, pageNum++, false);
}

// ===== Slides: Lesson Vocabulary Review -- purple background bookend,
// only words repeated in Closing =====
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
