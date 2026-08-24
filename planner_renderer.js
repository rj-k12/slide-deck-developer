const pptxgen = require("pptxgenjs");

// ===== Palette (same as existing generators) =====
const NAVY_INK = "0E0142";
const PURPLE = "3928AA";
const CORAL = "F19A65";
const MUTED = "6B6478";
const WHITE = "FFFFFF";
const BODY = "33322E";
const PEACH = "FFDBC5";
const TABLE_BORDER_BODY = "FFAB7B";
const SECTION_BG = "F4F2FC"; // light lavender, matches existing Language Goal box fill

const PAGE_W = 13.3, PAGE_H = 7.5;

/**
 * PROTOTYPE: renders a lesson's genre-specific "planner" -- a variable-
 * shape graphic organizer that differs per Writing/Project genre. See
 * scoping notes: this is the one piece that can't reuse anything from
 * generate_deck_literature_response.js, because Realistic Narrative,
 * Historical Figure Fact Card, and Personal Identity Poem planners all
 * have genuinely different shapes.
 *
 * Schema (proposed addition to extract_lesson.py's output):
 *
 *   planner_sections: [
 *     { type: "text",  label?: string, content: string },
 *     { type: "list",  label?: string, items: string[] },
 *     { type: "table", label?: string, columns: string[],
 *       rows: (string | string[])[][] }  // a cell is either a plain
 *                                         // string, or an array of
 *                                         // strings rendered as a
 *                                         // mini bulleted list inside
 *                                         // that cell (needed for the
 *                                         // Figurative Language chart's
 *                                         // "Examples" column)
 *   ]
 *
 * Layout: sections stack top-to-bottom, each auto-sized based on content
 * (a rough line-count heuristic for text/list, a fixed per-row height
 * for tables). This is a first-pass, generically-correct layout -- NOT
 * verified against a real Writing/Project template, since none exists
 * yet. Whether two short "list" sections should sit side-by-side (as
 * the actual template likely does for "Protagonist" + "Change
 * Protagonist Faces") is a real open design question -- flagged at the
 * bottom of this file rather than guessed at silently.
 */
function estimateSectionHeight(section, w) {
  if (section.type === "text") {
    const charsPerLine = w * 11; // rough estimate at 14pt Arial
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
    const rowH = 0.5;
    return 0.35 + (section.rows.length + 1) * rowH + 0.15;
  }
  return 0.5;
}

function renderTextSection(slide, section, x, y, w) {
  let cursorY = y;
  if (section.label) {
    slide.addText(section.label, { x, y: cursorY, w, h: 0.3, fontFace: "Arial", fontSize: 16, bold: true, color: PURPLE, margin: 0 });
    cursorY += 0.35;
  }
  slide.addText(section.content, { x, y: cursorY, w, h: 1, fontFace: "Arial", fontSize: 14, color: BODY, margin: 0, valign: "top" });
}

function renderListSection(slide, section, x, y, w) {
  let cursorY = y;
  if (section.label) {
    slide.addText(section.label, { x, y: cursorY, w, h: 0.3, fontFace: "Arial", fontSize: 16, bold: true, color: PURPLE, margin: 0 });
    cursorY += 0.35;
  }
  const bulletText = section.items.map(item => ({ text: item, options: { bullet: { code: "2022" }, breakLine: true } }));
  slide.addText(bulletText, { x, y: cursorY, w, h: 1.5, fontFace: "Arial", fontSize: 13, color: BODY, margin: 0, valign: "top", paraSpaceAfter: 4 });
}

function renderTableSection(slide, section, x, y, w) {
  let cursorY = y;
  if (section.label) {
    slide.addText(section.label, { x, y: cursorY, w, h: 0.3, fontFace: "Arial", fontSize: 16, bold: true, color: PURPLE, margin: 0 });
    cursorY += 0.35;
  }
  const nCols = section.columns.length;
  const headerRow = section.columns.map(c => ({
    text: c, options: { bold: true, color: WHITE, fill: { color: PURPLE }, align: "center", valign: "middle", fontFace: "Arial", fontSize: 12, border: { type: "solid", color: PURPLE, pt: 1 } }
  }));
  const bodyRows = section.rows.map((row, ri) => {
    const zebra = ri % 2 === 1;
    return row.map(cell => {
      // A cell can be a plain string, or an array of strings (a mini
      // bulleted list within the cell) -- needed for the Figurative
      // Language chart's "Examples" column.
      const cellText = Array.isArray(cell)
        ? cell.map(item => ({ text: item, options: { bullet: { code: "2022" }, breakLine: true } }))
        : cell;
      return { text: cellText, options: { color: BODY, fontFace: "Arial", fontSize: 11, valign: "top", fill: { color: zebra ? PEACH : WHITE }, border: { type: "solid", color: TABLE_BORDER_BODY, pt: 1 } } };
    });
  });
  const rowH = Math.max(0.5, (estimateSectionHeight(section, w) - 0.5) / (section.rows.length + 1));
  slide.addTable([headerRow, ...bodyRows], {
    x, y: cursorY, w, fontFace: "Arial", autoPage: false,
    rowH: [0.4, ...bodyRows.map(() => rowH)],
    colW: Array(nCols).fill(w / nCols),
  });
}

function renderPlannerSections(slide, sections, x, y, w, h) {
  let cursorY = y;
  sections.forEach(section => {
    const sectionH = estimateSectionHeight(section, w);
    if (section.type === "text") renderTextSection(slide, section, x, cursorY, w);
    else if (section.type === "list") renderListSection(slide, section, x, cursorY, w);
    else if (section.type === "table") renderTableSection(slide, section, x, cursorY, w);
    cursorY += sectionH;
  });
  return cursorY - y; // total height used, so the caller can detect overflow
}

module.exports = { renderPlannerSections, estimateSectionHeight };
