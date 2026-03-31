import {
  W_NS, parseXml, serializeXml,
  getEls, getFirstEl, createWEl, setWAttr,
  ensureChild, replaceChild, removeDirectChild
} from "./xml-parser.js";

// All known heading style ID variants (English + Chinese)
const HEADING_STYLE_PATTERN = /^(Heading\s*[1-6]|标题\s*[1-6]|heading\s*[1-6])$/i;

// List-related style IDs — skip pPr modification for these
const LIST_STYLE_PATTERN = /^(List|ListParagraph|List Paragraph|列表段落)/i;

function buildRFonts(doc, font) {
  const el = createWEl(doc, "rFonts");
  setWAttr(el, "ascii", font);
  setWAttr(el, "hAnsi", font);
  setWAttr(el, "eastAsia", font);
  setWAttr(el, "cs", font);
  return el;
}

function buildSz(doc, localName, pt) {
  const el = createWEl(doc, localName);
  setWAttr(el, "val", String(Math.round(pt * 2)));
  return el;
}

/**
 * Get the style ID of a paragraph (from its w:pPr > w:pStyle).
 */
function getParaStyleId(para) {
  const pPr = getFirstEl(para, "pPr");
  if (!pPr) return null;
  const pStyle = getFirstEl(pPr, "pStyle");
  if (!pStyle) return null;
  return pStyle.getAttributeNS(W_NS, "val") || pStyle.getAttribute("w:val") || null;
}

/**
 * Patch spacing on a paragraph's pPr (not style, but direct paragraph spacing).
 */
function patchParaSpacing(para, doc, { lineValue, lineRule, beforePt, afterPt }) {
  const pPr = ensureChild(para, "pPr", doc, getFirstEl(para, "r") || null);

  // Get existing spacing element to merge attributes
  removeDirectChild(pPr, "spacing");
  const spacingEl = createWEl(doc, "spacing");
  if (lineValue != null) {
    setWAttr(spacingEl, "line", String(lineValue));
    setWAttr(spacingEl, "lineRule", lineRule || "auto");
  }
  if (beforePt != null) setWAttr(spacingEl, "before", String(Math.round(beforePt * 20)));
  if (afterPt != null)  setWAttr(spacingEl, "after",  String(Math.round(afterPt * 20)));
  pPr.appendChild(spacingEl);

  // Remove contextual spacing
  removeDirectChild(pPr, "contextualSpacing");
}

/**
 * Patch all runs in a paragraph:
 * - For body paragraphs: stamp font + size directly onto each run's rPr
 * - For heading paragraphs: strip font/size overrides (let styles.xml take effect)
 */
function patchRuns(para, doc, { isHeading, font, sizePt }) {
  const runs = getEls(para, "r");
  for (const run of runs) {
    // Skip field instruction runs
    const instrText = getFirstEl(run, "instrText");
    if (instrText) continue;

    const rPr = ensureChild(run, "rPr", doc, getFirstEl(run, "t") || null);

    if (!isHeading) {
      // Body: apply font and size directly
      if (font) replaceChild(rPr, "rFonts", buildRFonts(doc, font));
      if (sizePt) {
        replaceChild(rPr, "sz",   buildSz(doc, "sz", sizePt));
        replaceChild(rPr, "szCs", buildSz(doc, "szCs", sizePt));
      }
      // Remove bold from body runs (unless it's intentional inline bold — keep it)
      // Actually, we should NOT strip intentional bold from body runs.
      // Only strip if the style would have set it (but body style shouldn't have bold).
    } else {
      // Heading: remove direct font/size overrides so styles.xml definition applies
      removeDirectChild(rPr, "rFonts");
      removeDirectChild(rPr, "sz");
      removeDirectChild(rPr, "szCs");
      // Remove rStyle references that may override heading style
      removeDirectChild(rPr, "rStyle");
    }
  }
}

/**
 * Patch page margins in the document's sectPr.
 * margins: { top, bottom, left, right } in cm
 */
function patchMargins(docXml, margins) {
  const CM_TO_TWIPS = 567; // 1 cm ≈ 567 twips
  const sectPr = getFirstEl(docXml, "sectPr");
  if (!sectPr) return;

  removeDirectChild(sectPr, "pgMar");
  const pgMar = createWEl(docXml, "pgMar");
  setWAttr(pgMar, "top",    String(Math.round(margins.top    * CM_TO_TWIPS)));
  setWAttr(pgMar, "bottom", String(Math.round(margins.bottom * CM_TO_TWIPS)));
  setWAttr(pgMar, "left",   String(Math.round(margins.left   * CM_TO_TWIPS)));
  setWAttr(pgMar, "right",  String(Math.round(margins.right  * CM_TO_TWIPS)));
  setWAttr(pgMar, "header", "851");
  setWAttr(pgMar, "footer", "992");
  setWAttr(pgMar, "gutter", "0");
  sectPr.appendChild(pgMar);
}

/**
 * Main entry point.
 * config = {
 *   body: { font, sizePt, lineValue, lineRule, beforePt, afterPt },
 *   margins: null | { top, bottom, left, right }  // cm
 * }
 */
export function patchDocument(docXmlString, config) {
  const doc = parseXml(docXmlString);

  const paragraphs = getEls(doc, "p");

  for (const para of paragraphs) {
    const styleId = getParaStyleId(para) || "Normal";
    const isHeading = HEADING_STYLE_PATTERN.test(styleId);
    const isList = LIST_STYLE_PATTERN.test(styleId);

    // Patch spacing on body paragraphs (not headings — heading spacing is in styles.xml)
    if (!isHeading) {
      patchParaSpacing(para, doc, {
        lineValue: config.body.lineValue,
        lineRule: config.body.lineRule,
        beforePt: config.body.beforePt,
        afterPt: config.body.afterPt,
      });
    }

    // Patch runs
    patchRuns(para, doc, {
      isHeading,
      font: config.body.font,
      sizePt: config.body.sizePt,
    });
  }

  // Patch page margins if configured
  if (config.margins) {
    patchMargins(doc, config.margins);
  }

  return serializeXml(doc);
}
