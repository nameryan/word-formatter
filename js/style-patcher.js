import {
  W_NS, parseXml, serializeXml,
  getEls, getFirstEl, createWEl, setWAttr,
  ensureChild, replaceChild, removeDirectChild
} from "./xml-parser.js";

// Chinese document font size map (pt values)
export const FONT_SIZES = {
  "初号": 42, "小初": 36,
  "一号": 26, "小一": 24,
  "二号": 22, "小二": 18,
  "三号": 16, "小三": 15,
  "四号": 14, "小四": 12,
  "五号": 10.5, "小五": 9,
};

// English <-> Chinese heading style ID mapping
const HEADING_IDS = {
  "Heading1": ["Heading1", "标题1", "标题 1"],
  "Heading2": ["Heading2", "标题2", "标题 2"],
  "Heading3": ["Heading3", "标题3", "标题 3"],
  "Heading4": ["Heading4", "标题4", "标题 4"],
  "Heading5": ["Heading5", "标题5", "标题 5"],
  "Heading6": ["Heading6", "标题6", "标题 6"],
};

const NORMAL_IDS = ["Normal", "正文", "默认段落字体"];

/**
 * Find a <w:style> node by any of the candidate style IDs.
 */
function findStyle(stylesDoc, candidates) {
  for (const styleEl of getEls(stylesDoc, "style")) {
    const idAttr = styleEl.getAttributeNS(W_NS, "styleId")
      || styleEl.getAttribute("w:styleId");
    if (candidates.includes(idAttr)) return styleEl;
  }
  return null;
}

/**
 * Get or create a <w:style> node for the given canonical styleId.
 * If creating, uses the first candidate as the styleId value.
 */
function getOrCreateStyle(stylesDoc, canonicalId, candidates, type = "paragraph") {
  let styleEl = findStyle(stylesDoc, candidates);
  if (!styleEl) {
    styleEl = createWEl(stylesDoc, "style");
    setWAttr(styleEl, "type", type);
    setWAttr(styleEl, "styleId", canonicalId);
    const stylesRoot = getFirstEl(stylesDoc, "styles") || stylesDoc.documentElement;
    stylesRoot.appendChild(styleEl);
  }
  return styleEl;
}

/**
 * Build a <w:rFonts> element with all four font attributes set.
 */
function buildRFonts(doc, font) {
  const el = createWEl(doc, "rFonts");
  setWAttr(el, "ascii", font);
  setWAttr(el, "hAnsi", font);
  setWAttr(el, "eastAsia", font);
  setWAttr(el, "cs", font);
  return el;
}

/**
 * Build <w:sz> or <w:szCs> element (pt → half-points).
 */
function buildSz(doc, localName, pt) {
  const el = createWEl(doc, localName);
  setWAttr(el, "val", String(Math.round(pt * 2)));
  return el;
}

/**
 * Patch the rPr of a style element.
 */
function patchRpr(styleEl, doc, { font, sizePt, bold, color }) {
  const rPr = ensureChild(styleEl, "rPr", doc);

  if (font) {
    replaceChild(rPr, "rFonts", buildRFonts(doc, font));
  }
  if (sizePt) {
    replaceChild(rPr, "sz", buildSz(doc, "sz", sizePt));
    replaceChild(rPr, "szCs", buildSz(doc, "szCs", sizePt));
  }

  removeDirectChild(rPr, "b");
  removeDirectChild(rPr, "bCs");
  if (bold) {
    rPr.appendChild(createWEl(doc, "b"));
    rPr.appendChild(createWEl(doc, "bCs"));
  }

  removeDirectChild(rPr, "color");
  if (color && color !== "#000000" && color !== "000000") {
    const colorEl = createWEl(doc, "color");
    setWAttr(colorEl, "val", color.replace("#", ""));
    rPr.appendChild(colorEl);
  }
}

/**
 * Patch the pPr spacing of a style element.
 * lineValue: 240=single, 360=1.5x, 480=double, or custom twips
 * beforePt / afterPt: paragraph spacing in pt (converted to twips = pt*20)
 */
function patchPpr(styleEl, doc, { lineValue, lineRule, beforePt, afterPt, indent }) {
  const pPr = ensureChild(styleEl, "pPr", doc);

  // Remove existing spacing, rebuild with merged attributes
  removeDirectChild(pPr, "spacing");

  const spacingEl = createWEl(doc, "spacing");
  if (lineValue != null) {
    setWAttr(spacingEl, "line", String(lineValue));
    setWAttr(spacingEl, "lineRule", lineRule || "auto");
  }
  if (beforePt != null) setWAttr(spacingEl, "before", String(Math.round(beforePt * 20)));
  if (afterPt != null)  setWAttr(spacingEl, "after",  String(Math.round(afterPt * 20)));
  pPr.appendChild(spacingEl);

  // Remove contextual spacing for cleaner output
  removeDirectChild(pPr, "contextualSpacing");

  if (indent) {
    removeDirectChild(pPr, "ind");
    const indEl = createWEl(doc, "ind");
    if (indent.firstLine != null) setWAttr(indEl, "firstLine", String(Math.round(indent.firstLine * 20)));
    if (indent.left != null)      setWAttr(indEl, "left",      String(Math.round(indent.left * 20)));
    pPr.appendChild(indEl);
  }

  return pPr;
}

/**
 * Main entry point.
 * config = {
 *   body: { font, sizePt, lineValue, lineRule, beforePt, afterPt },
 *   headings: [ { font, sizePt, bold, color, beforePt, afterPt }, ... ]  // index 0 = H1
 * }
 */
export function patchStyles(stylesXmlString, config) {
  const doc = parseXml(stylesXmlString);

  // --- Patch Normal / body style ---
  const normalEl = getOrCreateStyle(doc, "Normal", NORMAL_IDS);
  patchPpr(normalEl, doc, {
    lineValue: config.body.lineValue,
    lineRule: config.body.lineRule,
    beforePt: config.body.beforePt,
    afterPt: config.body.afterPt,
  });

  // --- Patch Heading1-6 ---
  const headingKeys = Object.keys(HEADING_IDS); // ["Heading1" ... "Heading6"]
  config.headings.forEach((hCfg, i) => {
    if (!hCfg) return;
    const canonicalId = headingKeys[i];
    const candidates = HEADING_IDS[canonicalId];
    const styleEl = getOrCreateStyle(doc, canonicalId, candidates);

    // Ensure name element exists
    let nameEl = getFirstEl(styleEl, "name");
    if (!nameEl) {
      nameEl = createWEl(doc, "name");
      setWAttr(nameEl, "val", canonicalId.toLowerCase().replace("heading", "heading "));
      styleEl.insertBefore(nameEl, styleEl.firstChild);
    }

    const headingPPr = patchPpr(styleEl, doc, {
      lineValue: hCfg.lineValue || config.body.lineValue,
      lineRule: hCfg.lineRule || config.body.lineRule,
      beforePt: hCfg.beforePt != null ? hCfg.beforePt : 6,
      afterPt: hCfg.afterPt != null ? hCfg.afterPt : 6,
    });
    if (config.headingNumbering && config.headingNumbering !== "none") {
      removeDirectChild(headingPPr, "numPr");
    }
  });

  return serializeXml(doc);
}
