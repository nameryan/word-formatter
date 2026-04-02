import {
  W_NS, parseXml, serializeXml,
  getEls, getFirstEl, createWEl, setWAttr,
  ensureChild, replaceChild, removeDirectChild
} from "./xml-parser.js";

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

function getDirectChild(parent, localName) {
  for (const child of parent.childNodes) {
    if (child.nodeType === 1 && child.localName === localName && child.namespaceURI === W_NS) {
      return child;
    }
  }
  return null;
}

function getParaStyleId(para) {
  const pPr = getFirstEl(para, "pPr");
  if (!pPr) return null;
  const pStyle = getFirstEl(pPr, "pStyle");
  if (!pStyle) return null;
  return pStyle.getAttributeNS(W_NS, "val") || pStyle.getAttribute("w:val") || null;
}

function removeAllDirectChildren(parent, localName) {
  for (const child of Array.from(parent.childNodes)) {
    if (child.nodeType === 1 && child.localName === localName && child.namespaceURI === W_NS) {
      child.remove();
    }
  }
}

function patchParaSpacing(para, doc, { lineValue, lineRule, beforePt, afterPt, indent }) {
  const pPr = ensureChild(para, "pPr", doc, para.firstChild || null);

  removeAllDirectChildren(pPr, "spacing");
  const spacingEl = createWEl(doc, "spacing");
  if (lineValue != null) {
    setWAttr(spacingEl, "line", String(lineValue));
    setWAttr(spacingEl, "lineRule", lineRule || "auto");
  }
  if (beforePt != null) setWAttr(spacingEl, "before", String(Math.round(beforePt * 20)));
  if (afterPt != null)  setWAttr(spacingEl, "after",  String(Math.round(afterPt * 20)));
  pPr.appendChild(spacingEl);

  removeAllDirectChildren(pPr, "contextualSpacing");

  const hasNumPr = !!getFirstEl(pPr, "numPr");
  const styleId = getParaStyleId(para) || "";
  const isSpecialStyle = /^(TOC|目录|toc|Bibliography|参考文献|Caption|题注|Index|索引|Footnote|脚注|Endnote|尾注|Header|页眉|Footer|页脚)/i.test(styleId);
  if (indent === false || hasNumPr || isInTableCell(para) || isSpecialStyle) return;

  let indEl = getDirectChild(pPr, "ind");
  if (indent && indent.twips > 0) {
    if (!indEl) {
      indEl = createWEl(doc, "ind");
      pPr.appendChild(indEl);
    }
    indEl.removeAttribute("w:hanging");
    indEl.removeAttributeNS(W_NS, "hanging");
    indEl.removeAttribute("w:firstLineChars");
    indEl.removeAttributeNS(W_NS, "firstLineChars");
    indEl.setAttribute("w:firstLine", String(indent.twips));
  } else if (indent === null && indEl) {
    indEl.removeAttribute("w:firstLine");
    indEl.removeAttributeNS(W_NS, "firstLine");
    indEl.removeAttribute("w:firstLineChars");
    indEl.removeAttributeNS(W_NS, "firstLineChars");
    if (indEl.attributes.length === 0) indEl.remove();
  }
}

function isInTableCell(para) {
  let node = para.parentNode;
  while (node) {
    if (node.nodeType === 1 && node.namespaceURI === W_NS) {
      if (node.localName === "tc") return true;
      if (node.localName === "body") return false;
    }
    node = node.parentNode;
  }
  return false;
}

function getHeadingLevel(styleId) {
  const match = styleId && styleId.match(/^(?:Heading|标题)\s*([1-6])$/i);
  return match ? parseInt(match[1], 10) : 0;
}

function hasPageBreakBefore(para) {
  const pPr = getFirstEl(para, "pPr");
  return pPr ? !!getFirstEl(pPr, "pageBreakBefore") : false;
}

function hasExplicitPageBreak(para) {
  const runs = getEls(para, "r");
  for (const run of runs) {
    const br = getFirstEl(run, "br");
    if (!br) continue;
    const type = br.getAttribute("w:type") || br.getAttributeNS(W_NS, "type");
    if (type === "page") return true;
  }
  return false;
}

function hasLastRenderedPageBreak(para) {
  const runs = getEls(para, "r");
  return runs.some(run => !!getFirstEl(run, "lastRenderedPageBreak"));
}

function getParaSectPr(para) {
  const pPr = getFirstEl(para, "pPr");
  return pPr ? getFirstEl(pPr, "sectPr") : null;
}

function getSectionBreakType(sectPr) {
  if (!sectPr) return null;
  const typeEl = getDirectChild(sectPr, "type");
  const value = typeEl
    ? (typeEl.getAttributeNS(W_NS, "val") || typeEl.getAttribute("w:val") || "")
    : "";
  return value || "nextPage";
}

function endsWithSectionPageBreak(para) {
  const sectPr = getParaSectPr(para);
  if (!sectPr) return false;
  return getSectionBreakType(sectPr) !== "continuous";
}

function getParaText(para) {
  return getEls(para, "t").map(t => t.textContent || "").join("");
}

function applyRunFormatting(rPr, doc, font, sizePt, bold, color) {
  if (font) replaceChild(rPr, "rFonts", buildRFonts(doc, font));
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

function stripHeadingNumberPrefix(para) {
  const textRuns = [];
  for (const run of getEls(para, "r")) {
    if (getFirstEl(run, "instrText")) continue;
    const tEl = getFirstEl(run, "t");
    if (tEl) textRuns.push({ run, tEl });
  }

  const fullText = textRuns.map(item => item.tEl.textContent).join("");
  const match = fullText.match(
    /^(?:(?:\d+\.)*\d+(?:[\s\.\u3000、，,\uff0c]+|(?=[\u4e00-\u9fff\uff00-\uffef]))|第[\d○零一二三四五六七八九十百千]+[章节条款篇部附]\s*[\s\u3000]*|[A-Za-z][\.\)]\s*|[（(]\d+[)）]\s*|[一二三四五六七八九十]+[、\.\s\u3000]+)/
  );
  if (!match) return;

  let charsLeft = match[0].length;
  for (const item of textRuns) {
    if (charsLeft <= 0) break;
    const text = item.tEl.textContent;
    if (text.length <= charsLeft) {
      charsLeft -= text.length;
      item.tEl.textContent = "";
      if (!item.tEl.textContent && item.run.parentNode) item.run.parentNode.removeChild(item.run);
    } else {
      item.tEl.textContent = text.slice(charsLeft);
      charsLeft = 0;
    }
  }
}

function applyHeadingNumbering(doc, headings, config) {
  const counters = [0, 0, 0, 0, 0, 0];
  const startNum = config.headingNumberingStart || 1;
  let firstH1 = true;

  for (const { para, level } of headings) {
    const idx = level - 1;
    const headingText = getParaText(para);
    if (!headingText || !headingText.trim()) continue;
    if (firstH1 && idx > 0) continue;
    if (idx === 0 && firstH1) {
      counters[0] = startNum - 1;
      firstH1 = false;
    }
    counters[idx]++;
    for (let i = idx + 1; i < 6; i++) counters[i] = 0;

    const parts = counters.slice(0, idx + 1);
    let numberText;
    if (config.headingNumbering === "chinese" && idx === 0) {
      numberText = "第" + parts[0] + "章";
    } else {
      numberText = parts.join(".");
    }
    numberText += "\u3000";

    stripHeadingNumberPrefix(para);

    const numberRun = createWEl(doc, "r");
    numberRun.appendChild(createWEl(doc, "rPr"));
    const textEl = createWEl(doc, "t");
    textEl.setAttribute("xml:space", "preserve");
    textEl.textContent = numberText;
    numberRun.appendChild(textEl);

    const pPr = getFirstEl(para, "pPr");
    para.insertBefore(numberRun, pPr ? pPr.nextSibling : para.firstChild);
  }
}

function patchMargins(docXml, margins) {
  const CM_TO_TWIPS = 567;
  const sectPrs = getEls(docXml, "sectPr");
  for (const sectPr of sectPrs) {
    removeAllDirectChildren(sectPr, "pgMar");
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
}

export function patchDocument(docXmlString, config) {
  const doc = parseXml(docXmlString);
  const paragraphs = getEls(doc, "p");
  const headingsToNumber = [];

  let pastFirstPage = !config.skipCover;

  for (const para of paragraphs) {
    const styleId = getParaStyleId(para) || "Normal";
    const headingLevel = getHeadingLevel(styleId);
    const isHeading = headingLevel > 0;

    if (!pastFirstPage) {
      if (hasPageBreakBefore(para)) {
        pastFirstPage = true;
      } else if (hasExplicitPageBreak(para) || hasLastRenderedPageBreak(para) || endsWithSectionPageBreak(para)) {
        pastFirstPage = true;
        continue;
      } else {
        if (!isHeading) {
          patchParaSpacing(para, doc, {
            lineValue: config.body.lineValue,
            lineRule: config.body.lineRule,
            beforePt: config.body.beforePt,
            afterPt: config.body.afterPt,
            indent: false,
          });
        }
        continue;
      }
    }

    if (isHeading) {
      const headingConfig = config.headings[headingLevel - 1] || {};
      patchParaSpacing(para, doc, {
        lineValue: headingConfig.lineValue || config.body.lineValue,
        lineRule: headingConfig.lineRule || config.body.lineRule,
        beforePt: headingConfig.beforePt != null ? headingConfig.beforePt : 6,
        afterPt: headingConfig.afterPt != null ? headingConfig.afterPt : 6,
        indent: false,
      });
      if (config.headingNumbering && config.headingNumbering !== "none") {
        const pPr = ensureChild(para, "pPr", doc, para.firstChild || null);
        removeAllDirectChildren(pPr, "numPr");
        headingsToNumber.push({ para, level: headingLevel, text: getParaText(para) });
      }
    } else {
      patchParaSpacing(para, doc, {
        lineValue: config.body.lineValue,
        lineRule: config.body.lineRule,
        beforePt: config.body.beforePt,
        afterPt: config.body.afterPt,
        indent: config.body.indent,
      });
    }

    for (const run of getEls(para, "r")) {
      if (getFirstEl(run, "instrText")) continue;

      const rPr = ensureChild(run, "rPr", doc, run.firstChild || null);
      if (!isHeading) {
        applyRunFormatting(rPr, doc, config.body.font, config.body.sizePt, null, null);
      } else {
        const headingConfig = config.headings[headingLevel - 1] || {};
        applyRunFormatting(rPr, doc, headingConfig.font, headingConfig.sizePt, headingConfig.bold !== false, headingConfig.color || null);
        removeDirectChild(rPr, "rStyle");
      }
    }
  }

  if (config.headingNumbering && config.headingNumbering !== "none") {
    applyHeadingNumbering(doc, headingsToNumber, config);
  }

  if (config.margins) {
    patchMargins(doc, config.margins);
  }

  return serializeXml(doc);
}
