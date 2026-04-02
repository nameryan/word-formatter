// DOCX XML namespace
export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

export function parseXml(str) {
  const doc = new DOMParser().parseFromString(str, "application/xml");
  const err = doc.querySelector("parsererror");
  if (err) throw new Error("XML 解析失败: " + err.textContent.slice(0, 200));
  return doc;
}

export function serializeXml(doc) {
  return new XMLSerializer().serializeToString(doc);
}

// Get all elements by local name (namespace-safe)
export function getEls(node, localName) {
  return Array.from(node.getElementsByTagNameNS(W_NS, localName));
}

export function getFirstEl(node, localName) {
  return node.getElementsByTagNameNS(W_NS, localName)[0] || null;
}

// Remove all children with given local name
export function removeChildren(parent, localName) {
  for (const el of Array.from(parent.getElementsByTagNameNS(W_NS, localName))) {
    if (el.parentNode === parent) el.remove();
  }
}

// Create a w:xxx element
export function createWEl(doc, localName) {
  return doc.createElementNS(W_NS, "w:" + localName);
}

// Set w:xxx attribute
export function setWAttr(el, localName, value) {
  el.setAttribute("w:" + localName, String(value));
}

// Ensure a direct child element exists, create if missing
export function ensureChild(parent, localName, doc, insertBefore = null) {
  let el = null;
  for (const child of parent.childNodes) {
    if (child.nodeType === 1 && child.localName === localName && child.namespaceURI === W_NS) {
      el = child;
      break;
    }
  }
  if (!el) {
    el = createWEl(doc, localName);
    if (insertBefore) {
      parent.insertBefore(el, insertBefore);
    } else {
      parent.appendChild(el);
    }
  }
  return el;
}

// Replace or create a direct child element
export function replaceChild(parent, localName, newEl) {
  for (const child of Array.from(parent.childNodes)) {
    if (child.nodeType === 1 && child.localName === localName && child.namespaceURI === W_NS) {
      parent.replaceChild(newEl, child);
      return;
    }
  }
  parent.appendChild(newEl);
}

// Remove a direct child element by local name
export function removeDirectChild(parent, localName) {
  for (const child of Array.from(parent.childNodes)) {
    if (child.nodeType === 1 && child.localName === localName && child.namespaceURI === W_NS) {
      child.remove();
      return;
    }
  }
}
