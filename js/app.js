import { loadDocx } from "./docx-reader.js";
import { patchStyles, FONT_SIZES } from "./style-patcher.js";
import { patchDocument } from "./doc-patcher.js";
import { downloadFormatted } from "./docx-writer.js";

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
let state = {
  zip: null,
  files: null,
  fileName: null,
};

// ---------------------------------------------------------------------------
// DOM references
// ---------------------------------------------------------------------------
const dropZone       = document.getElementById("drop-zone");
const fileInput      = document.getElementById("file-input");
const fileInfo       = document.getElementById("file-info");
const fileName       = document.getElementById("file-name");
const changeFileBtn  = document.getElementById("change-file");
const actionBtn      = document.getElementById("action-btn");
const statusEl       = document.getElementById("status");
const resetBtn       = document.getElementById("reset-btn");
const marginsToggle  = document.getElementById("margins-toggle");
const marginsPanel   = document.getElementById("margins-panel");
const lineCustomWrap = document.getElementById("line-custom-wrap");

function getSelectedLineSpacing() {
  const checked = document.querySelector('input[name="line-spacing"]:checked');
  return checked ? checked.value : "1.5";
}

// ---------------------------------------------------------------------------
// Defaults
// ---------------------------------------------------------------------------
const DEFAULTS = {
  bodyFont:    "宋体",
  bodySize:    "小四",
  lineSpacing: "1.5",
  customLine:  "28",
  beforePt:    "0",
  afterPt:     "0",
  marginTop:    "2.54",
  marginBottom: "2.54",
  marginLeft:   "3.18",
  marginRight:  "3.18",
  headings: [
    { font: "黑体", size: "二号",   bold: true,  color: "#000000" },
    { font: "黑体", size: "三号",   bold: true,  color: "#000000" },
    { font: "黑体", size: "小三",   bold: true,  color: "#000000" },
    { font: "黑体", size: "四号",   bold: true,  color: "#000000" },
    { font: "黑体", size: "小四",   bold: true,  color: "#000000" },
    { font: "黑体", size: "小四",   bold: true,  color: "#000000" },
  ],
};

// ---------------------------------------------------------------------------
// Line spacing value (twips)
// ---------------------------------------------------------------------------
function getLineValue(spacingVal, customVal) {
  switch (spacingVal) {
    case "1.0": return { lineValue: 240, lineRule: "auto" };
    case "1.5": return { lineValue: 360, lineRule: "auto" };
    case "2.0": return { lineValue: 480, lineRule: "auto" };
    case "custom": {
      const pt = parseFloat(customVal) || 28;
      return { lineValue: Math.round(pt * 20), lineRule: "exact" };
    }
    default: return { lineValue: 360, lineRule: "auto" };
  }
}

// ---------------------------------------------------------------------------
// Collect config from form
// ---------------------------------------------------------------------------
function collectConfig() {
  const bodyFont  = document.getElementById("body-font").value;
  const bodySizeLabel = document.getElementById("body-size").value;
  const bodySizePt = FONT_SIZES[bodySizeLabel] || 12;
  const lineSpacing = getSelectedLineSpacing();
  const customLineVal = document.getElementById("custom-line").value;
  const { lineValue, lineRule } = getLineValue(lineSpacing, customLineVal);
  const beforePt = parseFloat(document.getElementById("before-pt").value) || 0;
  const afterPt  = parseFloat(document.getElementById("after-pt").value)  || 0;

  const headings = [1,2,3,4,5,6].map(i => {
    const font   = document.getElementById(`h${i}-font`).value;
    const sizeLabel = document.getElementById(`h${i}-size`).value;
    const sizePt = FONT_SIZES[sizeLabel] || 14;
    const bold   = document.getElementById(`h${i}-bold`).checked;
    const color  = document.getElementById(`h${i}-color`).value;
    return { font, sizePt, bold, color };
  });

  let margins = null;
  if (document.getElementById("apply-margins").checked) {
    margins = {
      top:    parseFloat(document.getElementById("margin-top").value)    || 2.54,
      bottom: parseFloat(document.getElementById("margin-bottom").value) || 2.54,
      left:   parseFloat(document.getElementById("margin-left").value)   || 3.18,
      right:  parseFloat(document.getElementById("margin-right").value)  || 3.18,
    };
  }

  return {
    body: { font: bodyFont, sizePt: bodySizePt, lineValue, lineRule, beforePt, afterPt },
    headings,
    margins,
  };
}

// ---------------------------------------------------------------------------
// Status helpers
// ---------------------------------------------------------------------------
function setStatus(msg, type = "info") {
  statusEl.textContent = msg;
  statusEl.className = "status " + type;
}

// ---------------------------------------------------------------------------
// File loading
// ---------------------------------------------------------------------------
async function handleFile(file) {
  setStatus("正在读取文件...", "info");
  try {
    const result = await loadDocx(file);
    state.zip = result.zip;
    state.files = result.files;
    state.fileName = file.name;

    dropZone.style.display = "none";
    fileInfo.style.display = "flex";
    fileName.textContent = file.name;
    actionBtn.disabled = false;
    setStatus("文件已加载，请配置格式后点击「格式化并下载」", "success");
  } catch (e) {
    setStatus("加载失败：" + e.message, "error");
  }
}

// ---------------------------------------------------------------------------
// Format + download
// ---------------------------------------------------------------------------
async function handleFormat() {
  if (!state.zip) return;

  actionBtn.disabled = true;
  const config = collectConfig();

  try {
    setStatus("正在修改样式定义...", "info");
    await tick();
    const newStyles = patchStyles(state.files.get("word/styles.xml"), config);

    setStatus("正在修改文档段落格式...", "info");
    await tick();
    const newDoc = patchDocument(state.files.get("word/document.xml"), config);

    const patchedFiles = new Map([
      ["word/styles.xml", newStyles],
      ["word/document.xml", newDoc],
    ]);

    setStatus("正在打包下载...", "info");
    await tick();
    await downloadFormatted(state.zip, patchedFiles, state.fileName);

    setStatus("完成！文件已下载。", "success");
  } catch (e) {
    console.error(e);
    setStatus("处理失败：" + e.message, "error");
  } finally {
    actionBtn.disabled = false;
  }
}

// Yield to browser so status updates render
function tick() {
  return new Promise(r => setTimeout(r, 30));
}

// ---------------------------------------------------------------------------
// Reset to defaults
// ---------------------------------------------------------------------------
function resetToDefaults() {
  document.getElementById("body-font").value = DEFAULTS.bodyFont;
  document.getElementById("body-size").value = DEFAULTS.bodySize;
  const radioToSelect = document.querySelector(`input[name="line-spacing"][value="${DEFAULTS.lineSpacing}"]`);
  if (radioToSelect) radioToSelect.checked = true;
  document.getElementById("custom-line").value = DEFAULTS.customLine;
  document.getElementById("before-pt").value = DEFAULTS.beforePt;
  document.getElementById("after-pt").value = DEFAULTS.afterPt;
  document.getElementById("margin-top").value = DEFAULTS.marginTop;
  document.getElementById("margin-bottom").value = DEFAULTS.marginBottom;
  document.getElementById("margin-left").value = DEFAULTS.marginLeft;
  document.getElementById("margin-right").value = DEFAULTS.marginRight;

  lineCustomWrap.style.display = "none";

  DEFAULTS.headings.forEach((h, i) => {
    const n = i + 1;
    document.getElementById(`h${n}-font`).value  = h.font;
    document.getElementById(`h${n}-size`).value  = h.size;
    document.getElementById(`h${n}-bold`).checked = h.bold;
    document.getElementById(`h${n}-color`).value = h.color;
  });
}

// ---------------------------------------------------------------------------
// Event wiring
// ---------------------------------------------------------------------------

// Drag & drop
dropZone.addEventListener("dragover", e => {
  e.preventDefault();
  dropZone.classList.add("drag-over");
});
dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"));
dropZone.addEventListener("drop", e => {
  e.preventDefault();
  dropZone.classList.remove("drag-over");
  const file = e.dataTransfer.files[0];
  if (file) handleFile(file);
});
dropZone.addEventListener("click", () => fileInput.click());

fileInput.addEventListener("change", () => {
  if (fileInput.files[0]) handleFile(fileInput.files[0]);
});

changeFileBtn.addEventListener("click", () => {
  dropZone.style.display = "";
  fileInfo.style.display = "none";
  actionBtn.disabled = true;
  state = { zip: null, files: null, fileName: null };
  fileInput.value = "";
  setStatus("请上传 .docx 文件", "info");
});

actionBtn.addEventListener("click", handleFormat);
resetBtn.addEventListener("click", resetToDefaults);

// Toggle custom line spacing input
document.querySelectorAll('input[name="line-spacing"]').forEach(radio => {
  radio.addEventListener("change", function () {
    lineCustomWrap.style.display = this.value === "custom" ? "flex" : "none";
  });
});

// Toggle margins panel
marginsToggle.addEventListener("click", () => {
  const open = marginsPanel.style.display !== "none";
  marginsPanel.style.display = open ? "none" : "block";
  marginsToggle.textContent = open ? "▶ 页边距设置（可选）" : "▼ 页边距设置（可选）";
});

// Initialize
resetToDefaults();
