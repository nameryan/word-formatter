import { loadDocx } from "./docx-reader.js";
import { patchStyles, FONT_SIZES } from "./style-patcher.js";
import { patchDocument } from "./doc-patcher.js";
import { downloadFormatted } from "./docx-writer.js";

let state = {
  zip: null,
  files: null,
  fileName: null,
};

const dropZone = document.getElementById("drop-zone");
const fileInput = document.getElementById("file-input");
const fileInfo = document.getElementById("file-info");
const fileName = document.getElementById("file-name");
const changeFileBtn = document.getElementById("change-file");
const actionBtn = document.getElementById("action-btn");
const statusEl = document.getElementById("status");
const resetBtn = document.getElementById("reset-btn");
const marginsToggle = document.getElementById("margins-toggle");
const marginsPanel = document.getElementById("margins-panel");
const lineCustomWrap = document.getElementById("line-custom-wrap");
const indentCustomWrap = document.getElementById("indent-custom-wrap");

function getSelectedLineSpacing() {
  const checked = document.querySelector('input[name="line-spacing"]:checked');
  return checked ? checked.value : "1.5";
}

function getLineValue(value, customPt) {
  if (value === "1.0") return { lineValue: 240, lineRule: "auto" };
  if (value === "1.5") return { lineValue: 360, lineRule: "auto" };
  if (value === "2.0") return { lineValue: 480, lineRule: "auto" };
  if (value === "custom") {
    return {
      lineValue: Math.round((parseFloat(customPt) || 28) * 20),
      lineRule: "exact",
    };
  }
  return { lineValue: 360, lineRule: "auto" };
}

function collectConfig() {
  const bodyFont = document.getElementById("body-font").value;
  const bodySizePt = FONT_SIZES[document.getElementById("body-size").value] || 12;
  const spacing = getLineValue(getSelectedLineSpacing(), document.getElementById("custom-line").value);
  const beforePt = parseFloat(document.getElementById("before-pt").value) || 0;
  const afterPt = parseFloat(document.getElementById("after-pt").value) || 0;

  const indentType = document.getElementById("indent-type").value;
  const indentChars = indentType === "none"
    ? 0
    : indentType === "custom"
      ? (parseFloat(document.getElementById("indent-chars").value) || 2)
      : parseFloat(indentType);
  const indent = indentChars > 0
    ? { twips: Math.round(indentChars * bodySizePt * 20) }
    : null;

  const headings = [1, 2, 3, 4, 5, 6].map(level => ({
    font: document.getElementById(`h${level}-font`).value,
    sizePt: FONT_SIZES[document.getElementById(`h${level}-size`).value] || 14,
    bold: document.getElementById(`h${level}-bold`).checked,
    color: document.getElementById(`h${level}-color`).value,
  }));

  let margins = null;
  if (document.getElementById("apply-margins").checked) {
    margins = {
      top: parseFloat(document.getElementById("margin-top").value) || 2.54,
      bottom: parseFloat(document.getElementById("margin-bottom").value) || 2.54,
      left: parseFloat(document.getElementById("margin-left").value) || 3.18,
      right: parseFloat(document.getElementById("margin-right").value) || 3.18,
    };
  }

  return {
    body: {
      font: bodyFont,
      sizePt: bodySizePt,
      lineValue: spacing.lineValue,
      lineRule: spacing.lineRule,
      beforePt,
      afterPt,
      indent,
    },
    headings,
    margins,
    skipCover: document.getElementById("skip-cover").checked,
    headingNumbering: document.getElementById("heading-numbering").value,
    headingNumberingStart: parseInt(document.getElementById("heading-start").value, 10) || 1,
  };
}

function setStatus(message, type = "info") {
  statusEl.textContent = message;
  statusEl.className = "status " + type;
}

async function handleFile(file) {
  setStatus("正在读取文件...", "info");
  try {
    const result = await loadDocx(file);
    state = {
      zip: result.zip,
      files: result.files,
      fileName: file.name,
    };

    dropZone.style.display = "none";
    fileInfo.style.display = "flex";
    fileName.textContent = file.name;
    actionBtn.disabled = false;
    setStatus("文件已加载，请配置格式后点击「格式化并下载」", "success");
  } catch (error) {
    setStatus("加载失败：" + error.message, "error");
  }
}

function tick() {
  return new Promise(resolve => setTimeout(resolve, 30));
}

async function handleFormat() {
  if (!state.zip) return;

  actionBtn.disabled = true;
  const config = collectConfig();

  try {
    setStatus("正在修改样式定义...", "info");
    await tick();
    const newStyles = patchStyles(state.files.get("word/styles.xml"), config);

    setStatus("正在修改段落格式...", "info");
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
  } catch (error) {
    console.error(error);
    setStatus("处理失败：" + error.message, "error");
  } finally {
    actionBtn.disabled = false;
  }
}

function resetToDefaults() {
  document.getElementById("body-font").value = "宋体";
  document.getElementById("body-size").value = "小四";

  const defaultSpacing = document.querySelector('input[name="line-spacing"][value="1.5"]');
  if (defaultSpacing) defaultSpacing.checked = true;
  document.getElementById("custom-line").value = "28";
  document.getElementById("before-pt").value = "0";
  document.getElementById("after-pt").value = "0";
  document.getElementById("indent-type").value = "2";
  document.getElementById("indent-chars").value = "2";
  indentCustomWrap.style.display = "none";
  document.getElementById("skip-cover").checked = true;
  document.getElementById("heading-numbering").value = "numeric";
  document.getElementById("heading-start").value = "1";
  document.getElementById("margin-top").value = "2.54";
  document.getElementById("margin-bottom").value = "2.54";
  document.getElementById("margin-left").value = "3.18";
  document.getElementById("margin-right").value = "3.18";
  lineCustomWrap.style.display = "none";

  const headingDefaults = [
    { font: "黑体", size: "二号" },
    { font: "黑体", size: "三号" },
    { font: "黑体", size: "小三" },
    { font: "黑体", size: "四号" },
    { font: "黑体", size: "小四" },
    { font: "黑体", size: "小四" },
  ];
  headingDefaults.forEach((heading, index) => {
    const level = index + 1;
    document.getElementById(`h${level}-font`).value = heading.font;
    document.getElementById(`h${level}-size`).value = heading.size;
    document.getElementById(`h${level}-bold`).checked = true;
    document.getElementById(`h${level}-color`).value = "#000000";
  });
}

dropZone.addEventListener("dragover", event => {
  event.preventDefault();
  event.stopPropagation();
  dropZone.classList.add("drag-over");
});

dropZone.addEventListener("dragleave", event => {
  if (!dropZone.contains(event.relatedTarget)) {
    dropZone.classList.remove("drag-over");
  }
});

dropZone.addEventListener("drop", event => {
  event.preventDefault();
  event.stopPropagation();
  dropZone.classList.remove("drag-over");
  const file = event.dataTransfer.files[0];
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

document.getElementById("indent-type").addEventListener("change", function () {
  indentCustomWrap.style.display = this.value === "custom" ? "" : "none";
});

document.querySelectorAll('input[name="line-spacing"]').forEach(radio => {
  radio.addEventListener("change", function () {
    lineCustomWrap.style.display = this.value === "custom" ? "flex" : "none";
  });
});

marginsToggle.addEventListener("click", () => {
  const open = marginsPanel.style.display !== "none";
  marginsPanel.style.display = open ? "none" : "block";
  marginsToggle.textContent = open ? "▶ 页边距设置（可选）" : "▼ 页边距设置（可选）";
});

resetToDefaults();
