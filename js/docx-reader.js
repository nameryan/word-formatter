/**
 * Load a .docx file (File object) and extract relevant XML files.
 * Returns { zip, files: Map<string, string> }
 */
export async function loadDocx(file) {
  if (!file.name.toLowerCase().endsWith(".docx")) {
    throw new Error("请上传 .docx 格式的文件");
  }

  let zip;
  try {
    zip = await JSZip.loadAsync(file);
  } catch (e) {
    if (e.message && e.message.includes("password")) {
      throw new Error("该文档已加密，无法处理");
    }
    throw new Error("文件解析失败，请确认是有效的 .docx 文件");
  }

  const filesToLoad = [
    "word/document.xml",
    "word/styles.xml",
    "word/footnotes.xml",   // optional
    "word/endnotes.xml",    // optional
  ];

  const files = new Map();
  for (const path of filesToLoad) {
    const entry = zip.file(path);
    if (entry) {
      files.set(path, await entry.async("string"));
    }
  }

  if (!files.has("word/document.xml")) {
    throw new Error("无效的 Word 文档（缺少 document.xml）");
  }

  return { zip, files };
}
