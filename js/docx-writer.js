/**
 * Update patched XML files in the zip and trigger browser download.
 * @param {JSZip} zip - the original JSZip object
 * @param {Map<string, string>} patchedFiles - map of path → new XML string
 * @param {string} originalName - original filename
 */
export async function downloadFormatted(zip, patchedFiles, originalName) {
  for (const [path, xmlString] of patchedFiles) {
    zip.file(path, xmlString);
  }

  const blob = await zip.generateAsync({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });

  const safeName = originalName.replace(/\.docx$/i, "") + "_formatted.docx";
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = safeName;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}
