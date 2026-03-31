# JS 模块说明

这个目录是模块化的参考源码，**不是运行时代码**。

由于 `file://` 协议下 ES 模块会被浏览器 CORS 策略拦截，
实际运行的代码已合并内联在根目录的 `index.html` 中。

如需在本地 HTTP 服务器（如 `python3 -m http.server`）下运行，
可将 `index.html` 底部的内联 `<script>` 替换为：

```html
<script type="module" src="js/app.js"></script>
```

## 模块职责

| 文件 | 职责 |
|------|------|
| `app.js` | 主控制器：事件绑定、流程编排、UI 状态管理 |
| `xml-parser.js` | XML 解析/序列化工具，命名空间安全的辅助函数 |
| `style-patcher.js` | 修改 `word/styles.xml`：Normal + Heading1-6 样式定义 |
| `doc-patcher.js` | 遍历 `word/document.xml` 所有段落，直接覆盖/清除格式 |
| `docx-reader.js` | 用 JSZip 读取 `.docx` 文件，提取 XML 字符串 |
| `docx-writer.js` | 将修改后的 XML 重新打包为 `.docx` 并触发浏览器下载 |

## 数据流

```
File (drag/click)
  └── docx-reader.js  →  { zip, files: Map<path, xmlString> }
                               │
              ┌────────────────┼────────────────┐
              ↓                                 ↓
      style-patcher.js                   doc-patcher.js
      (修改 styles.xml)               (修改 document.xml)
              │                                 │
              └────────────────┬────────────────┘
                               ↓
                       docx-writer.js
                   (重打包 → 触发下载)
```
