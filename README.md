# Word 文档格式化工具

一个纯前端的网页工具，用于批量修复 `.docx` 文件的排版格式。文件在本地处理，**不上传服务器**。

## 功能

- 拖拽或点击上传 `.docx` 文件
- 统一正文字体和字号（支持中文标准字号：小四、四号、三号等）
- 分级设置标题格式（H1–H6）：字体、字号、加粗、颜色
- 统一行间距（单倍 / 1.5 倍 / 双倍 / 固定磅值）
- 设置段前段后间距
- 可选：统一页边距
- 一键下载格式化后的文档

## 使用方法

直接用浏览器打开 `index.html`，无需安装任何工具或服务器。

```
open index.html   # macOS
```

> 注意：需要网络连接以加载 JSZip CDN。如需离线使用，可下载 JSZip 并替换 CDN 链接。

## 技术实现

### 核心原理

`.docx` 文件本质上是一个 ZIP 压缩包，内部包含 XML 文件：

```
word/
├── document.xml   # 文档正文内容（段落、文字、表格）
└── styles.xml     # 样式定义（Normal、Heading1-6 等）
```

格式修复分两层：
1. **styles.xml**：修改 Normal 和 Heading1-6 的样式定义（全局基准）
2. **document.xml**：遍历所有段落，对正文段落直接写入字体/字号覆盖，对标题段落清除直接格式让样式定义生效

### 依赖

- [JSZip 3.10.1](https://stuk.github.io/jszip/)（CDN）：读写 `.docx` ZIP 包
- 无其他依赖，无构建工具

### 项目结构

```
index.html          # 主页面（含全部 JS 逻辑内联）
css/
└── style.css       # 界面样式
js/                 # 模块化源码（供阅读参考，运行时未使用）
├── xml-parser.js   # XML 工具函数
├── style-patcher.js # styles.xml 修改逻辑
├── doc-patcher.js  # document.xml 段落遍历逻辑
├── docx-reader.js  # 文件读取
└── docx-writer.js  # 重打包下载
```

> `js/` 目录下是模块化的参考源码。实际运行的代码已合并内联在 `index.html` 中（解决 `file://` 协议下 ES 模块 CORS 限制）。

### 关键 OOXML 格式说明

| 属性 | XML 写法 | 说明 |
|------|---------|------|
| 字体 | `<w:rFonts w:ascii="宋体" w:eastAsia="宋体"/>` | 中西文字体分别设置 |
| 字号 | `<w:sz w:val="24"/>` | 半磅值，24 = 12pt |
| 行距 | `<w:spacing w:line="360" w:lineRule="auto"/>` | 240=单倍, 360=1.5倍, 480=双倍 |
| 段距 | `<w:spacing w:before="0" w:after="0"/>` | 单位：twips（1pt = 20 twips） |
| 页边距 | `<w:pgMar w:top="1440"/>` | 单位：twips（1cm ≈ 567 twips） |

## 已知限制

- 脚注/尾注内容（`footnotes.xml`）暂不处理
- 密码保护的文档无法处理
- 图片、表格结构不受影响，仅修改文字格式
