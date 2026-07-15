# Markdown Hub

<!-- BADGES -->
![Visual Studio Marketplace Version](https://img.shields.io/badge/version-0.3.6-blue)
![License: MIT](https://img.shields.io/badge/license-MIT-green)
![VS Code Engine](https://img.shields.io/badge/VS%20Code-1.80+-37AAFF)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

**Markdown 文档转换的瑞士军刀。**
支持 Markdown 与 DOCX / PDF / HTML / PPTX 互转，Office → Markdown，图表转 PNG，批量转换，自定义模板等——全部通过 VS Code 右键菜单完成。

[English](./README.md) · [更新日志](./CHANGELOG.md) · [反馈问题](https://github.com/ywfhighlo/markdown-hub/issues)

---

## ✨ 功能特性

### Markdown → Office
- **Markdown → DOCX**（支持自定义模板）
- **Markdown → PDF**（通过 Word / LibreOffice）
- **Markdown → HTML**（GitHub 风格 CSS，自动目录）
- **Markdown → PPTX**（标题页 + 内容页）
- SVG / Mermaid / PlantUML / Draw.io 代码块自动转为 PNG

### Office → Markdown
- **DOCX → Markdown**
- **XLSX → Markdown**（感知多 sheet，自动转义 `|`）
- **PDF → Markdown**（PyMuPDF 文本提取 + 扫描件 OCR 回退）
- **PPTX → Markdown**

### 图表 → PNG
- SVG / Mermaid / PlantUML / Draw.io → 高质量 PNG（Batik + mmdc）

### 批量转换
- 一键将整个目录的 Markdown 文件批量转为 DOCX / PDF / HTML / PPTX
- 一键将目录中的 PDF / DOCX / PPTX / XLSX 文件批量转为 Markdown

---

## 📊 为什么选择 Markdown Hub？

| 功能 | Markdown Hub | vscode-pandoc | MarkItDown | Markdown Preview Enhanced |
|---------|:---:|:---:|:---:|:---:|
| Markdown → DOCX/PDF/HTML/PPTX | ✅ | ✅（部分）| ❌ | ✅（通过 pandoc）|
| Office → Markdown（反向）| ✅ | ❌ | ✅ | ❌ |
| **双向转换** | ✅ | ❌ | ❌（单向）| ❌ |
| **批量转换** | ✅ | ❌ | ❌ | ❌ |
| **图表 → PNG** | ✅ | ❌ | ❌ | 部分 |
| **自定义模板** | ✅ | ❌ | ❌ | ❌ |
| PDF OCR 回退 | ✅ | ❌ | ❌ | ❌ |

**Markdown Hub 是唯一一个同时覆盖双向转换、批量处理、图表渲染和自定义模板的扩展。**

---

## 🚀 快速开始

1. 从 VS Code 应用商店安装本扩展。
2. 在资源管理器中右键任意 `.md` 文件 → **Convert to DOCX**（或 PDF / HTML / PPTX）。
3. 右键任意 `.docx` / `.xlsx` / `.pdf` / `.pptx` → **Convert to Markdown**。
4. 右键任意 `.svg` / `.drawio` / `.puml` → **Convert to PNG**。
5. 右键**文件夹** → 批量转换其中所有匹配的文件。

转换后的文件保存到配置的输出目录（默认：`./converted_markdown_files`）。

---

## 📋 系统要求

Markdown Hub 使用 Python 执行转换逻辑。请安装 Python 3.8+ 以及你需要的依赖（各功能相互独立——只装你用得到的即可）。

### 最小安装（仅 PDF → Markdown）
```bash
pip install PyMuPDF
```

### 按功能安装
| 功能 | 安装命令 |
|---------|---------|
| Word(.docx) → Markdown | `pip install docx2txt` |
| Excel → Markdown | `pip install pandas tabulate openpyxl` |
| PPTX → Markdown | `pip install python-pptx` |
| Markdown → DOCX | `pip install python-docx docxtpl docxcompose docx2txt` |
| Markdown → PDF | `pip install markdown` + [Pandoc](https://pandoc.org/installing.html) + Word/LibreOffice |
| Markdown → HTML | `pip install markdown` |
| Markdown → PPTX | `pip install python-pptx Pillow` |
| 图表 → PNG | `pip install Pillow` |

### 系统工具（可选，按功能需要）
- **[Pandoc](https://pandoc.org/installing.html)** —— Markdown → DOCX / PDF / HTML 必需
- **Microsoft Word**（Windows）或 **LibreOffice**（macOS/Linux）—— 用于 DOCX 模板渲染和 PDF 导出
- **[Tesseract OCR](https://github.com/tesseract-ocr/tesseract)** —— 用于扫描版 PDF → Markdown
- **[draw.io 桌面版](https://github.com/jgraph/drawio-desktop/releases)** —— 用于 Draw.io 图表转换
- **[Mermaid CLI](https://github.com/mermaid-js/mermaid-cli)** —— `npm install -g @mermaid-js/mermaid-cli`
- **Java** —— 用于 SVG 转换和 PlantUML（已内置 Batik）

> 💡 运行 **`Markdown Hub: Check Dependencies`** 命令，可查看已安装和缺失的依赖。

---

## ⚙️ 配置选项

在 VS Code 设置中搜索 `markdown-hub`。主要选项：

| 配置项 | 说明 | 默认值 |
|---------|-------------|---------|
| `markdown-hub.outputDirectory` | 转换后文件的输出目录 | `./converted_markdown_files` |
| `markdown-hub.pythonPath` | Python 解释器路径（留空则自动检测）| `""` |
| `markdown-hub.useDocxTemplate` | 是否启用 DOCX 模板 | `true` |
| `markdown-hub.docxTemplatePath` | 自定义 `.docx` 模板路径 | `""` |
| `markdown-hub.promoteHeadings` | 标题级别自动提升一级（封面页式写作）| `true` |
| `markdown-hub.codeHighlightTheme` | 代码块高亮主题（pygments/tango/.../off）| `pygments` |
| `markdown-hub.svgDpi` | SVG 转 PNG 的 DPI | `300` |

完整配置：打开 VS Code 设置，搜索 `markdown-hub`。

---

## 🛠️ 从源码构建

```bash
git clone https://github.com/ywfhighlo/markdown-hub.git
cd markdown-hub
npm install
npm run compile
# 打包 .vsix
npx @vscode/vsce package
# 本地安装
code --install-extension markdown-hub-0.3.6.vsix
```

---

## 🤝 参与贡献

欢迎贡献！详见 [CONTRIBUTING.md](./.github/CONTRIBUTING.md)。

- 🐛 [报告 Bug](https://github.com/ywfhighlo/markdown-hub/issues/new?template=bug_report.md)
- 💡 [功能建议](https://github.com/ywfhighlo/markdown-hub/issues/new?template=feature_request.md)
- 🔧 [提交 PR](https://github.com/ywfhighlo/markdown-hub/compare)

---

## 👨‍💻 作者

**余文锋 (Yu Wenfeng)** · 📧 909188787@qq.com

## 📄 许可证

[MIT](./LICENSE)
