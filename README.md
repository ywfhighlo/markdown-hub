# Markdown Hub

<!-- BADGES -->
![Visual Studio Marketplace Version](https://img.shields.io/badge/version-0.3.6-blue)
![License: MIT](https://img.shields.io/badge/license-MIT-green)
![VS Code Engine](https://img.shields.io/badge/VS%20Code-1.80+-37AAFF)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey)

**The Swiss Army Knife for Markdown conversion.**
Convert Markdown ↔ DOCX / PDF / HTML / PPTX, Office → Markdown, diagrams → PNG, with batch conversion, custom templates, and more — all from the VS Code right-click menu.

[中文文档](./README_zh.md) · [Changelog](./CHANGELOG.md) · [Report Issue](https://github.com/ywfhighlo/markdown-hub/issues)

---

## ✨ Features

### Markdown → Office
- **Markdown → DOCX** (with custom template support)
- **Markdown → PDF** (via Word / LibreOffice)
- **Markdown → HTML** (GitHub-style CSS, auto TOC)
- **Markdown → PPTX** (title + content slides)
- SVG / Mermaid / PlantUML / Draw.io code blocks auto-converted to PNG

### Office → Markdown
- **DOCX → Markdown**
- **XLSX → Markdown** (multi-sheet aware, pipe-escaped)
- **PDF → Markdown** (PyMuPDF text extraction + OCR fallback for scanned PDFs)
- **PPTX → Markdown**

### Diagrams → PNG
- SVG / Mermaid / PlantUML / Draw.io → high-quality PNG (Batik + mmdc)

### Batch Conversion
- Convert an entire folder of Markdown files to DOCX / PDF / HTML / PPTX in one click
- Convert a folder of PDF / DOCX / PPTX / XLSX files to Markdown in one click

---

## 📊 Why Markdown Hub?

| Feature | Markdown Hub | vscode-pandoc | MarkItDown | Markdown Preview Enhanced |
|---------|:---:|:---:|:---:|:---:|
| Markdown → DOCX/PDF/HTML/PPTX | ✅ | ✅ (partial) | ❌ | ✅ (via pandoc) |
| Office → Markdown (reverse) | ✅ | ❌ | ✅ | ❌ |
| **Bidirectional** | ✅ | ❌ | ❌ (one-way) | ❌ |
| **Batch conversion** | ✅ | ❌ | ❌ | ❌ |
| **Diagram → PNG** | ✅ | ❌ | ❌ | partial |
| **Custom templates** | ✅ | ❌ | ❌ | ❌ |
| PDF OCR fallback | ✅ | ❌ | ❌ | ❌ |

**Markdown Hub is the only extension that combines bidirectional conversion, batch processing, diagram rendering, and custom templates.**

---

## 🚀 Quick Start

1. Install the extension from the VS Code Marketplace.
2. Right-click any `.md` file in the Explorer → **Convert to DOCX** (or PDF / HTML / PPTX).
3. Right-click any `.docx` / `.xlsx` / `.pdf` / `.pptx` → **Convert to Markdown**.
4. Right-click any `.svg` / `.drawio` / `.puml` → **Convert to PNG**.
5. Right-click a **folder** → batch-convert all matching files inside.

Converted files are saved to the configured output directory (default: `./converted_markdown_files`).

---

## 📋 Prerequisites

Markdown Hub uses Python for conversion logic. Install Python 3.8+ and the dependencies you need (each feature is independent — install only what you use).

### Minimal install (PDF → Markdown only)
```bash
pip install PyMuPDF
```

### Per-feature install
| Feature | Install |
|---------|---------|
| Word(.docx) → Markdown | `pip install docx2txt` |
| Excel → Markdown | `pip install pandas tabulate openpyxl` |
| PPTX → Markdown | `pip install python-pptx` |
| Markdown → DOCX | `pip install python-docx docxtpl docxcompose docx2txt` |
| Markdown → PDF | `pip install markdown` + [Pandoc](https://pandoc.org/installing.html) + Word/LibreOffice |
| Markdown → HTML | `pip install markdown` |
| Markdown → PPTX | `pip install python-pptx Pillow` |
| Diagram → PNG | `pip install Pillow` |

### System tools (optional, per feature)
- **[Pandoc](https://pandoc.org/installing.html)** — required for Markdown → DOCX / PDF / HTML
- **Microsoft Word** (Windows) or **LibreOffice** (macOS/Linux) — for DOCX template rendering and PDF export
- **[Tesseract OCR](https://github.com/tesseract-ocr/tesseract)** — for scanned PDF → Markdown
- **[draw.io desktop](https://github.com/jgraph/drawio-desktop/releases)** — for Draw.io diagram conversion
- **[Mermaid CLI](https://github.com/mermaid-js/mermaid-cli)** — `npm install -g @mermaid-js/mermaid-cli`
- **Java** — for SVG conversion and PlantUML (Batik is bundled)

> 💡 Run the **`Markdown Hub: Check Dependencies`** command to see exactly what's installed and what's missing.

---

## ⚙️ Configuration

Search `markdown-hub` in VS Code Settings. Key options:

| Setting | Description | Default |
|---------|-------------|---------|
| `markdown-hub.outputDirectory` | Output directory for converted files | `./converted_markdown_files` |
| `markdown-hub.pythonPath` | Python executable path (auto-detected if empty) | `""` |
| `markdown-hub.useDocxTemplate` | Enable DOCX template | `true` |
| `markdown-hub.docxTemplatePath` | Custom `.docx` template path | `""` |
| `markdown-hub.promoteHeadings` | Shift heading levels up by one (cover-page style) | `true` |
| `markdown-hub.codeHighlightTheme` | Code block highlight theme (pygments/tango/.../off) | `pygments` |
| `markdown-hub.svgDpi` | DPI for SVG → PNG | `300` |

Full settings: open VS Code Settings and search `markdown-hub`.

---

## 🛠️ Build from Source

```bash
git clone https://github.com/ywfhighlo/markdown-hub.git
cd markdown-hub
npm install
npm run compile
# Package a .vsix
npx @vscode/vsce package
# Install locally
code --install-extension markdown-hub-0.3.6.vsix
```

---

## 🤝 Contributing

Contributions are welcome! See [CONTRIBUTING.md](./.github/CONTRIBUTING.md) for guidelines.

- 🐛 [Report a bug](https://github.com/ywfhighlo/markdown-hub/issues/new?template=bug_report.md)
- 💡 [Request a feature](https://github.com/ywfhighlo/markdown-hub/issues/new?template=feature_request.md)
- 🔧 [Open a PR](https://github.com/ywfhighlo/markdown-hub/compare)

---

## 👨‍💻 Author

**Yu Wenfeng** · 📧 909188787@qq.com

## 📄 License

[MIT](./LICENSE)
