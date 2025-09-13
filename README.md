# Markdown Hub - VSCode 扩展

Markdown文档转换的瑞士军刀。

## 👨‍💻 作者信息

**作者**: 余文锋  
**邮箱**: 909188787@qq.com  
**项目地址**：https://github.com/ywfhighlo/markdown-hub 

## 🎯 功能特性

### Markdown 转换
- **Markdown → DOCX**: 将 `.md` 文件转换为带有自定义模板的 Word 文档。
- **Markdown → DOCX (SVG支持)**: 将包含SVG代码块的 `.md` 文件转换为 Word 文档，SVG自动转换为PNG图片。
- **Markdown → PDF**: 将 `.md` 文件转换为 PDF 文档。
- **Markdown → HTML**: 将 `.md` 文件转换为带样式的 HTML 网页。
- **Markdown → PPTX**: 将 `.md` 文件转换为 PPTX 演示文稿。

### Office 与其他格式转换
- **DOCX → Markdown**: 将 Word 文档转换为 `.md` 文件。
- **XLSX → Markdown**: 将 Excel 表格转换为 `.md` 文件。
- **PDF → Markdown**: 将 PDF 文档转换为 `.md` 文件。
- **图表 → PNG**: 将SVG、Mermaid、Draw.io、PlantUML等图表文件转换为高质量PNG图片。

### 批量转换
- **Markdown批量转DOCX/PDF/HTML/PPTX**: 批量转换目录中所有 `.md` 文件。
- **PDF/DOCX/PPTX/Excel/All Files批量转Markdown**: 批量转换目录中指定类型或所有支持的文件为 `.md` 文件。

## 📋 系统要求

在使用本扩展前，请确保您的系统已安装以下依赖：

### Windows
- Python 3.8 或更高版本
- Microsoft Word（用于DOCX转换）
- Pandoc（[下载安装包](https://pandoc.org/installing.html)）
- Tesseract OCR（用于PDF文字识别）
- Cairo图形库（用于SVG转换，可选）
- Inkscape（用于SVG转换，可选）

### macOS
```bash
# 使用 Homebrew 安装系统依赖
brew install pandoc
brew install tesseract
brew install cairo
brew install inkscape
```
- Python 3.8 或更高版本
- LibreOffice（用于DOCX转换）

### Linux
- Python 3.8 或更高版本
- LibreOffice
- Pandoc
- Tesseract OCR
- Cairo图形库
- Inkscape（用于SVG转换）

## 🛠️ 安装

1. 在 VS Code 中安装本扩展
2. 安装 Python 依赖：
```bash
cd backend
pip install -r requirements.txt
```

## 🚀 使用方法

1. 在 VS Code 的资源管理器中，右键点击任何支持的文件或包含这些文件的文件夹。
2. 在弹出的上下文菜单中，选择您需要的转换命令 (例如 "Convert to DOCX")。
3. 对于批量转换：
   - 右键点击文件夹。
   - 选择如 "Markdown批量转PDF" 或 "PDF批量转Markdown" 等选项。
   - 系统会自动处理目录中的所有匹配文件。
4. 转换后的文件将出现在您配置的输出目录中（默认 `./converted_markdown_files`）。
5. 若要配置模板、作者信息等，请右键选择 **"Template Settings..."**，这会直接带您到 VS Code 的设置页面。

## ⚙️ 配置选项

您可以在 VS Code 的 `设置(Settings)` 中搜索 `markdown-hub` 来找到所有配置项。

- **`markdown-hub.outputDirectory`**: 所有转换后文件的输出目录。
- **`markdown-hub.pythonPath`**: Python 解释器的路径或命令。
- **`markdown-hub.useTemplate`**: 是否为 `Markdown → DOCX` 的转换启用模板功能。
- **`markdown-hub.templatePath`**: 自定义 `.docx` 模板文件的完整路径。
- **`markdown-hub.projectName`**: 模板中使用的项目名称。
- **`markdown-hub.author`**: 模板中使用的作者姓名。
- **`markdown-hub.email`**: 模板中使用的邮箱地址。
- **`markdown-hub.mobilephone`**: 模板中使用的联系电话。
- **`markdown-hub.promoteHeadings`**: 自动提升Markdown文档的标题级别，以适配"封面页"式的写作习惯。

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！

## 📄 许可证

本项目采用 MIT 许可证。详见 [LICENSE](LICENSE) 文件。
