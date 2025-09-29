# Change Log

All notable changes to the "Markdown Hub" extension will be documented in this file.

## [0.3.4] - 2025-9-29

### Added
- 支持在markdown中插入 PlantUML 图表链接
- 支持drawio图表转化为PNG图片

### Fixed
- 修复了没有指定模板时，docx文档字体设置为斜体的问题
- 修复了md非标准列表格式导致所有列表内容显示在同一行的问题

## [0.3.3] - 2025-9-21

### Added
- Markdown文档转换的瑞士军刀功能
- 支持 Markdown → DOCX/PDF/HTML/PPTX 转换
- 支持 Office 文档 → Markdown 转换
- 支持图表文件（SVG、PlantUML等）→ PNG 转换
- 批量转换功能
- 自定义模板支持
- 多平台支持（Windows、macOS、Linux）

### Features
- **Markdown 转换**
  - Markdown → DOCX（支持自定义模板）
  - Markdown → PDF
  - Markdown → HTML
  - Markdown → PPTX
  - SVG 代码块自动转换为 PNG 图片

- **Office 转换**
  - DOCX → Markdown
  - XLSX → Markdown  
  - PDF → Markdown（支持 OCR）
  - PPTX → Markdown

- **图表转换**
  - SVG → PNG
  - Mermaid → PNG
  - PlantUML → PNG

- **批量处理**
  - 批量 Markdown 转换
  - 批量 Office 文档转换
  - 文件夹级别的批量操作

### Configuration
- 可配置输出目录
- 可配置 Python 路径
- 可配置模板文件路径
- 可配置作者信息
- 可配置转换参数

### Dependencies
- Python 3.8+
- Pandoc
- Tesseract OCR（用于 PDF OCR）
- LibreOffice/Microsoft Word（用于 Office 转换）

---

## 版本说明

本扩展遵循 [语义化版本](https://semver.org/) 规范。

### 版本格式
- **主版本号**：不兼容的 API 修改
- **次版本号**：向下兼容的功能性新增
- **修订号**：向下兼容的问题修正

### 变更类型
- `Added` - 新增功能
- `Changed` - 功能变更
- `Deprecated` - 即将移除的功能
- `Removed` - 已移除的功能
- `Fixed` - 问题修复
- `Security` - 安全相关修复