# Change Log

All notable changes to the "Markdown Hub" extension will be documented in this file.

## [0.3.6] - 2025-06-26

### Added
- 代码块语法高亮：默认主题 pygments；新增设置 `markdown-hub.codeHighlightTheme`，可选 8 个 pandoc 内置主题（pygments / tango / espresso / zenburn / kate / monochrome / breezedark / haddock）或 `off`。作用于 Markdown → DOCX / PDF / HTML；PPTX 不受影响。

### Changed
- 实现依赖懒加载和解耦，单个依赖缺失不影响其他功能正常使用
- 依赖检查改为按功能维度独立检查，前端依赖面板按功能展示状态
- 在 `_normalize_unordered_lists` 与 `_ensure_list_spacing` 中显式跳过 pipe table 行（防御性早退），降低未来正则调整时的回归风险
- 用户可见文案全面英文化（命令标题、错误提示、进度条、依赖检查面板）；中文文档保留在 `README_zh.md`
- `package.json` 新增 16 个 keywords、补全 categories、声明 `license: MIT`、description 改英文，提升 Marketplace 可发现性
- README 改为英文主文档 + 中文版 `README_zh.md`，新增竞品对比表与徽章
- 移除 requirements.txt 安装流程，改为 README.md 中提供分层安装说明

### Fixed
- pipe table 对齐冒号（`:---:`、`---:`、`:---`）被忽略的回归：分隔行原始冒号不再被剥离，DOCX / PDF / HTML 输出恢复列对齐语义（居中 / 右对齐 / 左对齐）
- 代码块高亮长期失效：删除反向作用的 `_process_code_highlighting`（旧实现将 ` ```python ` 改写为无语言标注并插入 `[python代码块]` 文本），改由 pandoc `--highlight-style` 原生渲染
- 单元格里含合法转义 `\|` 时被误切多列：新增 `_split_table_row` 工具按未转义 `|` 切分（支持奇偶反斜杠语义），`_optimize_table_column_widths` 加入列数严格校验；分隔行含 `\|` 不再触发优化（按 GFM/Pandoc 语义）
- XLSX → Markdown 反向路径补齐 cell 内 `|` 转义（pandas `to_markdown` 不自动转义，原路径会破坏跨列），同时把单元格内换行替换为 `<br>`
- Office → Markdown 后处理路径（PDF 表格合并 `_should_merge_table_rows` / `_merge_table_rows` / `_is_definitive_table_row`）共享同一个 `_split_table_row`，避免表格合并时再次误切 `\|`
- `cli.py` 传给前端的 progress stage 值与 `STAGE_LABELS` key 不匹配（中文 stage 无法翻译），统一为英文 key
- 修复错误报告笼统不具体的问题，提供精确的依赖缺失原因和解决方案
- 修复 PyMuPDF 已安装但因 DLL 加载失败导致的误报"缺少依赖"问题
- 修复 Markdown → PDF 转换失败时错误提示误导用户安装 markdown 库的问题，改为正确提示缺少 Pandoc 或 Word/LibreOffice
- 修复 PDF → MD 转换失败时提示不准确的问题，区分主路径(PyMuPDF)和回退路径(pypdf+OCR)并展示具体加载错误

## [0.3.5] - 2025-9-30

### Fixed
- 明确指定输出文件名，防止PlantUML使用@startuml的标题作为文件名
- 增加svg预处理，修复不兼容的svg2.0语法
- 启用Docx模板，但是用户未指定模板文件时，使用默认模板

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