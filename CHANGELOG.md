# Change Log

All notable changes to the "Markdown Hub" extension will be documented in this file.

## [0.3.6] - 2026-07-15

### Added
- Code block syntax highlighting in DOCX/PDF/HTML output. New setting `markdown-hub.codeHighlightTheme` (pygments, tango, zenburn, ... or off).

### Changed
- Dependencies are now lazy-loaded and decoupled: a missing library only disables its own feature
- Per-feature dependency check with a clearer status panel
- All user-facing UI switched to English (Chinese docs kept in `README_zh.md`)
- Improved Marketplace discoverability (keywords, categories, license, badges, comparison table)

### Fixed
- Pipe-table column alignment (`:---:` / `---:` / `:---`) now respected in output
- Escaped pipes (`\|`) inside table cells no longer break column structure, across MD→Office and XLSX/PDF→MD paths
- Progress-bar labels now translate correctly (stage keys aligned between frontend and backend)
- Clearer, more accurate error messages for missing dependencies (Pandoc, Word/LibreOffice, PyMuPDF)

### Added
- Math formula rendering in DOCX/PDF/HTML output ($...$, $$...$$, \[...\], \begin{equation}...): native OMML in DOCX, MathJax-compatible spans in HTML — replaces the previous placeholder text
- HTML → Markdown conversion (right-click any `.html` or `.htm` file)
- DOCX → Markdown now preserves headings, bold/italic, tables, and lists (was plain text only)

## [0.3.5] - 2025-09-30

### Fixed
- Explicitly specify the output filename to prevent PlantUML from using the `@startuml` title as the filename
- Added SVG preprocessing to fix incompatible SVG 2.0 syntax
- When the DOCX template is enabled but no template file is specified, the default template is now used

## [0.3.4] - 2025-09-29

### Added
- Support inserting PlantUML diagram links in Markdown
- Support converting Draw.io diagrams to PNG images

### Fixed
- Fixed DOCX font being set to italic when no template was specified
- Fixed non-standard Markdown list formatting causing all list items to render on a single line

## [0.3.3] - 2025-09-21

### Added
- The "Swiss Army Knife" Markdown conversion feature set
- Markdown → DOCX / PDF / HTML / PPTX conversion
- Office document → Markdown conversion
- Diagram files (SVG, PlantUML, etc.) → PNG conversion
- Batch conversion
- Custom template support
- Cross-platform support (Windows, macOS, Linux)

### Features
- **Markdown conversion**
  - Markdown → DOCX (with custom template support)
  - Markdown → PDF
  - Markdown → HTML
  - Markdown → PPTX
  - SVG code blocks auto-converted to PNG images

- **Office conversion**
  - DOCX → Markdown
  - XLSX → Markdown
  - PDF → Markdown (with OCR support)
  - PPTX → Markdown

- **Diagram conversion**
  - SVG → PNG
  - Mermaid → PNG
  - PlantUML → PNG

- **Batch processing**
  - Batch Markdown conversion
  - Batch Office document conversion
  - Folder-level batch operations

### Configuration
- Configurable output directory
- Configurable Python path
- Configurable template file path
- Configurable author info
- Configurable conversion parameters

### Dependencies
- Python 3.8+
- Pandoc
- Tesseract OCR (for PDF OCR)
- LibreOffice / Microsoft Word (for Office conversion)

---

## Versioning

This extension follows [Semantic Versioning](https://semver.org/).

### Version format
- **Major**: incompatible API changes
- **Minor**: backward-compatible feature additions
- **Patch**: backward-compatible bug fixes

### Change types
- `Added` - new features
- `Changed` - changes in existing functionality
- `Deprecated` - soon-to-be removed features
- `Removed` - removed features
- `Fixed` - bug fixes
- `Security` - security-related fixes
