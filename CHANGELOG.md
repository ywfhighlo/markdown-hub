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
- DOCX → Markdown: horizontally-merged table cells no longer emit duplicate content (python-docx returns the same cell for each grid column of a merge; we now dedupe by `<w:tc>` element)
- Dependency detection: `python-pptx` and `Pillow` are now correctly recognized when installed via pip (their pip name differs from their import name — `pptx` and `PIL` respectively; previously reported as missing, blocking PPTX and image-related conversions)
- DOCX → Markdown numbered lists now renumber consecutive items (1./2./3. instead of 1./1./1. — python-docx returns every numbered item as "1.")
- DOCX → Markdown nested lists now preserve indentation from Word's level suffix (`List Bullet 2` becomes a 2-space-indented sub-item)
- DOCX → Markdown inline images are now extracted to `<docx>_images/` and referenced with `![](path)` — previously dropped entirely
- Batch conversion: emits a per-file summary to the output channel showing succeeded/failed counts and the reason for each failure — previously failed files were silently skipped
- PDF OCR (`DOCX→MD _ocr_pdf`) and batch PDF→PNG: when no explicit `--poppler-path` / `POPPLER_PATH` is configured, the converters now consult the on-demand cache (auto-downloading Poppler on Windows) before letting pdf2image fall back to PATH. Closes a regression introduced when Poppler was removed from the VSIX — Windows users without a system poppler install had silently lost PDF OCR and PDF→PNG.
- PDF OCR and batch PDF→PNG: when Poppler still can't be located (download failed on Windows, or not installed on macOS/Linux), converters now emit a platform-specific install hint instead of a bare opaque error — `brew install poppler` on macOS, `apt`/`dnf`/`pacman install poppler-utils` on Linux, or the manual-download URL + `--poppler-path` option on Windows. Batch PDF→PNG aborts on the first Poppler-missing failure rather than repeating the same error for every file.

### Changed
- HTML → Markdown: output style normalized to match the rest of the project (`*italic*` instead of `_italic_`, `- item` instead of `* item`, top-level 2-space indent removed)
- VSIX install size: PlantUML jar (21 MB) is no longer bundled — it's downloaded on-demand to `~/.markdown-hub/cache/plantuml/` on first use. Failed downloads fall back to the previous behavior (user-configured / system path). Disable auto-download by setting `MARKDOWN_HUB_NO_AUTO_DOWNLOAD=1`.
- VSIX install size: Batik (5 MB) is no longer bundled — same lazy-download pattern. Poppler (48 MB, Windows-only) is now also lazy-downloaded: on first PDF→PNG / PDF OCR use, the Windows build downloads to `~/.markdown-hub/cache/poppler/` and is reused afterwards; macOS/Linux keep using the system-installed poppler (brew/apt) since `_poppler_url()` only serves Windows. The same `MARKDOWN_HUB_NO_AUTO_DOWNLOAD=1` env var disables all three downloads. VSIX now sits at ~1.8 MB before any cache warm-up, down from 41 MB.
- On-demand downloads (PlantUML/Batik/Poppler) are now resilient: progress is logged every ~10% (or every 5 MB when the server omits Content-Length) so a 48 MB Poppler first-download no longer looks like a hung process; transient network errors trigger up to 3 retries with exponential backoff (partial files cleaned between attempts); short reads are rejected so a truncated download can't produce a corrupt archive.

### Added
- Math formula rendering in DOCX/PDF/HTML output ($...$, $$...$$, \[...\], \begin{equation}...): native OMML in DOCX, MathJax-compatible spans in HTML — replaces the previous placeholder text
- HTML → Markdown conversion (right-click any `.html` or `.htm` file)
- DOCX → Markdown now preserves headings, bold/italic, tables, and lists (was plain text only)
- Markdown → EPUB 3 conversion (right-click → "Convert to EPUB"): generates proper EPUB with TOC, navigation, and MathML math rendering for e-readers
- Markdown → PPTX now properly renders markdown syntax: H1/H2/H3 as sized headings, `**bold**` / `*italic*` / `` `code` `` as run-level formatting, GFM tables as real python-pptx tables, code blocks on a full-width black background, bullet/numbered lists, and blockquotes with a left bar — was previously dumping raw markdown text

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
