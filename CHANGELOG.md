# Change Log

All notable changes to the "Markdown Hub" extension will be documented in this file.

## [0.3.6] - 2026-11-03

### Added

- Code block syntax highlighting in DOCX/PDF/HTML output. New setting `markdown-hub.codeHighlightTheme` (pygments, tango, zenburn, ... or off).
- Math formula rendering in DOCX/PDF/HTML output ($...$, $$...$$, \[...\], \begin{equation}...): native OMML in DOCX, MathJax-compatible spans in HTML — replaces the previous placeholder text
- HTML → Markdown conversion (right-click any `.html` or `.htm` file)
- DOCX → Markdown now preserves headings, bold/italic, tables, and lists (was plain text only)
- Markdown → EPUB 3 conversion (right-click → "Convert to EPUB"): generates proper EPUB with TOC, navigation, and MathML math rendering for e-readers
- Markdown → PPTX now properly renders markdown syntax: H1/H2/H3 as sized headings, `**bold**` / `*italic*` / `` `code` `` as run-level formatting, GFM tables as real python-pptx tables, code blocks on a full-width black background, bullet/numbered lists, and blockquotes with a left bar — was previously dumping raw markdown text
- Python auto-discovery (`src/pythonLocator.ts`): when Python is installed but not on PATH — the single most common Windows setup failure — the extension now scans common install locations (python.org per-user/all-users, conda, Microsoft Store, `py` launcher), verifies `--version` ≥ 3.8, and uses the first hit. Conversions work out of the box, and a one-time hint shows the detected path with a "Set pythonPath" quick action to pin it.
- MD → PDF zero-install fallback: when neither Word nor LibreOffice is available, the intermediate DOCX (already diagram-preprocessed and border-injected) is rendered to PDF via a headless Chromium-family browser (Edge/Chrome) — no extra installation required. A4 print CSS with CJK fonts, table borders, and code styling is injected automatically.
- "Python not found" error now includes a copyable `markdown-hub.pythonPath` example (e.g. `C:\Users\<you>\AppData\Local\Programs\Python\Python312\python.exe`) plus **Open Settings** / **Download Python** actions.
- The Check Dependencies panel now shows where Python was found (`auto-detected: <path>`).

### Changed

- Dependencies are now lazy-loaded and decoupled: a missing library only disables its own feature
- Per-feature dependency check with a clearer status panel
- All user-facing UI switched to English (Chinese docs kept in `README_zh.md`)
- Improved Marketplace discoverability (keywords, categories, license, badges, comparison table)
- HTML → Markdown: output style normalized to match the rest of the project (`*italic*` instead of `_italic_`, `- item` instead of `* item`, top-level 2-space indent removed)
- VSIX install size: PlantUML jar (21 MB) is no longer bundled — it's downloaded on-demand to `~/.markdown-hub/cache/plantuml/` on first use. Failed downloads fall back to the previous behavior (user-configured / system path). Disable auto-download by setting `MARKDOWN_HUB_NO_AUTO_DOWNLOAD=1`.
- VSIX install size: Batik (5 MB) is no longer bundled — same lazy-download pattern. Poppler (48 MB, Windows-only) is now also lazy-downloaded: on first PDF→PNG / PDF OCR use, the Windows build downloads to `~/.markdown-hub/cache/poppler/` and is reused afterwards; macOS/Linux keep using the system-installed poppler (brew/apt) since `_poppler_url()` only serves Windows. The same `MARKDOWN_HUB_NO_AUTO_DOWNLOAD=1` env var disables all three downloads. VSIX now sits at ~1.8 MB before any cache warm-up, down from 41 MB.
- On-demand downloads (PlantUML/Batik/Poppler) are now resilient: progress is logged every ~10% (or every 5 MB when the server omits Content-Length) so a 48 MB Poppler first-download no longer looks like a hung process; transient network errors trigger up to 3 retries with exponential backoff (partial files cleaned between attempts); short reads are rejected so a truncated download can't produce a corrupt archive.
- README demo section: placeholder images replaced with text descriptions; the placeholder PNGs and the GIF-recording workflow (`media/RECORDING.md`, `scripts/record-gif.sh`) were removed.
- `TASK_SPEC.md` (internal task-tracking artifact) removed from main — it belongs on task branches; the roadmap change-log's stale `TBD A1–A10` row replaced with the actual B2/B4 follow-up entry.
- `.eslintrc.json` added so `npm run lint` actually runs (it previously failed with "ESLint couldn't find a configuration file").

### Fixed

- Preprocess no longer strips heading numbers (`1.2.3`) or `fig:` image caption prefixes (needed for pandoc-crossref / commercial docs)
- Footnotes: stop rewriting pandoc `[^id]` / `[^id]:` into a broken trailing `## 脚注` section — leave native footnotes alone so Word gets real footnotes
- Definition lists: stop smashing pandoc `Term` / `: Def` (and `Term: Def`) into `- **Term**: …` bullets — preserve Definition semantics

### Added

- Optional `pandoc-crossref` / citeproc wiring: CLI `--enable-crossref`, `--enable-citeproc`, `--bibliography`, `--csl` and matching VS Code settings; degrades with a clear log if the filter is missing
- Multi-file book assemble: `--book` or a `*.book.yml` / `.chapters` / `.book.json` chapter list → one DOCX/PDF/HTML/EPUB with `--toc`
- No-template DOCX reference doc now sets CJK `eastAsia` fonts (SimSun / SimHei) alongside Times New Roman

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
- Dependency install hints: when Java, Graphviz, or Tesseract is missing, the PlantUML/Batik/Office→MD converters now log a platform-specific install command (`brew install ...` / `apt install ...` / download URL) alongside the "未安装" message, instead of leaving the user to guess how to obtain the dependency. `dep_check.install_hint_for()` now covers graphviz and LibreOffice (`soffice`) in addition to pandoc/tesseract/java/mmdc/drawio.
- Dependency check panel (frontend): the "Check Dependencies" QuickPick panel now shows three-platform install hints for every tool (`brew` / `apt` / download URL, instead of macOS-only or a bare URL) and adds Graphviz to the checked optional tools (probed via `dot -V`). Previously Graphviz was missing from the panel entirely, and pandoc/tesseract/java only surfaced a macOS hint.
- Poppler install hint precision: on Windows the failure hint now says whether auto-download *failed* (check network/proxy) or was *disabled* (`MARKDOWN_HUB_NO_AUTO_DOWNLOAD=1`), instead of a generic "could not be located". macOS/Linux keep the plain hint since system install is the expected path there.
- MD → PPTX was completely broken: `_create_presentation_from_template` referenced `Presentation` which was never imported — the lazy-loader `_resolve_pptx` resolved `_Inches`/`_Pt`/`_RGBColor`/etc. but missed `Presentation` — so every conversion failed with `NameError: name 'Presentation' is not defined`, swallowed into the generic "PPTX 转换失败，请检查 Markdown 内容是否正确". Fixed by adding `_Presentation` to the lazy-loader. The two PPTX `except` blocks now also set `_last_failure_reason` with the real exception (and log with `exc_info=True`) so future failures surface the actual cause instead of a generic hint.
- DOCX template composition was broken on **every** platform: it was gated on pywin32/Word COM although composition only needs pure-Python libs (docxtpl + docxcompose + python-docx), and the compose function referenced undefined bare names (`DocxTemplate`/`Document`/`Composer`/`WD_SECTION_START` — NameError silently swallowed by `except`, always falling back to plain styles). Now gated on `_template_compose_available()`; works without Word on Windows and on macOS/Linux for the first time. Word COM remains an optional nicety (TOC refresh only, with a transparent log line when skipped).
- `--reproducible` md-to-docx is now truly byte-for-byte reproducible: `docProps/core.xml` `dcterms:created/modified` were left as wall-clock time, so two runs still differed by seconds. Both ZIP entry timestamps and core properties are now normalized via a unified `_finalize_output` pass (idempotent, honors `SOURCE_DATE_EPOCH`); unit test added (`backend/tests/test_reproducible_docx.py`).
- `--reproducible` version manifest: extension version now read from `package.json` (was always `'unknown'` — the backend is not a pip distribution); pandoc probe uses `resolve_command()` so it no longer falsely reports `not found` on machines where pandoc lives off-PATH; help text no longer overclaims PDF byte-reproducibility.
- PPTX → Markdown table cells: empty paragraphs no longer emit stray `<br>`; runs inside one paragraph are no longer split with `<br>` (only paragraphs are); GFM emphasis/code markers wrap only non-whitespace text (leading/trailing spaces moved outside the markers); monospace runs take precedence over bold/italic (code spans cannot carry emphasis); bold+italic now emits `***x***`; adjacent emphasis markers are separated with a zero-width space to avoid GFM parsing ambiguity.
- **PPTX → Markdown crashed for every deck containing a table**: the rich-text extractor read `run.bold` / `run.italic`, which do not exist on python-pptx runs (the API is `run.font.bold` / `run.font.italic`), so table-bearing files failed with an opaque "未知错误". Fixed, and locked down by a new end-to-end sample test (`backend/tests/test_pptx_table_richtext.py`) that was the first test to actually execute this code path.
- `--reproducible` md-to-docx still differed between runs when a DOCX template was used: the template's `document_no` / `date` fields are rendered into the document *body* from the current wall-clock time, so normalizing ZIP entries and core properties could not help. Both now use the same fixed time base (`SOURCE_DATE_EPOCH`, default 1980-01-01) in reproducible mode, and the duplicated timestamp stripping in the two DOCX re-zip hooks was consolidated into the single `_finalize_output` pass.
- Frontend dependency check now maps pip package names to real import names (`python-docx` → `docx`, `PyMuPDF` → `fitz`), in sync with the backend `_INSTALLABLE_HEAVY_LIBS` registry — previously `python-docx` was permanently reported as missing even when installed, blocking MD → DOCX with a misleading "Install Now" loop that could never succeed.
- The dependency checker now respects the `markdown-hub.pythonPath` setting instead of hard-coding bare `python` / `python3`.
- Windows Python detection also tries the `py` launcher (a bare `python3` on Windows is usually just the Microsoft Store alias stub).
- Batch PDF → PNG read the wrong configuration section (`markdownHub` instead of `markdown-hub`), so the `pythonPath` / `popplerPath` settings were silently ignored.
- Dependency install commands now run via `<python> -m pip install …` so they work even when `pip` itself is not on PATH.
- Misleading "Microsoft Word (bundled with Windows)" claim in the PDF-conversion error corrected — Windows does not bundle Word; the message now also mentions the headless-browser option.

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
