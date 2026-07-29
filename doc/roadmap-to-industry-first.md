# Roadmap: Becoming "Industry #1" in Markdown ↔ Office Conversion

**Goal**: Make Markdown Hub the best VS Code extension for Markdown ↔ Office conversion across three quality dimensions: **table rendering**, **list display**, and **reproducible output**.

**North Star**: Any user converting a complex academic / business / technical document (with multi-level lists, tables, code, math, images) gets a result that's "presentation-ready" out of the box — not a placeholder they need to fix up.

---

## 1. Current State (Baseline)

Verified by end-to-end testing on 2026-07-29, all 7 conversion paths + batch + lazy download:

### 1.1 What works (✅)

| Capability | Status |
|------------|--------|
| 9 conversion directions (MD ↔ Office, Mermaid, HTML/DOCX, etc.) | ✅ |
| Math: OMML in DOCX / MathJax in HTML / MathML in EPUB | ✅ |
| Code highlight: 8 pandoc themes + `off` | ✅ |
| Table alignment: `:`，`---:`，`:---:` preserved | ✅ |
| DOCX ← MD: bold/italic/code/strike inline runs | ✅ |
| PPTX ← MD: H1/H2/H3 sized, bullets, real tables, code blocks | ✅ |
| DOCX → MD: python-docx structured (preserves headings, lists, tables, images, numbered) | ✅ |
| Batch conversion: per-file summary with success/failure counts and reasons | ✅ |
| Diagrams → PNG: PlantUML / Batik / Poppler lazy download | ✅ |
| VSIX size: 1.8 MB (down from 41 MB, 96% reduction) | ✅ |
| 72 pytest passing, 0 TS errors | ✅ |

### 1.2 Critical gaps (🔴)

| # | Gap | Evidence | Impact |
|---|-----|----------|--------|
| **G1** | **DOCX output has no visible table borders** | `_inject_table_borders` only runs in PDF pipeline (L324); plain DOCX path skips it | High — tables look invisible in Word |
| **G2** | **HTML/EPUB table CSS is 3 lines** | `_get_html_theme_css("github_floating_toc")` only has `border-collapse` + 1px border + gray header | High — looks nothing like GitHub |
| **G3** | **Ordered lists get demoted to bullet** | `_normalize_unordered_lists` (L3342) also rewrites `1.` to `-` | Critical — destroys semantics |
| **G4** | **PPTX tables don't support cell inline formatting** | `_render_table` (L1255) only sets size/bold/color; ignores italic/code/link/strike | High — promises made, not kept |
| **G5** | **PPTX lists: only 0/1 level nesting** | `_process_html_node` (L841) only checks grandparent is `<li>` for level 1 | Medium — list depth flattens |
| **G6** | **PPTX lists: same bullet `•` for all levels** | `_render_list` L1156 hardcodes `char='•'` | Low — visual boredom |
| **G7** | **Office → MD: PPTX tables/lists/checkbox ignored** | `_extract_text_from_powerpoint` only dumps `shape.text` | Medium — feature gap |
| **G8** | **Office → MD: DOCX cell rich runs lost** | `_docx_to_markdown_structured` uses `cell.text` only (no runs) | Medium — **bold in tables** lost |
| **G9** | **Dead code clutters the codebase** | `_post_process_html`, `_get_github_theme_css`, `_process_strikethrough` (replaces `~~text~~` with literal `[删除线: text]`!), `_process_superscript_subscript`, `_process_keyboard_keys`, `_remove_image_captions` (stub) | Low — confusing + one is actively wrong |
| **G10** | **No DOCX/HTML column-width control** | `_optimize_table_column_widths` pads with spaces; doesn't actually set columns | High — promised feature is fake |
| **G11** | **PPTX column widths always equal-split** | `_render_table` calls `add_table` without column widths | Medium — wide column dominates |
| **G12** | **DOCX template path has no integration tests** | All 72 pytest tests are utility functions | Critical — regression risk on every release |

### 1.3 Architectural reality

- **md_to_office.py = 3,756 lines** (God Class carrying 5 output formats + preprocessing + templates + XML injection + HTML themes)
- **office_to_md.py = 1,834 lines** (5 input formats + OCR + table detection)
- **No end-to-end test on any conversion path** — every change is manually verified
- **Tests only cover helpers**: `_split_table_row` (escape), `_high_light_style_args` (8 themes), `_optimize_table_column_widths` (column padding math)

---

## 2. Competitive Landscape

| Dimension | Markdown Hub | vscode-pandoc | Markdown Preview Enhanced | MarkItDown |
|-----------|:---:|:---:|:---:|:---:|
| Product scope | Bidirectional + batch + diagrams | Thin Pandoc wrapper | Preview + Puppeteer export | LLM-focused |
| Reverse conversion | ✅ Full bidirectional | ❌ One-way | ❌ One-way | ✅ (semantic only) |
| Batch + aggregates | ✅ Per-file summary | ❌ | ❌ | ❌ |
| **HTML visual styling** | **🟡 3 lines CSS** | 🔴 (Pandoc default) | **🟢 Full Less/CSS** | ❌ |
| **Table complex (colspan)** | 🔴 Pipe only | 🟢 grid tables | 🟡 extended | 🔴 Flatten |
| **List semantics** | **🔴 1. → -** | 🟢 Pandoc AST | 🟢 Pandoc-based | ❌ |
| **Reproducible output** | 🔴 No commitment | 🟡 Version-dependent | 🟡 Time-dependent | 🟢 Semantic only |
| PPTX output | 🟡 Basic (this is what we're fixing) | ❌ | ✅ via Puppeteer HTML path | ❌ |
| Math | ✅ Native via pandoc | ✅ Pandoc | ✅ KaTeX/MathJax | ❌ |
| **Markdown Hub unique: complete bidirectional + batch + lazy downloads = 1.8 MB** | | | | |

**Conclusion**: Nobody does *all* of bidirectional + batch + diagrams + math in one VS Code extension. Our moat is **integration**. Visual quality is average at best.

---

## 3. Strategy: "Industry #1" Path

### 3.1 What "Industry #1" actually means for us

We will NOT beat MPE on preview CSS (their decade of CSS themes + 41k★ users). We will NOT beat Pandoc Universal on raw format coverage. We will win on:

> **"Right-click → 1 second → presentable document"** — including complex tables, multi-level lists, code, math, images.

This is a narrower claim but **no one else delivers it**.

### 3.2 Phased deliverables

#### Phase A — Polish what we have (2 weeks)
The "polish" phase is about **closing the gap between promise and reality** in the current feature set, before adding any new features.

| ID | Title | Severity | Est. |
|----|-------|----------|------|
| **A1** | Fix ordered-list demotion (`1. 2. 3.` → `-`) | 🔴 Critical | 30 min |
| **A2** | Inject complete GitHub-style table CSS in HTML/EPUB | 🔴 High | 4 hours |
| **A3** | PPTX table cell inline formatting (italic/code/link/strike) | 🔴 High | 2 hours |
| **A4** | PPTX list multi-level bullet symbols (• ◦ ▪ -) | 🟡 Medium | 1 hour |
| **A5** | _inject_table_borders runs in DOCX pipeline too (not just PDF) | 🔴 High | 30 min |
| **A6** | PPTX column widths by content ratio (not equal-split) | 🟡 Medium | 2 hours |
| **A7** | Remove dead/contrary code (`_process_strikethrough` etc.) | 🟡 Medium | 1 hour |
| **A8** | DOCX → MD: extract cell runs (preserve bold/italic in table cells) | 🟡 Medium | 3 hours |
| **A9** | Force `code_highlight_theme` even on empty `<style>` insertion — fix HTML CSS injection | 🟡 Medium | 1 hour |
| **A10** | Office → MD: basic PPTX list/table detection | 🟡 Medium | 4 hours |

**Phase A exit criteria**: All 10 items merged. Manual visual review on 3 sample documents (academic, business, technical) shows no regressions and visible improvements.

#### Phase B — Reproducible output (2 weeks)
The "reproducible" phase makes outputs **byte-stable** for given inputs.

| ID | Title | Est. |
|----|-------|------|
| **B1** | Pin pandoc / LibreOffice versions in output manifest | 1 day |
| **B2** | Strip timestamps from DOCX ZIP entries | 0.5 day |
| **B3** | Embed fonts in DOCX/PPTX (no system fallback) | 1 day |
| **B4** | Lock reference-doc per output format | 0.5 day |
| **B5** | `--reproducible` flag in CLI; surfaces version manifest | 0.5 day |

#### Phase C — Real table richer-than-Pipe (3 weeks, optional)
Only if Phase A+B demonstrate user demand. Otherwise skip — this is the highest-risk / highest-cost section.

| ID | Title | Est. |
|----|-------|------|
| C1 | Switch DOCX tables to pandoc grid_tables via preprocessor | 2 days |
| C2 | Preserve colspan/rowspan in DOCX → MD | 3 days |
| C3 | PPTX native table with rich text + cell merging | 3 days |

---

## 4. Non-Goals (explicit)

To prevent scope creep:

- ❌ **Image OCR**: pdf2image → tesseract is bulky (~50MB), accuracy < 90% on CJK. Skip.
- ❌ **Real Markdown editor features** (live preview, autosave): The user already rejected this.
- ❌ **VS Code sidebar / webview editor**: VS Code has built-in Markdown Preview.
- ❌ **LSP / hover hints / IntelliSense**: This is a converter, not a language server.
- ❌ **Cloud sync / multi-device**: Out of scope for a local tool.
- ❌ **Beating Markdown Preview Enhanced on preview CSS**: Stronger contenders already dominate; we focus on conversion output.

---

## 5. Execution Plan

### Phase A (this sprint, 2 weeks)

| Week | Focus | Deliverable |
|------|-------|-------------|
| Week 1 | A1, A2, A3, A5, A7 | Quick fixes + most visible wins |
| Week 2 | A4, A6, A8, A9, A10 | Remaining polish |

### Phase B (next sprint)

Detailed plan emerges after Phase A feedback.

### Phase C (only if justified)

Gate: ship Phase A, gather usage data, see if pipeline struggles demand richer table support.

---

## 6. Quality Gates

Before each release:

- [ ] All pytest pass (`python -m pytest backend/tests/`)
- [ ] TypeScript compiles (`npm run compile`)
- [ ] VSIX rebuilds (`npm run package`)
- [ ] Multi-layer E2E test (the one in `_dbg/multitest2.py`) — extend as features grow
- [ ] Visual diff against Phase A baseline (manual)

---

## 7. Roles & Decision Rights

- **Author**: keep going on Phase A
- **Reviewer**: user — final say on scope, tradeoffs, deferrals
- **Pilot**: after Phase A, ask 3 users (academic / technical / business) to test for 1 week
- **Decision trigger for Phase B**: pilot feedback + bug count + GH issues over 30 days
- **Decision trigger for Phase C**: explicit demand from >= 5 users

---

## 8. Anti-Goals (Things We Explicitly Reject)

- ❌ Adding a UI dialog for every setting. **Too much surface area = bugs.**
- ❌ "Make it work offline" fallacy. We're a converter; if you need offline, use the CLI.
- ❌ Backward compatibility for `pptx_svg_mode='title_and_svg'` — it's a dead code path.
- ❌ Support for PPTX `<p:grpSp>` group shapes in Office→MD — too much variance, low ROI.
- ❌ Renaming the extension to feel "more VS Code native". Brand recognition has value.

---

## 9. Change Log

| Date | Phase | Status |
|------|-------|--------|
| 2026-07-29 | Phase A planning | Document drafted |
| TBD | A1–A10 | Each item moved to "Done" when merged |
