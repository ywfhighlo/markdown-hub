from typing import List, Optional, Dict, Tuple
from collections import Counter
import os
import unicodedata
from .base_converter import BaseConverter
from .dep_check import lib_available, lib_error, command_available, ensure_pymupdf
import re
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime


# ─────────────────────────────────────────
# 懒加载的依赖解析器
# 设计目标：单个依赖缺失不污染其他依赖
# ─────────────────────────────────────────

def _safe_import(name: str):
    """
    动态导入一个第三方库。失败时返回 None。
    使用 importlib 而非顶层 import，避免任何 ImportError 影响整个模块加载。
    """
    import importlib
    try:
        return importlib.import_module(name)
    except Exception:
        return None


# 模块级懒加载占位
fitz = None
pypdf = None
pytesseract = None
convert_from_path = None
docx2txt = None
pd = None
tabulate = None
Presentation = None
html2text = None
psutil_mod = None

# 各功能是否可用的缓存标志（按需探测，缺失不影响其他功能）
_flags_resolved = False


def _resolve_dependencies():
    """按需解析所有依赖（仅在第一次被调用时执行）"""
    global _flags_resolved
    if _flags_resolved:
        return
    _flags_resolved = True

    global fitz, pypdf, pytesseract, convert_from_path
    global docx2txt, pd, tabulate, Presentation, html2text, psutil_mod

    if lib_available("PyMuPDF"):
        fitz = _safe_import("fitz")
    else:
        # 首次自动下载 PyMuPDF（带 C 扩展，无法内置到 vendor）
        ok, msg = ensure_pymupdf()
        if ok:
            fitz = _safe_import("fitz")
    if lib_available("pypdf"):
        pypdf = _safe_import("pypdf")
    if lib_available("pytesseract"):
        pytesseract = _safe_import("pytesseract")
    if lib_available("pdf2image"):
        convert_from_path = _safe_import("pdf2image").convert_from_path
    if lib_available("docx2txt"):
        docx2txt = _safe_import("docx2txt")
    if lib_available("pandas"):
        pd = _safe_import("pandas")
    if lib_available("tabulate"):
        tabulate = _safe_import("tabulate")
    if lib_available("python-pptx"):
        Presentation = _safe_import("pptx").Presentation
    if lib_available("html2text"):
        html2text = _safe_import("html2text")
    if lib_available("psutil"):
        psutil_mod = _safe_import("psutil")


# 便捷标志（首次使用后才会被解析）
def _fitz_available() -> bool:
    _resolve_dependencies()
    return fitz is not None


def _pypdf_available() -> bool:
    _resolve_dependencies()
    return pypdf is not None and pytesseract is not None and convert_from_path is not None


def _docx_available() -> bool:
    _resolve_dependencies()
    return docx2txt is not None


def _pandas_available() -> bool:
    _resolve_dependencies()
    return pd is not None and tabulate is not None


def _pptx_available() -> bool:
    _resolve_dependencies()
    return Presentation is not None


def _html2text_available() -> bool:
    _resolve_dependencies()
    return html2text is not None


def _pdf_available() -> bool:
    """只要 PyMuPDF 或 pypdf 三件套之一就够"""
    return _fitz_available() or _pypdf_available()


class OfficeToMdConverter(BaseConverter):
    """
    Office 文档到 Markdown 转换器

    支持 PDF/DOCX/XLSX/PPTX/HTML -> Markdown 转换
    PDF 转换使用 PyMuPDF，根据字体大小和加粗样式自动识别标题层级
    """

    def __init__(self, output_dir: str, **kwargs):
        super().__init__(output_dir, **kwargs)
        self.poppler_path = kwargs.get('poppler_path')
        self.tesseract_cmd = kwargs.get('tesseract_cmd')
        self._check_dependencies()

    def _check_dependencies(self):
        """在创建实例时记录依赖情况。任意依赖缺失都不会抛出异常。"""
        # 触发一次懒加载
        _resolve_dependencies()

        if not _fitz_available():
            self.logger.warning("PyMuPDF未安装，PDF智能转换不可用，将尝试 pypdf 兜底")

        if not _pypdf_available():
            self.logger.info("pypdf/pytesseract/pdf2image未安装，扫描版PDF的OCR回退不可用")

        if not _pdf_available():
            self.logger.error("PDF处理库均未安装：请 pip install PyMuPDF")

        if not _docx_available():
            self.logger.warning("docx2txt未安装，Word(.docx)转换将跳过")

        if not _pandas_available():
            self.logger.warning("pandas/tabulate未安装，Excel转换将跳过")

        if not _pptx_available():
            self.logger.warning("python-pptx未安装，PPTX转换将跳过")

        if not _html2text_available():
            self.logger.warning("html2text未安装，HTML转换将跳过")

        if psutil_mod is None:
            self.logger.info("psutil未安装，批量并行处理功能不可用，将使用串行处理")

        try:
            subprocess.run(["tesseract", "--version"],
                          stdout=subprocess.PIPE,
                          stderr=subprocess.PIPE,
                          check=True)
        except (subprocess.SubprocessError, FileNotFoundError):
            self.logger.info("tesseract未安装，扫描版PDF的OCR功能不可用")

    def convert(self, input_path: str) -> List[str]:
        supported_extensions = ['.pdf', '.docx', '.doc', '.xlsx', '.xls', '.xlsm',
                              '.pptx', '.ppt', '.html', '.htm']

        if not self._is_valid_input(input_path, supported_extensions):
            raise ValueError(f"无效的输入文件或目录: {input_path}")

        output_files = []
        skipped_reasons = []

        if os.path.isfile(input_path):
            result = self._convert_single_file(input_path)
            if result is None:
                # 单文件模式：None 说明依赖缺失或处理失败，需要给出原因
                skipped_reasons.append(self._last_skip_reason or "未知错误")
            else:
                output_files.append(result)
        else:
            office_files = self._get_files_by_extension(input_path, supported_extensions)
            if not office_files:
                raise ValueError(f"目录中未找到支持的Office文件: {input_path}")

            for office_file in office_files:
                result = self._convert_single_file(office_file)
                if result:
                    output_files.append(result)
                # 批量模式：跳过缺依赖的文件是正常的，不阻断其他文件

        # 单文件模式下如果有跳过原因且无输出，抛出明确错误
        if not output_files and skipped_reasons:
            raise RuntimeError(f"依赖缺失：{'; '.join(skipped_reasons)}")

        return output_files

    def _convert_single_file(self, file_path: str) -> Optional[str]:
        self._last_skip_reason = None
        file_path_obj = Path(file_path)
        file_type = self._get_file_type(file_path_obj)

        if not file_type:
            self.logger.warning(f"不支持的文件类型: {file_path}")
            return None

        # 早期按文件类型检查依赖：缺则跳过该文件，不影响其他文件
        _resolve_dependencies()
        missing = self._missing_deps_for(file_type)
        if missing:
            reason = f"依赖缺失：\n" + "\n".join(f"  - {m}" for m in missing)
            self.logger.warning(f"跳过 {file_path_obj.name}：{reason}")
            self._last_skip_reason = reason
            return None

        try:
            self.logger.info(f"正在处理{file_type}文件: {file_path_obj.name}...")

            if file_type == 'pdf':
                md_text = self._extract_text_from_pdf(file_path_obj)
            elif file_type == 'word':
                text = self._extract_text_from_word(file_path_obj)
                md_text = self._convert_to_markdown(text) if text else ""
            elif file_type == 'excel':
                md_text = self._extract_text_from_excel(file_path_obj)
            elif file_type == 'powerpoint':
                md_text = self._extract_text_from_powerpoint(file_path_obj)
            elif file_type == 'html':
                md_text = self._extract_text_from_html(file_path_obj)
            else:
                self.logger.warning(f"未知文件类型: {file_type}")
                return None

            if md_text:
                md_text = self._optimize_markdown(md_text)
                output_file = self._save_markdown(md_text, file_path_obj)
                return output_file
            else:
                self.logger.warning(f"无法从 {file_path_obj.name} 提取文本")
                return None

        except Exception as e:
            err_msg = f"{type(e).__name__}: {e}"
            self.logger.error(f"处理文件 {file_path} 失败: {err_msg}")
            self._last_skip_reason = f"处理失败: {err_msg}"
            return None

    def _missing_deps_for(self, file_type: str) -> List[str]:
        """返回处理 file_type 所需但当前缺失的依赖库列表（附带具体错误原因）"""
        type_to_check = {
            # PDF 智能路径是 PyMuPDF，但允许 pypdf 兜底（需 pypdf+pytesseract+pdf2image 三件套）
            'pdf':         ('PyMuPDF', 'pypdf', 'pytesseract', 'pdf2image'),
            'word':        ('docx2txt',),
            'excel':       ('pandas', 'tabulate'),
            'powerpoint':  ('python-pptx',),
            'html':        ('html2text',),
        }
        deps = type_to_check.get(file_type, ())
        missing = [d for d in deps if not lib_available(d)]

        if file_type == 'pdf':
            primary_ok = 'PyMuPDF' not in missing
            fallback_ok = not ('pypdf' in missing or 'pytesseract' in missing or 'pdf2image' in missing)
            if primary_ok or fallback_ok:
                # 功能可用，清空 missing
                return []
            # 两条路径都不可用，构建详细提示
            if not primary_ok and not fallback_ok:
                return [
                    f"PyMuPDF (主路径) 或 pypdf+pytesseract+pdf2image (OCR回退路径) 均不可用",
                    f"  PyMuPDF: {lib_error('PyMuPDF') or '未安装'}",
                    f"  pypdf: {lib_error('pypdf') or '未安装'}",
                    f"  pytesseract: {lib_error('pytesseract') or '未安装'}",
                    f"  pdf2image: {lib_error('pdf2image') or '未安装'}",
                    "推荐: pip install PyMuPDF",
                ]

        # 非 PDF 类型：为每个缺失依赖附带错误原因
        detailed = []
        for dep in missing:
            err = lib_error(dep)
            if err:
                detailed.append(f"{dep} ({err})")
            else:
                detailed.append(dep)
        return detailed

    def _get_file_type(self, file_path: Path) -> Optional[str]:
        suffix = file_path.suffix.lower()

        if suffix in ['.pdf']:
            return 'pdf'
        elif suffix in ['.docx', '.doc']:
            return 'word'
        elif suffix in ['.xlsx', '.xls', '.xlsm']:
            return 'excel'
        elif suffix in ['.pptx', '.ppt']:
            return 'powerpoint'
        elif suffix in ['.html', '.htm']:
            return 'html'
        else:
            return None

    # ─────────────────────────────────────────────
    # PDF 提取 - PyMuPDF 智能转换
    # ─────────────────────────────────────────────

    def _extract_text_from_pdf(self, pdf_path: Path) -> str:
        if not _pdf_available():
            self.logger.error("PDF处理库未安装，无法处理PDF文件")
            return "## PDF内容提取失败\n\n需安装PDF处理库: pip install PyMuPDF"

        if _fitz_available():
            return self._extract_pdf_with_fitz(pdf_path)
        else:
            return self._extract_pdf_with_pypdf(pdf_path)

    def _extract_pdf_with_fitz(self, pdf_path: Path) -> str:
        _resolve_dependencies()
        if fitz is None:
            return ""

        try:
            doc = fitz.open(pdf_path)
        except Exception as e:
            self.logger.error(f"无法打开PDF文件 {pdf_path}: {e}")
            return ""

        try:
            assets_dir_name = f"{pdf_path.stem}_assets"
            assets_dir = Path(self.output_dir) / assets_dir_name
            assets_dir_created = False

            all_page_spans = []
            page_heights = []
            page_count = len(doc)

            for page_idx in range(page_count):
                page = doc[page_idx]
                page_dict = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
                page_spans = []
                page_heights.append(page.rect.height)

                for block in page_dict.get("blocks", []):
                    if block.get("type") != 0:
                        continue
                    block_bbox = block.get("bbox", (0, 0, 0, 0))
                    for line in block.get("lines", []):
                        line_text = ""
                        line_spans = []
                        for span in line.get("spans", []):
                            text = span.get("text", "").strip()
                            if not text:
                                continue
                            line_text += span.get("text", "")
                            line_spans.append({
                                "text": text,
                                "size": round(span.get("size", 12), 1),
                                "font": span.get("font", ""),
                                "flags": span.get("flags", 0),
                                "bbox": span.get("bbox", (0, 0, 0, 0)),
                                "color": span.get("color", 0),
                            })

                        if line_text.strip():
                            line_bbox = line.get("bbox", block_bbox)
                            page_spans.append({
                                "text": line_text.strip(),
                                "spans": line_spans,
                                "bbox": line_bbox,
                                "page_idx": page_idx,
                            })

                all_page_spans.append(page_spans)

            header_footer_texts = self._detect_headers_footers(all_page_spans, page_count, page_heights)

            font_sizes = []
            for page_spans in all_page_spans:
                for span_info in page_spans:
                    for s in span_info["spans"]:
                        font_sizes.append(s["size"])

            heading_thresholds = self._compute_heading_thresholds(font_sizes)

            md_lines = []
            total_text_len = 0

            bookmarks_text = self._extract_pdf_bookmarks(doc)
            if bookmarks_text:
                md_lines.append("\n## 目录\n\n")
                md_lines.append(bookmarks_text)
                md_lines.append("\n")

            metadata = self._extract_pdf_metadata(doc)
            if metadata:
                md_lines.append(self._format_pdf_metadata(metadata))
                md_lines.append("\n")

            for page_idx, page_spans in enumerate(all_page_spans):
                page = doc[page_idx]
                has_figure = False

                for span_info in page_spans:
                    text = span_info["text"]
                    bbox = span_info["bbox"]

                    if self._is_header_footer(text, bbox, header_footer_texts, page_dict_for_size=None):
                        continue

                    text = self._clean_text(text)
                    if not text:
                        continue

                    heading_level = self._detect_heading_level(span_info["spans"], heading_thresholds)

                    if heading_level:
                        prefix = "#" * heading_level
                        md_lines.append(f"\n{prefix} {text}\n")
                    else:
                        md_lines.append(f"{text}\n")

                    total_text_len += len(text)

                    if re.search(r'(Figure\s+\d+[:.])|(?<![如见看])(图\s*\d+)', text):
                        has_figure = True

                extracted_images_md = ""
                try:
                    image_list = page.get_images(full=True)
                    if image_list:
                        if not assets_dir_created:
                            assets_dir.mkdir(parents=True, exist_ok=True)
                            assets_dir_created = True

                        for img_idx, img_info in enumerate(image_list):
                            xref = img_info[0]
                            try:
                                base_image = doc.extract_image(xref)
                                if not base_image:
                                    continue

                                image_data = base_image.get("image", b"")
                                image_ext = base_image.get("ext", "png")
                                image_width = base_image.get("width", 0)
                                image_height = base_image.get("height", 0)
                                image_name = f"page_{page_idx+1}_img_{img_idx+1}.{image_ext}"

                                image_info = self._analyze_pdf_image(image_data, image_width, image_height)
                                if image_info.get("skip_reason"):
                                    self.logger.info(f"跳过PDF图片: {image_name} - {image_info['skip_reason']}")
                                    continue

                                image_path = assets_dir / image_name
                                with open(image_path, "wb") as fp:
                                    fp.write(image_data)

                                relative_path = f"{assets_dir_name}/{image_name}"
                                caption = self._generate_image_caption(image_info, page_idx + 1, img_idx + 1)

                                if image_info.get("is_chart"):
                                    extracted_images_md += f"\n> 📊 {caption}\n![{caption}]({relative_path})\n\n"
                                else:
                                    extracted_images_md += f"\n![{caption}]({relative_path})\n\n"

                                self.logger.info(f"提取图片: {image_name} ({image_info.get('size_str', 'unknown')})")
                            except Exception as img_e:
                                self.logger.warning(f"提取图片 xref={xref} 失败: {img_e}")
                            except Exception as img_e:
                                self.logger.warning(f"提取图片 xref={xref} 失败: {img_e}")
                except Exception as img_e:
                    self.logger.warning(f"提取第 {page_idx+1} 页图片时出错: {img_e}")

                if has_figure:
                    try:
                        self.logger.info(f"第 {page_idx+1} 页包含Figure，尝试渲染页面快照...")
                        if not assets_dir_created:
                            assets_dir.mkdir(parents=True, exist_ok=True)
                            assets_dir_created = True

                        pix = page.get_pixmap(dpi=150)
                        render_filename = f"page_{page_idx+1}_render.png"
                        render_path = assets_dir / render_filename
                        pix.save(str(render_path))

                        relative_path = f"{assets_dir_name}/{render_filename}"
                        extracted_images_md += f"\n> **Page {page_idx+1} Snapshot (Contains Figure)**\n\n![Page {page_idx+1} Render]({relative_path})\n\n"
                    except Exception as render_e:
                        self.logger.warning(f"渲染第 {page_idx+1} 页失败: {render_e}")

                if extracted_images_md:
                    md_lines.append(f"\n## 提取的图片\n\n{extracted_images_md}")

            md_text = "\n".join(md_lines)

            if total_text_len < 100 and _pypdf_available():
                self.logger.info(f"{pdf_path.name} 可能是扫描版PDF，尝试使用OCR...")
                ocr_result = self._ocr_pdf(pdf_path)
                if ocr_result and len(ocr_result.strip()) > total_text_len:
                    md_text = ocr_result

            return md_text

        finally:
            doc.close()

    def _extract_pdf_bookmarks(self, doc) -> str:
        _resolve_dependencies()
        if fitz is None:
            return ""

        try:
            toc = doc.get_toc(simple=False)
            if not toc:
                return ""

            self.logger.info(f"提取到 {len(toc)} 个书签/大纲")

            lines = []
            for item in toc:
                level = item[0]
                title = item[1]
                page_num = item[2] if len(item) > 2 else 0

                indent = "  " * (level - 1)
                page_ref = f" (第 {page_num} 页)" if page_num > 0 else ""
                title_clean = self._clean_markdown_title(title)
                lines.append(f"{indent}- [{title_clean}{page_ref}](#{self._generate_anchor(title_clean)})")

            return "\n".join(lines)

        except Exception as e:
            self.logger.warning(f"提取PDF书签失败: {e}")
            return ""

    def _clean_markdown_title(self, title: str) -> str:
        title = title.strip()
        title = re.sub(r'\s+', ' ', title)
        title = re.sub(r'[\[\]]', '', title)
        return title

    def _generate_anchor(self, title: str) -> str:
        anchor = title.lower()
        anchor = unicodedata.normalize('NFKC', anchor)
        parts = re.split(r'[\s\-_\.,;:!?()（）【】《》\[\]]+', anchor)
        anchor_parts = []
        for part in parts:
            part = part.strip()
            if not part:
                continue
            if re.match(r'^[\w]+$', part):
                anchor_parts.append(part)
            else:
                chinese_part = re.sub(r'[^\u4e00-\u9fff]', '', part)
                if chinese_part:
                    anchor_parts.append(chinese_part)
        anchor = '-'.join(anchor_parts)
        anchor = re.sub(r'-+', '-', anchor)
        anchor = anchor.strip('-')
        return anchor if anchor else title.lower().replace(' ', '-')

    def _extract_pdf_metadata(self, doc) -> Optional[Dict]:
        _resolve_dependencies()
        if fitz is None:
            return None

        try:
            metadata = doc.metadata
            if not metadata:
                return None

            useful_fields = {}
            if metadata.get('title'):
                useful_fields['title'] = metadata['title']
            if metadata.get('author'):
                useful_fields['author'] = metadata['author']
            if metadata.get('subject'):
                useful_fields['subject'] = metadata['subject']
            if metadata.get('creator'):
                useful_fields['creator'] = metadata['creator']
            if metadata.get('producer'):
                useful_fields['producer'] = metadata['producer']
            if metadata.get('creationDate'):
                useful_fields['creation_date'] = metadata['creationDate']
            if metadata.get('modDate'):
                useful_fields['modification_date'] = metadata['modDate']

            return useful_fields if useful_fields else None

        except Exception as e:
            self.logger.warning(f"提取PDF元数据失败: {e}")
            return None

    def _format_pdf_metadata(self, metadata: Dict) -> str:
        lines = ["\n## 文档信息\n\n"]
        lines.append("| 属性 | 值 |\n")
        lines.append("|------|----|\n")

        field_names = {
            'title': '标题',
            'author': '作者',
            'subject': '主题',
            'creator': '创建程序',
            'producer': 'PDF生成器',
            'creation_date': '创建日期',
            'modification_date': '修改日期'
        }

        for key, value in metadata.items():
            display_name = field_names.get(key, key)
            lines.append(f"| {display_name} | {value} |\n")

        return "".join(lines)

    def _extract_pdf_with_pypdf(self, pdf_path: Path) -> str:
        _resolve_dependencies()
        if pypdf is None:
            self.logger.error("pypdf也未安装，无法处理PDF文件")
            return ""

        try:
            reader = pypdf.PdfReader(pdf_path)
            text = ""

            metadata = {}
            if reader.metadata:
                metadata = {
                    k: v for k, v in reader.metadata.items()
                    if v and k in ['/Title', '/Author', '/Subject', '/Creator', '/Producer']
                }

            for page in reader.pages:
                page_text = page.extract_text()
                if page_text and page_text.strip():
                    page_text = self._clean_text(page_text)
                    text += page_text + "\n\n"

            if metadata:
                formatted_meta = self._format_pypdf_metadata(metadata)
                text = formatted_meta + "\n" + text

            if len(text.strip()) < 100:
                self.logger.info(f"{pdf_path.name} 可能是扫描版PDF，尝试使用OCR...")
                ocr_result = self._ocr_pdf(pdf_path)
                if ocr_result and len(ocr_result.strip()) > len(text.strip()):
                    text = ocr_result

            return self._convert_to_markdown(text) if text else ""

        except Exception as e:
            self.logger.error(f"处理 {pdf_path} 时出错: {str(e)}")
            return ""

    def _format_pypdf_metadata(self, metadata: Dict) -> str:
        lines = ["\n## 文档信息\n\n"]
        lines.append("| 属性 | 值 |\n")
        lines.append("|------|----|\n")

        field_names = {
            '/Title': '标题',
            '/Author': '作者',
            '/Subject': '主题',
            '/Creator': '创建程序',
            '/Producer': 'PDF生成器'
        }

        for key, value in metadata.items():
            display_name = field_names.get(key, key)
            lines.append(f"| {display_name} | {value} |\n")

        return "".join(lines)

    def _compute_heading_thresholds(self, font_sizes: List[float]) -> Dict[int, float]:
        if not font_sizes:
            return {}

        size_counts = Counter(font_sizes)
        body_size = size_counts.most_common(1)[0][0]

        unique_sizes = sorted(set(font_sizes), reverse=True)

        thresholds = {}
        heading_sizes = [s for s in unique_sizes if s > body_size]

        for i, size in enumerate(heading_sizes):
            level = i + 1
            if level <= 5:
                thresholds[level] = size

        return thresholds

    def _detect_heading_level(self, spans: List[Dict], heading_thresholds: Dict[int, float]) -> Optional[int]:
        if not spans or not heading_thresholds:
            return None

        max_size = max(s["size"] for s in spans)
        any_bold = any(self._is_bold_span(s) for s in spans)

        text = "".join(s["text"] for s in spans).strip()
        if len(text) > 200:
            return None

        if len(text.split()) < 2 and not any_bold:
            return None

        for level in sorted(heading_thresholds.keys()):
            if max_size >= heading_thresholds[level]:
                if level == 1 and not any_bold and max_size < heading_thresholds.get(1, 999) * 1.1:
                    pass
                return level

        if any_bold and max_size >= list(heading_thresholds.values())[-1] if heading_thresholds else False:
            return max(heading_thresholds.keys())

        return None

    def _is_bold_span(self, span: Dict) -> bool:
        font_name = span.get("font", "").lower()
        flags = span.get("flags", 0)
        if flags & 2 ** 4:
            return True
        bold_keywords = ["bold", "black", "heavy", "demi", "粗"]
        return any(kw in font_name for kw in bold_keywords)

    def _detect_headers_footers(self, all_page_spans: List[List[Dict]], page_count: int, page_heights: List[float] = None) -> Dict[str, int]:
        if page_count < 3:
            return {}

        text_position_map = Counter()

        for page_idx, page_spans in enumerate(all_page_spans):
            for span_info in page_spans:
                text = span_info["text"].strip()
                if len(text) > 100:
                    continue
                if re.match(r'^\d+$', text):
                    continue

                bbox = span_info["bbox"]
                y_pos = round(bbox[1], 0)
                key = f"{text}||{y_pos}"
                text_position_map[key] += 1

        header_footer_texts = {}
        threshold = max(3, page_count * 0.4)

        for key, count in text_position_map.items():
            if count >= threshold:
                text = key.split("||")[0]
                header_footer_texts[text] = count

        if page_heights:
            header_zone = set()
            footer_zone = set()
            for page_idx, page_spans in enumerate(all_page_spans):
                if page_idx >= len(page_heights):
                    break
                page_height = page_heights[page_idx]
                top_threshold = page_height * 0.1
                bottom_threshold = page_height * 0.9
                for span_info in page_spans:
                    bbox = span_info["bbox"]
                    y_top = bbox[1]
                    y_bottom = bbox[3]
                    text = span_info["text"].strip()
                    if not text or len(text) > 100:
                        continue
                    if y_top < top_threshold or y_bottom > bottom_threshold:
                        if text not in header_footer_texts:
                            header_footer_texts[text] = 1
                        else:
                            header_footer_texts[text] += 1

        return header_footer_texts

    def _is_header_footer(self, text: str, bbox: Tuple, header_footer_texts: Dict[str, int], page_dict_for_size=None) -> bool:
        if not header_footer_texts:
            return False

        stripped = text.strip()
        if stripped in header_footer_texts:
            return True

        for hf_text in header_footer_texts:
            if stripped == hf_text:
                return True

        return False

    def _analyze_pdf_image(self, image_data: bytes, width: int, height: int) -> Dict:
        info = {
            "width": width,
            "height": height,
            "size": len(image_data),
            "size_str": "",
            "is_chart": False,
            "is_icon": False,
            "skip_reason": None,
        }

        if info["size"] < 5 * 1024:
            info["skip_reason"] = f"图片过小 ({info['size']} bytes)"
            return info

        size_kb = info["size"] / 1024
        size_mb = info["size"] / (1024 * 1024)
        if size_mb >= 1:
            info["size_str"] = f"{size_mb:.1f} MB"
        else:
            info["size_str"] = f"{size_kb:.0f} KB"

        if width > 0 and height > 0:
            if width < 50 or height < 50:
                info["skip_reason"] = f"尺寸过小 ({width}x{height})"
                return info
            if height < 200 and (width / height) > 3.0:
                info["skip_reason"] = f"宽高比过大 ({width}x{height}), 可能是横幅"
                return info
            if height < 120 and (width / height) > 1.5:
                info["skip_reason"] = f"宽高比过大 ({width}x{height}), 可能是分隔线"
                return info

            ratio = width / height if height > 0 else 0
            area = width * height

            if area > 50000 and 0.3 < ratio < 3.0:
                info["is_chart"] = True

            if area < 10000 and (ratio > 3.0 or ratio < 0.3):
                info["is_icon"] = True

        return info

    def _generate_image_caption(self, image_info: Dict, page_num: int, img_num: int) -> str:
        if image_info.get("is_chart"):
            return f"图表 {page_num}-{img_num}"
        if image_info.get("is_icon"):
            return f"图标 {page_num}-{img_num}"
        return f"图片 {page_num}-{img_num}"

    def _clean_text(self, text: str) -> str:
        try:
            text = text.encode('utf-8', errors='ignore').decode('utf-8')
            text = ''.join(char for char in text if ord(char) >= 32 or char in '\n\t\r')
        except Exception:
            text = text.encode('ascii', errors='ignore').decode('ascii')
        return text

    def _detect_language_for_ocr(self, image) -> str:
        try:
            lang_config = pytesseract.get_languages(config='')
            has_chinese = any('chi_sim' in lang or 'chi_tra' in lang for lang in lang_config)

            if has_chinese:
                return 'chi_sim+eng'
            else:
                return 'eng'
        except Exception:
            return 'chi_sim+eng'

    def _ocr_pdf(self, pdf_path: Path) -> str:
        _resolve_dependencies()
        if pypdf is None or pytesseract is None or convert_from_path is None:
            self.logger.error("OCR回退库未安装(pypdf/pytesseract/pdf2image)")
            return ""

        try:
            if self.tesseract_cmd:
                pytesseract.pytesseract.tesseract_cmd = self.tesseract_cmd

            images = convert_from_path(pdf_path, poppler_path=self.poppler_path)
            text_parts = []
            detected_lang = None

            for i, image in enumerate(images):
                self.logger.info(f"正在OCR第 {i+1}/{len(images)} 页...")

                if detected_lang is None:
                    detected_lang = self._detect_language_for_ocr(image)
                    self.logger.info(f"检测到OCR语言: {detected_lang}")

                text = pytesseract.image_to_string(image, lang=detected_lang)
                text_parts.append(text)

            combined_text = "\n\n".join(text_parts)

            combined_text = self._merge_split_table_cells(combined_text)

            return combined_text

        except Exception as e:
            self.logger.error(f"OCR处理失败: {str(e)}")
            return ""

    def _merge_split_table_cells(self, text: str) -> str:
        lines = text.split('\n')
        if not lines:
            return text

        table_regions = self._detect_table_regions(lines)
        if not table_regions:
            return text

        merged_lines = []
        last_end = 0

        for start, end in table_regions:
            merged_lines.extend(lines[last_end:start])
            region_lines = lines[start:end]
            merged_region = self._merge_table_region(region_lines)
            merged_lines.extend(merged_region)
            last_end = end

        merged_lines.extend(lines[last_end:])
        return '\n'.join(merged_lines)

    def _detect_table_regions(self, lines: List[str]) -> List[Tuple[int, int]]:
        if len(lines) < 2:
            return []

        regions = []
        i = 0
        while i < len(lines) - 1:
            line = lines[i].strip()
            next_line = lines[i + 1].strip() if i + 1 < len(lines) else ""

            if self._is_definitive_table_row(line) and self._is_definitive_table_row(next_line):
                pipe_count = line.count('|')
                next_pipe_count = next_line.count('|')
                if abs(pipe_count - next_pipe_count) <= 1:
                    start = i
                    end = i + 2
                    while end < len(lines) and self._is_definitive_table_row(lines[end].strip()):
                        end += 1
                    if end - start >= 2:
                        regions.append((start, end))
                        i = end
                        continue

            i += 1

        return regions

    def _is_definitive_table_row(self, line: str) -> bool:
        line = line.strip()
        if not line:
            return False

        if line.startswith('|') and line.endswith('|'):
            pipe_count = line.count('|')
            if pipe_count < 3:
                return False

            cells = [c.strip() for c in line.split('|')[1:-1]]
            if not cells:
                return False

            cell_lengths = [len(c) for c in cells]
            avg_len = sum(cell_lengths) / len(cell_lengths) if cell_lengths else 0

            if avg_len > 60:
                return False

            non_empty = sum(1 for c in cells if c and not re.match(r'^[\s\-:]+$', c))
            if non_empty == 0:
                return False

            return True

        return False

    def _merge_table_region(self, region_lines: List[str]) -> List[str]:
        merged_lines = []
        i = 0

        while i < len(region_lines):
            line = region_lines[i].strip()

            if self._is_table_separator(line):
                merged_lines.append(region_lines[i])
                i += 1
                continue

            if not self._is_definitive_table_row(line):
                merged_lines.append(region_lines[i])
                i += 1
                continue

            merged_row = line
            j = i + 1

            while j < len(region_lines):
                next_line = region_lines[j].strip()
                if not self._is_definitive_table_row(next_line):
                    break

                if self._should_merge_table_rows(merged_row, next_line):
                    merged_row = self._merge_table_rows(merged_row, next_line)
                    j += 1
                else:
                    break

            merged_lines.append(merged_row)
            i = j

        return merged_lines

    def _is_table_separator(self, line: str) -> bool:
        separator_pattern = r'^\|[\s\-:]+\|[\s\-:]+\|'
        return bool(re.match(separator_pattern, line)) or line.strip() == '---' or line.strip() == '---'

    def _looks_like_table_row(self, line: str) -> bool:
        if not line.startswith('|') and not line.endswith('|'):
            return False
        pipe_count = line.count('|')
        return pipe_count >= 3

    def _should_merge_table_rows(self, row1: str, row2: str) -> bool:
        if not row1.startswith('|') or not row2.startswith('|'):
            return False

        cells1 = [c.strip() for c in row1.split('|')[1:-1]]
        cells2 = [c.strip() for c in row2.split('|')[1:-1]]

        if len(cells1) != len(cells2):
            return False

        non_empty_count1 = sum(1 for c in cells1 if c and not re.match(r'^[\s\-:]+$', c))
        non_empty_count2 = sum(1 for c in cells2 if c and not re.match(r'^[\s\-:]+$', c))

        return non_empty_count2 > non_empty_count1 * 0.5

    def _merge_table_rows(self, row1: str, row2: str) -> str:
        cells1 = [c.strip() for c in row1.split('|')[1:-1]]
        cells2 = [c.strip() for c in row2.split('|')[1:-1]]

        merged_cells = []
        for c1, c2 in zip(cells1, cells2):
            if c2 and not re.match(r'^[\s\-:]+$', c2):
                if c1:
                    merged_cells.append(f"{c1} {c2}")
                else:
                    merged_cells.append(c2)
            else:
                merged_cells.append(c1)

        return '|' + '|'.join(merged_cells) + '|'

    # ─────────────────────────────────────────────
    # 其他文档类型提取
    # ─────────────────────────────────────────────

    def _extract_text_from_word(self, docx_path: Path) -> str:
        _resolve_dependencies()
        if docx2txt is None:
            self.logger.error("docx2txt库未安装，无法处理Word文件")
            return "## Word内容提取失败\n\n需安装docx2txt库: pip install docx2txt"

        try:
            text = docx2txt.process(docx_path)
            return text

        except Exception as e:
            self.logger.error(f"处理Word文档 {docx_path} 时出错: {str(e)}")
            return ""

    def _extract_text_from_excel(self, excel_path: Path) -> str:
        _resolve_dependencies()
        if pd is None or tabulate is None:
            self.logger.error("pandas/tabulate库未安装，无法处理Excel文件")
            return "## Excel内容提取失败\n\n需安装相关库: pip install pandas tabulate openpyxl"

        try:
            xls = pd.ExcelFile(excel_path)
            md_text = ""

            for sheet_name in xls.sheet_names:
                df = pd.read_excel(excel_path, sheet_name=sheet_name)
                md_text += f"## Sheet: {sheet_name}\n\n"
                md_table = df.to_markdown(index=False)
                md_text += md_table + "\n\n"

            return md_text

        except ImportError as e:
            if "openpyxl" in str(e):
                self.logger.error(f"处理Excel文档 {excel_path} 时出错: 缺少openpyxl依赖")
                return "## Excel内容提取失败\n\n需安装openpyxl库: pip install openpyxl\n\n错误信息: " + str(e)
            elif "tabulate" in str(e):
                self.logger.error(f"处理Excel文档 {excel_path} 时出错: 缺少tabulate依赖")
                return "## Excel内容提取失败\n\n需安装tabulate库: pip install tabulate\n\n错误信息: " + str(e)
            else:
                self.logger.error(f"处理Excel文档 {excel_path} 时出错: {str(e)}")
                return ""
        except Exception as e:
            self.logger.error(f"处理Excel文档 {excel_path} 时出错: {str(e)}")
            return ""

    def _extract_text_from_powerpoint(self, pptx_path: Path) -> str:
        _resolve_dependencies()
        if Presentation is None:
            self.logger.error("python-pptx库未安装，无法处理PowerPoint文件")
            return "## PowerPoint内容提取失败\n\n需安装python-pptx库: pip install python-pptx"

        try:
            prs = Presentation(pptx_path)
            md_text = ""

            for i, slide in enumerate(prs.slides):
                md_text += f"## 幻灯片 {i+1}\n\n"

                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text:
                        md_text += f"{shape.text}\n\n"

                md_text += "---\n\n"

            return md_text

        except Exception as e:
            self.logger.error(f"处理PowerPoint文档 {pptx_path} 时出错: {str(e)}")
            return ""

    def _extract_text_from_html(self, html_path: Path) -> str:
        _resolve_dependencies()
        if html2text is None:
            self.logger.error("html2text库未安装，无法处理HTML文件")
            return "## HTML内容提取失败\n\n需安装html2text库: pip install html2text"

        try:
            with open(html_path, "r", encoding="utf-8") as f:
                html_content = f.read()
            h = html2text.HTML2Text()
            h.ignore_links = False
            md_text = h.handle(html_content)
            return md_text
        except Exception as e:
            self.logger.error(f"处理HTML文档 {html_path} 时出错: {str(e)}")
            return ""

    def _convert_to_markdown(self, text: str) -> str:
        md_text = text

        md_text = re.sub(r'^\s*(\d+)\.\s+', r'\1. ', md_text, flags=re.MULTILINE)

        md_text = re.sub(r'\*([^*]+)\*', r'**\1**', md_text)

        return md_text

    def _optimize_markdown(self, md_text: str) -> str:
        md_text = self._preserve_tables(md_text)
        md_text = self._preserve_code_blocks(md_text)
        md_text = self._preserve_inline_code(md_text)
        md_text = self._preserve_links_and_images(md_text)
        md_text = self._preserve_math_formulas(md_text)
        md_text = self._preserve_blockquotes(md_text)
        md_text = self._detect_local_file_links(md_text)

        return md_text

    def _preserve_tables(self, text: str) -> str:
        lines = text.split('\n')
        result_lines = []
        i = 0

        while i < len(lines):
            line = lines[i]
            stripped = line.strip()

            if self._is_markdown_table_row(stripped):
                table_rows = [line]
                i += 1

                while i < len(lines):
                    next_line = lines[i]
                    next_stripped = next_line.strip()

                    if self._is_markdown_table_separator(next_stripped):
                        table_rows.append(next_line)
                        i += 1
                        break
                    elif self._is_markdown_table_row(next_stripped):
                        table_rows.append(next_line)
                        i += 1
                    else:
                        break

                while i < len(lines):
                    next_line = lines[i]
                    next_stripped = next_line.strip()

                    if self._is_markdown_table_row(next_stripped):
                        table_rows.append(next_line)
                        i += 1
                    else:
                        break

                for table_line in table_rows:
                    result_lines.append(table_line)
            else:
                result_lines.append(line)
                i += 1

        return '\n'.join(result_lines)

    def _is_markdown_table_row(self, line: str) -> bool:
        if not line.strip().startswith('|'):
            return False
        if '---' in line or ':--' in line or '--:' in line:
            return False
        return '|' in line[1:]

    def _is_markdown_table_separator(self, line: str) -> bool:
        cleaned = line.strip().strip('|').replace(' ', '')
        if not cleaned:
            return False
        segments = cleaned.split('|')
        return all(re.match(r'^:?-+:?$', seg) for seg in segments)

    def _preserve_code_blocks(self, text: str) -> str:
        lines = text.split('\n')
        result_lines = []
        i = 0
        in_code_block = False

        while i < len(lines):
            line = lines[i]

            if line.strip().startswith('```'):
                result_lines.append(line)
                in_code_block = not in_code_block
                i += 1
                continue

            if not in_code_block:
                if self._is_indented_code_block(line):
                    result_lines.append(line)
                else:
                    result_lines.append(line)
            else:
                result_lines.append(line)

            i += 1

        return '\n'.join(result_lines)

    def _is_indented_code_block(self, line: str) -> bool:
        return line.startswith('    ') or line.startswith('\t')

    def _preserve_inline_code(self, text: str) -> str:
        return text

    def _preserve_links_and_images(self, text: str) -> str:
        return text

    def _preserve_math_formulas(self, text: str) -> str:
        return text

    def _preserve_blockquotes(self, text: str) -> str:
        return text

    def _detect_local_file_links(self, text: str) -> str:
        pattern = r'\[([^\]]+)\]\(([^\)]+\.(?:pdf|docx?|xlsx?|pptx?|html?|txt))\)'
        matches = re.finditer(pattern, text)

        local_links = []
        for match in matches:
            link_text = match.group(1)
            link_path = match.group(2)

            if not link_path.startswith(('http://', 'https://', 'ftp://', 'mailto:')):
                local_links.append((link_text, link_path))

        if local_links:
            self.logger.info(f"检测到 {len(local_links)} 个本地文件链接:")
            for link_text, link_path in local_links:
                self.logger.info(f"  - [{link_text}]({link_path})")

        return text

    def _save_markdown(self, md_text: str, file_path: Path) -> str:
        output_path = Path(self.output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        file_name = file_path.stem
        file_type = file_path.suffix.lower()[1:]

        output_file = output_path / f"{file_name}.md"

        try:
            md_text = md_text.encode('utf-8', errors='ignore').decode('utf-8')
            md_text = ''.join(char for char in md_text if ord(char) >= 32 or char in '\n\t\r')
        except Exception as e:
            self.logger.warning(f"文本清理时出错: {e}，将使用替换策略")
            md_text = md_text.encode('ascii', errors='ignore').decode('ascii')

        header = f"""---
title: {file_name} Document
source_file: {file_path.name}
file_type: {file_type}
converted_date: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
---

"""
        md_text = header + md_text

        try:
            with open(output_file, 'w', encoding='utf-8', errors='replace') as f:
                f.write(md_text)
        except Exception as e:
            self.logger.error(f"保存文件时出错: {e}")
            try:
                with open(output_file, 'w', encoding='ascii', errors='replace') as f:
                    f.write(md_text)
                self.logger.warning(f"使用ASCII编码保存文件: {output_file}")
            except Exception as e2:
                self.logger.error(f"使用ASCII编码保存文件也失败: {e2}")
                raise e2

        self.logger.info(f"已保存到 {output_file}")
        return str(output_file)
