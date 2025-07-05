from typing import List, Optional
import os
from .base_converter import BaseConverter
import json
import subprocess
import platform
import re
import shutil
from datetime import datetime
from pathlib import Path
import logging

# 平台检测和条件导入（迁移自tools/md_to_docx.py）
IS_WINDOWS = platform.system() == "Windows"
if IS_WINDOWS:
    try:
        from win32com.client import Dispatch
        from docx.enum.section import WD_SECTION_START
        from docxcompose.composer import Composer
        from docx import Document
        from docxtpl import DocxTemplate
        WIN32COM_AVAILABLE = True
    except ImportError:
        WIN32COM_AVAILABLE = False
else:
    WIN32COM_AVAILABLE = False
    WD_SECTION_START = None

class MdToOfficeConverter(BaseConverter):
    """
    This class encapsulates the logic from the original md_to_docx.py script,
    refactored to be reusable and integrate into the VS Code extension backend.
    """

    def __init__(self, output_dir: str, **kwargs):
        super().__init__(output_dir, **kwargs)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Get config from kwargs, with defaults
        self.output_format = kwargs.get('output_format', 'docx')
        self.docx_template_path = kwargs.get('docx_template_path')
        self.pptx_template_path = kwargs.get('pptx_template_path')
        self.project_name = kwargs.get('project_name', '')
        self.author = kwargs.get('author', '')
        self.mobilephone = kwargs.get('mobilephone', '')
        self.email = kwargs.get('email', '')
        self.promote_headings = kwargs.get('promote_headings', False)
        
        # 模板路径现在由前端直接提供，后端不再进行复杂的查找
        self.template_path = None
        if self.output_format in ['docx', 'pdf']:
            self.template_path = self.docx_template_path
        elif self.output_format == 'pptx':
            self.template_path = self.pptx_template_path

    def convert(self, input_path: str) -> List[str]:
        """
        Main conversion entry point. Handles both single files and directories.
        """
        if not self._is_valid_input(input_path, ['.md']):
            raise ValueError(f"Invalid input file or directory: {input_path}")

        output_files = []
        if os.path.isfile(input_path):
            output_file = self._convert_single_file(input_path)
            if output_file:
                output_files.append(output_file)
        else:
            md_files = self._get_files_by_extension(input_path, ['.md'])
            for md_file in md_files:
                output_file = self._convert_single_file(md_file)
                if output_file:
                    output_files.append(output_file)
        
        return output_files

    def _convert_single_file(self, input_file: str) -> Optional[str]:
        """
        Routes a single file to the correct conversion method based on output format.
        """
        if not Path(input_file).exists():
            self.logger.error(f"Input file not found: {input_file}")
            return None

        if self.output_format == 'docx':
            return self._convert_to_docx(input_file)
        elif self.output_format == 'pdf':
            # 1. 定义最终的PDF输出路径
            final_pdf_path = str(self.output_dir / f"{Path(input_file).stem}.pdf")
            
            # 2. 创建临时的DOCX文件
            docx_path = self._convert_to_docx(input_file, to_pdf=True)
            if not docx_path:
                self.logger.error(f"Failed to create intermediate DOCX for PDF conversion from {input_file}")
                return None
            
            # 3. 将临时DOCX转换为最终的PDF
            pdf_path_result = self._convert_docx_to_pdf(docx_path, final_pdf_path)
            
            # 4. 清理临时的DOCX文件
            if pdf_path_result and os.path.exists(docx_path):
                try:
                    os.remove(docx_path)
                    self.logger.info(f"Removed intermediate file: {docx_path}")
                except OSError as e:
                    self.logger.warning(f"Failed to remove intermediate file {docx_path}: {e}")

            return pdf_path_result
        elif self.output_format == 'html':
            return self._convert_to_html(input_file)
        elif self.output_format == 'pptx':
            return self._convert_to_pptx(input_file)
        else:
            self.logger.error(f"Unsupported output format: {self.output_format}")
            return None

    def _convert_to_pptx(self, input_file: str) -> Optional[str]:
        """Converts a Markdown file to PPTX."""
        input_path = Path(input_file)
        output_file_path = self.output_dir / f"{input_path.stem}.pptx"

        processed_content, temp_images = self._preprocess_markdown(input_file)
        if processed_content is None:
            return None

        processed_md_file = input_path.with_name(f"{input_path.stem}_processed_{os.getpid()}.md")
        processed_md_file.write_text(processed_content, encoding='utf-8')
        
        all_temp_files = temp_images + [str(processed_md_file)]

        try:
            if not self._check_tool_availability("pandoc"):
                self.logger.error("Pandoc not found. Please install pandoc and add it to your PATH.")
                raise FileNotFoundError("Pandoc not found. Please install pandoc and add it to your system's PATH.")

            cmd = [
                'pandoc', str(processed_md_file),
                '-o', str(output_file_path),
                '--resource-path=' + str(input_path.parent),
                '--quiet'
            ]
            
            # 只有在提供了模板路径时才使用 --reference-doc
            if self.template_path and Path(self.template_path).exists():
                self.logger.info(f"使用PPTX模板: {self.template_path}")
                cmd.extend(['--reference-doc', self.template_path])
            else:
                self.logger.info("未提供PPTX模板，使用Pandoc默认样式")

            if self.promote_headings:
                cmd.append('--shift-heading-level-by=-1')
            
            # Pandoc不支持为pptx注入变量，但我们保留这个结构以备将来扩展
            # title = self._get_title_from_md(processed_content, input_path)

            subprocess.run(cmd, check=True, capture_output=True, text=True, encoding='utf-8')
            
            self.logger.info(f"Successfully converted {input_file} to {output_file_path}")
            return str(output_file_path)

        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            error_message = e.stderr if hasattr(e, 'stderr') else str(e)
            self.logger.error(f"Failed during PPTX conversion: {error_message}")
            return None
        finally:
            self._cleanup_temp_files(all_temp_files, str(processed_md_file), input_file)

    def _check_tool_availability(self, tool_name: str) -> bool:
        """Checks if an external tool is available in the system's PATH."""
        return shutil.which(tool_name) is not None

    def _get_title_from_md(self, content: str, fallback_path: Path) -> str:
        """Extracts title from Markdown content."""
        try:
            pandoc_title_match = re.search(r'^---\s*\ntitle:\s*(.+?)\n', content, re.DOTALL)
            if pandoc_title_match:
                return pandoc_title_match.group(1).strip()
            
            first_heading_match = re.search(r'^#\s+(.+)', content, re.MULTILINE)
            if first_heading_match:
                return first_heading_match.group(1).strip()
        except Exception as e:
            self.logger.warning(f"Could not extract title due to error: {e}")
        
        return fallback_path.stem

    def _preprocess_markdown(self, md_file_path: str) -> (Optional[str], List[str]):
        """
        Pre-processes Markdown content for conversion.
        """
        try:
            with open(md_file_path, 'r', encoding='utf-8') as f:
                content = f.read()
        except Exception as e:
            self.logger.error(f"Cannot read Markdown file {md_file_path}: {e}")
            return None, []

        temp_files = []
        md_dir = Path(md_file_path).parent

        content = re.sub(r'^(#+)\s*(\d+(\.\d+)*\s+)', r'\1 ', content, flags=re.MULTILINE)
        content = re.sub(r'^(#+)\s*(\d+(\.\d+)*\.\s+)', r'\1 ', content, flags=re.MULTILINE)
        content = re.sub(r'(!\[)(fig:.*?)(\])', r'\1\3', content)

        if self._check_tool_availability("mmdc"):
            def replace_mermaid(match):
                code = match.group(1)
                img_path = md_dir / f"mermaid-generated-{os.urandom(4).hex()}.png"
                try:
                    subprocess.run(['mmdc', '-i', '-', '-o', str(img_path)], input=code.encode('utf-8'), check=True, capture_output=True)
                    temp_files.append(str(img_path))
                    return f"![Mermaid Diagram]({img_path.name})"
                except (subprocess.CalledProcessError, FileNotFoundError) as e:
                    self.logger.error(f"Mermaid conversion failed: {e.stderr if hasattr(e, 'stderr') else e}")
                    return f"```mermaid\n{code}\n```"
            content = re.sub(r'```mermaid\n(.*?)\n```', replace_mermaid, content, flags=re.DOTALL)

        return content, temp_files
    
    def _cleanup_temp_files(self, temp_files: List[str], processed_file: str = None, original_file: str = None):
        """清理临时文件"""
        for temp_file in temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except Exception as e:
                self.logger.warning(f"无法删除临时文件 {temp_file}: {e}")
        
        # 清理处理过的文件（如果与原文件不同）
        if processed_file and original_file and processed_file != original_file:
            try:
                if os.path.exists(processed_file):
                    os.remove(processed_file)
            except Exception as e:
                self.logger.warning(f"无法删除临时文件 {processed_file}: {e}")

    def _convert_to_docx(self, input_file: str, to_pdf: bool = False) -> Optional[str]:
        """Converts a Markdown file to DOCX using a unified pandoc approach."""
        input_path = Path(input_file)
        
        processed_content, temp_images = self._preprocess_markdown(input_file)
        if processed_content is None:
            return None

        processed_md_file = input_path.with_name(f"{input_path.stem}_processed_{os.getpid()}.md")
        processed_md_file.write_text(processed_content, encoding='utf-8')
        
        all_temp_files = temp_images + [str(processed_md_file)]

        try:
            if not self._check_tool_availability("pandoc"):
                self.logger.error("Pandoc not found. Please install pandoc and add it to your PATH.")
                raise FileNotFoundError("Pandoc not found. Please install pandoc and add it to your system's PATH.")

            # Decide whether to use the advanced template feature
            use_advanced_template = (
                self.template_path and 
                Path(self.template_path).exists() and 
                WIN32COM_AVAILABLE
            )

            if use_advanced_template:
                # --- Advanced Template Path (Windows Only) ---
                self.logger.info(f"使用高级模板功能: {self.template_path}")
                
                # 1. Create a temporary content-only DOCX
                temp_content_docx = self.output_dir / f"{input_path.stem}_content_{os.getpid()}.docx"
                all_temp_files.append(str(temp_content_docx))
                
                cmd = [
                    'pandoc', str(processed_md_file),
                    '-o', str(temp_content_docx),
                    '--resource-path=' + str(input_path.parent),
                    '--quiet'
                ]
                if self.promote_headings:
                    cmd.append('--shift-heading-level-by=-1')
                    
                subprocess.run(cmd, check=True, capture_output=True, text=True, encoding='utf-8')

                # 2. Get title and compose final document
                title = self._get_title_from_md(processed_content, input_path)
                final_output_path = self._copy_template_and_append_content(
                    self.template_path,
                    str(temp_content_docx),
                    title,
                    original_input_file=input_file
                )

                # 3. Update TOC if composition was successful
                if final_output_path and Path(final_output_path).exists() and final_output_path != str(temp_content_docx):
                    self._update_toc(final_output_path)
                    self.logger.info(f"成功转换并应用模板: {input_file} -> {final_output_path}")
                    
                    # If converting to PDF, this is the intermediate file.
                    if to_pdf:
                        return final_output_path
                    
                    # Otherwise, it's the final product, we can clean up the content docx
                    # Note: all_temp_files will be cleaned up in the finally block.
                    return final_output_path
                else:
                    self.logger.warning("模板合成失败，返回无模板的DOCX文件。")
                    # It failed and returned the original content_path. We must rename it to a non-temp name.
                    final_path_on_failure = self.output_dir / f"{input_path.stem}.docx"
                    shutil.move(temp_content_docx, final_path_on_failure)
                    all_temp_files.remove(str(temp_content_docx)) # Don't delete it
                    return str(final_path_on_failure)
            else:
                # --- Simple/Cross-Platform Path ---
                if self.template_path:
                    if not WIN32COM_AVAILABLE:
                        self.logger.warning("检测到非Windows环境，模板功能受限（仅应用样式）。")
                    self.logger.info(f"使用DOCX模板进行样式转换 (reference-doc): {self.template_path}")
                else:
                    self.logger.info("未提供DOCX模板，使用Pandoc默认样式")

                if to_pdf:
                    output_file_path = self.output_dir / f"{input_path.stem}_temp_for_pdf_{os.getpid()}.docx"
                else:
                    output_file_path = self.output_dir / f"{input_path.stem}.docx"
                
                cmd = [
                    'pandoc', str(processed_md_file),
                    '-o', str(output_file_path),
                    '--resource-path=' + str(input_path.parent),
                    '--quiet'
                ]
                # Use --reference-doc for styling, same as PPTX conversion
                if self.template_path and Path(self.template_path).exists():
                    cmd.extend(['--reference-doc', self.template_path])

                if self.promote_headings:
                    cmd.append('--shift-heading-level-by=-1')
                
                subprocess.run(cmd, check=True, capture_output=True, text=True, encoding='utf-8')
                
                self.logger.info(f"成功转换 {input_file} to {output_file_path}")
                return str(output_file_path)

        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            error_message = e.stderr if hasattr(e, 'stderr') else str(e)
            self.logger.error(f"Failed during DOCX conversion: {error_message}")
            return None
        finally:
            self._cleanup_temp_files(all_temp_files)

    def _update_toc(self, docx_path: str):
        """Updates the Table of Contents in a DOCX file using Word COM object."""
        if not WIN32COM_AVAILABLE:
            return
        
        word = None
        try:
            word = Dispatch('Word.Application')
            doc = word.Documents.Open(str(Path(docx_path).resolve()))
            doc.Fields.Update()
            if hasattr(doc, 'TablesOfContents'):
                for toc in doc.TablesOfContents:
                    toc.Update()
            doc.Save()
            self.logger.info(f"Updated TOC for {docx_path}")
        except Exception as e:
            self.logger.error(f"Failed to update TOC for {docx_path}: {e}")
        finally:
            if word:
                try:
                    if 'doc' in locals() and doc:
                        doc.Close(False)
                    word.Quit()
                except:
                    pass
    
    def _convert_docx_to_pdf(self, docx_path: str, pdf_path: str) -> Optional[str]:
        """Converts a DOCX file to PDF."""
        final_pdf_path = Path(pdf_path)

        if WIN32COM_AVAILABLE:
            word = None
            try:
                word = Dispatch('Word.Application')
                doc = word.Documents.Open(str(Path(docx_path).resolve()))
                doc.SaveAs(str(final_pdf_path.resolve()), FileFormat=17)
                self.logger.info(f"Successfully created PDF with Word: {final_pdf_path}")
                return str(final_pdf_path)
            except Exception as e:
                self.logger.error(f"Word PDF conversion failed: {e}")
            finally:
                if word:
                    try:
                        if 'doc' in locals() and doc:
                            doc.Close(False)
                        word.Quit()
                    except:
                        pass

        if self._check_tool_availability("soffice"):
            try:
                # soffice 会自动处理输出文件名，我们只需提供目录
                cmd = ["soffice", "--headless", "--convert-to", "pdf", "--outdir", str(final_pdf_path.parent), docx_path]
                subprocess.run(cmd, check=True, capture_output=True)
                
                # LibreOffice/soffice 会创建与输入文件同名的PDF，但可能与我们期望的命名不同，所以需要重命名
                expected_soffice_output = Path(docx_path).with_suffix('.pdf')
                if expected_soffice_output.exists() and str(expected_soffice_output) != str(final_pdf_path):
                    shutil.move(str(expected_soffice_output), str(final_pdf_path))

                self.logger.info(f"Successfully created PDF with LibreOffice: {final_pdf_path}")
                return str(final_pdf_path)
            except (subprocess.CalledProcessError, FileNotFoundError) as e:
                self.logger.warning(f"LibreOffice conversion failed: {e.stderr if hasattr(e, 'stderr') else e}")

        self.logger.error("No suitable tool (Word/LibreOffice) found for PDF conversion.")
        return None

    def _convert_to_html(self, input_file: str) -> Optional[str]:
        """Converts a Markdown file to a styled HTML file."""
        input_path = Path(input_file)
        output_file_path = self.output_dir / f"{input_path.stem}.html"

        processed_content, temp_images = self._preprocess_markdown(input_file)
        if processed_content is None:
            return None

        processed_md_file = input_path.with_name(f"{input_path.stem}_processed_{os.getpid()}.md")
        processed_md_file.write_text(processed_content, encoding='utf-8')
        
        all_temp_files = temp_images + [str(processed_md_file)]

        try:
            if not self._check_tool_availability("pandoc"):
                self.logger.error("Pandoc not found. Please install it to convert files.")
                return None
            
            resource_path_arg = '--resource-path=' + str(input_path.parent)
            cmd = ['pandoc', str(processed_md_file), '--from', 'markdown+smart', '--to', 'html', resource_path_arg]
            result = subprocess.run(cmd, check=True, capture_output=True, text=True, encoding='utf-8')
            html_body = result.stdout
            
            heading_counts = {}
            def add_anchor_to_heading(match):
                level, title = len(match.group(1)), match.group(2).strip()
                base_id = re.sub(r'[^\w\s-]', '', title).strip().lower()
                base_id = re.sub(r'[\s-]+', '-', base_id)
                count = heading_counts.get(base_id, 0)
                heading_counts[base_id] = count + 1
                anchor_id = f"{base_id}-{count}" if count > 0 else base_id
                return f'<h{level} id="{anchor_id}">{title}</h{level}>'
            html_body = re.sub(r'<h([1-6])>(.*?)</h\1>', add_anchor_to_heading, html_body)

            toc_html = self._generate_html_toc(processed_content)
            title = self._get_title_from_md(processed_content, input_path)
            css = self._get_html_theme_css("github_floating_toc")
            
            final_html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>{css}</style>
</head>
<body>
    <div class="container">
        <div class="toc-container">{toc_html}</div>
        <div class="content-container">{html_body}</div>
    </div>
</body>
</html>"""

            output_file_path.write_text(final_html, encoding='utf-8')
            self.logger.info(f"Successfully created HTML: {output_file_path}")
            return str(output_file_path)

        except (subprocess.CalledProcessError, FileNotFoundError) as e:
            self.logger.error(f"Failed during HTML conversion: {e.stderr if hasattr(e, 'stderr') else e}")
            return None
        finally:
            self._cleanup_temp_files(all_temp_files, str(processed_md_file), input_file)

    def _generate_html_toc(self, content: str) -> str:
        """Generates a nested HTML list for the Table of Contents."""
        toc_lines = ['<nav class="toc"><ul>']
        heading_counts = {}
        for line in content.splitlines():
            match = re.match(r'^(#+)\s+(.*)', line)
            if match:
                level, title = len(match.group(1)), match.group(2).strip()
                base_id = re.sub(r'[^\w\s-]', '', title).strip().lower()
                base_id = re.sub(r'[\s-]+', '-', base_id)
                count = heading_counts.get(base_id, 0)
                heading_counts[base_id] = count + 1
                anchor_id = f"{base_id}-{count}" if count > 0 else base_id
                toc_lines.append(f'<li class="toc-level-{level}"><a href="#{anchor_id}">{title}</a></li>')
        toc_lines.append('</ul></nav>')
        return '\n'.join(toc_lines)

    def _get_html_theme_css(self, theme_name: str) -> str:
        """Returns CSS for the HTML output."""
        github_floating_toc = """
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Helvetica, Arial, sans-serif; line-height: 1.6; color: #333; background-color: #fff; margin: 0; padding: 0; }
        .container { max-width: 1200px; margin: 20px auto; display: flex; flex-direction: row; align-items: flex-start; }
        .toc-container { width: 250px; flex-shrink: 0; position: -webkit-sticky; position: sticky; top: 20px; height: calc(100vh - 40px); overflow-y: auto; padding-right: 20px; border-right: 1px solid #e1e4e8; }
        .content-container { flex-grow: 1; padding-left: 30px; max-width: 800px; }
        .toc ul { list-style: none; padding-left: 0; } .toc li a { color: #0366d6; text-decoration: none; display: block; padding: 4px 0; font-size: 14px; }
        .toc li a:hover { text-decoration: underline; }
        .toc-level-1 { padding-left: 5px; font-weight: 600; } .toc-level-2 { padding-left: 20px; } .toc-level-3 { padding-left: 35px; } .toc-level-4 { padding-left: 50px; }
        h1, h2, h3, h4, h5, h6 { font-weight: 600; line-height: 1.25; margin-top: 24px; margin-bottom: 16px; border-bottom: 1px solid #eaecef; padding-bottom: .3em; }
        h1 { font-size: 2em; } h2 { font-size: 1.5em; } h3 { font-size: 1.25em; }
        p { margin-top: 0; margin-bottom: 16px; } a { color: #0366d6; text-decoration: none; } a:hover { text-decoration: underline; }
        code, pre { font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, Courier, monospace; font-size: 13px; }
        pre { word-wrap: normal; padding: 16px; overflow: auto; line-height: 1.45; background-color: #f6f8fa; border-radius: 3px; }
        code { background-color: rgba(27,31,35,.05); padding: .2em .4em; margin: 0; border-radius: 3px; }
        pre > code { padding: 0; margin: 0; background-color: transparent; border: 0; }
        table { border-collapse: collapse; } th, td { border: 1px solid #ddd; padding: 8px; } th { background-color: #f2f2f2; }
        img { max-width: 100%; } blockquote { color: #6a737d; border-left: .25em solid #dfe2e5; padding: 0 1em; margin-left: 0; }
        """
        return github_floating_toc

    def _remove_title_numbers(self, input_file: str) -> str:
        """
        处理Markdown文件，去掉标题前面的序号（如1.1、2.3.4等格式），
        并删除图片标题，同时确保列表前有空行以便Pandoc正确识别。
        """
        try:
            # 读取原始文件内容
            with open(input_file, 'r', encoding='utf-8') as f:
                original_content = f.read()
            
            # 1. 正则表达式匹配标题前的序号
            pattern_title_numbers = r'^(#+)\s+(\d+(\.\d+)*)\s+(.+)$'
            processed_content = re.sub(pattern_title_numbers, r'\1 \4', original_content, flags=re.MULTILINE)
            
            # 2. 删除图片标题
            processed_content = self._remove_image_captions(processed_content)

            # 3. 确保列表前有空行以便Pandoc正确识别
            lines = processed_content.splitlines()
            new_processed_lines = []
            list_markers = ("- ", "* ", "+ ")
            
            for i, current_line_text in enumerate(lines):
                stripped_line = current_line_text.lstrip()
                is_list_item = any(stripped_line.startswith(marker) for marker in list_markers)

                if is_list_item:
                    if i > 0:
                        previous_line_text = lines[i-1]
                        stripped_previous_line = previous_line_text.lstrip()
                        is_previous_list_item = any(stripped_previous_line.startswith(marker) for marker in list_markers)
                        
                        if previous_line_text.strip() and not is_previous_list_item:
                            new_processed_lines.append("")
                
                new_processed_lines.append(current_line_text)
            
            processed_content = "\n".join(new_processed_lines)
            
            # 确保处理后的内容以换行符结尾
            if original_content.endswith('\n') and not processed_content.endswith('\n'):
                processed_content += '\n'
            elif new_processed_lines and any(new_processed_lines[-1].lstrip().startswith(marker) for marker in list_markers) and not processed_content.endswith('\n'):
                processed_content += '\n'

            # 如果内容没有变化，直接返回原文件路径
            if processed_content == original_content:
                return input_file
            
            # 创建临时文件保存处理后的内容
            temp_file = input_file + '.temp.md'
            with open(temp_file, 'w', encoding='utf-8') as f:
                f.write(processed_content)
            
            return temp_file
        except Exception as e:
            self.logger.error(f"处理Markdown预处理时出错 (file: {input_file}): {str(e)}")
            return input_file
    
    def _remove_image_captions(self, content: str) -> str:
        """删除图片标题（占位符实现）"""
        # 这里可以添加具体的图片标题删除逻辑
        return content
    
    def _copy_template_and_append_content(self, template_path: str, content_path: str, title: str, original_input_file: str) -> str:
        """
        Applies a template by rendering variables and composing it with the content document.
        This uses `DocxTemplate` and `docxcompose`.
        """
        if not WIN32COM_AVAILABLE:
            self.logger.warning("在非Windows系统上无法使用模板功能，将使用简单转换")
            return content_path
            
        # Create a deterministic final output path based on the original input file.
        original_input_path = Path(original_input_file)
        output_path = str(self.output_dir / f"{original_input_path.stem}.docx")
        
        try:
            # 获取模板上下文数据
            context = {
                'project_name': self.project_name or '',
                'title': title,
                'document_no': "P" + datetime.now().strftime("%Y%m%d%H%M%S"),
                'date': datetime.now().strftime("%Y-%m-%d"),
                'author': self.author or '',
                'mobilephone': self.mobilephone or '',
                'email': self.email or ''
            }
            
            self.logger.info(f"使用模板: {template_path}")
            self.logger.info(f"模板上下文: {context}")
            
            # 使用DocxTemplate渲染模板
            doc_tpl = DocxTemplate(template_path)
            doc_tpl.render(context)
            doc_tpl.save(output_path)

            # 加载渲染后的模板文档
            master = Document(output_path)
            
            # 创建composer对象
            composer = Composer(master)
            
            # 加载内容文档
            content_doc = Document(content_path)
            
            # 在模板文档末尾添加连续分节符
            section = master.add_section()
            section.start_type = WD_SECTION_START.CONTINUOUS
            
            # 合并文档，保留样式
            composer.append(content_doc)
            
            # 更新文档属性
            master.core_properties.title = title
            
            # 保存合并后的文档
            composer.save(output_path)
            
            self.logger.info(f"模板处理成功: {output_path}")
            return output_path
            
        except Exception as e:
            self.logger.error(f"模板处理失败: {e}")
            import traceback
            self.logger.error(f"详细错误信息: {traceback.format_exc()}")
            # On failure, return the path to the original content so it can be handled upstream
            return content_path
    
    def _post_process_html(self, html_file: str, processed_md_file: str):
        """后处理HTML文件，添加样式和目录"""
        try:
            # 读取生成的HTML文件
            with open(html_file, 'r', encoding='utf-8') as f:
                html_content = f.read()
            
            # 获取GitHub主题CSS
            theme_css = self._get_github_theme_css()
            
            # 添加锚点ID到标题
            import re
            def add_anchor_to_heading(match):
                level = len(match.group(1))
                title = match.group(2)
                anchor_id = re.sub(r'[^\w\s-]', '', title).strip()
                anchor_id = re.sub(r'[\s_-]+', '-', anchor_id).lower()
                return f'<h{level} id="{anchor_id}">{title}</h{level}>'
            
            html_content = re.sub(r'<h([1-6])>(.*?)</h[1-6]>', add_anchor_to_heading, html_content)
            
            # 插入CSS样式
            css_insert = f'<style>\n{theme_css}\n</style>'
            if '</head>' in html_content:
                html_content = html_content.replace('</head>', f'{css_insert}\n</head>')
            else:
                html_content = f'<head>\n{css_insert}\n</head>\n{html_content}'
            
            # 写回HTML文件
            with open(html_file, 'w', encoding='utf-8') as f:
                f.write(html_content)
                
        except Exception as e:
            self.logger.warning(f"HTML后处理失败: {e}")
    
    def _get_github_theme_css(self) -> str:
        """获取GitHub主题的CSS样式"""
        return """
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            line-height: 1.6;
            color: #24292e;
            max-width: 980px;
            margin: 0 auto;
            padding: 45px;
            background-color: #ffffff;
        }
        h1, h2, h3, h4, h5, h6 {
            margin-top: 24px;
            margin-bottom: 16px;
            font-weight: 600;
            line-height: 1.25;
        }
        h1 { font-size: 2em; border-bottom: 1px solid #eaecef; padding-bottom: 10px; }
        h2 { font-size: 1.5em; border-bottom: 1px solid #eaecef; padding-bottom: 8px; }
        h3 { font-size: 1.25em; }
        h4 { font-size: 1em; }
        h5 { font-size: 0.875em; }
        h6 { font-size: 0.85em; color: #6a737d; }
        p { margin-top: 0; margin-bottom: 16px; }
        blockquote {
            padding: 0 1em;
            color: #6a737d;
            border-left: 0.25em solid #dfe2e5;
            margin: 0 0 16px 0;
        }
        ul, ol { padding-left: 2em; margin-top: 0; margin-bottom: 16px; }
        li { word-wrap: break-all; }
        code {
            padding: 0.2em 0.4em;
            margin: 0;
            font-size: 85%;
            background-color: rgba(27,31,35,0.05);
            border-radius: 3px;
        }
        pre {
            padding: 16px;
            overflow: auto;
            font-size: 85%;
            line-height: 1.45;
            background-color: #f6f8fa;
            border-radius: 3px;
        }
        """
    
 