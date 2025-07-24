#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SVG处理器模块
用于处理Markdown文件中的SVG代码，将其转换为PNG图片

作者：AI Assistant
日期：2025年1月27日
"""

import os
import re
import logging
import tempfile
import time
from pathlib import Path
from typing import List, Dict, Tuple, Optional
import hashlib

# 可选依赖：psutil用于内存监控
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

# 图形处理库
try:
    import cairosvg
    CAIROSVG_AVAILABLE = True
except ImportError:
    CAIROSVG_AVAILABLE = False

try:
    from svglib.svglib import renderSVG
    from reportlab.graphics import renderPM
    SVGLIB_AVAILABLE = True
except ImportError:
    SVGLIB_AVAILABLE = False

try:
    import subprocess
    import shutil
    SUBPROCESS_AVAILABLE = True
except ImportError:
    SUBPROCESS_AVAILABLE = False


class SVGProcessor:
    """
    SVG处理器类
    
    负责从Markdown内容中检测、提取SVG代码块，并将其转换为PNG图片。
    支持多种转换方法，包括Python库和外部命令行工具。
    
    主要功能：
    1. 检测和提取Markdown中的SVG代码块
    2. 将SVG代码转换为PNG图片
    3. 替换Markdown中的SVG为图片引用
    4. 管理临时文件的创建和清理
    
    支持的转换方法（按优先级排序）：
    1. cairosvg (Python库，跨平台)
    2. inkscape (命令行工具，高质量)
    3. rsvg-convert (命令行工具，Linux/macOS)
    4. svglib + reportlab (Python库，备选)
    """
    
    # 默认配置
    DEFAULT_DPI = 300
    DEFAULT_OUTPUT_WIDTH = 800
    
    # SVG匹配正则表达式
    SVG_PATTERN = r'<svg[^>]*>.*?</svg>'
    
    def __init__(self, output_dir: str, dpi: int = None, **kwargs):
        """
        初始化SVG处理器
        
        Args:
            output_dir (str): 输出目录路径
            dpi (int, optional): PNG输出DPI，默认300
            **kwargs: 其他配置参数
                - conversion_method (str): 指定转换方法 ('auto', 'cairosvg', 'inkscape', 'rsvg-convert', 'svglib')
                - output_width (int): 输出图片宽度，默认800px
                - fallback_enabled (bool): 是否启用降级处理，默认True
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        self.dpi = dpi or self.DEFAULT_DPI
        self.output_width = kwargs.get('output_width', self.DEFAULT_OUTPUT_WIDTH)
        self.conversion_method = kwargs.get('conversion_method', 'auto')
        self.fallback_enabled = kwargs.get('fallback_enabled', True)
        
        # 临时文件列表
        self.temp_files = []
        
        # 日志记录器
        self.logger = logging.getLogger(self.__class__.__name__)
        
        # 性能监控
        self.performance_stats = {
            'total_conversions': 0,
            'successful_conversions': 0,
            'failed_conversions': 0,
            'total_processing_time': 0.0,
            'average_processing_time': 0.0,
            'memory_usage_peak': 0.0
        }
        
        # 错误统计
        self.error_stats = {
            'dependency_errors': 0,
            'conversion_errors': 0,
            'file_io_errors': 0,
            'timeout_errors': 0
        }
        
        # 检查依赖和工具可用性
        self._check_dependencies()
        
        self.logger.info(f"SVGProcessor初始化完成，输出目录: {self.output_dir}")
    
    def _check_dependencies(self):
        """
        检查依赖库和外部工具的可用性
        """
        self.available_methods = {}
        
        # 检查Python库
        self.available_methods['cairosvg'] = CAIROSVG_AVAILABLE
        self.available_methods['svglib'] = SVGLIB_AVAILABLE
        
        # 检查外部工具
        if SUBPROCESS_AVAILABLE:
            self.available_methods['inkscape'] = self._check_tool_availability('inkscape')
            self.available_methods['rsvg-convert'] = self._check_tool_availability('rsvg-convert')
        else:
            self.available_methods['inkscape'] = False
            self.available_methods['rsvg-convert'] = False
        
        # 记录可用方法
        available_list = [method for method, available in self.available_methods.items() if available]
        if available_list:
            self.logger.info(f"可用的SVG转换方法: {', '.join(available_list)}")
        else:
            self.logger.warning("未检测到可用的SVG转换方法")
        
        # 警告缺失的依赖
        if not CAIROSVG_AVAILABLE:
            self.logger.warning("cairosvg库未安装，建议安装: pip install cairosvg")
        if not SVGLIB_AVAILABLE:
            self.logger.warning("svglib库未安装，建议安装: pip install svglib reportlab")
    
    def _check_tool_availability(self, tool_name: str) -> bool:
        """
        检查外部工具是否可用
        
        Args:
            tool_name (str): 工具名称
            
        Returns:
            bool: 工具是否可用
        """
        try:
            return shutil.which(tool_name) is not None
        except Exception:
            return False
    
    def process_markdown_content(self, content: str, base_dir: Path, output_format: str = 'docx', markdown_filename: Optional[str] = None) -> Tuple[str, List[str]]:
        """
        处理Markdown内容中的SVG代码块
        
        Args:
            content (str): 原始Markdown内容
            base_dir (Path): Markdown文件所在目录
            output_format (str): 输出格式 ('docx', 'pdf', 'html', 'pptx')
            markdown_filename (Optional[str]): Markdown文件名（不含扩展名），用于生成PNG文件名
            
        Returns:
            Tuple[str, List[str]]: (处理后的内容, 生成的临时文件列表)
        """
        self.logger.info("开始处理Markdown内容中的SVG")
        
        # 提取SVG代码块
        svg_blocks = self.extract_svg_blocks(content)
        
        if not svg_blocks:
            self.logger.info("未检测到SVG代码块")
            return content, []
        
        self.logger.info(f"检测到 {len(svg_blocks)} 个SVG代码块")
        
        # 处理每个SVG
        processed_content = content
        temp_files = []
        
        for i, svg_info in enumerate(svg_blocks):
            svg_content = svg_info['content']
            
            # 生成PNG文件名
            base_name = markdown_filename if markdown_filename else base_dir.stem
            png_filename = self._generate_png_filename(i + 1, base_name)
            png_path = self.output_dir / png_filename
            
            # 转换SVG为PNG
            success = self.convert_svg_to_png(svg_content, png_path)
            
            if success:
                # 根据输出格式决定图片引用路径
                if output_format == 'html':
                    # HTML格式：图片在svg_temp子目录中，使用相对路径
                    img_reference = f"![图{i+1}](svg_temp/{png_filename})"
                    temp_files.append(str(png_path))  # HTML格式保留在svg_temp目录
                    self.logger.info(f"HTML格式：PNG文件保存在svg_temp目录: {png_path}")
                else:
                    # DOCX/PDF/PPTX格式：将PNG文件复制到Markdown文件所在目录，确保Pandoc能找到
                    target_png_path = base_dir / png_filename
                    try:
                        import shutil
                        shutil.copy2(str(png_path), str(target_png_path))
                        temp_files.append(str(target_png_path))  # 记录复制后的文件用于清理
                        self.logger.info(f"PNG文件已复制到: {target_png_path}")
                    except Exception as e:
                        self.logger.warning(f"复制PNG文件失败: {e}，使用绝对路径")
                        target_png_path = png_path
                        temp_files.append(str(png_path))
                    
                    # 使用相对路径，因为图片已复制到MD文件目录
                    img_reference = f"![图{i+1}]({png_filename})"
                
                processed_content = processed_content.replace(svg_content, img_reference, 1)
                self.logger.info(f"成功转换SVG {i+1}: {png_filename}")
            else:
                # 转换失败的处理
                if self.fallback_enabled:
                    # 保留原始SVG作为代码块
                    fallback_content = f"```svg\n{svg_content}\n```"
                    processed_content = processed_content.replace(svg_content, fallback_content, 1)
                    self.logger.warning(f"SVG {i+1} 转换失败，保留为代码块")
                else:
                    self.logger.error(f"SVG {i+1} 转换失败")
        
        # 更新临时文件列表
        self.temp_files.extend(temp_files)
        
        self.logger.info(f"SVG处理完成，生成了 {len(temp_files)} 个PNG文件")
        return processed_content, temp_files
    
    def extract_svg_blocks(self, content: str) -> List[Dict[str, any]]:
        """
        从Markdown内容中提取SVG代码块
        
        Args:
            content (str): Markdown内容
            
        Returns:
            List[Dict]: SVG信息列表，每个字典包含:
                - content: SVG代码内容
                - start: 开始位置
                - end: 结束位置
        """
        svg_blocks = []
        
        # 使用正则表达式查找所有SVG代码块
        pattern = re.compile(self.SVG_PATTERN, re.DOTALL | re.IGNORECASE)
        
        for match in pattern.finditer(content):
            svg_info = {
                'content': match.group(0),
                'start': match.start(),
                'end': match.end()
            }
            svg_blocks.append(svg_info)
        
        return svg_blocks
    
    def convert_svg_to_png(self, svg_content: str, output_path: Path) -> bool:
        """
        将SVG内容转换为PNG文件
        
        Args:
            svg_content (str): SVG代码内容
            output_path (Path): 输出PNG文件路径
            
        Returns:
            bool: 转换是否成功
        """
        start_time = time.time()
        start_memory = self._get_memory_usage()
        
        try:
            self.performance_stats['total_conversions'] += 1
            
            if not svg_content or not svg_content.strip():
                self.logger.warning("SVG内容为空")
                self.error_stats['conversion_errors'] += 1
                return False
            
            # 确保输出目录存在
            try:
                output_path.parent.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                self.logger.error(f"创建输出目录失败: {e}")
                self.error_stats['file_io_errors'] += 1
                return False
            
            # 增强SVG内容
            enhanced_svg = self._enhance_svg_content(svg_content)
            
            # 根据配置选择转换方法
            if self.conversion_method == 'auto':
                # 自动选择最佳可用方法
                methods = [
                    ('cairosvg', self._convert_with_cairosvg),
                    ('inkscape', self._convert_with_inkscape),
                    ('rsvg-convert', self._convert_with_rsvg),
                    ('svglib', self._convert_with_svglib)
                ]
            else:
                # 使用指定方法
                method_map = {
                    'cairosvg': self._convert_with_cairosvg,
                    'inkscape': self._convert_with_inkscape,
                    'rsvg-convert': self._convert_with_rsvg,
                    'svglib': self._convert_with_svglib
                }
                if self.conversion_method in method_map:
                    methods = [(self.conversion_method, method_map[self.conversion_method])]
                else:
                    self.logger.error(f"不支持的转换方法: {self.conversion_method}")
                    self.error_stats['conversion_errors'] += 1
                    return False
            
            # 尝试转换
            last_error = None
            for method_name, method_func in methods:
                if not self.available_methods.get(method_name, False):
                    continue
                    
                try:
                    method_start_time = time.time()
                    success = method_func(enhanced_svg, output_path)
                    method_end_time = time.time()
                    
                    if success:
                        self.logger.debug(f"使用 {method_name} 成功转换SVG，耗时: {method_end_time - method_start_time:.3f}秒")
                        self.performance_stats['successful_conversions'] += 1
                        return True
                    else:
                        self.logger.debug(f"{method_name} 转换返回失败")
                        
                except Exception as e:
                    last_error = e
                    self.logger.debug(f"{method_name} 转换异常: {str(e)}")
                    if 'dependency' in str(e).lower() or 'not found' in str(e).lower():
                        self.error_stats['dependency_errors'] += 1
                    else:
                        self.error_stats['conversion_errors'] += 1
                    continue
            
            # 所有方法都失败
            self.logger.error(f"所有转换方法都失败了，最后错误: {last_error}")
            self.performance_stats['failed_conversions'] += 1
            self.error_stats['conversion_errors'] += 1
            return False
            
        except Exception as e:
            self.logger.error(f"转换过程中发生未预期错误: {e}")
            self.performance_stats['failed_conversions'] += 1
            self.error_stats['conversion_errors'] += 1
            return False
            
        finally:
            # 记录性能指标
            end_time = time.time()
            processing_time = end_time - start_time
            self.performance_stats['total_processing_time'] += processing_time
            self.performance_stats['average_processing_time'] = (
                self.performance_stats['total_processing_time'] / 
                self.performance_stats['total_conversions']
            )
            
            # 记录内存使用峰值
            current_memory = self._get_memory_usage()
            memory_used = current_memory - start_memory
            if memory_used > self.performance_stats['memory_usage_peak']:
                self.performance_stats['memory_usage_peak'] = memory_used
    
    def _enhance_svg_content(self, svg_content: str) -> str:
        """
        增强SVG内容，添加必要的属性和中文字体支持
        
        Args:
            svg_content (str): 原始SVG内容
            
        Returns:
            str: 增强后的SVG内容
        """
        # 确保SVG有正确的命名空间
        if '<svg' in svg_content and 'xmlns=' not in svg_content:
            svg_content = svg_content.replace(
                '<svg',
                '<svg xmlns="http://www.w3.org/2000/svg"'
            )
        
        # 检查是否已有font-family设置
        if 'font-family' not in svg_content:
            # 为text元素添加中文字体支持
            chinese_fonts = "Microsoft YaHei, SimHei, SimSun, Arial, sans-serif"
            svg_content = svg_content.replace(
                '<text',
                f'<text font-family="{chinese_fonts}"'
            )
        
        # 确保SVG有合适的尺寸
        if 'width=' not in svg_content and 'viewBox=' not in svg_content:
            svg_content = svg_content.replace(
                '<svg',
                f'<svg width="{self.output_width}" height="{self.output_width}"'
            )
        
        return svg_content
    
    def _generate_png_filename(self, svg_index: int, base_name: str) -> str:
        """
        生成PNG文件名
        
        Args:
            svg_index (int): SVG索引（从1开始）
            base_name (str): 基础文件名（Markdown文件名）
            
        Returns:
            str: PNG文件名，格式为 {base_name}_{svg_index:02d}.png
        """
        return f"{base_name}_{svg_index:02d}.png"
    
    def _convert_with_cairosvg(self, svg_content: str, output_path: Path) -> bool:
        """
        使用cairosvg转换SVG
        
        Args:
            svg_content (str): SVG内容
            output_path (Path): 输出路径
            
        Returns:
            bool: 转换是否成功
        """
        if not CAIROSVG_AVAILABLE:
            return False
        
        try:
            # 动态导入以避免NameError
            import cairosvg
            cairosvg.svg2png(
                bytestring=svg_content.encode('utf-8'),
                write_to=str(output_path),
                output_width=self.output_width,
                dpi=self.dpi
            )
            return True
        except Exception as e:
            self.logger.debug(f"cairosvg转换失败: {str(e)}")
            return False
    
    def _convert_with_inkscape(self, svg_content: str, output_path: Path) -> bool:
        """
        使用inkscape转换SVG
        
        Args:
            svg_content (str): SVG内容
            output_path (Path): 输出路径
            
        Returns:
            bool: 转换是否成功
        """
        if not self.available_methods.get('inkscape', False):
            return False
        
        # 创建临时SVG文件
        temp_svg = self.output_dir / f"temp_{os.getpid()}_{int(time.time())}.svg"
        
        try:
            # 写入临时SVG文件
            with open(temp_svg, 'w', encoding='utf-8') as f:
                f.write(svg_content)
            
            # 调用inkscape
            cmd = [
                "inkscape",
                str(temp_svg),
                "--export-type=png",
                f"--export-filename={output_path}",
                f"--export-dpi={self.dpi}"
            ]
            
            result = subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            return result.returncode == 0
            
        except Exception as e:
            self.logger.debug(f"inkscape转换失败: {str(e)}")
            return False
        finally:
            # 清理临时文件
            if temp_svg.exists():
                try:
                    temp_svg.unlink()
                except Exception:
                    pass
    
    def _convert_with_rsvg(self, svg_content: str, output_path: Path) -> bool:
        """
        使用rsvg-convert转换SVG
        
        Args:
            svg_content (str): SVG内容
            output_path (Path): 输出路径
            
        Returns:
            bool: 转换是否成功
        """
        if not self.available_methods.get('rsvg-convert', False):
            return False
        
        try:
            cmd = [
                "rsvg-convert",
                "-o", str(output_path),
                "-d", str(self.dpi),
                "-p", str(self.dpi)
            ]
            
            result = subprocess.run(
                cmd,
                input=svg_content,
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE
            )
            return result.returncode == 0
            
        except Exception as e:
            self.logger.debug(f"rsvg-convert转换失败: {str(e)}")
            return False
    
    def _convert_with_svglib(self, svg_content: str, output_path: Path) -> bool:
        """
        使用svglib转换SVG
        
        Args:
            svg_content (str): SVG内容
            output_path (Path): 输出路径
            
        Returns:
            bool: 转换是否成功
        """
        if not SVGLIB_AVAILABLE:
            return False
        
        # 创建临时SVG文件
        temp_svg = self.output_dir / f"temp_{os.getpid()}_{int(time.time())}.svg"
        
        try:
            # 写入临时SVG文件
            with open(temp_svg, 'w', encoding='utf-8') as f:
                f.write(svg_content)
            
            # 动态导入以避免NameError
            from svglib.svglib import renderSVG
            from reportlab.graphics import renderPM
            
            # 使用svglib转换
            drawing = renderSVG.renderSVG(str(temp_svg))
            renderPM.drawToFile(drawing, str(output_path), fmt="PNG", dpi=self.dpi)
            return True
            
        except Exception as e:
            self.logger.debug(f"svglib转换失败: {str(e)}")
            return False
        finally:
            # 清理临时文件
            if temp_svg.exists():
                try:
                    temp_svg.unlink()
                except Exception:
                    pass
    
    def get_performance_report(self) -> dict:
        """
        获取性能报告
        
        Returns:
            dict: 包含性能统计信息的字典
        """
        total_conversions = self.performance_stats['total_conversions']
        success_rate = (self.performance_stats['successful_conversions'] / total_conversions * 100) if total_conversions > 0 else 0
        
        return {
            'performance': {
                'total_conversions': total_conversions,
                'successful_conversions': self.performance_stats['successful_conversions'],
                'failed_conversions': self.performance_stats['failed_conversions'],
                'success_rate': f"{success_rate:.2f}%",
                'total_processing_time': f"{self.performance_stats['total_processing_time']:.3f}s",
                'average_processing_time': f"{self.performance_stats['average_processing_time']:.3f}s",
                'memory_usage_peak': f"{self.performance_stats['memory_usage_peak']:.2f}MB"
            },
            'errors': {
                'dependency_errors': self.error_stats['dependency_errors'],
                'conversion_errors': self.error_stats['conversion_errors'],
                'file_io_errors': self.error_stats['file_io_errors'],
                'timeout_errors': self.error_stats['timeout_errors']
            }
        }
    
    def reset_statistics(self):
        """
        重置所有统计信息
        """
        self.performance_stats = {
            'total_conversions': 0,
            'successful_conversions': 0,
            'failed_conversions': 0,
            'total_processing_time': 0.0,
            'average_processing_time': 0.0,
            'memory_usage_peak': 0.0
        }
        
        self.error_stats = {
            'dependency_errors': 0,
            'conversion_errors': 0,
            'file_io_errors': 0,
            'timeout_errors': 0
        }
        
        self.logger.info("SVGProcessor统计信息已重置")
    
    def cleanup_temp_files(self):
        """
        清理所有临时文件
        """
        cleaned_count = 0
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                    cleaned_count += 1
            except Exception as e:
                self.logger.warning(f"无法删除临时文件 {temp_file}: {e}")
        
        if cleaned_count > 0:
            self.logger.info(f"清理了 {cleaned_count} 个临时文件")
        
        # 清空临时文件列表
        self.temp_files.clear()
    
    def _get_memory_usage(self) -> float:
        """
        获取当前内存使用量（MB）
        
        Returns:
            float: 内存使用量（MB）
        """
        if not PSUTIL_AVAILABLE:
            return 0.0
            
        try:
            # 动态导入以避免NameError
            import psutil
            process = psutil.Process()
            return process.memory_info().rss / 1024 / 1024  # 转换为MB
        except Exception:
            return 0.0
    
    def get_conversion_stats(self) -> Dict[str, any]:
        """
        获取转换统计信息
        
        Returns:
            Dict: 统计信息
        """
        return {
            'available_methods': self.available_methods,
            'temp_files_count': len(self.temp_files),
            'output_dir': str(self.output_dir),
            'dpi': self.dpi,
            'conversion_method': self.conversion_method,
            'performance_stats': self.performance_stats.copy(),
            'error_stats': self.error_stats.copy()
        }