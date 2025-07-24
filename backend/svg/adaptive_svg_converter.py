#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
自适应SVG转换器模块
用于在各种文档转换过程中自动检测和转换SVG内容

作者：AI Assistant
日期：2025年1月27日
"""

import re
import os
import logging
from pathlib import Path
from typing import List, Dict, Tuple, Optional, Union
import hashlib
import tempfile
import shutil
import time
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
import json

try:
    from .svg_detector import SVGDetector, SVGBlock
    from ..converters.svg_processor import SVGProcessor
except ImportError:
    # 处理相对导入问题
    import sys
    from pathlib import Path
    backend_path = Path(__file__).parent.parent
    if str(backend_path) not in sys.path:
        sys.path.insert(0, str(backend_path))
    
    from .svg_detector import SVGDetector, SVGBlock


class AdaptiveSVGConverter:
    """
    自适应SVG转换器
    
    在文档转换过程中自动检测Markdown内容中的SVG，
    并将其转换为PNG图片，支持多种输出格式（PDF、HTML、DOCX、PPTX等）。
    
    主要功能：
    1. 自动检测Markdown中的SVG内容
    2. 智能选择最佳转换方法
    3. 管理临时文件和清理
    4. 提供统一的转换接口
    5. 支持批量处理
    """
    
    def __init__(self, 
                 output_dir: Optional[str] = None,
                 dpi: int = 300,
                 **kwargs):
        """
        初始化自适应SVG转换器
        
        Args:
            output_dir (str, optional): 输出目录，默认使用临时目录
            dpi (int): PNG输出DPI，默认300
            **kwargs: 其他配置参数
                - conversion_method (str): 转换方法 ('auto', 'cairosvg', 'inkscape', 'rsvg-convert', 'svglib')
                - include_xml_blocks (bool): 是否包含XML代码块中的SVG，默认True
                - case_sensitive (bool): 是否大小写敏感，默认False
                - validate_svg_content (bool): 是否验证SVG内容，默认True
                - cleanup_on_exit (bool): 是否在退出时自动清理，默认True
        """
        # 设置输出目录
        if output_dir:
            self.output_dir = Path(output_dir)
            self.output_dir.mkdir(parents=True, exist_ok=True)
            self._temp_dir_created = False
        else:
            self.output_dir = Path(tempfile.mkdtemp(prefix='svg_convert_'))
            self._temp_dir_created = True
        
        self.dpi = dpi
        self.cleanup_on_exit = kwargs.get('cleanup_on_exit', True)
        
        # 初始化SVG检测器
        detector_config = {
            'include_xml_blocks': kwargs.get('include_xml_blocks', True),
            'case_sensitive': kwargs.get('case_sensitive', False),
            'validate_svg_content': kwargs.get('validate_svg_content', True)
        }
        self.svg_detector = SVGDetector(**detector_config)
        
        # 延迟导入SVG处理器以避免循环导入
        try:
            from ..converters.svg_processor import SVGProcessor
        except ImportError:
            import sys
            sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
            from converters.svg_processor import SVGProcessor
        
        # 初始化SVG处理器
        processor_config = {
            'conversion_method': kwargs.get('conversion_method', 'auto'),
            'output_width': kwargs.get('output_width', 800),
            'fallback_enabled': kwargs.get('fallback_enabled', True)
        }
        self.svg_processor = SVGProcessor(
            output_dir=str(self.output_dir),
            dpi=self.dpi,
            **processor_config
        )
        
        # 转换统计
        self.conversion_stats = {
            'total_processed': 0,
            'svg_detected': 0,
            'svg_converted': 0,
            'conversion_failed': 0,
            'files_created': [],
            'processing_time': 0.0,
            'cache_hits': 0,
            'cache_misses': 0
        }
        
        # 性能监控
        self.performance_metrics = {
            'start_time': None,
            'end_time': None,
            'svg_processing_times': [],
            'memory_usage': []
        }
        
        # SVG缓存（基于内容哈希）
        self.svg_cache = {}
        self.cache_enabled = kwargs.get('cache_enabled', True)
        self.max_cache_size = kwargs.get('max_cache_size', 100)
        
        # 并行处理配置
        self.parallel_enabled = kwargs.get('parallel_enabled', True)
        self.max_workers = kwargs.get('max_workers', min(4, os.cpu_count() or 1))
        
        # 线程锁
        self._cache_lock = threading.Lock()
        self._stats_lock = threading.Lock()
        
        # 日志记录器
        self.logger = logging.getLogger(self.__class__.__name__)
        self.logger.info(f"自适应SVG转换器初始化完成，输出目录: {self.output_dir}")
    
    def __enter__(self):
        """上下文管理器入口"""
        return self
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器退出，自动清理"""
        if self.cleanup_on_exit:
            self.cleanup()
    
    def process_markdown(self, 
                        content: str, 
                        base_dir: Optional[Path] = None,
                        markdown_filename: Optional[str] = None,
                        target_format: str = 'png') -> Tuple[str, Dict]:
        """
        处理Markdown内容，自动转换其中的SVG
        
        Args:
            content (str): Markdown内容
            base_dir (Path, optional): 基础目录，用于相对路径处理
            markdown_filename (str, optional): Markdown文件名（不含扩展名），用于生成PNG文件名
            target_format (str): 目标格式，默认'png'
            
        Returns:
            Tuple[str, Dict]: (处理后的内容, 转换信息)
        """
        start_time = time.time()
        self.performance_metrics['start_time'] = start_time
        
        with self._stats_lock:
            self.conversion_stats['total_processed'] += 1
        
        if not content or not content.strip():
            return content, {'svg_count': 0, 'converted_files': []}
        
        # 快速检测是否包含SVG
        if not self.svg_detector.has_svg_content(content):
            self.logger.debug("未检测到SVG内容，跳过处理")
            return content, {'svg_count': 0, 'converted_files': []}
        
        # 详细检测SVG块
        svg_blocks = self.svg_detector.detect_svg_blocks(content)
        if not svg_blocks:
            return content, {'svg_count': 0, 'converted_files': []}
        
        with self._stats_lock:
            self.conversion_stats['svg_detected'] += len(svg_blocks)
        self.logger.info(f"检测到 {len(svg_blocks)} 个SVG块")
        
        # 处理每个SVG块
        processed_content = content
        converted_files = []
        
        # 按位置倒序处理，避免位置偏移问题
        svg_blocks_reversed = sorted(svg_blocks, key=lambda x: x.start_pos, reverse=True)
        
        # 选择处理方式：并行或串行
        if self.parallel_enabled and len(svg_blocks_reversed) > 1:
            converted_files = self._process_svg_blocks_parallel(svg_blocks_reversed, base_dir, markdown_filename)
        else:
            converted_files = self._process_svg_blocks_sequential(svg_blocks_reversed, base_dir, markdown_filename)
        
        # 替换内容
        for i, svg_block in enumerate(svg_blocks_reversed):
            file_index = len(svg_blocks) - i
            if i < len(converted_files) and converted_files[i]:
                final_path = converted_files[i]
                img_reference = self._create_image_reference(
                    Path(final_path), base_dir, file_index
                )
                
                # 使用位置信息精确替换
                processed_content = (
                    processed_content[:svg_block.start_pos] + 
                    img_reference + 
                    processed_content[svg_block.end_pos:]
                )
        
        # 记录处理时间
        end_time = time.time()
        processing_time = end_time - start_time
        self.performance_metrics['end_time'] = end_time
        
        with self._stats_lock:
            self.conversion_stats['processing_time'] += processing_time
        
        # 过滤有效的转换文件
        valid_converted_files = [f for f in converted_files if f]
        
        conversion_info = {
            'svg_count': len(svg_blocks),
            'converted_files': valid_converted_files,
            'success_count': len(valid_converted_files),
            'failed_count': len(svg_blocks) - len(valid_converted_files),
            'processing_time': processing_time,
            'cache_hits': self.conversion_stats['cache_hits'],
            'cache_misses': self.conversion_stats['cache_misses']
        }
        
        return processed_content, conversion_info

    def _process_svg_blocks_sequential(self, svg_blocks: List, base_dir: Optional[Path], markdown_filename: Optional[str] = None) -> List[Optional[str]]:
        """
        串行处理SVG块
        
        Args:
            svg_blocks: SVG块列表
            base_dir: 基础目录
            markdown_filename: Markdown文件名
            
        Returns:
            List[Optional[str]]: 转换后的文件路径列表
        """
        converted_files = []
        
        for i, svg_block in enumerate(svg_blocks):
            file_index = len(svg_blocks) - i
            converted_file = self._process_single_svg_block(svg_block, file_index, base_dir, markdown_filename)
            converted_files.append(converted_file)
            
        return converted_files
    
    def _process_svg_blocks_parallel(self, svg_blocks: List, base_dir: Optional[Path], markdown_filename: Optional[str] = None) -> List[Optional[str]]:
        """
        并行处理SVG块
        
        Args:
            svg_blocks: SVG块列表
            base_dir: 基础目录
            markdown_filename: Markdown文件名
            
        Returns:
            List[Optional[str]]: 转换后的文件路径列表
        """
        converted_files = [None] * len(svg_blocks)
        
        with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
            # 提交所有任务
            future_to_index = {}
            for i, svg_block in enumerate(svg_blocks):
                file_index = len(svg_blocks) - i
                future = executor.submit(self._process_single_svg_block, svg_block, file_index, base_dir, markdown_filename)
                future_to_index[future] = i
            
            # 收集结果
            for future in as_completed(future_to_index):
                index = future_to_index[future]
                try:
                    result = future.result()
                    converted_files[index] = result
                except Exception as e:
                    self.logger.error(f"并行处理SVG块 {index} 时出错: {e}")
                    converted_files[index] = None
        
        return converted_files
    
    def _process_single_svg_block(self, svg_block, file_index: int, base_dir: Optional[Path], markdown_filename: Optional[str] = None) -> Optional[str]:
        """
        处理单个SVG块
        
        Args:
            svg_block: SVG块对象
            file_index: 文件索引
            base_dir: 基础目录
            markdown_filename: Markdown文件名
            
        Returns:
            Optional[str]: 转换后的文件路径，失败时返回None
        """
        try:
            # 检查缓存
            svg_hash = self._get_svg_hash(svg_block.content)
            cached_file = self._get_from_cache(svg_hash)
            
            if cached_file:
                with self._stats_lock:
                    self.conversion_stats['cache_hits'] += 1
                self.logger.debug(f"SVG {file_index} 命中缓存: {cached_file}")
                return cached_file
            
            with self._stats_lock:
                self.conversion_stats['cache_misses'] += 1
            
            # 生成唯一的文件名
            png_filename = self._generate_unique_filename(file_index, base_dir, markdown_filename)
            png_path = self.output_dir / png_filename
            
            # 转换SVG为PNG
            svg_start_time = time.time()
            success = self.svg_processor.convert_svg_to_png(
                svg_block.content, 
                png_path
            )
            svg_end_time = time.time()
            
            # 记录处理时间
            self.performance_metrics['svg_processing_times'].append(svg_end_time - svg_start_time)
            
            if success:
                # 处理文件路径
                final_path = self._handle_output_file(png_path, base_dir)
                final_path_str = str(final_path)
                
                with self._stats_lock:
                    self.conversion_stats['files_created'].append(final_path_str)
                    self.conversion_stats['svg_converted'] += 1
                
                # 添加到缓存
                self._add_to_cache(svg_hash, final_path_str)
                
                self.logger.info(f"成功转换SVG {file_index}: {png_filename}")
                return final_path_str
            else:
                with self._stats_lock:
                    self.conversion_stats['conversion_failed'] += 1
                self.logger.warning(f"SVG转换失败: {svg_block.content[:50]}...")
                return None
                
        except Exception as e:
            with self._stats_lock:
                self.conversion_stats['conversion_failed'] += 1
            self.logger.error(f"处理SVG块时出错: {e}")
            return None
    
    def _get_svg_hash(self, svg_content: str) -> str:
        """
        获取SVG内容的哈希值
        
        Args:
            svg_content: SVG内容
            
        Returns:
            str: 哈希值
        """
        return hashlib.md5(svg_content.encode('utf-8')).hexdigest()
    
    def _get_from_cache(self, svg_hash: str) -> Optional[str]:
        """
        从缓存中获取文件路径
        
        Args:
            svg_hash: SVG哈希值
            
        Returns:
            Optional[str]: 缓存的文件路径，未找到时返回None
        """
        if not self.cache_enabled:
            return None
            
        with self._cache_lock:
            cached_info = self.svg_cache.get(svg_hash)
            if cached_info and Path(cached_info['file_path']).exists():
                # 更新访问时间
                cached_info['last_accessed'] = time.time()
                return cached_info['file_path']
            elif cached_info:
                # 文件不存在，从缓存中移除
                del self.svg_cache[svg_hash]
        
        return None
    
    def _add_to_cache(self, svg_hash: str, file_path: str):
        """
        添加文件到缓存
        
        Args:
            svg_hash: SVG哈希值
            file_path: 文件路径
        """
        if not self.cache_enabled:
            return
            
        with self._cache_lock:
            # 检查缓存大小限制
            if len(self.svg_cache) >= self.max_cache_size:
                self._cleanup_cache()
            
            self.svg_cache[svg_hash] = {
                'file_path': file_path,
                'created_time': time.time(),
                'last_accessed': time.time()
            }
    
    def _cleanup_cache(self):
        """
        清理缓存，移除最旧的条目
        """
        if not self.svg_cache:
            return
            
        # 按最后访问时间排序，移除最旧的条目
        sorted_items = sorted(
            self.svg_cache.items(),
            key=lambda x: x[1]['last_accessed']
        )
        
        # 移除最旧的25%条目
        remove_count = max(1, len(sorted_items) // 4)
        for i in range(remove_count):
            svg_hash, _ = sorted_items[i]
            del self.svg_cache[svg_hash]

    def _generate_unique_filename(self, index: int, base_dir: Optional[Path] = None, markdown_filename: Optional[str] = None) -> str:
        """
        生成PNG文件名
        
        Args:
            index (int): SVG索引
            base_dir (Path, optional): 基础目录
            markdown_filename (str, optional): Markdown文件名
            
        Returns:
            str: PNG文件名，格式为 {markdown_filename}_{index:02d}.png
        """
        # 基础名称优先使用markdown_filename
        if markdown_filename:
            base_name = markdown_filename
        elif base_dir:
            if isinstance(base_dir, str):
                base_name = Path(base_dir).stem
            else:
                base_name = base_dir.stem
        else:
            base_name = "svg"
        
        return f"{base_name}_{index:02d}.png"
    
    def _handle_output_file(self, png_path: Path, base_dir: Optional[Path] = None) -> Path:
        """
        处理输出文件，根据需要复制到目标目录
        
        Args:
            png_path (Path): 原始PNG路径
            base_dir (Path, optional): 目标基础目录
            
        Returns:
            Path: 最终文件路径
        """
        if not base_dir:
            return png_path
            
        # 确保base_dir是Path对象
        if isinstance(base_dir, str):
            base_dir = Path(base_dir)
            
        if not base_dir.exists():
            return png_path
        
        try:
            # 复制到Markdown文件所在目录
            target_path = base_dir / png_path.name
            shutil.copy2(str(png_path), str(target_path))
            self.logger.debug(f"文件已复制到: {target_path}")
            return target_path
        except Exception as e:
            self.logger.warning(f"复制文件失败: {e}，使用原始路径")
            return png_path
    
    def _create_image_reference(self, 
                               image_path: Path, 
                               base_dir: Optional[Path] = None,
                               index: int = 1) -> str:
        """
        创建图片引用的Markdown语法
        
        Args:
            image_path (Path): 图片路径
            base_dir (Path, optional): 基础目录
            index (int): 图片索引
            
        Returns:
            str: Markdown图片引用
        """
        # 计算相对路径
        if base_dir:
            # 确保base_dir是Path对象
            if isinstance(base_dir, str):
                base_dir = Path(base_dir)
                
            if base_dir.exists():
                try:
                    rel_path = image_path.relative_to(base_dir)
                    path_str = str(rel_path).replace('\\', '/')
                except ValueError:
                    # 无法计算相对路径，使用绝对路径
                    path_str = str(image_path).replace('\\', '/')
            else:
                path_str = str(image_path).replace('\\', '/')
        else:
            path_str = str(image_path).replace('\\', '/')
        
        return f"![图{index}]({path_str})"
    
    def batch_process_files(self, 
                           file_paths: List[Union[str, Path]],
                           output_dir: Optional[Path] = None) -> Dict[str, Dict]:
        """
        批量处理多个Markdown文件
        
        Args:
            file_paths (List[Union[str, Path]]): 文件路径列表
            output_dir (Path, optional): 输出目录
            
        Returns:
            Dict[str, Dict]: 处理结果，键为文件路径，值为处理信息
        """
        results = {}
        
        for file_path in file_paths:
            file_path = Path(file_path)
            
            if not file_path.exists() or file_path.suffix.lower() != '.md':
                results[str(file_path)] = {
                    'success': False,
                    'error': 'File not found or not a Markdown file'
                }
                continue
            
            try:
                # 读取文件内容
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                # 处理内容
                processed_content, conversion_info = self.process_markdown(
                    content, file_path.parent, file_path.stem
                )
                
                # 保存处理后的内容（如果指定了输出目录）
                if output_dir:
                    output_dir = Path(output_dir)
                    output_dir.mkdir(parents=True, exist_ok=True)
                    output_file = output_dir / file_path.name
                    
                    with open(output_file, 'w', encoding='utf-8') as f:
                        f.write(processed_content)
                    
                    conversion_info['output_file'] = str(output_file)
                
                results[str(file_path)] = {
                    'success': True,
                    'conversion_info': conversion_info
                }
                
            except Exception as e:
                results[str(file_path)] = {
                    'success': False,
                    'error': str(e)
                }
                self.logger.error(f"处理文件 {file_path} 时出错: {e}")
        
        return results
    
    def get_conversion_statistics(self) -> Dict:
        """
        获取转换统计信息
        
        Returns:
            Dict: 统计信息
        """
        stats = self.conversion_stats.copy()
        stats['success_rate'] = (
            stats['svg_converted'] / max(stats['svg_detected'], 1) * 100
        )
        return stats
    
    def cleanup(self):
        """
        清理临时文件和目录
        """
        try:
            # 清理SVG处理器的临时文件
            if hasattr(self.svg_processor, 'cleanup_temp_files'):
                self.svg_processor.cleanup_temp_files()
            
            # 清理临时目录（如果是我们创建的）
            if self._temp_dir_created and self.output_dir.exists():
                shutil.rmtree(str(self.output_dir))
                self.logger.info(f"已清理临时目录: {self.output_dir}")
            
        except Exception as e:
            self.logger.warning(f"清理时出错: {e}")
    
    def cleanup_selective(self, preserve_png: bool = False, remove_svg: bool = False):
        """
        选择性清理临时文件
        
        Args:
            preserve_png (bool): 是否保留PNG文件
            remove_svg (bool): 是否删除SVG文件
        """
        try:
            # 选择性清理SVG处理器的临时文件
            if hasattr(self.svg_processor, 'cleanup_temp_files_selective'):
                self.svg_processor.cleanup_temp_files_selective(preserve_png=preserve_png, remove_svg=remove_svg)
            elif hasattr(self.svg_processor, 'temp_files'):
                # 手动进行选择性清理
                files_to_remove = []
                for temp_file in self.svg_processor.temp_files:
                    if preserve_png and temp_file.lower().endswith('.png'):
                        self.logger.info(f"保留PNG文件: {temp_file}")
                        continue
                    if remove_svg and temp_file.lower().endswith('.svg'):
                        self.logger.info(f"删除SVG文件: {temp_file}")
                        files_to_remove.append(temp_file)
                    elif not preserve_png:
                        files_to_remove.append(temp_file)
                
                # 删除选定的文件
                for temp_file in files_to_remove:
                    try:
                        if os.path.exists(temp_file):
                            os.remove(temp_file)
                            self.logger.debug(f"已删除临时文件: {temp_file}")
                    except Exception as e:
                        self.logger.warning(f"无法删除临时文件 {temp_file}: {e}")
                
                # 从临时文件列表中移除已删除的文件
                self.svg_processor.temp_files = [
                    f for f in self.svg_processor.temp_files if f not in files_to_remove
                ]
            
            # 如果不保留PNG文件，则清理临时目录
            if not preserve_png and self._temp_dir_created and self.output_dir.exists():
                shutil.rmtree(str(self.output_dir))
                self.logger.info(f"已清理临时目录: {self.output_dir}")
            
        except Exception as e:
            self.logger.warning(f"选择性清理时出错: {e}")
    
    def is_svg_content_present(self, content: str) -> bool:
        """
        快速检测内容是否包含SVG
        
        Args:
            content (str): 内容字符串
            
        Returns:
            bool: 是否包含SVG
        """
        return self.svg_detector.has_svg_content(content)
    
    def get_svg_statistics(self, content: str) -> Dict[str, int]:
        """
        获取内容中的SVG统计信息
        
        Args:
            content (str): 内容字符串
            
        Returns:
            Dict[str, int]: 统计信息
        """
        svg_blocks = self.svg_detector.detect_svg_blocks(content)
        return {
            'total_svg_blocks': len(svg_blocks),
            'code_blocks': len([b for b in svg_blocks if b.block_type == 'code_block']),
            'inline_svg': len([b for b in svg_blocks if b.block_type == 'inline'])
        }
    
    def get_performance_report(self) -> Dict:
        """
        获取性能报告
        
        Returns:
            Dict: 性能统计信息
        """
        with self._stats_lock:
            stats = self.conversion_stats.copy()
        
        # 计算平均处理时间
        avg_svg_time = 0
        if self.performance_metrics['svg_processing_times']:
            avg_svg_time = sum(self.performance_metrics['svg_processing_times']) / len(self.performance_metrics['svg_processing_times'])
        
        # 计算缓存命中率
        total_requests = stats['cache_hits'] + stats['cache_misses']
        cache_hit_rate = (stats['cache_hits'] / total_requests * 100) if total_requests > 0 else 0
        
        # 计算成功率
        total_svg = stats['svg_detected']
        success_rate = (stats['svg_converted'] / total_svg * 100) if total_svg > 0 else 0
        
        return {
            'conversion_stats': stats,
            'performance_metrics': {
                'total_processing_time': stats['processing_time'],
                'average_svg_processing_time': avg_svg_time,
                'cache_hit_rate': cache_hit_rate,
                'success_rate': success_rate,
                'parallel_enabled': self.parallel_enabled,
                'max_workers': self.max_workers,
                'cache_size': len(self.svg_cache)
            },
            'cache_info': {
                'enabled': self.cache_enabled,
                'current_size': len(self.svg_cache),
                'max_size': self.max_cache_size,
                'hit_rate': cache_hit_rate
            }
        }
    
    def clear_cache(self):
        """
        清空缓存
        """
        with self._cache_lock:
            self.svg_cache.clear()
        self.logger.info("SVG缓存已清空")
    
    def reset_statistics(self):
        """
        重置统计信息
        """
        with self._stats_lock:
            self.conversion_stats = {
                'total_processed': 0,
                'svg_detected': 0,
                'svg_converted': 0,
                'conversion_failed': 0,
                'files_created': [],
                'processing_time': 0.0,
                'cache_hits': 0,
                'cache_misses': 0
            }
        
        self.performance_metrics = {
            'start_time': None,
            'end_time': None,
            'svg_processing_times': [],
            'memory_usage': []
        }
        
        self.logger.info("统计信息已重置")