#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SVG检测器模块
用于自动检测Markdown文件中的SVG代码块

作者：AI Assistant
日期：2025年1月27日
"""

import re
import logging
import time
from typing import List, Dict, Optional, Tuple
from dataclasses import dataclass


@dataclass
class SVGBlock:
    """
    SVG代码块信息
    """
    content: str  # SVG代码内容
    start_pos: int  # 开始位置
    end_pos: int  # 结束位置
    block_type: str  # 块类型：'code_block' 或 'inline'
    original_text: str  # 原始文本（包括代码块标记）


class SVGDetector:
    """
    SVG检测器类
    
    负责从Markdown内容中快速检测和定位SVG代码块。
    支持多种SVG格式的检测，包括代码块和内联SVG。
    
    主要功能：
    1. 快速检测Markdown中是否包含SVG内容
    2. 提取SVG代码块的详细信息
    3. 支持多种SVG格式（```svg、```xml、内联<svg>）
    4. 提供位置信息便于后续处理
    """
    
    # SVG检测正则表达式模式
    PATTERNS = {
        # SVG代码块模式
        'svg_code_block': r'```svg\s*\n(.*?)\n```',
        # XML代码块模式（可能包含SVG）
        'xml_code_block': r'```xml\s*\n(.*?)\n```',
        # 内联SVG模式
        'inline_svg': r'<svg[^>]*>.*?</svg>',
        # 快速检测模式（用于has_svg_content）
        'quick_detect': r'(?:```svg|```xml|<svg)'
    }
    
    def __init__(self, **kwargs):
        """
        初始化SVG检测器
        
        Args:
            **kwargs: 配置参数
                - case_sensitive (bool): 是否区分大小写，默认False
                - include_xml_blocks (bool): 是否包含XML代码块，默认True
                - validate_svg_content (bool): 是否验证SVG内容有效性，默认True
        """
        self.case_sensitive = kwargs.get('case_sensitive', False)
        self.include_xml_blocks = kwargs.get('include_xml_blocks', True)
        self.validate_svg_content = kwargs.get('validate_svg_content', True)
        
        # 编译正则表达式
        flags = re.DOTALL
        if not self.case_sensitive:
            flags |= re.IGNORECASE
            
        self.compiled_patterns = {}
        for name, pattern in self.PATTERNS.items():
            self.compiled_patterns[name] = re.compile(pattern, flags)
        
        # 日志记录器
        self.logger = logging.getLogger(self.__class__.__name__)
        
        # 性能统计
        self.performance_stats = {
            'total_detections': 0,
            'successful_detections': 0,
            'failed_detections': 0,
            'total_processing_time': 0.0,
            'average_processing_time': 0.0,
            'blocks_detected': 0
        }
        
        # 错误统计
        self.error_stats = {
            'regex_errors': 0,
            'validation_errors': 0,
            'content_errors': 0
        }
        
        self.logger.debug(f"SVGDetector初始化完成，配置: case_sensitive={self.case_sensitive}, "
                         f"include_xml_blocks={self.include_xml_blocks}, "
                         f"validate_svg_content={self.validate_svg_content}")
    
    def has_svg_content(self, content: str) -> bool:
        """
        快速检测Markdown内容中是否包含SVG
        
        这是一个高性能的检测方法，用于快速判断是否需要进行详细的SVG处理。
        
        Args:
            content (str): Markdown内容
            
        Returns:
            bool: 是否包含SVG内容
        """
        if not content:
            return False
        
        # 使用快速检测模式
        return bool(self.compiled_patterns['quick_detect'].search(content))
    
    def detect_svg_blocks(self, content: str) -> List[SVGBlock]:
        """
        检测并提取Markdown中的所有SVG代码块
        
        Args:
            content (str): Markdown内容
            
        Returns:
            List[SVGBlock]: SVG块信息列表
        """
        start_time = time.time()
        self.performance_stats['total_detections'] += 1
        
        if not content:
            self.performance_stats['failed_detections'] += 1
            self.error_stats['content_errors'] += 1
            return []
        
        svg_blocks = []
        
        try:
             # 首先检测代码块（SVG和XML），记录它们的位置以避免重复检测
             code_block_ranges = []
             
             # 检测SVG代码块
             svg_code_blocks = self._detect_code_blocks(content, 'svg')
             svg_blocks.extend(svg_code_blocks)
             code_block_ranges.extend([(block.start_pos, block.end_pos) for block in svg_code_blocks])
             
             # 检测XML代码块（如果启用）
             if self.include_xml_blocks:
                 xml_blocks = self._detect_code_blocks(content, 'xml')
                 # 过滤出包含SVG的XML块
                 svg_xml_blocks = [block for block in xml_blocks if self._is_svg_content(block.content)]
                 svg_blocks.extend(svg_xml_blocks)
                 code_block_ranges.extend([(block.start_pos, block.end_pos) for block in svg_xml_blocks])
             else:
                 # 即使不包含XML块，也要记录所有XML代码块的位置，避免内联检测误检
                 xml_blocks = self._detect_code_blocks(content, 'xml')
                 code_block_ranges.extend([(block.start_pos, block.end_pos) for block in xml_blocks])
             
             # 检测内联SVG（排除已在代码块中的SVG）
             inline_svg_blocks = self._detect_inline_svg(content, code_block_ranges)
             svg_blocks.extend(inline_svg_blocks)
             
             # 按位置排序
             svg_blocks.sort(key=lambda x: x.start_pos)
             
             # 更新统计信息
             self.performance_stats['successful_detections'] += 1
             self.performance_stats['blocks_detected'] += len(svg_blocks)
             
             self.logger.debug(f"检测到 {len(svg_blocks)} 个SVG块")
            
        except Exception as e:
            self.performance_stats['failed_detections'] += 1
            self.error_stats['regex_errors'] += 1
            self.logger.error(f"SVG检测过程中发生错误: {str(e)}")
            svg_blocks = []
        
        finally:
            # 记录处理时间
            end_time = time.time()
            processing_time = end_time - start_time
            self.performance_stats['total_processing_time'] += processing_time
            
            # 计算平均处理时间
            if self.performance_stats['total_detections'] > 0:
                self.performance_stats['average_processing_time'] = (
                    self.performance_stats['total_processing_time'] / 
                    self.performance_stats['total_detections']
                )
        
        return svg_blocks
    
    def _detect_code_blocks(self, content: str, block_type: str) -> List[SVGBlock]:
        """
        检测指定类型的代码块
        
        Args:
            content (str): Markdown内容
            block_type (str): 代码块类型（'svg' 或 'xml'）
            
        Returns:
            List[SVGBlock]: SVG块列表
        """
        pattern_key = f'{block_type}_code_block'
        pattern = self.compiled_patterns[pattern_key]
        
        blocks = []
        for match in pattern.finditer(content):
            svg_content = match.group(1).strip()
            
            # 验证SVG内容（如果启用）
            if self.validate_svg_content and not self._is_svg_content(svg_content):
                continue
            
            block = SVGBlock(
                content=svg_content,
                start_pos=match.start(),
                end_pos=match.end(),
                block_type='code_block',
                original_text=match.group(0)
            )
            blocks.append(block)
        
        return blocks
    
    def _detect_inline_svg(self, content: str, exclude_ranges: List[Tuple[int, int]] = None) -> List[SVGBlock]:
        """
        检测内联SVG标签
        
        Args:
            content (str): Markdown内容
            exclude_ranges (List[Tuple[int, int]]): 要排除的位置范围（代码块位置）
            
        Returns:
            List[SVGBlock]: SVG块列表
        """
        if exclude_ranges is None:
            exclude_ranges = []
            
        pattern = self.compiled_patterns['inline_svg']
        
        blocks = []
        for match in pattern.finditer(content):
            start_pos = match.start()
            end_pos = match.end()
            
            # 检查是否在排除范围内（代码块内）
            is_in_code_block = False
            for exclude_start, exclude_end in exclude_ranges:
                if start_pos >= exclude_start and end_pos <= exclude_end:
                    is_in_code_block = True
                    break
            
            if is_in_code_block:
                continue
                
            svg_content = match.group(0)
            
            block = SVGBlock(
                content=svg_content,
                start_pos=start_pos,
                end_pos=end_pos,
                block_type='inline',
                original_text=svg_content
            )
            blocks.append(block)
        
        return blocks
    
    def _is_svg_content(self, content: str) -> bool:
        """
        判断内容是否为有效的SVG
        
        Args:
            content (str): 内容字符串
            
        Returns:
            bool: 是否为SVG内容
        """
        if not content:
            return False
        
        content = content.strip()
        
        # 检查是否包含SVG标签
        svg_tag_pattern = re.compile(r'<svg[^>]*>', re.IGNORECASE | re.DOTALL)
        return bool(svg_tag_pattern.search(content))
    

    
    def get_svg_statistics(self, content: str) -> Dict[str, int]:
        """
        获取SVG统计信息
        
        Args:
            content (str): Markdown内容
            
        Returns:
            Dict[str, int]: 统计信息
        """
        if not content:
            return {
                'total_svg_blocks': 0,
                'code_blocks': 0,
                'inline_svg': 0,
                'xml_blocks_with_svg': 0
            }
        
        svg_blocks = self.detect_svg_blocks(content)
        
        stats = {
            'total_svg_blocks': len(svg_blocks),
            'code_blocks': 0,
            'inline_svg': 0,
            'xml_blocks_with_svg': 0
        }
        
        for block in svg_blocks:
            if block.block_type == 'code_block':
                if block.original_text.startswith('```xml'):
                    stats['xml_blocks_with_svg'] += 1
                else:
                    stats['code_blocks'] += 1
            elif block.block_type == 'inline':
                stats['inline_svg'] += 1
        
        return stats
    
    def extract_svg_content_only(self, content: str) -> List[str]:
        """
        仅提取SVG内容，不包含位置信息
        
        Args:
            content (str): Markdown内容
            
        Returns:
            List[str]: SVG内容列表
        """
        svg_blocks = self.detect_svg_blocks(content)
        return [block.content for block in svg_blocks]
    
    def get_performance_report(self) -> Dict[str, any]:
        """
        获取性能报告
        
        Returns:
            Dict[str, any]: 包含性能统计信息的字典
        """
        success_rate = 0
        if self.performance_stats['total_detections'] > 0:
            success_rate = (self.performance_stats['successful_detections'] / 
                          self.performance_stats['total_detections'] * 100)
        
        return {
            'performance': {
                'total_detections': self.performance_stats['total_detections'],
                'successful_detections': self.performance_stats['successful_detections'],
                'failed_detections': self.performance_stats['failed_detections'],
                'success_rate': f"{success_rate:.2f}%",
                'total_processing_time': f"{self.performance_stats['total_processing_time']:.3f}s",
                'average_processing_time': f"{self.performance_stats['average_processing_time']:.3f}s",
                'blocks_detected': self.performance_stats['blocks_detected']
            },
            'errors': {
                'regex_errors': self.error_stats['regex_errors'],
                'validation_errors': self.error_stats['validation_errors'],
                'content_errors': self.error_stats['content_errors']
            },
            'configuration': {
                'case_sensitive': self.case_sensitive,
                'include_xml_blocks': self.include_xml_blocks,
                'validate_svg_content': self.validate_svg_content
            }
        }
    
    def reset_statistics(self):
        """
        重置所有统计信息
        """
        self.performance_stats = {
            'total_detections': 0,
            'successful_detections': 0,
            'failed_detections': 0,
            'total_processing_time': 0.0,
            'average_processing_time': 0.0,
            'blocks_detected': 0
        }
        
        self.error_stats = {
            'regex_errors': 0,
            'validation_errors': 0,
            'content_errors': 0
        }
        
        self.logger.info("SVGDetector统计信息已重置")