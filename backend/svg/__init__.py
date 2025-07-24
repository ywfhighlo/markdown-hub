#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SVG处理模块

包含SVG检测、转换和处理相关的功能模块。

模块说明：
- adaptive_svg_converter: 自适应SVG转换器
- svg_detector: SVG检测器
"""

from .adaptive_svg_converter import AdaptiveSVGConverter
from .svg_detector import SVGDetector, SVGBlock

__all__ = [
    'AdaptiveSVGConverter',
    'SVGDetector', 
    'SVGBlock'
]

__version__ = '1.0.0'
__author__ = 'AI Assistant'