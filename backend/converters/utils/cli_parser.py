#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CLI Parameter Interface for Excel to Code Converter

提供命令行参数解析功能，将原始config.json参数映射到CLI参数
"""

import argparse
import os
import sys
from pathlib import Path
from typing import Optional

from .autocoder_core import AutocoderConfig


def create_cli_parser() -> argparse.ArgumentParser:
    """
    创建命令行参数解析器
    
    Returns:
        argparse.ArgumentParser: 配置好的参数解析器
    """
    parser = argparse.ArgumentParser(
        description='Excel to Code Converter - Convert Excel register definitions to AUTOSAR code',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s input.xlsx -o ./output
  %(prog)s input.xlsx --language chinese --debug-level debug
  %(prog)s input.xlsx --mask-style nxp5777m --sysinfo-json sysinfo.json

Supported mask styles:
  - nxp5777m (default)
  - custom

Supported languages:
  - english (default)
  - chinese

Supported debug levels:
  - debug
  - info (default)
  - warning
  - error
        """
    )
    
    # 必需参数
    parser.add_argument(
        'input_file',
        help='Input Excel file path (e.g., ArkLin_regdef.xlsx)'
    )
    
    # 可选参数
    parser.add_argument(
        '-o', '--output-dir',
        default='./build-header-files',
        help='Output directory for generated code files (default: ./build-header-files)'
    )
    
    parser.add_argument(
        '--debug-level',
        choices=['debug', 'info', 'warning', 'error'],
        default='info',
        help='Set logging level (default: info)'
    )
    
    parser.add_argument(
        '--language',
        choices=['english', 'chinese'],
        default='english',
        help='Output language for generated code comments (default: english)'
    )
    
    parser.add_argument(
        '--mask-style',
        default='nxp5777m',
        help='Mask style for register definitions (default: nxp5777m)'
    )
    
    parser.add_argument(
        '--reg-short-description',
        action='store_true',
        default=True,
        help='Use short descriptions for registers (default: True)'
    )
    
    parser.add_argument(
        '--no-reg-short-description',
        dest='reg_short_description',
        action='store_false',
        help='Use full descriptions for registers'
    )
    
    parser.add_argument(
        '--sysinfo-json',
        help='Path to system information JSON file (optional)'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version='Excel to Code Converter 1.2.5'
    )
    
    return parser


def validate_arguments(args: argparse.Namespace) -> None:
    """
    验证命令行参数
    
    Args:
        args: 解析后的命令行参数
        
    Raises:
        SystemExit: 参数验证失败时退出
    """
    # 验证输入文件
    if not os.path.exists(args.input_file):
        print(f"Error: Input file '{args.input_file}' does not exist.", file=sys.stderr)
        sys.exit(1)
    
    if not args.input_file.lower().endswith('.xlsx'):
        print(f"Error: Input file '{args.input_file}' must be an Excel file (.xlsx).", file=sys.stderr)
        sys.exit(1)
    
    # 验证系统信息JSON文件（如果提供）
    if args.sysinfo_json and not os.path.exists(args.sysinfo_json):
        print(f"Error: SysInfo JSON file '{args.sysinfo_json}' does not exist.", file=sys.stderr)
        sys.exit(1)
    
    # 确保输出目录存在
    try:
        os.makedirs(args.output_dir, exist_ok=True)
    except OSError as e:
        print(f"Error: Cannot create output directory '{args.output_dir}': {e}", file=sys.stderr)
        sys.exit(1)


def args_to_config(args: argparse.Namespace) -> AutocoderConfig:
    """
    将命令行参数转换为AutocoderConfig对象
    
    Args:
        args: 解析后的命令行参数
        
    Returns:
        AutocoderConfig: 配置对象
    """
    return AutocoderConfig(
        debug_level=args.debug_level,
        language=args.language,
        reg_short_description=args.reg_short_description,
        mask_style=args.mask_style,
        input_file=os.path.abspath(args.input_file),
        output_dir=os.path.abspath(args.output_dir),
        sysinfo_json=os.path.abspath(args.sysinfo_json) if args.sysinfo_json else ''
    )


def parse_cli_arguments(argv: Optional[list] = None) -> AutocoderConfig:
    """
    解析命令行参数并返回配置对象
    
    Args:
        argv: 命令行参数列表，None表示使用sys.argv
        
    Returns:
        AutocoderConfig: 配置对象
        
    Raises:
        SystemExit: 参数解析或验证失败时退出
    """
    parser = create_cli_parser()
    args = parser.parse_args(argv)
    
    # 验证参数
    validate_arguments(args)
    
    # 转换为配置对象
    config = args_to_config(args)
    
    return config


def main():
    """
    CLI入口点，用于测试
    """
    try:
        config = parse_cli_arguments()
        print(f"Configuration: {config}")
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()