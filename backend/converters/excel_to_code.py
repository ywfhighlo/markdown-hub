#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Excel to Code Converter

将Excel寄存器定义文件转换为AUTOSAR标准的C代码文件
"""

import os
import sys
import logging
from typing import List, Dict, Any, Optional
from pathlib import Path

# 添加utils目录到Python路径
utils_dir = os.path.join(os.path.dirname(__file__), 'utils')
if utils_dir not in sys.path:
    sys.path.insert(0, utils_dir)

from .base_converter import BaseConverter
from .utils.autocoder_core import (
    AutocoderConfig, 
    init_autocoder, 
    run_autocoder, 
    reset_autocoder
)
from .utils.sysinfo_extractor import SysInfoExtractor
from .utils.cli_parser import parse_cli_arguments


class ExcelToCodeConverter(BaseConverter):
    """
    Excel到代码转换器
    
    将Excel寄存器定义文件转换为AUTOSAR标准的C代码文件
    """
    
    def __init__(self, output_dir: str = None, **config):
        """初始化Excel转代码转换器
        
        Args:
            output_dir: 输出目录
            **config: 其他配置参数
        """
        super().__init__(output_dir)
        
        # 获取Excel专用输出目录，如果没有则使用默认输出目录
        excel_output_dir = config.get('excel_output_dir') or output_dir or './output'
        
        # 创建AutocoderConfig实例
        self.config = AutocoderConfig(
            debug_level=config.get('debug_level', 'info'),
            language=config.get('language', 'english'),
            reg_short_description=config.get('reg_short_description', True),
            mask_style=config.get('mask_style', 'nxp'),
            input_file='',  # 将在convert方法中设置
            output_dir=excel_output_dir,
            sysinfo_json=config.get('sysinfo_json', '')
        )
        
        # 设置日志
        self._setup_logging()
        
    def _setup_logging(self):
        """
        设置日志配置
        """
        numeric_level = getattr(logging, self.config.debug_level.upper(), None)
        if not isinstance(numeric_level, int):
            raise ValueError(f'Invalid log level: {self.config.debug_level}')
        
        self.logger.setLevel(numeric_level)
        
        # 如果没有handler，添加一个
        if not self.logger.handlers:
            handler = logging.StreamHandler()
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            self.logger.addHandler(handler)
    
    def convert(self, input_path: str) -> List[str]:
        """
        执行Excel到代码的转换
        
        Args:
            input_path: 输入Excel文件路径
            
        Returns:
            List[str]: 生成的输出文件路径列表
            
        Raises:
            Exception: 转换过程中的任何错误
        """
        try:
            # 验证输入文件
            if not self._is_valid_input(input_path, ['.xlsx']):
                raise ValueError(f"Invalid input file: {input_path}")
            
            self.logger.info(f"Starting Excel to Code conversion for: {input_path}")
            
            # 更新配置对象的输入文件路径
            self.config.input_file = os.path.abspath(input_path)
            
            # 确保输出目录存在
            os.makedirs(self.config.output_dir, exist_ok=True)
            
            # 初始化autocoder
            init_autocoder(self.config)
            
            # 运行转换
            generated_files = run_autocoder()
            
            self.logger.info(f"Conversion completed. Generated {len(generated_files or [])} files.")
            
            return generated_files or []
            
        except Exception as e:
            self.logger.error(f"Conversion failed: {str(e)}")
            raise
        finally:
            # 重置状态
            reset_autocoder()
    
    def extract_sysinfo(self, excel_path: str, output_json_path: str) -> str:
        """
        从Excel文件提取系统信息到JSON文件
        
        Args:
            excel_path: Excel文件路径
            output_json_path: 输出JSON文件路径
            
        Returns:
            str: 生成的JSON文件路径
        """
        try:
            self.logger.info(f"Extracting system info from: {excel_path}")
            
            extractor = SysInfoExtractor()
            json_path = extractor.extract_to_json(excel_path, output_json_path)
            
            self.logger.info(f"System info extracted to: {json_path}")
            return json_path
            
        except Exception as e:
            self.logger.error(f"SysInfo extraction failed: {str(e)}")
            raise
    
    def convert_with_sysinfo_extraction(self, input_path: str, sysinfo_excel_path: Optional[str] = None) -> List[str]:
        """
        执行完整的转换流程，包括系统信息提取
        
        Args:
            input_path: 输入Excel文件路径
            sysinfo_excel_path: 系统信息Excel文件路径（可选）
            
        Returns:
            List[str]: 生成的输出文件路径列表
        """
        try:
            # 如果提供了系统信息Excel文件，先提取系统信息
            if sysinfo_excel_path and os.path.exists(sysinfo_excel_path):
                sysinfo_json_path = os.path.join(self.config.output_dir, 'sysinfo.json')
                self.extract_sysinfo(sysinfo_excel_path, sysinfo_json_path)
                self.config.sysinfo_json = sysinfo_json_path
            
            # 执行转换
            return self.convert(input_path)
            
        except Exception as e:
            self.logger.error(f"Complete conversion failed: {str(e)}")
            raise
    
    @classmethod
    def from_cli_args(cls, argv: Optional[List[str]] = None) -> 'ExcelToCodeConverter':
        """
        从命令行参数创建转换器实例
        
        Args:
            argv: 命令行参数列表，None表示使用sys.argv
            
        Returns:
            ExcelToCodeConverter: 转换器实例
        """
        config_args = parse_cli_arguments(argv)
        
        # 将配置参数打包成字典
        config_dict = {
            'debug_level': config_args.debug_level,
            'language': config_args.language,
            'mask_style': config_args.mask_style,
            'reg_short_description': config_args.reg_short_description,
            'sysinfo_json': config_args.sysinfo_json
        }
        
        return cls(
            output_dir=config_args.output_dir,
            **config_dict
        )
    
    def get_supported_extensions(self) -> List[str]:
        """返回支持的文件扩展名"""
        return ['.xlsx', '.xls']
    
    def get_description(self) -> str:
        """返回转换器描述"""
        return "将Excel寄存器描述文件转换为C/C++代码文件"


def convert_excel_to_code(input_file: str, output_dir: str, **kwargs) -> List[str]:
    """
    便捷函数：将Excel文件转换为代码
    
    Args:
        input_file: 输入Excel文件路径
        output_dir: 输出目录
        **kwargs: 其他配置参数
        
    Returns:
        List[str]: 生成的文件路径列表
    """
    converter = ExcelToCodeConverter(output_dir, **kwargs)
    return converter.convert(input_file)


def main():
    """
    CLI入口点
    """
    try:
        # 从命令行参数创建转换器
        converter = ExcelToCodeConverter.from_cli_args()
        
        # 获取输入文件路径（第一个位置参数）
        config = parse_cli_arguments()
        input_file = config.input_file
        
        # 执行转换
        generated_files = converter.convert(input_file)
        
        print(f"转换成功！生成了 {len(generated_files)} 个文件:")
        for file_path in generated_files:
            print(f"  - {file_path}")
            
    except KeyboardInterrupt:
        print("\n操作被用户取消。", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"错误: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()