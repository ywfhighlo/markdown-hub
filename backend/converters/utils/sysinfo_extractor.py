#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SysInfo Extractor

从Ark_sysinfo.xlsx提取系统信息并转换为JSON格式
"""

import json
import logging
import pandas as pd
from pathlib import Path
from typing import Dict, Any, Optional

logger = logging.getLogger(__name__)

class SysInfoExtractor:
    """从Excel文件提取系统信息并转换为JSON"""
    
    def __init__(self):
        self.logger = logger
    
    def extract_to_json(self, xlsx_path: str, output_path: str) -> str:
        """
        从Excel文件提取数据并保存为JSON
        
        Args:
            xlsx_path: Excel文件路径
            output_path: 输出JSON文件路径
            
        Returns:
            str: 生成的JSON文件路径
            
        Raises:
            FileNotFoundError: Excel文件不存在
            Exception: 提取过程中的任何错误
        """
        try:
            xlsx_path = Path(xlsx_path)
            if not xlsx_path.exists():
                raise FileNotFoundError(f"Excel文件不存在: {xlsx_path}")
            
            self.logger.info(f"开始提取 {xlsx_path} 的系统信息")
            
            # 读取所有sheet
            excel_data = pd.read_excel(xlsx_path, sheet_name=None)
            
            # 转换为JSON格式
            json_data = self._convert_to_json_structure(excel_data)
            
            # 保存JSON文件
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(json_data, f, ensure_ascii=False, indent=2)
            
            self.logger.info(f"系统信息已保存到: {output_path}")
            return str(output_path)
            
        except Exception as e:
            self.logger.error(f"提取系统信息失败: {str(e)}")
            raise
    
    def _convert_to_json_structure(self, excel_data: Dict[str, pd.DataFrame]) -> Dict[str, Any]:
        """
        将Excel数据转换为JSON结构
        
        Args:
            excel_data: Excel数据字典
            
        Returns:
            Dict: JSON结构的数据
        """
        json_data = {}
        
        for sheet_name, df in excel_data.items():
            self.logger.debug(f"处理sheet: {sheet_name}")
            
            if sheet_name == 'Baseinfo':
                json_data['baseinfo'] = self._process_baseinfo_sheet(df)
            elif sheet_name == 'AutosarModleList':
                json_data['autosar_module_list'] = self._process_autosar_module_list_sheet(df)
            else:
                # 其他sheet直接转换为记录列表
                json_data[sheet_name.lower()] = df.to_dict('records')
        
        return json_data
    
    def _process_baseinfo_sheet(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        处理Baseinfo sheet
        
        Args:
            df: Baseinfo DataFrame
            
        Returns:
            Dict: 处理后的数据结构
        """
        # 清理数据
        df = df.dropna(how='all')  # 删除全空行
        
        # 转换为字典，以Module Name为key
        baseinfo_dict = {}
        for _, row in df.iterrows():
            module_name = row.get('Module Name')
            if pd.notna(module_name):
                baseinfo_dict[str(module_name)] = {
                    'module_id': row.get('Module ID'),
                    'module_sw_version': row.get('Module SW Version'),
                    'autosar_release': row.get('AUTOSAR Release'),
                    'vendor_id': row.get('Vendor ID'),
                    'regdef_filename': row.get('Regdef Filename')
                }
        
        return baseinfo_dict
    
    def _process_autosar_module_list_sheet(self, df: pd.DataFrame) -> Dict[str, Any]:
        """
        处理AutosarModleList sheet
        
        Args:
            df: AutosarModleList DataFrame
            
        Returns:
            Dict: 处理后的数据结构
        """
        # 清理数据
        df = df.dropna(how='all')  # 删除全空行
        
        # 转换为记录列表
        records = []
        for _, row in df.iterrows():
            # 过滤掉空值
            record = {}
            for col, value in row.items():
                if pd.notna(value) and str(value).strip():
                    record[col] = str(value).strip()
            
            if record:  # 只添加非空记录
                records.append(record)
        
        return {
            'modules': records,
            'total_count': len(records)
        }
    
    def get_module_info(self, json_path: str, module_name: str) -> Optional[Dict[str, Any]]:
        """
        从JSON文件获取特定模块信息
        
        Args:
            json_path: JSON文件路径
            module_name: 模块名称
            
        Returns:
            Dict: 模块信息，如果不存在返回None
        """
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            baseinfo = data.get('baseinfo', {})
            return baseinfo.get(module_name)
            
        except Exception as e:
            self.logger.error(f"读取模块信息失败: {str(e)}")
            return None

def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description='提取Ark_sysinfo.xlsx到JSON格式')
    parser.add_argument('input_file', help='输入的Excel文件路径')
    parser.add_argument('-o', '--output', default='sysinfo.json', help='输出JSON文件路径')
    parser.add_argument('-v', '--verbose', action='store_true', help='详细输出')
    
    args = parser.parse_args()
    
    # 配置日志
    level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=level, format='%(asctime)s - %(levelname)s - %(message)s')
    
    # 执行提取
    extractor = SysInfoExtractor()
    try:
        output_path = extractor.extract_to_json(args.input_file, args.output)
        print(f"成功提取系统信息到: {output_path}")
    except Exception as e:
        print(f"提取失败: {str(e)}")
        return 1
    
    return 0

if __name__ == '__main__':
    exit(main())