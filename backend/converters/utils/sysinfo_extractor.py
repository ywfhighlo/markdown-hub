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
    
    def extract_to_json(self, xlsx_path: str, output_path: str, delete_excel: bool = True) -> str:
        """
        从Excel文件提取数据并保存为JSON
        
        Args:
            xlsx_path: Excel文件路径
            output_path: 输出JSON文件路径
            delete_excel: 是否删除原Excel文件
            
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
            
            # 删除原Excel文件
            if delete_excel:
                self._cleanup_excel_file(xlsx_path)
            
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
        baseinfo_data = None
        
        # 先处理Baseinfo以获取模块信息
        for sheet_name, df in excel_data.items():
            if sheet_name == 'Baseinfo':
                baseinfo_data = self._process_baseinfo_sheet(df)
                json_data['baseinfo'] = baseinfo_data
                break
        
        # 处理其他sheet
        for sheet_name, df in excel_data.items():
            self.logger.debug(f"处理sheet: {sheet_name}")
            
            if sheet_name == 'Baseinfo':
                continue  # 已经处理过了
            elif sheet_name == 'AutosarModleList':
                json_data['autosar_module_list'] = self._process_autosar_module_list_sheet(df, baseinfo_data)
            else:
                # 其他sheet直接转换为记录列表
                json_data[sheet_name.lower()] = df.to_dict('records')
        
        return json_data
    
    def _process_baseinfo_sheet(self, df: pd.DataFrame, autosar_module_data: Dict[str, Any] = None) -> Dict[str, Any]:
        """
        处理Baseinfo sheet
        
        Args:
            df: Baseinfo DataFrame
            autosar_module_data: AutosarModleList数据，用于查找module_id
            
        Returns:
            Dict: 处理后的数据结构
        """
        # 清理数据
        df = df.dropna(how='all')  # 删除全空行
        
        # 创建module_name到module_id的映射
        module_id_map = {}
        if autosar_module_data and 'modules' in autosar_module_data:
            for module in autosar_module_data['modules']:
                module_name = module.get('module_name')
                module_id = module.get('module_id')
                if module_name and module_id is not None:
                    module_id_map[module_name] = module_id
        
        # 转换为字典，以Module Name为key
        baseinfo_dict = {}
        for _, row in df.iterrows():
            module_name = row.get('Module Name')
            if pd.notna(module_name):
                module_name_str = str(module_name)
                # 从autosar_module_list中查找对应的module_id
                module_id = module_id_map.get(module_name_str)
                
                baseinfo_dict[module_name_str] = {
                    'module_sw_version': row.get('Module SW Version'),
                    'autosar_release': row.get('AUTOSAR Release'),
                    'vendor_id': row.get('Vendor ID'),
                    'module_id': module_id
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
        
        # 获取正确的列名（根据实际Excel文件结构）
        module_name_detail_col = 'Module name detail'
        module_name_col = 'Module Name'
        module_id_col = 'Module ID'
        autosar_layer_col = 'AUTOSAR SW Layer'
        spec_doc_col = 'Specification document'
        
        # 转换为记录列表
        records = []
        for _, row in df.iterrows():
            module_name_detail = row.get(module_name_detail_col)
            if pd.notna(module_name_detail):
                module_name_detail_str = str(module_name_detail).strip()
                module_name = str(row.get(module_name_col, '')).strip() if pd.notna(row.get(module_name_col)) else ''
                module_id = row.get(module_id_col)
                autosar_layer = str(row.get(autosar_layer_col, '')).strip() if pd.notna(row.get(autosar_layer_col)) else ''
                spec_doc = str(row.get(spec_doc_col, '')).strip() if pd.notna(row.get(spec_doc_col)) else ''
                
                # 转换Module ID为整数
                module_id_int = int(module_id) if pd.notna(module_id) else None
                
                record = {
                    'module_name_detail': module_name_detail_str,
                    'module_name': module_name,
                    'module_id': module_id_int,
                    'specification_document': spec_doc,
                    'autosar_sw_layer': autosar_layer
                }
                
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
    
    def _is_standard_module(self, module_name: str) -> bool:
        """
        判断是否为标准AUTOSAR模块
        
        Args:
            module_name: 模块名称
            
        Returns:
            bool: 是否为标准模块
        """
        # 标准AUTOSAR模块名列表
        standard_modules = {
            'ADC', 'CAN', 'DIO', 'GPT', 'ICU', 'MCU', 'PORT', 'PWM', 'SPI', 'WDG',
            'CANIF', 'COMM', 'DCM', 'DEM', 'DET', 'ECUM', 'FIM', 'NVM', 'OS', 'RTE',
            'CANTP', 'PDUR', 'COM', 'IPDUM', 'CANNM', 'NM', 'CANSM', 'BSWM', 'WDGM'
        }
        
        module_upper = module_name.upper()
        
        # 如果完全匹配标准模块名，则为标准模块
        if module_upper in standard_modules:
            return True
        
        # 如果模块名与标准模块相似但不完全相同，则认为是非标准模块
        for standard in standard_modules:
            similarity = self._calculate_similarity(module_upper, standard)
            if similarity > 0.7 and module_upper != standard:
                self.logger.debug(f"模块 {module_name} 与标准模块 {standard} 相似度 {similarity:.2f}，标记为非标准模块")
                return False
        
        return True
    
    def _calculate_similarity(self, str1: str, str2: str) -> float:
        """
        计算两个字符串的相似度
        
        Args:
            str1: 字符串1
            str2: 字符串2
            
        Returns:
            float: 相似度 (0-1)
        """
        # 使用Jaccard相似度
        set1 = set(str1.lower())
        set2 = set(str2.lower())
        
        intersection = len(set1.intersection(set2))
        union = len(set1.union(set2))
        
        return intersection / union if union > 0 else 0
    
    def _cleanup_excel_file(self, xlsx_path: Path) -> None:
        """
        删除原始Excel文件
        
        Args:
            xlsx_path: Excel文件路径
        """
        try:
            if xlsx_path.exists():
                self.logger.info(f"删除原始Excel文件: {xlsx_path}")
                xlsx_path.unlink()
                self.logger.info("Excel文件删除成功")
            else:
                self.logger.warning("Excel文件不存在，无需删除")
                
        except Exception as e:
            self.logger.error(f"删除Excel文件失败: {e}")
            raise

def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description='提取Ark_sysinfo.xlsx到JSON格式')
    parser.add_argument('input_file', help='输入的Excel文件路径')
    parser.add_argument('-o', '--output', default='sysinfo.json', help='输出JSON文件路径')
    parser.add_argument('-v', '--verbose', action='store_true', help='详细输出')
    parser.add_argument('--delete-excel', action='store_true', help='提取完成后删除原始Excel文件')
    
    args = parser.parse_args()
    
    # 配置日志
    level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=level, format='%(asctime)s - %(levelname)s - %(message)s')
    
    # 执行提取
    extractor = SysInfoExtractor()
    try:
        output_path = extractor.extract_to_json(args.input_file, args.output, delete_excel=args.delete_excel)
        print(f"成功提取系统信息到: {output_path}")
        if args.delete_excel:
            print(f"原始Excel文件已删除: {args.input_file}")
    except Exception as e:
        print(f"提取失败: {str(e)}")
        return 1
    
    return 0

if __name__ == '__main__':
    exit(main())