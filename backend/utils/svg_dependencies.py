#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
SVG转换依赖检查工具

该模块提供SVG转换所需依赖的检查和安装指导功能。
"""

import subprocess
import sys
import shutil
from typing import Dict, List, Tuple, Optional
import logging

logger = logging.getLogger(__name__)

class SVGDependencyChecker:
    """SVG转换依赖检查器"""
    
    def __init__(self):
        self.dependencies = {
            'cairosvg': {
                'type': 'python',
                'install_cmd': 'pip install cairosvg',
                'description': 'Python库，基于Cairo图形库，质量高但依赖较多'
            },
            'inkscape': {
                'type': 'system',
                'install_cmd': 'https://inkscape.org/release/',
                'description': '专业矢量图形编辑器，转换质量最高'
            },
            'rsvg-convert': {
                'type': 'system', 
                'install_cmd': 'librsvg工具包',
                'description': 'GNOME项目的SVG渲染库，Linux下常用'
            },
            'svglib': {
                'type': 'python',
                'install_cmd': 'pip install svglib reportlab',
                'description': 'Python库，轻量级但功能有限'
            }
        }
    
    def check_python_package(self, package_name: str) -> bool:
        """检查Python包是否已安装"""
        try:
            __import__(package_name)
            return True
        except ImportError:
            return False
    
    def check_system_command(self, command: str) -> bool:
        """检查系统命令是否可用"""
        return shutil.which(command) is not None
    
    def check_dependency(self, dep_name: str) -> Tuple[bool, str]:
        """检查单个依赖
        
        Returns:
            Tuple[bool, str]: (是否可用, 状态信息)
        """
        if dep_name not in self.dependencies:
            return False, f"未知依赖: {dep_name}"
        
        dep_info = self.dependencies[dep_name]
        
        if dep_info['type'] == 'python':
            if self.check_python_package(dep_name):
                return True, f"{dep_name} 已安装"
            else:
                return False, f"{dep_name} 未安装"
        
        elif dep_info['type'] == 'system':
            if dep_name == 'inkscape':
                available = self.check_system_command('inkscape')
            elif dep_name == 'rsvg-convert':
                available = self.check_system_command('rsvg-convert')
            else:
                available = self.check_system_command(dep_name)
            
            if available:
                return True, f"{dep_name} 已安装"
            else:
                return False, f"{dep_name} 未安装"
        
        return False, f"未知依赖类型: {dep_info['type']}"
    
    def check_all_dependencies(self) -> Dict[str, Tuple[bool, str]]:
        """检查所有SVG转换依赖
        
        Returns:
            Dict[str, Tuple[bool, str]]: 依赖名称 -> (是否可用, 状态信息)
        """
        results = {}
        for dep_name in self.dependencies:
            results[dep_name] = self.check_dependency(dep_name)
        return results
    
    def get_available_methods(self) -> List[str]:
        """获取可用的转换方法列表"""
        available = []
        for dep_name, (is_available, _) in self.check_all_dependencies().items():
            if is_available:
                available.append(dep_name)
        return available
    
    def get_recommended_method(self) -> Optional[str]:
        """获取推荐的转换方法
        
        优先级: inkscape > cairosvg > rsvg-convert > svglib
        """
        priority_order = ['inkscape', 'cairosvg', 'rsvg-convert', 'svglib']
        available_methods = self.get_available_methods()
        
        for method in priority_order:
            if method in available_methods:
                return method
        
        return None
    
    def generate_installation_guide(self) -> str:
        """生成安装指导"""
        results = self.check_all_dependencies()
        guide = ["SVG转换依赖安装指导:\n"]
        
        for dep_name, (is_available, status) in results.items():
            dep_info = self.dependencies[dep_name]
            
            if is_available:
                guide.append(f"✓ {dep_name}: {status}")
            else:
                guide.append(f"✗ {dep_name}: {status}")
                guide.append(f"  安装方法: {dep_info['install_cmd']}")
                guide.append(f"  说明: {dep_info['description']}")
            guide.append("")
        
        # 添加推荐
        recommended = self.get_recommended_method()
        if recommended:
            guide.append(f"推荐使用: {recommended}")
        else:
            guide.append("警告: 没有可用的SVG转换方法！")
            guide.append("建议至少安装以下之一:")
            guide.append("1. pip install cairosvg (推荐，Python库)")
            guide.append("2. 安装 Inkscape (https://inkscape.org/)")
        
        return "\n".join(guide)
    
    def install_python_dependencies(self) -> bool:
        """尝试自动安装Python依赖"""
        python_deps = ['cairosvg', 'svglib']
        success = True
        
        for dep in python_deps:
            if not self.check_python_package(dep):
                try:
                    logger.info(f"正在安装 {dep}...")
                    if dep == 'svglib':
                        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'svglib', 'reportlab'])
                    else:
                        subprocess.check_call([sys.executable, '-m', 'pip', 'install', dep])
                    logger.info(f"{dep} 安装成功")
                except subprocess.CalledProcessError as e:
                    logger.error(f"{dep} 安装失败: {e}")
                    success = False
        
        return success

def main():
    """命令行入口"""
    logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
    
    checker = SVGDependencyChecker()
    
    print("检查SVG转换依赖...\n")
    
    # 检查所有依赖
    results = checker.check_all_dependencies()
    
    # 显示结果
    for dep_name, (is_available, status) in results.items():
        status_icon = "✓" if is_available else "✗"
        print(f"{status_icon} {dep_name}: {status}")
    
    print("\n" + "="*50)
    
    # 显示可用方法
    available = checker.get_available_methods()
    if available:
        print(f"可用的转换方法: {', '.join(available)}")
        recommended = checker.get_recommended_method()
        if recommended:
            print(f"推荐使用: {recommended}")
    else:
        print("警告: 没有可用的SVG转换方法！")
    
    print("\n" + "="*50)
    print(checker.generate_installation_guide())
    
    # 询问是否自动安装Python依赖
    if not available:
        response = input("\n是否尝试自动安装Python依赖 (cairosvg, svglib)? [y/N]: ")
        if response.lower() in ['y', 'yes']:
            if checker.install_python_dependencies():
                print("\n依赖安装完成，请重新检查:")
                # 重新检查
                available = checker.get_available_methods()
                if available:
                    print(f"现在可用的方法: {', '.join(available)}")
                else:
                    print("仍然没有可用的方法，请手动安装系统依赖")

if __name__ == '__main__':
    main()