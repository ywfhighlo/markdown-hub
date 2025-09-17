#!/usr/bin/env python3
"""
依赖检查脚本 - 诊断Markdown Hub后端依赖问题
"""

import sys
import subprocess
import shutil
from typing import List, Tuple

def check_python_packages() -> List[Tuple[str, bool, str]]:
    """检查Python包依赖"""
    required_packages = [
        'pandoc_attributes',  # 注意：包名是pandoc-attributes，但导入名是pandocattributes
        'PIL',  # Pillow
        'docx',  # python-docx
        'docxtpl',
        'svglib', 
        'reportlab',
        'psutil',
        'lxml',
        'markdown'
    ]
    
    # 特殊处理的包映射
    package_mapping = {
        'pandoc_attributes': 'pandocattributes',  # 实际导入名
        'PIL': 'PIL',
        'docx': 'docx'
    }
    
    results = []
    for package in required_packages:
        import_name = package_mapping.get(package, package)
        try:
            __import__(import_name)
            results.append((package, True, "已安装"))
        except ImportError as e:
            results.append((package, False, f"未安装: {str(e)}"))
    
    return results

def check_system_tools() -> List[Tuple[str, bool, str]]:
    """检查系统工具依赖"""
    required_tools = [
        'pandoc',
        'python',
        'pip'
    ]
    
    results = []
    for tool in required_tools:
        if shutil.which(tool):
            try:
                if tool == 'pandoc':
                    result = subprocess.run([tool, '--version'], capture_output=True, text=True)
                    version = result.stdout.split('\n')[0] if result.returncode == 0 else "未知版本"
                    results.append((tool, True, f"已安装: {version}"))
                else:
                    results.append((tool, True, "已安装"))
            except Exception as e:
                results.append((tool, False, f"检查失败: {str(e)}"))
        else:
            results.append((tool, False, "未找到"))
    
    return results

def install_missing_packages(missing_packages: List[str]):
    """安装缺失的包"""
    if not missing_packages:
        print("✅ 所有Python包都已安装")
        return
    
    print(f"\n🔧 发现 {len(missing_packages)} 个缺失的包，正在安装...")
    
    # 包名映射（pip安装名 vs 导入名）
    pip_package_mapping = {
        'pandoc_attributes': 'pandoc-attributes',
        'PIL': 'Pillow',
        'docx': 'python-docx'
    }
    
    for package in missing_packages:
        pip_name = pip_package_mapping.get(package, package)
        try:
            print(f"正在安装 {pip_name}...")
            subprocess.run([sys.executable, '-m', 'pip', 'install', pip_name], 
                         check=True, capture_output=True)
            print(f"✅ {pip_name} 安装成功")
        except subprocess.CalledProcessError as e:
            print(f"❌ {pip_name} 安装失败: {e}")

def main():
    print("🔍 Markdown Hub 依赖检查工具\n")
    
    print("Python 版本:", sys.version)
    print("Python 路径:", sys.executable)
    print()
    
    # 检查Python包
    print("📦 检查Python包依赖:")
    package_results = check_python_packages()
    missing_packages = []
    
    for package, installed, status in package_results:
        status_icon = "✅" if installed else "❌"
        print(f"  {status_icon} {package}: {status}")
        if not installed:
            missing_packages.append(package)
    
    print()
    
    # 检查系统工具
    print("🛠️  检查系统工具:")
    tool_results = check_system_tools()
    missing_tools = []
    
    for tool, available, status in tool_results:
        status_icon = "✅" if available else "❌"
        print(f"  {status_icon} {tool}: {status}")
        if not available:
            missing_tools.append(tool)
    
    print()
    
    # 提供解决方案
    if missing_packages or missing_tools:
        print("🚨 发现问题:")
        
        if missing_packages:
            print(f"\n缺失的Python包: {', '.join(missing_packages)}")
            response = input("是否自动安装缺失的Python包? (y/n): ")
            if response.lower() in ['y', 'yes', '是']:
                install_missing_packages(missing_packages)
            else:
                print("\n手动安装命令:")
                pip_mapping = {
                    'pandoc_attributes': 'pandoc-attributes',
                    'PIL': 'Pillow', 
                    'docx': 'python-docx'
                }
                for pkg in missing_packages:
                    pip_name = pip_mapping.get(pkg, pkg)
                    print(f"  pip install {pip_name}")
        
        if missing_tools:
            print(f"\n缺失的系统工具: {', '.join(missing_tools)}")
            print("\n安装指南:")
            for tool in missing_tools:
                if tool == 'pandoc':
                    print("  pandoc: 访问 https://pandoc.org/installing.html")
                    print("    Windows: 下载安装包或使用 chocolatey: choco install pandoc")
                    print("    macOS: brew install pandoc")
                    print("    Linux: sudo apt-get install pandoc")
    else:
        print("🎉 所有依赖都已正确安装！")
        print("\n✨ 您可以正常使用Markdown Hub了")

if __name__ == '__main__':
    main()