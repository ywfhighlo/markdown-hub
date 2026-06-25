#!/usr/bin/env python3
"""
下载纯 Python 依赖到 vendor/ 目录，供插件内置分发。
用法: python setup_vendor.py
"""
import os
import sys
import subprocess
import tempfile
import shutil
import zipfile
from pathlib import Path

VENDOR_DIR = Path(__file__).parent / 'vendor'

# 需要内置的纯 Python 库（pip install 名）
# 不含 C 扩展的库才适合内置
PACKAGES = [
    'markdown',
    'html2text',
    'docx2txt',
    'tabulate',
    'pypdf',
    'python-pptx',
    'python-docx',
    'docxtpl',
    'docxcompose',
]

def download_and_extract(package_name: str, vendor_dir: Path):
    """下载纯 Python wheel 并解压到 vendor 目录。

    只接受 wheel (.whl) —— 纯 Python wheel 是 zip 格式，可直接解压。
    如果 pip 只能拿到 sdist (.tar.gz)，说明该包可能不是纯 Python 或当前环境无 wheel，
    会跳过它（由调用者决定如何处理）。
    """
    print(f"  下载 {package_name}...")

    with tempfile.TemporaryDirectory() as tmpdir:
        # --only-binary=:all:  强制只要 wheel，避免拿到 sdist 后还要处理打包目录
        # --no-deps            只取该包本身，不递归依赖
        result = subprocess.run(
            [sys.executable, '-m', 'pip', 'download', package_name,
             '--only-binary=:all:',
             '--no-deps',
             '-d', tmpdir],
            capture_output=True, text=True
        )
        if result.returncode != 0:
            print(f"  警告: 下载 {package_name} 失败: {result.stderr.strip()}")
            return False

        wheels = [f for f in os.listdir(tmpdir) if f.endswith('.whl')]
        if not wheels:
            print(f"  警告: {package_name} 没有可用的纯 Python wheel（可能含 C 扩展），跳过")
            return False

        wheel_path = os.path.join(tmpdir, wheels[0])
        print(f"  解压 {wheels[0]}...")

        with zipfile.ZipFile(wheel_path, 'r') as zf:
            zf.extractall(str(vendor_dir))

    return True

def main():
    print(f"目标目录: {VENDOR_DIR}")
    VENDOR_DIR.mkdir(parents=True, exist_ok=True)
    
    # 确保 __init__.py 存在
    init_file = VENDOR_DIR / '__init__.py'
    if not init_file.exists():
        init_file.write_text('# vendor packages\n')
    
    success = []
    failed = []
    
    for pkg in PACKAGES:
        if download_and_extract(pkg, VENDOR_DIR):
            success.append(pkg)
        else:
            failed.append(pkg)
    
    print(f"\n=== 完成 ===")
    print(f"成功: {', '.join(success)}")
    if failed:
        print(f"失败: {', '.join(failed)}")
    
    # 统计 vendor 目录大小
    total_size = sum(f.stat().st_size for f in VENDOR_DIR.rglob('*') if f.is_file())
    print(f"vendor/ 总大小: {total_size / 1024:.0f} KB ({total_size / 1024 / 1024:.1f} MB)")

if __name__ == '__main__':
    main()
