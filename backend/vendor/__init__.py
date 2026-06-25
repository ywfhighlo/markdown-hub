"""
vendor 目录 - 内置的纯 Python 第三方库

这些库随插件一起分发，用户无需单独安装。
仅在用户系统没有安装对应库时作为 fallback 使用。
"""

import os
import sys
import logging

logger = logging.getLogger(__name__)

_VENDOR_DIR = os.path.dirname(os.path.abspath(__file__))

# 内置库的包名 -> 目录名映射
# key: pip install 名, value: vendor/ 下的实际目录名
# 只列出 setup_vendor.py 实际会打包的纯 Python 库。
# 带 C 扩展的库 (Pillow / pandas / openpyxl / PyMuPDF 等) 不在这里，
# 它们由用户系统安装或由 dep_check 在首次使用时自动 pip install。
_BUILTIN_PACKAGES = {
    'markdown': 'markdown',
    'html2text': 'html2text',
    'docx2txt': 'docx2txt',
    'tabulate': 'tabulate',
    'pypdf': 'pypdf',
    'python-pptx': 'pptx',
    'python-docx': 'docx',
    'docxtpl': 'docxtpl',
    'docxcompose': 'docxcompose',
}

_INITIALIZED = False


def init_vendor_path():
    """
    将 vendor 目录插入 sys.path，使内置库可被 import。
    插入位置靠前，但仍在用户 site-packages 之后，
    以确保用户自己安装的版本优先。
    """
    global _INITIALIZED
    if _INITIALIZED:
        return
    _INITIALIZED = True

    if _VENDOR_DIR not in sys.path:
        # 插入到 sys.path 的前面（但 index 0 通常是脚本自身目录）
        sys.path.insert(0, _VENDOR_DIR)
        logger.debug(f"vendor 路径已注入: {_VENDOR_DIR}")


def is_builtin(lib_name: str) -> bool:
    """检查某个库是否在 vendor 目录中内置"""
    dir_name = _BUILTIN_PACKAGES.get(lib_name, lib_name.replace('-', '_'))
    return os.path.isdir(os.path.join(_VENDOR_DIR, dir_name))


def list_builtins() -> list:
    """返回所有已内置的库名称"""
    return [name for name, dir_name in _BUILTIN_PACKAGES.items()
            if os.path.isdir(os.path.join(_VENDOR_DIR, dir_name))]
