"""
统一依赖探测模块
为每个功能提供独立的依赖检查入口，避免某个依赖缺失影响其他功能。
所有检测都是惰性的、按需的，单个 ImportError 不会污染整个模块。
"""

from __future__ import annotations
from typing import Dict, List, Tuple
import shutil
import logging
import importlib
import sys
import os
import subprocess

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────
# Python 库可用性（懒加载，按需探测）
# ─────────────────────────────────────────

_LIB_CACHE: Dict[str, bool] = {}
_LIB_ERROR_CACHE: Dict[str, str] = {}


def lib_available(lib_name: str) -> bool:
    """检查 Python 第三方库是否可导入。"""
    return _check_lib(lib_name)[0]


def lib_error(lib_name: str) -> str:
    """
    返回导入失败的原因字符串。如果库可用则返回空字符串。
    用于在错误消息中展示真正的失败原因（如 DLL 加载失败、版本不兼容等）。
    """
    return _check_lib(lib_name)[1]


def _check_lib(lib_name: str) -> Tuple[bool, str]:
    """
    检查 Python 第三方库是否可导入，并记录失败原因。

    Returns:
        (是否可导入, 失败原因或空字符串)
    """
    if lib_name in _LIB_CACHE:
        return _LIB_CACHE[lib_name], _LIB_ERROR_CACHE.get(lib_name, "")

    import_name = lib_name.replace('-', '_')

    # 第一次尝试：直接导入（用户自己装的版本优先）
    try:
        importlib.import_module(import_name)
        _LIB_CACHE[lib_name] = True
        _LIB_ERROR_CACHE[lib_name] = ""
        return True, ""
    except Exception as e:
        error_msg = f"{type(e).__name__}: {e}"

    # 第二次尝试：注入 vendor 路径后重试
    # 防止 cli.py 之外的入口忘记 init vendor，让 dep_check 自洽
    try:
        import vendor
        vendor.init_vendor_path()
        importlib.import_module(import_name)
        _LIB_CACHE[lib_name] = True
        _LIB_ERROR_CACHE[lib_name] = ""
        return True, ""
    except Exception as e:
        error_msg = f"{type(e).__name__}: {e}"
        logger.warning(f"依赖库加载失败: {lib_name} — {error_msg}")
        _LIB_CACHE[lib_name] = False
        _LIB_ERROR_CACHE[lib_name] = error_msg
    return False, error_msg


# ─────────────────────────────────────────
# 外部命令可用性
# ─────────────────────────────────────────

# cmd -> 实际可执行路径（PATH 命中时为 cmd 本身，常见路径命中时为绝对路径，都没找到时为 None）
_CMD_PATH_CACHE: Dict[str, str] = {}
# cmd -> 诊断信息（命令不可用时展示给用户）
_CMD_INFO_CACHE: Dict[str, str] = {}


def command_available(cmd: str) -> bool:
    """
    检查外部命令是否可用。
    PATH 命中 OR 在常见安装路径下找到 exe，都算可用。
    """
    return _check_command(cmd)[0]


def command_info(cmd: str) -> str:
    """
    返回命令不可用时的诊断信息。如果命令可用则返回空字符串。
    用于在错误消息中展示具体的诊断信息（如 PATH 问题、常见安装路径等）。
    """
    return _check_command(cmd)[1]


def resolve_command(cmd: str) -> str:
    """
    返回 cmd 实际可执行路径，供 subprocess 调用使用。
    - 在 PATH 中找到 → 返回 cmd 本身（让 subprocess 自己解析）
    - PATH 找不到但在常见安装路径找到 → 返回绝对路径（绕过 PATH 依赖）
    - 都找不到 → 返回空字符串

    调用方应在用到 cmd 的地方替换为 resolve_command(cmd)，
    这样用户用安装包装到 Program Files 等非 PATH 路径时也能直接用。
    """
    return _check_command(cmd)[2]


def _check_command(cmd: str) -> Tuple[bool, str, str]:
    """
    检查外部命令是否可用，并记录诊断信息。

    Returns:
        (是否可用, 诊断信息或空字符串, 实际可执行路径或空字符串)
    """
    if cmd in _CMD_PATH_CACHE:
        path = _CMD_PATH_CACHE[cmd]
        return bool(path), _CMD_INFO_CACHE.get(cmd, ""), path

    # 1. 先看 PATH
    found_in_path = shutil.which(cmd)
    if found_in_path:
        _CMD_PATH_CACHE[cmd] = cmd  # 用 cmd 本身即可，subprocess 会解析
        return True, "", cmd

    # 2. PATH 找不到，再扫常见安装路径
    probed = _probe_common_paths(cmd)
    if probed:
        # 在常见路径找到了 exe —— 视为可用，调用方用绝对路径
        _CMD_PATH_CACHE[cmd] = probed
        note = (f"在常见路径找到 {cmd}: {probed}，但不在 PATH 中。"
                f"已自动使用该路径，建议将其加入 PATH 以避免每次提示。")
        _CMD_INFO_CACHE[cmd] = note
        logger.info(f"外部命令 {cmd} 不在 PATH，使用常见路径: {probed}")
        return True, note, probed

    # 3. 都没找到
    _CMD_PATH_CACHE[cmd] = ""
    if sys.platform == "win32":
        path_note = f"当前 PATH 中未找到 '{cmd}'。当前 PATH: {os.environ.get('PATH', '(空)')[:300]}"
    else:
        path_note = f"当前 PATH 中未找到 '{cmd}' 命令"
    _CMD_INFO_CACHE[cmd] = path_note
    logger.warning(f"外部命令不可用: {cmd} — {path_note}")
    return False, path_note, ""


def _probe_common_paths(cmd: str) -> str:
    """
    在常见安装路径中探测命令。
    返回找到的第一个 exe 绝对路径；找不到返回空字符串。
    """
    candidates: list = []
    if sys.platform == "win32":
        if cmd == "pandoc":
            user_local = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Pandoc", "pandoc.exe")
            user_profile_local = os.path.join(os.environ.get("USERPROFILE", ""), "AppData", "Local", "Pandoc", "pandoc.exe")
            program_files = os.path.join(os.environ.get("ProgramFiles", ""), "Pandoc", "pandoc.exe")
            program_files_x86 = os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Pandoc", "pandoc.exe")
            candidates = [user_local, user_profile_local, program_files, program_files_x86]
        elif cmd == "soffice":
            program_files = os.path.join(os.environ.get("ProgramFiles", ""), "LibreOffice", "program", "soffice.exe")
            program_files_x86 = os.path.join(os.environ.get("ProgramFiles(x86)", ""), "LibreOffice", "program", "soffice.exe")
            candidates = [program_files, program_files_x86]
        elif cmd == "tesseract":
            program_files = os.path.join(os.environ.get("ProgramFiles", ""), "Tesseract-OCR", "tesseract.exe")
            program_files_x86 = os.path.join(os.environ.get("ProgramFiles(x86)", ""), "Tesseract-OCR", "tesseract.exe")
            candidates = [program_files, program_files_x86]
    elif sys.platform == "darwin":
        if cmd == "pandoc":
            candidates = ["/usr/local/bin/pandoc", "/opt/homebrew/bin/pandoc"]
        elif cmd == "soffice":
            candidates = ["/Applications/LibreOffice.app/Contents/MacOS/soffice"]
        elif cmd == "tesseract":
            candidates = ["/usr/local/bin/tesseract", "/opt/homebrew/bin/tesseract"]
    else:
        if cmd == "pandoc":
            candidates = ["/usr/bin/pandoc", "/usr/local/bin/pandoc"]
        elif cmd == "soffice":
            candidates = ["/usr/bin/soffice", "/usr/local/bin/soffice", "/opt/libreoffice/program/soffice"]
        elif cmd == "tesseract":
            candidates = ["/usr/bin/tesseract", "/usr/local/bin/tesseract"]

    for p in candidates:
        if p and os.path.isfile(p):
            return p
    return ""


# ─────────────────────────────────────────
# 功能维度依赖矩阵
# 每个功能独立检查依赖，缺失只会让该功能不可用，不会波及其他
# ─────────────────────────────────────────

# 格式 -> 解析所需 Python 库
_PARSER_DEPS: Dict[str, List[str]] = {
    "pdf":  ["PyMuPDF"],                # 主路径
    "pdf_ocr": ["pypdf", "pytesseract", "pdf2image"],   # OCR 回退
    "word": ["docx2txt"],
    "excel": ["pandas", "tabulate", "openpyxl"],
    "pptx":  ["python-pptx"],
    "html":  ["html2text"],
    "image": ["Pillow"],
}

# 输出格式 -> 生成所需 Python 库
_GENERATOR_DEPS: Dict[str, List[str]] = {
    "docx": ["python-docx", "docxtpl", "docxcompose", "docx2txt"],
    "pdf":  ["markdown", "pandoc-attributes"],   # 还需要系统级 pandoc
    "html": ["markdown"],
    "pptx": ["python-pptx", "Pillow"],
}

# 输出格式 -> 所需外部命令
_GENERATOR_CMDS: Dict[str, List[str]] = {
    "pdf": ["pandoc"],
}

# 命令 -> 安装提示
_CMD_INSTALL_HINTS: Dict[str, str] = {
    "pandoc": {
        "win32":   "下载安装包 https://pandoc.org/installing.html",
        "darwin":  "brew install pandoc",
        "linux":   "sudo apt install pandoc   # 或从 https://pandoc.org/installing.html 下载",
    },
    "tesseract": {
        "win32":   "下载 https://github.com/UB-Mannheim/tesseract/wiki",
        "darwin":  "brew install tesseract",
        "linux":   "sudo apt install tesseract-ocr",
    },
    "java": {
        "win32":   "下载 https://adoptium.net/",
        "darwin":  "brew install openjdk",
        "linux":   "sudo apt install openjdk-11-jdk",
    },
    "mmdc": {
        "win32":   "npm install -g @mermaid-js/mermaid-cli",
        "darwin":  "npm install -g @mermaid-js/mermaid-cli",
        "linux":   "npm install -g @mermaid-js/mermaid-cli",
    },
    "drawio": {
        "win32":   "下载 https://github.com/jgraph/drawio-desktop/releases",
        "darwin":  "brew install --cask drawio",
        "linux":   "下载 https://github.com/jgraph/drawio-desktop/releases",
    },
}


def install_hint_for(cmd: str) -> str:
    """获取外部命令的平台相关安装提示"""
    import sys
    plat = sys.platform
    if plat.startswith("win"):
        plat_key = "win32"
    elif plat == "darwin":
        plat_key = "darwin"
    else:
        plat_key = "linux"
    hints = _CMD_INSTALL_HINTS.get(cmd, {})
    return hints.get(plat_key, hints.get("linux", "请参考官方文档"))


# ─────────────────────────────────────────
# 对外 API：按功能检查
# ─────────────────────────────────────────

def parser_ready(file_kind: str) -> Tuple[bool, List[str]]:
    """
    检查解析某种文件所需的依赖是否就绪。

    Args:
        file_kind: pdf / word / excel / pptx / html / image

    Returns:
        (是否就绪, 缺失的库列表)
    """
    required = _PARSER_DEPS.get(file_kind, [])
    missing = [name for name in required if not lib_available(name)]
    return (len(missing) == 0, missing)


def generator_ready(fmt: str) -> Tuple[bool, List[str], List[str]]:
    """
    检查生成某种输出格式所需的依赖是否就绪。

    Args:
        fmt: docx / pdf / html / pptx

    Returns:
        (是否就绪, 缺失的Python库列表, 缺失的外部命令列表)
    """
    py_required = _GENERATOR_DEPS.get(fmt, [])
    py_missing = [name for name in py_required if not lib_available(name)]

    cmd_required = _GENERATOR_CMDS.get(fmt, [])
    cmd_missing = [name for name in cmd_required if not command_available(name)]

    return (len(py_missing) == 0 and len(cmd_missing) == 0, py_missing, cmd_missing)


def feature_snapshot() -> Dict[str, Dict]:
    """
    一次性返回所有功能的依赖状态，供前端展示。

    Returns:
        {
          "pdf_to_md":  {"available": bool, "missing_libs": [...], "missing_cmds": [...]},
          "word_to_md": {...},
          ...
        }
    """
    snapshot: Dict[str, Dict] = {}

    # 入方向（Office -> MD）
    for kind, label in [
        ("pdf",  "pdf_to_md"),
        ("word", "word_to_md"),
        ("excel", "excel_to_md"),
        ("pptx", "pptx_to_md"),
        ("html", "html_to_md"),
    ]:
        ok, missing = parser_ready(kind)
        snapshot[label] = {
            "available": ok,
            "missing_libs": missing,
            "missing_cmds": [],
        }

    # 出方向（MD -> Office）
    for fmt, label in [
        ("docx", "md_to_docx"),
        ("pdf",  "md_to_pdf"),
        ("html", "md_to_html"),
        ("pptx", "md_to_pptx"),
    ]:
        ok, py_miss, cmd_miss = generator_ready(fmt)
        snapshot[label] = {
            "available": ok,
            "missing_libs": py_miss,
            "missing_cmds": cmd_miss,
        }

    return snapshot


def reset_cache() -> None:
    """清空缓存，方便测试和热重载"""
    _LIB_CACHE.clear()
    _LIB_ERROR_CACHE.clear()
    _CMD_PATH_CACHE.clear()
    _CMD_INFO_CACHE.clear()


# ─────────────────────────────────────────
# 首次自动安装（用于带 C 扩展、不能塞进 vendor 的库）
# ─────────────────────────────────────────

# 这些库体积大或有平台特定二进制，不适合打包到 vendor/
# 但又是核心功能（如 PyMuPDF 之于 PDF 解析），所以提供首次自动下载能力
_INSTALLABLE_HEAVY_LIBS = {
    "PyMuPDF": {
        "pip_name": "PyMuPDF",
        "import_name": "fitz",
        "hint": "PDF 智能转换所需",
    },
}


def _pip_subprocess(args: list, timeout: int = 300):
    """统一的 pip 子进程调用，复用当前解释器"""
    return subprocess.run(
        [sys.executable, '-m', 'pip'] + args,
        capture_output=True, text=True,
        timeout=timeout,
    )


def try_install_lib(lib_name: str, timeout: int = 300) -> Tuple[bool, str]:
    """
    用当前 Python 解释器 pip install 某个库到用户 site-packages。
    仅用于 _INSTALLABLE_HEAVY_LIBS 中登记的库（白名单）。

    Args:
        lib_name: 库名（与 pip install 名一致，如 "PyMuPDF"）
        timeout: 子进程超时（秒）

    Returns:
        (是否成功, 失败原因或安装日志摘要)
    """
    spec = _INSTALLABLE_HEAVY_LIBS.get(lib_name)
    if not spec:
        return False, f"{lib_name} 不在允许自动安装的白名单内"

    pip_name = spec["pip_name"]
    try:
        result = _pip_subprocess(["install", "--user", pip_name], timeout=timeout)
    except subprocess.TimeoutExpired:
        return False, f"pip install {pip_name} 超时（>{timeout}s）"
    except Exception as e:
        return False, f"pip 子进程异常: {type(e).__name__}: {e}"

    if result.returncode != 0:
        return False, (result.stderr or result.stdout or "").strip()[-500:]

    # 安装后清缓存，让 _check_lib 能重新探测
    _LIB_CACHE.pop(lib_name, None)
    _LIB_ERROR_CACHE.pop(lib_name, None)

    # 验证能否真正 import
    ok, err = _check_lib(lib_name)
    if not ok:
        return False, f"安装完成但 import 失败: {err}"
    return True, result.stdout.strip()[-200:]


def ensure_pymupdf(timeout: int = 300) -> Tuple[bool, str]:
    """
    确保 PyMuPDF 可用：先看是否已导入，缺失则首次自动下载安装。

    Returns:
        (是否最终可用, 失败原因或成功消息)
    """
    if lib_available("PyMuPDF"):
        return True, "PyMuPDF 已就绪"

    logger.info("PyMuPDF 未安装，尝试首次自动下载...")
    ok, msg = try_install_lib("PyMuPDF", timeout=timeout)
    if ok:
        logger.info("PyMuPDF 自动安装成功")
        return True, "PyMuPDF 自动安装成功"
    logger.error(f"PyMuPDF 自动安装失败: {msg}")
    return False, msg


def install_hint_for_lib(lib_name: str) -> str:
    """获取某个 Python 库在当前平台的安装提示"""
    spec = _INSTALLABLE_HEAVY_LIBS.get(lib_name)
    if not spec:
        return f"pip install {lib_name}"
    return f"pip install {spec['pip_name']}  # {spec['hint']}"
