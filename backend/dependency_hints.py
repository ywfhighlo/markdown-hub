"""Platform-specific install hints for optional native dependencies.

Some converters rely on native tools (Poppler, Tesseract, Graphviz) that
we either can't bundle (license / size) or only auto-download on certain
platforms. When such a tool is missing, bare exception text tells the user
nothing actionable. These helpers produce a human-readable string
explaining what's needed and how to install it on the current OS, plus a
classifier so callers can detect "dependency missing" errors and route
them to the right hint.
"""
import platform
from typing import Optional


def poppler_install_hint(context: str = "") -> str:
    """Return platform-specific instructions for obtaining Poppler.

    Poppler is required by ``pdf2image`` for PDF → image conversion (used
    by PDF OCR and batch PDF → PNG). ``context`` optionally describes
    where the failure originated (e.g. ``"auto-download failed"``) and is
    folded into the header so the message explains both *what* went wrong
    and *how to fix it*.
    """
    p = platform.system()
    header = "Poppler is required for PDF→image conversion but could not be located"
    if context:
        header += f" ({context})"
    header += "."

    if p == "Windows":
        body = (
            "Options to fix:\n"
            "  1. Let Markdown Hub auto-download it — make sure the machine has\n"
            "     network access (configure a proxy if needed), then re-run.\n"
            "     If you set MARKDOWN_HUB_NO_AUTO_DOWNLOAD=1, unset it to re-enable.\n"
            "  2. Install manually: download from\n"
            "     https://github.com/oschwartz10612/poppler-windows/releases,\n"
            "     unzip, and add the inner 'Library\\bin' folder to your PATH.\n"
            "  3. Point to it explicitly via the --poppler-path CLI option\n"
            "     or the POPPLER_PATH environment variable."
        )
    elif p == "Darwin":
        body = (
            "Install via Homebrew:\n"
            "    brew install poppler\n"
            "Or point to it explicitly via the --poppler-path CLI option or\n"
            "the POPPLER_PATH environment variable."
        )
    else:  # Linux and other Unix-likes
        body = (
            "Install via your package manager:\n"
            "    Debian/Ubuntu:  sudo apt install poppler-utils\n"
            "    Fedora/RHEL:    sudo dnf install poppler-utils\n"
            "    Arch/Manjaro:   sudo pacman -S poppler\n"
            "Or point to it explicitly via the --poppler-path CLI option or\n"
            "the POPPLER_PATH environment variable."
        )
    return header + "\n" + body


def is_poppler_missing_error(exc: BaseException) -> bool:
    """True if ``exc`` indicates pdf2image couldn't find a Poppler binary.

    Uses an exception-type check first (pdf2image exposes dedicated
    classes), then falls back to a string match on the message so we still
    recognise the failure on older pdf2image versions or wrapped errors.
    """
    try:
        from pdf2image.exceptions import (
            PopplerNotInstalledError,
            PDFInfoNotInstalledError,
        )
        if isinstance(exc, (PopplerNotInstalledError, PDFInfoNotInstalledError)):
            return True
    except Exception:
        pass
    msg = str(exc).lower()
    return any(k in msg for k in ("poppler", "pdfinfo", "pdftoppm"))
