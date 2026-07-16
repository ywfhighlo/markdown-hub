"""On-demand resource downloader for large bundled dependencies.

Some third-party tools bundled with this extension (PlantUML jar, Poppler
binaries, Batik jar) are too large to ship inside the VSIX (they'd bloat
the install from ~5MB to ~70MB). Instead, we download them on first use
to a per-user cache directory and reuse them on subsequent calls.

Cache layout:
    ~/.markdown-hub/cache/
        plantuml/
            plantuml.jar
            .version         # Version string for invalidation
        poppler/
            poppler-24.02.0/  # Platform-specific subtree
            .platform        # e.g. "win-x64"
        batik/
            batik-all.jar
            .version

Downloads are best-effort: a failed download returns None and the caller
falls back to a system-installed tool if available.
"""

import hashlib
import logging
import os
import platform
import shutil
import sys
import tarfile
import time
import zipfile
from pathlib import Path
from typing import Optional
from urllib.error import URLError
from urllib.request import Request, urlopen

logger = logging.getLogger(__name__)

# Root of the per-user cache: ~/.markdown-hub/cache
CACHE_ROOT = Path.home() / ".markdown-hub" / "cache"


# ─────────────────────────────────────────
# Resource definitions
# ─────────────────────────────────────────

class ResourceSpec:
    """Description of a downloadable resource."""

    def __init__(self, name: str, url: str, version: str, archive: str = "zip"):
        self.name = name           # subdir under CACHE_ROOT
        self.url = url
        self.version = version
        self.archive = archive     # "zip" | "tar.gz" | "raw" (single file)


# PlantUML: standalone jar (single file)
PLANTUML_SPEC = ResourceSpec(
    name="plantuml",
    url="https://github.com/plantuml/plantuml/releases/download/v1.2024.7/plantuml-1.2024.7.jar",
    version="1.2024.7",
    archive="raw",
)

# Poppler: platform-specific binary release
def _poppler_url() -> str:
    """Return the Poppler download URL for the current OS + arch."""
    p = platform.system()
    if p == "Windows":
        # Poppler for Windows: standard MinGW build by oschwartz10612
        # (using a stable, well-known mirror; in production this could be
        # parameterised or fall back to the user's own mirror)
        return "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.02.0-0/Release-24.02.0-0.zip"
    elif p == "Darwin":
        # macOS: brew formula is the most reliable path; we document but
        # don't auto-download a .pkg (which would require sudo).
        return ""  # Empty: caller will fall back to system install
    else:
        # Linux: distro packages are reliable; we don't bundle.
        return ""


POPPLER_SPEC = ResourceSpec(
    name="poppler",
    url=_poppler_url(),
    version="24.02.0",
    archive="zip",
)

# Batik: standalone jar (single file)
BATIK_SPEC = ResourceSpec(
    name="batik",
    url="https://archive.apache.org/dist/xmlgraphics/batik/batik-1.17/binaries/batik-bin-1.17.zip",
    version="1.17",
    archive="zip",
)


# ─────────────────────────────────────────
# Public API
# ─────────────────────────────────────────

def cache_path(spec: ResourceSpec) -> Path:
    """Return the path to the extracted resource in the user cache."""
    return CACHE_ROOT / spec.name


def version_marker(spec: ResourceSpec) -> Path:
    """Path to the .version file that signals the cache is up-to-date."""
    return cache_path(spec) / ".version"


def is_cached(spec: ResourceSpec) -> bool:
    """True if a usable copy of the resource is already on disk."""
    marker = version_marker(spec)
    if not marker.exists():
        return False
    if marker.read_text(encoding="utf-8").strip() != spec.version:
        return False
    # Quick existence check: any file under the cache dir
    return any(cache_path(spec).iterdir())


def is_auto_download_enabled() -> bool:
    """Allow user to opt out via env var (offline / corporate policy)."""
    return os.environ.get("MARKDOWN_HUB_NO_AUTO_DOWNLOAD", "0") != "1"


def _download(url: str, dest: Path, max_retries: int = 3) -> bool:
    """Download ``url`` to ``dest`` with progress logging and retries.

    Returns True on success, False after all retries are exhausted. A 48 MB
    Poppler download no longer looks like a hung process: progress is logged
    every ~10% (or every 5 MB when the server omits Content-Length).
    Transient network errors trigger up to ``max_retries`` attempts with
    exponential backoff; any partial file is removed between attempts so a
    retry starts clean.
    """
    if not url:
        return False

    for attempt in range(1, max_retries + 1):
        try:
            _download_once(url, dest)
            return True
        except (URLError, OSError) as e:
            logger.warning(f"Download attempt {attempt}/{max_retries} failed: {e}")
            dest.unlink(missing_ok=True)  # discard partial file before retry
            if attempt < max_retries:
                backoff = 2 ** (attempt - 1)  # 1s, 2s, 4s
                logger.info(f"Retrying in {backoff}s...")
                time.sleep(backoff)
    logger.warning(f"Download failed after {max_retries} attempts: {url}")
    return False


def _download_once(url: str, dest: Path) -> None:
    """Perform a single download attempt, logging progress.

    Raises ``URLError`` / ``OSError`` on failure so the caller can retry.
    Validates a short read (fewer bytes than Content-Length promised) by
    raising ``OSError`` — this catches truncated downloads that would
    otherwise produce a corrupt archive.
    """
    logger.info(f"Downloading {url} -> {dest}")
    req = Request(url, headers={"User-Agent": "markdown-hub-extension"})
    chunk_size = 64 * 1024  # 64 KB
    progress_step = 5 * 1024 * 1024  # log at least every 5 MB

    with urlopen(req, timeout=60) as resp, open(dest, "wb") as f:
        total_hdr = resp.getheader("Content-Length")
        total = int(total_hdr) if total_hdr else None
        downloaded = 0
        last_log_pct = -1
        next_log_bytes = progress_step

        while True:
            chunk = resp.read(chunk_size)
            if not chunk:
                break
            f.write(chunk)
            downloaded += len(chunk)

            if total:
                pct = downloaded * 100 // total
                if pct >= last_log_pct + 10:
                    last_log_pct = pct
                    logger.info(
                        f"Download progress: {downloaded / 1e6:.1f} / "
                        f"{total / 1e6:.1f} MB ({pct}%)"
                    )
            elif downloaded >= next_log_bytes:
                logger.info(f"Download progress: {downloaded / 1e6:.1f} MB (size unknown)")
                next_log_bytes = downloaded + progress_step

    if total is not None and downloaded != total:
        raise OSError(f"Short read: got {downloaded} of {total} bytes")
    logger.info(f"Download complete: {downloaded / 1e6:.1f} MB")


def _extract(archive: Path, dest: Path, fmt: str) -> None:
    """Extract a zip/tar.gz archive into dest."""
    if fmt == "raw":
        # Single file: archive path == dest path
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(archive), str(dest))
        return
    if fmt == "zip":
        with zipfile.ZipFile(archive, "r") as zf:
            zf.extractall(dest)
    elif fmt in ("tar.gz", "tgz"):
        with tarfile.open(archive, "r:gz") as tf:
            tf.extractall(dest)
    archive.unlink(missing_ok=True)


def ensure_resource(spec: ResourceSpec) -> Optional[Path]:
    """Return the path to the resource, downloading if needed.

    Returns None if the resource cannot be obtained automatically. The caller
    should then fall back to system-installed tools if available.
    """
    # 1. Already cached and at the right version
    if is_cached(spec):
        return cache_path(spec)

    # 2. Opt-out: don't auto-download
    if not is_auto_download_enabled():
        return None

    # 3. Try to download
    if not spec.url:
        logger.info(f"{spec.name}: no auto-download URL for this platform; "
                    f"please install via system package manager")
        return None

    cache_path(spec).mkdir(parents=True, exist_ok=True)
    tmp = CACHE_ROOT / f"{spec.name}.download"
    if not _download(spec.url, tmp):
        return None

    try:
        # For "raw" archive format, the downloaded file IS the resource.
        # For PlantUML, the URL serves a .jar but the download may arrive
        # with a different filename (e.g. "plantuml.download" from our tmp
        # file). We rename to a canonical "<spec.name>.<ext>" so callers
        # can find the file by predictable name.
        if spec.archive == "raw":
            # Preserve the source extension if it has one
            src_ext = Path(spec.url.split("/")[-1].split("?")[0]).suffix
            canonical_name = f"{spec.name}{src_ext}" if src_ext else f"{spec.name}.bin"
            target = cache_path(spec) / canonical_name
            shutil.move(str(tmp), str(target))
        else:
            _extract(tmp, cache_path(spec), spec.archive)
        version_marker(spec).write_text(spec.version, encoding="utf-8")
    except (zipfile.BadZipFile, tarfile.TarError, OSError) as e:
        logger.warning(f"Failed to extract {spec.name}: {e}")
        # Clean up partial state
        shutil.rmtree(cache_path(spec), ignore_errors=True)
        return None
    finally:
        tmp.unlink(missing_ok=True)

    logger.info(f"Resource {spec.name} {spec.version} ready at {cache_path(spec)}")
    return cache_path(spec)


def get_poppler_bin_path() -> Optional[Path]:
    """Return the path to the Poppler ``bin`` directory (containing pdftoppm).

    Triggers an auto-download on Windows when the resource is not yet cached.
    Returns ``None`` on non-Windows platforms (where ``_poppler_url()`` is
    empty and we rely on the system package manager), or if the download
    fails — the caller should then let ``pdf2image`` fall back to PATH.

    The oschwartz10612 Windows release extracts as
    ``poppler-<ver>/Library/bin/pdftoppm.exe``; we ``rglob`` for the
    executable rather than hard-coding the path so future layout changes
    don't break discovery.
    """
    cached = ensure_resource(POPPLER_SPEC)
    if not cached:
        return None
    exe_name = "pdftoppm.exe" if platform.system() == "Windows" else "pdftoppm"
    for candidate in cached.rglob(exe_name):
        if candidate.is_file():
            return candidate.parent
    logger.warning(f"Poppler cache populated but {exe_name} not found under {cached}")
    return None


def clear_cache() -> int:
    """Remove the entire cache directory. Returns number of bytes freed."""
    if not CACHE_ROOT.exists():
        return 0
    total = sum(p.stat().st_size for p in CACHE_ROOT.rglob("*") if p.is_file())
    shutil.rmtree(CACHE_ROOT, ignore_errors=True)
    return total
