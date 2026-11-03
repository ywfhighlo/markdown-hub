# -*- coding: utf-8 -*-
"""Multi-file Markdown book assemble -> single manuscript for pandoc."""
from __future__ import annotations

import json
import logging
import os
import re
from pathlib import Path
from typing import Any, Dict, List, Tuple

logger = logging.getLogger(__name__)

_BOOK_SUFFIXES = {".book.yml", ".book.yaml", ".book.json", ".chapters", ".book.txt"}


def looks_like_book_manifest(path: str) -> bool:
    p = Path(path)
    if not p.is_file():
        return False
    name = p.name.lower()
    for suf in _BOOK_SUFFIXES:
        if name.endswith(suf):
            return True
    # Generic YAML/JSON with chapters key
    if p.suffix.lower() in {".yml", ".yaml", ".json"}:
        try:
            meta = _load_manifest(p)
            return bool(meta.get("chapters"))
        except Exception:
            return False
    return False


def _load_manifest(path: Path) -> Dict[str, Any]:
    text = path.read_text(encoding="utf-8")
    suf = path.suffix.lower()
    name = path.name.lower()

    if name.endswith(".book.json") or suf == ".json":
        data = json.loads(text)
        if isinstance(data, list):
            return {"chapters": data}
        return data if isinstance(data, dict) else {"chapters": []}

    if name.endswith((".chapters", ".book.txt")) or (
        suf == ".txt" and "chapter" in name
    ):
        chapters = []
        title = None
        for raw in text.splitlines():
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            if line.lower().startswith("title:"):
                title = line.split(":", 1)[1].strip()
                continue
            chapters.append(line)
        meta: Dict[str, Any] = {"chapters": chapters}
        if title:
            meta["title"] = title
        return meta

    # YAML (minimal parser for chapters list + title — avoids PyYAML hard dep)
    return _parse_simple_book_yaml(text)


def _parse_simple_book_yaml(text: str) -> Dict[str, Any]:
    """Tiny YAML subset: title: ... and chapters: with - item lines."""
    meta: Dict[str, Any] = {"chapters": []}
    in_chapters = False
    for raw in text.splitlines():
        if not raw.strip() or raw.strip().startswith("#"):
            continue
        if re.match(r"^title\s*:", raw, re.I):
            meta["title"] = raw.split(":", 1)[1].strip().strip("'\"")
            in_chapters = False
            continue
        if re.match(r"^chapters\s*:", raw, re.I):
            rest = raw.split(":", 1)[1].strip()
            in_chapters = True
            if rest.startswith("[") and rest.endswith("]"):
                inner = rest[1:-1].strip()
                if inner:
                    meta["chapters"] = [c.strip().strip("'\"") for c in inner.split(",")]
                in_chapters = False
            continue
        if in_chapters:
            m = re.match(r"^\s*-\s+(.+)$", raw)
            if m:
                meta["chapters"].append(m.group(1).strip().strip("'\""))
            elif re.match(r"^\S", raw):
                in_chapters = False
    return meta


def _resolve_chapter(base: Path, entry: str) -> Path:
    p = Path(entry)
    if not p.is_absolute():
        p = (base / p).resolve()
    if not p.is_file():
        raise FileNotFoundError(f"Book chapter not found: {entry} (resolved: {p})")
    return p


def assemble_book_markdown(
    manifest_path: str,
    work_dir: str | None = None,
) -> Tuple[str, Dict[str, Any]]:
    """
    Read a chapter list / book manifest and write one combined Markdown file.

    Returns (assembled_md_path, meta).
    """
    manifest = Path(manifest_path).resolve()
    meta = _load_manifest(manifest)
    chapters = meta.get("chapters") or []
    if not chapters:
        raise ValueError(f"Book manifest has no chapters: {manifest_path}")

    base = manifest.parent
    resolved: List[Path] = [_resolve_chapter(base, c) for c in chapters]

    parts: List[str] = []
    title = meta.get("title")
    if title:
        # pandoc title metadata so TOC / EPUB / DOCX pick it up
        parts.append(f"% {title}\n")

    for idx, ch in enumerate(resolved):
        body = ch.read_text(encoding="utf-8").strip()
        if idx > 0:
            # Soft page break hint for pandoc/docx (works when raw_tex / openxml allowed;
            # otherwise still separates chapters clearly)
            parts.append("\n\n")
        parts.append(body)
        parts.append("\n")

    out_dir = Path(work_dir) if work_dir else base
    out_dir.mkdir(parents=True, exist_ok=True)
    stem = re.sub(r"[^\w\-]+", "_", (title or manifest.stem))[:80] or "book"
    out_path = out_dir / f"{stem}_assembled_{os.getpid()}.md"
    out_path.write_text("".join(parts), encoding="utf-8")
    meta = dict(meta)
    meta["chapters"] = [str(p) for p in resolved]
    meta["assembled"] = str(out_path)
    logger.info("Assembled %s chapters into %s", len(resolved), out_path)
    return str(out_path), meta
