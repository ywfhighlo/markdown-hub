"""
_make_docx_reproducible() 的单元测试（不需要 pandoc / VS Code）。

运行：
    python -m pytest backend/tests/test_reproducible_docx.py -q
"""
import sys
import zipfile
from datetime import datetime, timezone
from pathlib import Path

# 让 backend/ 可导入
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from converters.md_to_office import _make_docx_reproducible, _CORE_TS_RE  # noqa: E402


def _make_docx(path: Path) -> None:
    from docx import Document
    Document().save(str(path))


def _read(path: Path):
    """返回 (core.xml 文本, [每个 ZIP 条目的 date_time])。"""
    with zipfile.ZipFile(path) as z:
        core = z.read('docProps/core.xml').decode('utf-8')
        entries = [i.date_time for i in z.infolist()]
    return core, entries


def test_core_properties_and_zip_entries_normalized(tmp_path):
    f = tmp_path / 'a.docx'
    _make_docx(f)
    core1, _ = _read(f)
    assert 'dcterms:created' in core1  # python-docx 写入了当前时间

    assert _make_docx_reproducible(str(f)) is True

    core2, ts2 = _read(f)
    # 1) 所有 ZIP 条目时间戳 → 1980-01-01（SOURCE_DATE_EPOCH 未设置时）
    assert all(dt == (1980, 1, 1, 0, 0, 0) for dt in ts2), ts2
    # 2) core.xml 的 created/modified 都被替换为同一固定值
    stamps = _CORE_TS_RE.findall(core2.encode('utf-8'))
    assert stamps, 'core.xml 里应有 created/modified 时间戳字段'
    assert all(s[1] == b'1980-01-01T00:00:00Z' for s in stamps), stamps
    # 3) 幂等：再跑一遍结果不变
    assert _make_docx_reproducible(str(f)) is True
    core3, ts3 = _read(f)
    assert core3 == core2
    assert ts3 == ts2


def test_source_date_epoch_env(tmp_path, monkeypatch):
    monkeypatch.setenv('SOURCE_DATE_EPOCH', '1600000000')
    f = tmp_path / 'b.docx'
    _make_docx(f)
    assert _make_docx_reproducible(str(f)) is True
    expected = datetime.fromtimestamp(
        1600000000, tz=timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    core, _ = _read(f)
    stamps = _CORE_TS_RE.findall(core.encode('utf-8'))
    assert stamps and all(s[1] == expected.encode() for s in stamps), stamps
