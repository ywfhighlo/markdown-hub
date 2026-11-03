"""
PPTX → Markdown 表格单元格富文本提取的端到端样例测试。

覆盖此前评审发现的问题：
  - 空段落不再产生多余的 <br>
  - GFM 强调标记只包裹非空白文本（'前 **粗体** 后'，而不是 '** 粗体 **' 失效写法）
  - 等宽字体 run 优先输出代码 span（不再出现 `**x**` 嵌套在代码里）
  - 粗体+斜体 → ***x***
  - 单元格内的竖线转义为 \\|

运行：
    python -m pytest backend/tests/test_pptx_table_richtext.py -q
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from pptx import Presentation  # noqa: E402
from pptx.util import Inches  # noqa: E402


def _add_run(paragraph, text, bold=False, italic=False, font_name=None):
    r = paragraph.add_run()
    r.text = text
    if bold:
        r.font.bold = True
    if italic:
        r.font.italic = True
    if font_name:
        r.font.name = font_name
    return r


def _make_pptx(path: Path) -> None:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # blank layout
    tbl = slide.shapes.add_table(
        3, 2, Inches(0.5), Inches(0.5), Inches(8), Inches(2)
    ).table

    # (0,0) 表头
    tbl.cell(0, 0).text_frame.paragraphs[0].add_run().text = '功能'
    tbl.cell(0, 1).text_frame.paragraphs[0].add_run().text = '提取结果'

    # (1,0) 富文本混合 run：粗体 / 斜体 / 粗斜体 / 等宽 / 带空白边界的粗体
    c = tbl.cell(1, 0)
    p = c.text_frame.paragraphs[0]
    _add_run(p, '前 ')
    _add_run(p, '粗体', bold=True)
    _add_run(p, ' 中')
    _add_run(p, '斜体', italic=True)
    _add_run(p, '粗斜', bold=True, italic=True)
    _add_run(p, '码', font_name='Consolas')

    # (1,1) 第一段为空段落 + 第二段有内容 —— 不应产生开头的 <br>
    c2 = tbl.cell(1, 1)
    c2.text_frame.paragraphs[0]  # 留空（无 run）
    p2 = c2.text_frame.add_paragraph()
    _add_run(p2, 'second-line')

    # (2,0) 单元格内竖线（加粗），应转义
    c3 = tbl.cell(2, 0)
    _add_run(c3.text_frame.paragraphs[0], 'a|b', bold=True)

    prs.save(str(path))


def test_table_rich_text_extraction(tmp_path):
    pptx = tmp_path / 'rich.pptx'
    _make_pptx(pptx)

    from converters.office_to_md import OfficeToMdConverter
    out_dir = tmp_path / 'out'
    out_dir.mkdir()
    conv = OfficeToMdConverter(str(out_dir))
    result = conv.convert(str(pptx))
    assert result, '转换应成功'
    md = Path(result[0]).read_text(encoding='utf-8')

    # 各类行内格式
    assert '**粗体**' in md, md
    assert '*斜体*' in md, md
    assert '***粗斜***' in md, md
    # 相邻的强调标记用零宽空格断开（否则 GFM 解析歧义：*斜体***粗斜*** 会被拆错）
    zwsp = chr(0x200B)
    assert '*斜体*' + zwsp + '***粗斜***' in md, md
    assert '***粗斜***' + zwsp + '`码`' in md, md
    # 空白移到标记外：'前 **粗体** 中'（而不是 '**粗体 **中' 之类失效写法）
    assert '前 **粗体** 中' in md, md
    assert '** 粗体' not in md and '粗体 **' not in md, md
    # 代码 span 不叠加强调：等宽 run 只输出 `码`，不会出现 **码**
    assert '`码`' in md, md
    assert '**码**' not in md, md
    assert '`**码**`' not in md, md
    # 空段落不产生 <br>，单元格内容直接是 second-line
    assert 'second-line' in md, md
    assert '| <br>' not in md and '<br> |' not in md, md
    # 竖线转义
    assert r'**a\|b**' in md, md
