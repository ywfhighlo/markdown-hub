# -*- coding: utf-8 -*-
"""Phase C: pipe-table merge markers <-> pandoc grid_tables <-> DOCX merges.

MD -> DOCX path
---------------
Pandoc maps **grid_tables** (not raw HTML tables) to real Word ``gridSpan`` /
``vMerge``. Pipe tables cannot express merges, so we accept an explicit
marker syntax inside GFM pipe cells and rewrite those tables to grid_tables
before pandoc runs::

    | [[colspan:2]]AB |   | C |
    | --- | --- | --- |
    | [[rowspan:2]]A | B | C |
    |  | D | E |

Rules
- ``[[colspan:N]]`` / ``[[rowspan:N]]`` may appear (in either order) at the
  start of a cell; the remainder is the visible text.
- The next ``N-1`` cells on the same row for a colspan MUST be empty
  placeholders and are consumed.
- For rowspan, the covered cells in following rows MUST be empty
  placeholders and are omitted from the grid.
- Tables without any markers are left untouched (stay GFM pipe).
- Native pandoc grid_tables already in the source are left untouched.

DOCX -> MD path
---------------
When a Word table has any ``gridSpan`` / ``vMerge``, emit a pandoc
grid_table so a later MD->DOCX round-trip restores the merges. Plain
tables stay as GFM pipe tables.
"""
from __future__ import annotations

import re
from typing import List, Optional, Sequence, Tuple

# Cell marker: optional colspan / rowspan tags at start of cell text
_MARKER_RE = re.compile(
    r"^(?:\[\[colspan:(\d+)\]\]|\[\[rowspan:(\d+)\]\])+"
    r"(.*)$",
    re.DOTALL,
)
_TAG_RE = re.compile(r"\[\[(colspan|rowspan):(\d+)\]\]")


def parse_cell_markers(raw: str) -> Tuple[int, int, str]:
    """Return (colspan, rowspan, text). Defaults are 1, 1."""
    text = raw.strip()
    colspan = rowspan = 1
    if not text.startswith("[["):
        return 1, 1, text
    # Pull tags only from the prefix
    pos = 0
    while True:
        m = _TAG_RE.match(text, pos)
        if not m:
            break
        kind, n = m.group(1), int(m.group(2))
        if n < 1:
            n = 1
        if kind == "colspan":
            colspan = n
        else:
            rowspan = n
        pos = m.end()
    rest = text[pos:].strip()
    return colspan, rowspan, rest


def table_has_merge_markers(table_text: str) -> bool:
    return "[[colspan:" in table_text or "[[rowspan:" in table_text


def _split_pipe_rows(table_text: str) -> Optional[Tuple[List[str], List[str], List[List[str]]]]:
    """Parse a GFM pipe table into header, separator cells, data rows.

    Returns None if the block is not a well-formed pipe table.
    """
    # Local import to avoid circular deps at module load
    from .md_to_office import _split_table_row

    lines = [ln for ln in table_text.strip("\n").split("\n") if ln.strip()]
    if len(lines) < 2:
        return None
    header = _split_table_row(lines[0])
    if not header:
        return None
    sep = _split_table_row(lines[1])
    if len(sep) != len(header):
        return None
    for sc in sep:
        if not re.fullmatch(r":?-+:?", sc.strip()):
            return None
    data: List[List[str]] = []
    for ln in lines[2:]:
        cells = _split_table_row(ln)
        if len(cells) != len(header):
            return None
        data.append(cells)
    return header, sep, data


def _pad(s: str, width: int) -> str:
    if len(s) >= width:
        return s
    return s + (" " * (width - len(s)))


def build_grid_table(
    rows: Sequence[Sequence[Tuple[int, int, str]]],
    col_count: int,
) -> str:
    """Build a pandoc grid_table from logical rows of (colspan, rowspan, text).

    ``rows`` is a list of row-lists. Each row lists only the cells that
    *start* in that row (continuations of rowspans from above are omitted).
    """
    # Occupancy grid: None = free, (r0,c0) = covered by cell starting at r0,c0
    # First expand to absolute grid of display strings + spans
    # We need max row count including rowspan extents
    # Build sparse map (r,c) -> (colspan, rowspan, text) for cell origins
    origins = {}
    max_r = 0
    for r, row in enumerate(rows):
        c = 0
        for colspan, rowspan, text in row:
            # advance c over occupied slots
            while c < col_count and (r, c) in origins or _covered(origins, r, c):
                # find next free
                c += 1
            while c < col_count and _covered(origins, r, c):
                c += 1
            if c >= col_count:
                break
            origins[(r, c)] = (colspan, rowspan, text)
            max_r = max(max_r, r + rowspan)
            c += colspan
        max_r = max(max_r, r + 1)

    n_rows = max(max_r, len(rows))
    # Compute column widths from cell text (distribute multi-col spans later)
    widths = [3] * col_count  # min width for " x "
    for (r, c), (cs, rs, text) in origins.items():
        # text may be multiline
        lines = text.split("\n") if text else [""]
        max_line = max((len(x) for x in lines), default=0)
        # distribute across colspan
        per = max(3, (max_line + cs - 1) // cs)
        for k in range(cs):
            if c + k < col_count:
                widths[c + k] = max(widths[c + k], per)

    def hline(ch: str = "-") -> str:
        parts = ["+"]
        for w in widths:
            parts.append(ch * (w + 2))
            parts.append("+")
        return "".join(parts)

    def cell_hline(r_after: int) -> str:
        """Horizontal rule after row r_after, respecting vertical merges."""
        parts = ["+"]
        c = 0
        while c < col_count:
            # If a rowspan from above covers (r_after+1, c) continuing through,
            # use spaces instead of dashes under that column group... 
            # Pandoc rowspan syntax: under a rowspan cell, the hline uses spaces
            # in those columns and only has '+' at edges where a border exists.
            covering = None
            for (or_, oc), (cs, rs, _t) in origins.items():
                if oc <= c < oc + cs and or_ <= r_after < or_ + rs - 1:
                    # Still inside rowspan (not yet at last covered row boundary after last)
                    covering = (oc, cs)
                    break
            if covering is not None:
                oc, cs = covering
                # emit spaces for each col in this span until end of span from c
                while c < oc + cs and c < col_count:
                    parts.append(" " * (widths[c] + 2))
                    c += 1
                    if c < oc + cs and c < col_count:
                        parts.append(" ")  # no + inside rowspan
                    else:
                        parts.append("+")
            else:
                parts.append("-" * (widths[c] + 2))
                parts.append("+")
                c += 1
        # Fix: pandoc expects '+' between columns on normal hlines.
        # Rebuild more carefully.
        return _hline_for_row_boundary(origins, widths, col_count, r_after)

    out: List[str] = []
    out.append(hline("-"))
    for r in range(n_rows):
        # Build one or more physical lines for this row (multi-line cells)
        cell_lines_map = {}
        max_lines = 1
        c = 0
        while c < col_count:
            if _covered_by_other(origins, r, c):
                c += 1
                continue
            if (r, c) in origins:
                cs, rs, text = origins[(r, c)]
                lines = (text or "").split("\n") or [""]
                cell_lines_map[(c, cs)] = lines
                max_lines = max(max_lines, len(lines))
                c += cs
            else:
                # empty 1x1
                cell_lines_map[(c, 1)] = [""]
                c += 1

        for li in range(max_lines):
            parts = ["|"]
            c = 0
            while c < col_count:
                if _covered_by_other(origins, r, c):
                    # inside a rowspan from above — pandoc wants a single '|' 
                    # only at the edges; interior is spaces without extra pipes
                    # Actually in pandoc vertical merge rows, the left cell is
                    # omitted and the row starts with the next cell. The left
                    # border of the rowspan is a '|' from the continuing span.
                    # Simpler approach: print spaces for covered cols without
                    # adding a pipe until we hit a real cell.
                    # Find width of continuous covered run
                    start = c
                    while c < col_count and _covered_by_other(origins, r, c):
                        c += 1
                    # spaces including the '+' gaps — use width sum
                    # Pandoc format for rowspan continuation row:
                    # |   | C |
                    # where first cell is blank area under rowspan — actually
                    # the native example was:
                    # +---+---+
                    # | A | B |
                    # |   +---+
                    # |   | C |
                    # +---+---+
                    # So the row text line is: "|   | C |"
                    # meaning one pipe, spaces for col0, pipe, C, pipe
                    span_w = sum(widths[k] + 2 for k in range(start, c)) + (c - start - 1)
                    # Between columns normally we'd have '+' = 1 char in hline;
                    # in content rows separators are '|'. Covered region still
                    # uses '|' only at boundaries. The spaces for one col are
                    # width+2... For multi-col rowspan the content is just
                    # spaces of first col? Looking at pandoc output for vspan:
                    # content row 2: only cell C — AST had one cell.
                    # The textual form uses "|   | C |" with 3 spaces for col A width.
                    w0 = widths[start]
                    parts.append(" " + _pad("", w0) + " ")
                    parts.append("|")
                    # If rowspan covered multiple cols, we already only used first
                    # — for our builder rowspan is usually colspan 1.
                    continue
                # find cell starting here
                found = None
                for (cc, cs), lines in cell_lines_map.items():
                    if cc == c:
                        found = (cs, lines)
                        break
                if found is None:
                    parts.append(" " + _pad("", widths[c]) + " ")
                    parts.append("|")
                    c += 1
                    continue
                cs, lines = found
                chunk = lines[li] if li < len(lines) else ""
                total_w = sum(widths[c:c + cs]) + 3 * (cs - 1)  # widths + ' | ' between?
                # In grid tables, a colspan cell is written as one field spanning
                # multiple column widths including the interior '+' positions as spaces:
                # | AB        | C |
                # where AB spans two cols: width0+2+1+width1+2 interior... 
                # Content width = sum(widths[i]+2 for i in span) + (cs-1) for the
                # '+' chars that become spaces, minus the outer padding...
                # Easier: content area width = sum(widths[c:c+cs]) + 3*(cs-1)
                # because each col contributes w, and between cols ' | ' (3 chars)
                # is replaced by spaces inside a colspan.
                content_w = sum(widths[c:c + cs]) + 3 * (cs - 1)
                parts.append(" " + _pad(chunk, content_w) + " ")
                parts.append("|")
                c += cs
            out.append("".join(parts))
        # hline after row
        out.append(_hline_for_row_boundary(origins, widths, col_count, r))
    return "\n".join(out) + "\n"


def _covered(origins, r, c) -> bool:
    return _covered_by_other(origins, r, c)


def _covered_by_other(origins, r, c) -> bool:
    for (or_, oc), (cs, rs, _t) in origins.items():
        if or_ == r and oc == c:
            return False
        if or_ <= r < or_ + rs and oc <= c < oc + cs:
            if not (or_ == r and oc == c):
                return True
    return False


def _hline_for_row_boundary(origins, widths, col_count, r_after) -> str:
    """Horizontal line after data row ``r_after`` (0-based)."""
    # For each column, decide if the border below is solid or open (rowspan cont)
    open_cols = [False] * col_count
    for (or_, oc), (cs, rs, _t) in origins.items():
        # Rowspan continues below this boundary if r_after is strictly before
        # the last covered row index (or_+rs-1)
        last = or_ + rs - 1
        if or_ <= r_after < last:
            for k in range(oc, oc + cs):
                if k < col_count:
                    open_cols[k] = True

    parts = ["+"]
    c = 0
    while c < col_count:
        if open_cols[c]:
            # run of open columns — spaces, with ' ' instead of '+' between
            while c < col_count and open_cols[c]:
                parts.append(" " * (widths[c] + 2))
                c += 1
                if c < col_count and open_cols[c]:
                    parts.append(" ")
                else:
                    parts.append("+")
        else:
            parts.append("-" * (widths[c] + 2))
            parts.append("+")
            c += 1
    return "".join(parts)


def pipe_table_markers_to_grid(table_text: str) -> Optional[str]:
    """Convert one pipe table with merge markers into a grid_table string.

    Returns None if the table has no markers or cannot be parsed safely.
    """
    if not table_has_merge_markers(table_text):
        return None
    parsed = _split_pipe_rows(table_text)
    if parsed is None:
        return None
    header, _sep, data = parsed
    col_count = len(header)
    all_rows = [header] + data

    # Build occupancy to validate placeholders
    # logical_rows[r] = list of (colspan, rowspan, text) for cells that start at r
    logical: List[List[Tuple[int, int, str]]] = []
    occupied = [[False] * col_count for _ in range(len(all_rows) + 32)]

    for r, raw_cells in enumerate(all_rows):
        # Ensure occupancy grid long enough
        while len(occupied) < r + 16:
            occupied.append([False] * col_count)
        row_cells: List[Tuple[int, int, str]] = []
        c = 0
        while c < col_count:
            if occupied[r][c]:
                # must be empty placeholder in source
                if raw_cells[c].strip() not in ("",):
                    # conflicting content in rowspan/colspan continuation
                    return None
                c += 1
                continue
            cs, rs, text = parse_cell_markers(raw_cells[c])
            if c + cs > col_count:
                return None
            # placeholders for colspan: next cs-1 cells must be empty
            for k in range(1, cs):
                if raw_cells[c + k].strip():
                    return None
            # mark occupancy
            for rr in range(r, r + rs):
                while len(occupied) <= rr:
                    occupied.append([False] * col_count)
                for cc in range(c, c + cs):
                    if occupied[rr][cc] and not (rr == r and cc == c):
                        return None
                    occupied[rr][cc] = True
            row_cells.append((cs, rs, text))
            c += cs
        logical.append(row_cells)

    return build_grid_table(logical, col_count)


_PIPE_TABLE_BLOCK = re.compile(
    r"(^\|[^\n]+\|\n\|[-:\s|]+\|\n(?:\|[^\n]+\|\n?)*)",
    re.MULTILINE,
)


def rewrite_marked_pipe_tables(content: str) -> str:
    """Rewrite every pipe table that contains merge markers into a grid_table."""

    def repl(match: re.Match) -> str:
        block = match.group(1)
        if not table_has_merge_markers(block):
            return block
        grid = pipe_table_markers_to_grid(block)
        if grid is None:
            return block
        # Ensure surrounding blank lines for pandoc
        return grid if block.endswith("\n") else grid

    return _PIPE_TABLE_BLOCK.sub(repl, content)


# ── DOCX -> grid / pipe ─────────────────────────────────────

def _tc_grid_span(tc) -> int:
    from docx.oxml.ns import qn
    tcPr = tc.tcPr
    if tcPr is None:
        return 1
    el = tcPr.find(qn("w:gridSpan"))
    if el is None:
        return 1
    try:
        return max(1, int(el.get(qn("w:val"))))
    except (TypeError, ValueError):
        return 1


def _tc_vmerge(tc) -> Optional[str]:
    """Return 'restart', 'continue', or None."""
    from docx.oxml.ns import qn
    tcPr = tc.tcPr
    if tcPr is None:
        return None
    el = tcPr.find(qn("w:vMerge"))
    if el is None:
        return None
    val = el.get(qn("w:val"))
    if val in (None, "continue"):
        return "continue"
    return "restart"


def docx_table_has_merges(table) -> bool:
    seen = set()
    for row in table.rows:
        for cell in row.cells:
            tc = cell._tc
            tid = id(tc)
            if tid in seen:
                return True
            seen.add(tid)
            if _tc_grid_span(tc) > 1:
                return True
            if _tc_vmerge(tc) is not None:
                return True
    return False


def _cell_text_md(cell) -> str:
    parts = []
    for p in cell.paragraphs:
        t = p.text.strip()
        if t:
            parts.append(t)
    return "<br>".join(parts)


def docx_table_to_markdown(table) -> str:
    """Emit grid_table if merges exist, else a GFM pipe table."""
    if not docx_table_has_merges(table):
        return _docx_table_to_pipe(table)
    return _docx_table_to_grid(table)


def _docx_table_to_pipe(table) -> str:
    from .md_to_office import _render_table_cell

    rows = []
    for row in table.rows:
        cells = []
        seen = set()
        for cell in row.cells:
            tid = id(cell._tc)
            if tid in seen:
                cells.append("")
            else:
                seen.add(tid)
                cells.append(_render_table_cell(_cell_text_md(cell)))
        rows.append(cells)
    if not rows:
        return ""
    cols = max(len(r) for r in rows)
    for r in rows:
        while len(r) < cols:
            r.append("")
    out = [
        "| " + " | ".join(rows[0]) + " |",
        "| " + " | ".join(["---"] * cols) + " |",
    ]
    for r in rows[1:]:
        out.append("| " + " | ".join(r) + " |")
    return "\n".join(out)


def _docx_table_to_grid(table) -> str:
    """Build pandoc grid_table from a python-docx Table with merges."""
    # Determine grid column count from tblGrid if possible
    from docx.oxml.ns import qn

    tbl = table._tbl
    grid = tbl.find(qn("w:tblGrid"))
    if grid is not None:
        col_count = len(list(grid.findall(qn("w:gridCol"))))
    else:
        col_count = 0
        for row in table.rows:
            col_count = max(col_count, len(row.cells))

    # Walk unique tc elements per row in document order
    logical: List[List[Tuple[int, int, str]]] = []
    # Track vMerge restart cells to compute rowspan: map tc_id -> (row_idx, col_idx, colspan)
    vmerge_starts = {}  # id(tc) -> index in logical tracking

    # First pass: collect per-row starts with provisional rowspan=1
    row_specs: List[List[Tuple[int, int, str, object]]] = []  # + tc ref
    for r_idx, row in enumerate(table.rows):
        specs: List[Tuple[int, int, str, object]] = []
        seen = set()
        # Iterate cells but skip duplicate tc and vMerge continue
        # python-docx expands horizontally; we need grid positions
        # Use the row's tr content directly for accurate order
        tr = row._tr
        c_pos = 0
        for tc in tr.findall(qn("w:tc")):
            cs = _tc_grid_span(tc)
            vm = _tc_vmerge(tc)
            # Get text via cell object if possible
            # Find matching cell by _tc identity
            text = ""
            for cell in row.cells:
                if cell._tc is tc or id(cell._tc) == id(tc):
                    text = _cell_text_md(cell)
                    break
            if vm == "continue":
                # continuation — omit from logical row; rowspan bumped later
                c_pos += cs
                continue
            specs.append((cs, 1, text, tc))
            if vm == "restart":
                vmerge_starts[id(tc)] = (r_idx, len(specs) - 1)
            c_pos += cs
        row_specs.append(specs)

    # Compute rowspans: for each restart, count following continue in same column
    # Build column occupancy from tr walk again
    # Simpler: for each restart tc, look at subsequent rows' tc at same grid col
    def row_grid_map(r_idx):
        """Map starting grid column -> (tc, colspan, vmerge)."""
        m = {}
        tr = table.rows[r_idx]._tr
        c = 0
        for tc in tr.findall(qn("w:tc")):
            cs = _tc_grid_span(tc)
            m[c] = (tc, cs, _tc_vmerge(tc))
            c += cs
        return m

    maps = [row_grid_map(i) for i in range(len(table.rows))]

    # Find each restart's rowspan
    rowspan_at = {}  # (r, start_col) -> rowspan
    for r_idx, gmap in enumerate(maps):
        for c, (tc, cs, vm) in gmap.items():
            if vm != "restart" and _tc_grid_span(tc) == cs:
                if vm == "restart":
                    pass
            if vm == "restart":
                span = 1
                for r2 in range(r_idx + 1, len(maps)):
                    other = maps[r2].get(c)
                    if other and other[2] == "continue" and other[1] == cs:
                        span += 1
                    else:
                        break
                rowspan_at[(r_idx, c)] = span

    # Rebuild logical rows with correct rowspan
    logical = []
    for r_idx, gmap in enumerate(maps):
        row_cells: List[Tuple[int, int, str]] = []
        # iterate cols in order
        for c in sorted(gmap.keys()):
            tc, cs, vm = gmap[c]
            if vm == "continue":
                continue
            rs = rowspan_at.get((r_idx, c), 1)
            text = ""
            for cell in table.rows[r_idx].cells:
                if cell._tc is tc or id(cell._tc) == id(tc):
                    text = _cell_text_md(cell)
                    break
            row_cells.append((cs, rs, text))
        logical.append(row_cells)

    if not logical:
        return _docx_table_to_pipe(table)
    return build_grid_table(logical, col_count)


def repair_docx_vertical_merges(docx_path: str) -> bool:
    """Fix pandoc DOCX writer quirks for vertical merges.

    Pandoc correctly emits ``RowSpan`` in its AST for grid_tables, but the
    DOCX writer sometimes writes ``w:vMerge=restart`` on *every* row of a
    vertical merge (or omits vMerge and leaves a duplicate cell).  Detect
    consecutive same-column cells that clearly belong together and call
    ``python-docx`` ``cell.merge()`` so Word sees a real vertical merge.

    Returns True if the document was modified and saved.
    """
    from docx import Document
    from docx.oxml.ns import qn

    path = str(docx_path)
    doc = Document(path)
    changed = False

    for table in doc.tables:
        # Build per-row map: start_col -> (tc, colspan, vmerge, cell)
        nrows = len(table.rows)
        if nrows < 2:
            continue

        def row_map(r_idx):
            m = {}
            row = table.rows[r_idx]
            tr = row._tr
            c = 0
            # Match tc elements to cell objects
            tc_to_cell = {}
            for cell in row.cells:
                tc_to_cell[id(cell._tc)] = cell
            for tc in tr.findall(qn("w:tc")):
                cs = _tc_grid_span(tc)
                vm = _tc_vmerge(tc)
                cell = tc_to_cell.get(id(tc))
                if cell is None:
                    for cell2 in row.cells:
                        if cell2._tc is tc:
                            cell = cell2
                            break
                m[c] = (tc, cs, vm, cell)
                c += cs
            return m

        maps = [row_map(i) for i in range(nrows)]
        # Track which (r,c) already absorbed into a merge
        absorbed = set()

        for r_idx in range(nrows - 1):
            gmap = maps[r_idx]
            for c, (tc, cs, vm, cell) in list(gmap.items()):
                if cell is None or (r_idx, c) in absorbed:
                    continue
                # Candidate start: explicit restart, OR plain cell followed by
                # duplicate-text / empty restart-ish cells below.
                text0 = (cell.text or "").strip()
                # Look ahead
                merge_cells = []
                for r2 in range(r_idx + 1, nrows):
                    if (r2, c) in absorbed:
                        break
                    other = maps[r2].get(c)
                    if other is None:
                        break
                    tc2, cs2, vm2, cell2 = other
                    if cell2 is None or cs2 != cs:
                        break
                    # Already a proper continue → count but no action needed
                    if vm2 == "continue":
                        absorbed.add((r2, c))
                        continue
                    text2 = (cell2.text or "").strip()
                    # Pandoc bug: another restart with same text, or empty
                    # placeholder under a rowspan grid cell.
                    same = text2 == text0 and text0 != ""
                    empty_cont = text2 == "" and (vm == "restart" or vm2 == "restart" or text0 != "")
                    # Only merge when we have signal this was meant to be rowspan
                    if vm == "restart" and (vm2 == "restart" or empty_cont or same):
                        merge_cells.append(cell2)
                        absorbed.add((r2, c))
                    elif same and vm2 == "restart":
                        merge_cells.append(cell2)
                        absorbed.add((r2, c))
                    elif empty_cont and vm == "restart":
                        merge_cells.append(cell2)
                        absorbed.add((r2, c))
                    else:
                        break
                if merge_cells:
                    try:
                        anchor = cell
                        for other_cell in merge_cells:
                            # Skip if already same tc
                            if id(anchor._tc) == id(other_cell._tc):
                                continue
                            anchor = anchor.merge(other_cell)
                            changed = True
                    except Exception:
                        # Merge can fail on complex grids; leave as-is
                        pass

    if changed:
        doc.save(path)
    return changed
