"""作者：Wangyuan；日期：2026-10-29；职责：通用文档内容树、差异及渲染提取模块。"""
from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.opc.part import Part
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.shared import Mm, Pt


from manual_revision_models import AppendixBlock, ContentBlock, DiffRun, RenderedRevisionRow, RevisionMetadata


FONT = "SimSun"
HEADER_FILL = "CCCCCC"
PAGE_WIDTH_MM = 297
PAGE_HEIGHT_MM = 210
CELL_WIDTHS_MM = [12, 65, 65, 28, 28, 25, 25]
CELL_WIDTHS_WITH_APPENDIX_MM = CELL_WIDTHS_MM
APPENDIX_WIDTHS_BY_COLS_MM = {
    6: [41, 41, 41, 41, 41, 41],
}
HEADERS = ["序号", "修订前内容", "修订后内容", "章节名称", "涉及部门\n（科室）", "依据", "备注"]


def _set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def _set_cell_border(cell) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "6")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "000000")


def _set_cell_margins(cell, top: int = 80, start: int = 80, bottom: int = 80, end: int = 80) -> None:
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    margins = tc_pr.first_child_found_in("w:tcMar")
    if margins is None:
        margins = OxmlElement("w:tcMar")
        tc_pr.append(margins)
    for m, value in {"top": top, "start": start, "bottom": bottom, "end": end}.items():
        node = margins.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            margins.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _set_row_height(row, value_dxa: int, rule: str = "atLeast") -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tr_height = tr_pr.trHeight
    if tr_height is None:
        tr_height = OxmlElement("w:trHeight")
        tr_pr.append(tr_height)
    tr_height.set(qn("w:val"), str(value_dxa))
    tr_height.set(qn("w:hRule"), rule)


def _set_run_font(run, *, size_pt: float = 10.5, bold: bool = False, color: str | None = None, underline: bool = False, strike: bool = False) -> None:
    run.font.name = FONT
    run._element.rPr.rFonts.set(qn("w:eastAsia"), FONT)
    run.font.size = Pt(size_pt)
    run.bold = bold
    if color:
        run.font.color.rgb = _rgb(color)
    run.underline = underline
    run.font.strike = strike


def _rgb(hex_color: str):
    from docx.shared import RGBColor

    value = hex_color.strip().lstrip("#")
    return RGBColor(int(value[0:2], 16), int(value[2:4], 16), int(value[4:6], 16))


def _mm_to_dxa(value: float) -> int:
    return int(round(float(value) * 56.6929133858))


def _set_table_fixed_geometry(table, widths_mm: list[float], indent_dxa: int | None = None) -> None:
    widths_dxa = [_mm_to_dxa(width) for width in widths_mm]
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    if tbl_pr is None:
        tbl_pr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tbl_pr)

    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.append(tbl_w)
    tbl_w.set(qn("w:w"), str(sum(widths_dxa)))
    tbl_w.set(qn("w:type"), "dxa")

    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    existing_grid = table._tbl.tblGrid
    if existing_grid is not None:
        table._tbl.remove(existing_grid)
    grid = OxmlElement("w:tblGrid")
    for width in widths_dxa:
        grid_col = OxmlElement("w:gridCol")
        grid_col.set(qn("w:w"), str(width))
        grid.append(grid_col)
    table._tbl.insert(1, grid)

    if indent_dxa is not None:
        ind = tbl_pr.find(qn("w:tblInd"))
        if ind is None:
            ind = OxmlElement("w:tblInd")
            tbl_pr.append(ind)
        ind.set(qn("w:w"), str(indent_dxa))
        ind.set(qn("w:type"), "dxa")

    for row in table.rows:
        for idx, cell in enumerate(row.cells):
            if idx >= len(widths_dxa):
                continue
            cell.width = Mm(widths_mm[idx])
            tc_pr = cell._tc.get_or_add_tcPr()
            tc_w = tc_pr.tcW
            if tc_w is None:
                tc_w = OxmlElement("w:tcW")
                tc_pr.append(tc_w)
            tc_w.set(qn("w:w"), str(widths_dxa[idx]))
            tc_w.set(qn("w:type"), "dxa")


def _clear_cell(cell) -> None:
    cell.text = ""


def _paragraph(cell, align=WD_ALIGN_PARAGRAPH.LEFT):
    if not cell.paragraphs:
        cell.add_paragraph()
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.1
    return paragraph


def _add_text(paragraph, text: str, *, size_pt: float = 10.5, bold: bool = False) -> None:
    parts = str(text or "").split("\n")
    for index, part in enumerate(parts):
        if index:
            paragraph.add_run().add_break()
        run = paragraph.add_run(part)
        _set_run_font(run, size_pt=size_pt, bold=bold)


def _add_runs(paragraph, runs: list[DiffRun], *, size_pt: float = 10.5) -> None:
    for diff_run in runs:
        parts = diff_run.text.split("\n")
        for index, part in enumerate(parts):
            if index:
                paragraph.add_run().add_break()
            run = paragraph.add_run(part)
            _set_run_font(
                run,
                size_pt=size_pt,
                bold=diff_run.bold,
                color=diff_run.color,
                underline=diff_run.role == "underline",
                strike=diff_run.role == "strike",
            )


def _new_paragraph(cell, align=WD_ALIGN_PARAGRAPH.LEFT):
    paragraph = cell.add_paragraph()
    paragraph.alignment = align
    paragraph.paragraph_format.space_before = Pt(2)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.05
    return paragraph


def _style_cell(cell, align=WD_ALIGN_PARAGRAPH.LEFT, fill: str | None = None) -> None:
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    _set_cell_border(cell)
    _set_cell_margins(cell)
    if fill:
        _set_cell_shading(cell, fill)
    for paragraph in cell.paragraphs:
        paragraph.alignment = align
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)


def _scaled_widths(widths_mm: list[float], column_count: int, max_width_mm: float | None) -> list[float]:
    if not widths_mm or len(widths_mm) != column_count:
        total = max_width_mm or 80
        return [round(total / max(1, column_count), 2)] * max(1, column_count)
    widths = [max(4.0, float(width)) for width in widths_mm]
    total = sum(widths)
    if max_width_mm and total > max_width_mm:
        factor = max_width_mm / total
        widths = [round(max(4.0, width * factor), 2) for width in widths]
    return widths


def _style_nested_table(table, column_count: int, widths_mm: list[float] | None = None, max_width_mm: float | None = None) -> None:
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    _set_table_fixed_geometry(table, _scaled_widths(widths_mm or [], column_count, max_width_mm))
    font_size = 7 if column_count <= 3 else 5.5
    cell_margin = 20 if column_count > 3 else 40
    for row in table.rows:
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            _set_cell_border(cell)
            _set_cell_margins(cell, top=cell_margin, start=cell_margin, bottom=cell_margin, end=cell_margin)
            for paragraph in cell.paragraphs:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                paragraph.paragraph_format.space_before = Pt(0)
                paragraph.paragraph_format.space_after = Pt(0)
                paragraph.paragraph_format.line_spacing = 1.0
                for run in paragraph.runs:
                    _set_run_font(run, size_pt=font_size)


def _append_table_block(cell, block: ContentBlock, max_width_mm: float | None = None) -> None:
    rows = block.rows or [[""]]
    cols = max(1, max((len(row) for row in rows), default=1))
    nested = cell.add_table(rows=len(rows), cols=cols)
    for row_index, values in enumerate(rows):
        for col_index in range(cols):
            nested_cell = nested.rows[row_index].cells[col_index]
            nested_cell.text = values[col_index] if col_index < len(values) else ""
    _style_nested_table(nested, cols, block.column_widths_mm, max_width_mm)


def _append_image_block(cell, block: ContentBlock, max_width_mm: float | None = None) -> None:
    paragraph = _new_paragraph(cell, WD_ALIGN_PARAGRAPH.LEFT)
    image_path = Path(block.image_path) if block.image_path else None
    if _append_raw_drawing_if_possible(paragraph, block, image_path):
        return
    usable_path = _usable_image_path(image_path)
    if usable_path:
        try:
            run = paragraph.add_run()
            width_mm = block.width_mm or max_width_mm or 42
            if max_width_mm:
                width_mm = min(width_mm, max_width_mm)
            run.add_picture(str(usable_path), width=Mm(width_mm))
            return
        except Exception:  # noqa: BLE001
            pass
    _add_text(paragraph, f"[图片无法插入：{block.filename or '未命名'}]", size_pt=8.5)


def _append_raw_drawing_if_possible(paragraph, block: ContentBlock, image_path: Path | None) -> bool:
    if not image_path or not image_path.is_file() or not block.drawing_xml:
        return False
    suffix = image_path.suffix.lower()
    if suffix not in {".emf", ".wmf"}:
        return False
    try:
        package = paragraph.part.package
        partname = package.next_partname(f"/word/media/image%d{suffix}")
        media_part = Part(partname, block.content_type or ("image/x-emf" if suffix == ".emf" else "image/x-wmf"), image_path.read_bytes(), package)
        relationship_id = paragraph.part.relate_to(media_part, RT.IMAGE)
        drawing = parse_xml(block.drawing_xml)
        for blip in drawing.xpath(".//a:blip"):
            blip.set(qn("r:embed"), relationship_id)
            if blip.get(qn("r:link")):
                del blip.attrib[qn("r:link")]
        run = paragraph.add_run()
        run._r.append(drawing)
        return True
    except Exception:  # noqa: BLE001
        return False


def _usable_image_path(image_path: Path | None) -> Path | None:
    """仅使用显式输入的本地图像，不发现或启动外部转换器。"""
    if image_path and image_path.is_file() and image_path.suffix.lower() in {".png", ".jpg", ".jpeg", ".gif", ".bmp"}:
        return image_path
    return None


def _append_content_blocks(cell, blocks: list[ContentBlock], max_width_mm: float | None = None) -> None:
    if not blocks:
        return
    for block in blocks:
        if block.kind == "table":
            _append_table_block(cell, block, max_width_mm)
        elif block.kind == "image":
            _append_image_block(cell, block, max_width_mm)


def _add_page_field(paragraph, instruction: str) -> None:
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" {instruction} "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    text = OxmlElement("w:t")
    text.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instr, separate, text, end])
    _set_run_font(run, size_pt=9)


def _add_footer(section) -> None:
    paragraph = section.footer.paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_text(paragraph, "第 ", size_pt=9)
    _add_page_field(paragraph, "PAGE")
    _add_text(paragraph, " 页/共 ", size_pt=9)
    _add_page_field(paragraph, "NUMPAGES")
    _add_text(paragraph, " 页", size_pt=9)


def _apply_table_widths(table, widths_mm: list[float]) -> None:
    _set_table_fixed_geometry(table, widths_mm)


def _append_appendix_table(doc: Document, block: AppendixBlock) -> None:
    rows = block.rows or []
    if not rows:
        return
    cols = max(1, max((len(row) for row in rows), default=1))
    table = doc.add_table(rows=len(rows), cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row_index, values in enumerate(rows):
        for col_index in range(cols):
            cell = table.rows[row_index].cells[col_index]
            cell.text = values[col_index] if col_index < len(values) else ""
            align = WD_ALIGN_PARAGRAPH.CENTER if col_index != max(0, cols - 2) else WD_ALIGN_PARAGRAPH.LEFT
            _style_cell(cell, align, fill=HEADER_FILL if row_index == 0 else None)
            for paragraph in cell.paragraphs:
                paragraph.alignment = align
                paragraph.paragraph_format.line_spacing = 1.0
                for run in paragraph.runs:
                    _set_run_font(run, size_pt=10.5, bold=True)
    widths = APPENDIX_WIDTHS_BY_COLS_MM.get(cols) or _scaled_widths(block.column_widths_mm, cols, 245)
    _set_table_fixed_geometry(table, widths, indent_dxa=135 if cols == 6 else None)


def _append_appendix_blocks(doc: Document, blocks: list[AppendixBlock]) -> None:
    for block in blocks:
        if block.kind == "paragraph" and block.text.strip():
            paragraph = doc.add_paragraph()
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.space_before = Pt(8)
            paragraph.paragraph_format.space_after = Pt(2)
            _add_text(paragraph, block.text, size_pt=10.5, bold=True)
        elif block.kind == "table":
            _append_appendix_table(doc, block)


def render_revision_docx(
    metadata: RevisionMetadata,
    rows: list[RenderedRevisionRow],
    output_path: Path,
    profile: dict,
    appendix_blocks: list[AppendixBlock] | None = None,
) -> Path:
    """从数据模型生成中性可编辑修订表，不读取真实模板。"""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = Document()
    section = doc.sections[0]
    section.start_type = WD_SECTION.NEW_PAGE
    section.page_width = Mm(PAGE_WIDTH_MM)
    section.page_height = Mm(PAGE_HEIGHT_MM)
    section.left_margin = Mm(18.0)
    section.right_margin = Mm(14.5)
    section.top_margin = Mm(10.0)
    section.bottom_margin = Mm(16.0)
    _add_footer(section)
    appendix_blocks = appendix_blocks or []
    cell_widths = CELL_WIDTHS_WITH_APPENDIX_MM if appendix_blocks else CELL_WIDTHS_MM

    title = metadata.title or f"《{metadata.manual_name or 'XXXX手册'}》{metadata.revision_version or 'RXX-XX'}修订差异记录表"
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_text(title_p, title, size_pt=16, bold=True)

    table = doc.add_table(rows=0, cols=7)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    if profile.get("top_spacer_row"):
        row = table.add_row()
        _set_row_height(row, int(profile.get("top_spacer_height_dxa") or 284), "exact")
        for cell in row.cells:
            _clear_cell(cell)
            _style_cell(cell, WD_ALIGN_PARAGRAPH.CENTER, fill=HEADER_FILL)

    layout = profile.get("metadata_layout") or "single_row"
    date_label = profile.get("date_label") or "修订日期"
    if layout == "merged_multiline":
        row = table.add_row()
        _set_row_height(row, int(profile.get("metadata_row_height_dxa") or 1892), "atLeast")
        merged = row.cells[0].merge(row.cells[-1])
        _clear_cell(merged)
        paragraph = _paragraph(merged)
        lines = [
            f"差异表编写人：{metadata.writer or ''}",
            f"审核人： {metadata.reviewer or ''}",
            f"{date_label}：{metadata.revision_date or ''}",
        ]
        _add_text(paragraph, "\n".join(lines), size_pt=15, bold=True)
        _style_cell(merged)
    elif layout == "three_rows":
        for label, value in (
            ("差异表编写人", metadata.writer),
            ("审核人", metadata.reviewer),
            (date_label, metadata.revision_date),
        ):
            row = table.add_row()
            merged = row.cells[0].merge(row.cells[-1])
            _clear_cell(merged)
            paragraph = _paragraph(merged)
            _add_text(paragraph, f"{label}：{value or ''}", size_pt=15, bold=True)
            _style_cell(merged)
    else:
        row = table.add_row()
        merged = row.cells[0].merge(row.cells[-1])
        _clear_cell(merged)
        paragraph = _paragraph(merged)
        gap = "        "
        text = f"差异表编写人：{metadata.writer or ''}{gap}审核人：{metadata.reviewer or ''}{gap}{date_label}：{metadata.revision_date or ''}"
        if len(text) > 48:
            text = f"差异表编写人：{metadata.writer or ''}    审核人：{metadata.reviewer or ''}    {date_label}：{metadata.revision_date or ''}"
        _add_text(paragraph, text, size_pt=13.5 if len(text) > 48 else 15, bold=True)
        _style_cell(merged)

    header_row = table.add_row()
    if profile.get("header_row_height_dxa"):
        _set_row_height(header_row, int(profile["header_row_height_dxa"]), "atLeast")
    for idx, label in enumerate(HEADERS):
        cell = header_row.cells[idx]
        _clear_cell(cell)
        paragraph = _paragraph(cell, WD_ALIGN_PARAGRAPH.CENTER)
        _add_text(paragraph, label, bold=True)
        _style_cell(cell, WD_ALIGN_PARAGRAPH.CENTER, fill=HEADER_FILL)

    for rendered in rows:
        row = table.add_row()
        values = [
            str(rendered.index),
            rendered.before_runs,
            rendered.after_runs,
            rendered.chapter_runs,
            rendered.department,
            rendered.basis,
            rendered.remark,
        ]
        for idx, value in enumerate(values):
            cell = row.cells[idx]
            _clear_cell(cell)
            align = WD_ALIGN_PARAGRAPH.CENTER if idx in {0, 3, 4, 5, 6} else WD_ALIGN_PARAGRAPH.LEFT
            paragraph = _paragraph(cell, align)
            if isinstance(value, list):
                _add_runs(paragraph, value)
            else:
                _add_text(paragraph, value, bold=idx == 0)
            if idx == 1:
                _append_content_blocks(cell, rendered.before_blocks, max_width_mm=cell_widths[idx] - 2)
            elif idx == 2:
                _append_content_blocks(cell, rendered.after_blocks, max_width_mm=cell_widths[idx] - 2)
            _style_cell(cell, align)

    _apply_table_widths(table, cell_widths)
    _append_appendix_blocks(doc, appendix_blocks)
    doc.save(output_path)
    return output_path
