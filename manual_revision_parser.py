"""作者：Wangyuan；日期：2026-10-29；职责：通用文档内容树、差异及渲染提取模块。"""
from __future__ import annotations

import hashlib
import json
import re
import unicodedata
from pathlib import Path
from typing import Iterable

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from manual_revision_diff import classify_revision_row
from manual_revision_models import AppendixBlock, ContentBlock, ParsedRevisionDocument, RevisionMetadata, RevisionRow


STANDARD_HEADERS = ["序号", "修订前内容", "修订后内容", "涉及章节", "涉及部门", "依据", "备注"]

HEADER_ALIASES = {
    "序号": "index",
    "编号": "index",
    "修订前内容": "before_text",
    "修订前": "before_text",
    "原内容": "before_text",
    "原条款": "before_text",
    "修订后内容": "after_text",
    "修订后": "after_text",
    "新内容": "after_text",
    "新条款": "after_text",
    "涉及章节": "chapter",
    "章节名称": "chapter",
    "章节": "chapter",
    "条款": "chapter",
    "涉及部门": "department",
    "涉及部门科室": "department",
    "科室": "department",
    "部门": "department",
    "依据": "basis",
    "修订依据": "basis",
    "原因": "basis",
    "备注": "remark",
    "说明": "remark",
}


def _normalize_header(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    text = re.sub(r"[\s　()（）/、：:]+", "", text)
    return text


def _clean_cell_text(text: str) -> str:
    text = (text or "").replace("\r", "\n").replace("\x0b", "\n")
    lines = [re.sub(r"[ \t　]+", " ", line).strip() for line in text.splitlines()]
    return "\n".join(line for line in lines if line).strip()


def _plain_for_meta(text: str) -> str:
    return re.sub(r"\s+", " ", (text or "").replace("\n", " ")).strip()


def _unique_cell_texts(row) -> list[str]:
    values: list[str] = []
    for cell in row.cells:
        text = _clean_cell_text(cell.text)
        if text and text not in values:
            values.append(text)
    return values


def _parse_title(texts: Iterable[str]) -> tuple[str, str, str]:
    for raw in texts:
        text = _plain_for_meta(raw)
        if "修订差异记录表" not in text:
            continue
        manual = ""
        version = ""
        manual_match = re.search(r"《([^》]+)》", text)
        version_match = re.search(r"(R(?:\d{2}|XX)-(?:\d{2}|XX))", text, flags=re.I)
        if manual_match:
            manual = manual_match.group(1).strip()
        if version_match:
            version = version_match.group(1).upper()
        return text, manual, version
    return "", "", ""


def _extract_label(text: str, labels: tuple[str, ...], stop_labels: tuple[str, ...]) -> str:
    normalized = text.replace("\n", " ")
    for label in labels:
        match = re.search(re.escape(label) + r"\s*[：:]\s*", normalized)
        if not match:
            continue
        start = match.end()
        stop_positions = []
        for stop in stop_labels:
            found = re.search(re.escape(stop) + r"\s*[：:]", normalized[start:])
            if found:
                stop_positions.append(start + found.start())
        end = min(stop_positions) if stop_positions else len(normalized)
        value = normalized[start:end]
        return re.sub(r"\s+", " ", value).strip(" ：:;；")
    return ""


def _extract_metadata(texts: Iterable[str], title: str, manual_name: str, revision_version: str, source_filename: str) -> RevisionMetadata:
    merged = "\n".join(_plain_for_meta(text) for text in texts if _plain_for_meta(text))
    writer = _extract_label(merged, ("差异表编写人", "编写人"), ("审核人", "修订日期", "编写完成日期", "日期"))
    reviewer = _extract_label(merged, ("审核人",), ("修订日期", "编写完成日期", "日期"))
    revision_date = _extract_label(merged, ("修订日期", "编写完成日期", "日期"), ("差异表编写人", "编写人", "审核人"))
    return RevisionMetadata(
        title=title,
        manual_name=manual_name,
        revision_version=revision_version,
        writer=writer,
        reviewer=reviewer,
        revision_date=revision_date,
        source_filename=source_filename,
    )


def _field_for_header(text: str) -> str | None:
    return HEADER_ALIASES.get(_normalize_header(text))


def _locate_revision_table(doc: Document):
    best = None
    best_score = -1
    for table_index, table in enumerate(doc.tables):
        for row_index, row in enumerate(table.rows):
            header_map: dict[str, int] = {}
            for col_index, cell in enumerate(row.cells):
                field = _field_for_header(cell.text)
                if field and field not in header_map:
                    header_map[field] = col_index
            score = len(header_map)
            if {"before_text", "after_text"}.issubset(header_map) and score > best_score:
                best = (table_index, table, row_index, header_map)
                best_score = score
    if best is None:
        raise ValueError("未找到包含“修订前内容/修订后内容”的差异表。")
    return best


def _row_value(row, header_map: dict[str, int], field: str) -> str:
    idx = header_map.get(field)
    if idx is None or idx >= len(row.cells):
        return ""
    return _clean_cell_text(row.cells[idx].text)


def _table_digest(rows: list[list[str]]) -> str:
    payload = json.dumps(rows, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()[:16]


def _dxa_to_mm(value: str | int | None) -> float | None:
    try:
        return round(int(value) / 56.6929133858, 2)
    except (TypeError, ValueError):
        return None


def _grid_widths_mm(table) -> list[float]:
    widths: list[float] = []
    grid = getattr(table._tbl, "tblGrid", None)
    if grid is None:
        return widths
    for grid_col in grid.gridCol_lst:
        width = _dxa_to_mm(grid_col.get(qn("w:w")))
        if width and width > 0:
            widths.append(width)
    return widths


def _image_extension(content_type: str, fallback_name: str = "") -> str:
    content_type = (content_type or "").lower()
    if "png" in content_type:
        return ".png"
    if "jpeg" in content_type or "jpg" in content_type:
        return ".jpg"
    if "gif" in content_type:
        return ".gif"
    if "bmp" in content_type:
        return ".bmp"
    if "x-emf" in content_type or "emf" in content_type:
        return ".emf"
    if "x-wmf" in content_type or "wmf" in content_type:
        return ".wmf"
    if "svg" in content_type:
        return ".svg"
    suffix = Path(fallback_name or "").suffix.lower()
    return suffix if suffix in {".png", ".jpg", ".jpeg", ".gif", ".bmp"} else ".png"


def _extract_table_block(table) -> ContentBlock:
    if table._tbl.xpath(".//w:tbl") or table._tbl.xpath(".//w:drawing|.//w:pict"):raise ValueError("嵌套表内包含更深表格或图片，当前不支持无损重建，请人工处理")
    rows: list[list[str]] = []
    for table_row in table.rows:
        values = [_clean_cell_text(cell.text) for cell in table_row.cells]
        if any(values):
            rows.append(values)
    return ContentBlock(kind="table", rows=rows, column_widths_mm=_grid_widths_mm(table), digest=_table_digest(rows))


def _extract_image_blocks(doc: Document, cell, asset_dir: Path | None, prefix: str) -> list[ContentBlock]:
    blocks: list[ContentBlock] = []
    drawings = cell._tc.xpath(".//w:drawing")
    candidates = []
    for drawing in drawings:
        blips = drawing.xpath(".//a:blip")
        if not blips:
            continue
        extent = None
        extents = drawing.xpath(".//wp:extent")
        if extents:
            width_mm = _dxa_to_mm(str(round(int(extents[0].get("cx")) / 635)))
            height_mm = _dxa_to_mm(str(round(int(extents[0].get("cy")) / 635)))
            extent = (width_mm, height_mm)
        candidates.append((blips[0], extent, drawing.xml))
    if not candidates:
        candidates = [(blip, None, "") for blip in cell._tc.xpath(".//a:blip")]

    for blip, extent, drawing_xml in candidates:
        relationship_id = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
        if not relationship_id:
            continue
        part = doc.part.related_parts.get(relationship_id)
        if part is None or not hasattr(part, "blob"):
            continue
        blob = part.blob
        content_type = str(getattr(part, "content_type", "") or "")
        digest = hashlib.sha256(blob).hexdigest()[:16]
        original_name = Path(str(getattr(part, "partname", "") or "")).name
        ext = _image_extension(content_type, original_name)
        filename = f"{prefix}_img{len(blocks) + 1}{ext}"
        image_path = ""
        if asset_dir is not None:
            asset_dir.mkdir(parents=True, exist_ok=True)
            path = asset_dir / filename
            path.write_bytes(blob)
            image_path = str(path)
        blocks.append(
            ContentBlock(
                kind="image",
                image_path=image_path,
                filename=filename,
                content_type=content_type,
                digest=digest,
                width_mm=extent[0] if extent else None,
                height_mm=extent[1] if extent else None,
                drawing_xml=drawing_xml,
            )
        )
    return blocks


def _appendix_table_rows(table: Table) -> list[list[str]]:
    rows: list[list[str]] = []
    for table_row in table.rows:
        values = [_clean_cell_text(cell.text) for cell in table_row.cells]
        if any(values):
            rows.append(values)
    return rows


def _extract_appendix_blocks(doc: Document, revision_table) -> list[AppendixBlock]:
    blocks: list[AppendixBlock] = []
    after_revision_table = False
    for element in doc.element.body.iterchildren():
        if element is revision_table._tbl:
            after_revision_table = True
            continue
        if not after_revision_table:
            continue
        if element.tag == qn("w:sectPr"):
            break
        if element.xpath(".//w:drawing|.//w:pict") or element.tag==qn("w:tbl") and element.xpath(".//w:tbl"):
            raise ValueError("附录含图片或嵌套表格，当前不支持无损重建，请人工处理")
        if element.tag == qn("w:p"):
            text = _clean_cell_text(Paragraph(element, doc).text)
            if text:
                blocks.append(AppendixBlock(kind="paragraph", text=text))
        elif element.tag == qn("w:tbl"):
            table = Table(element, doc)
            rows = _appendix_table_rows(table)
            if rows:
                blocks.append(AppendixBlock(kind="table", rows=rows, column_widths_mm=_grid_widths_mm(table)))
    return blocks


def _row_content(doc: Document, row, header_map: dict[str, int], field: str, asset_dir: Path | None, source_row: int) -> tuple[str, list[ContentBlock]]:
    idx = header_map.get(field)
    if idx is None or idx >= len(row.cells):
        return "", []
    cell = row.cells[idx]
    text = _clean_cell_text(cell.text)
    blocks = [_extract_table_block(table) for table in cell.tables]
    blocks.extend(_extract_image_blocks(doc, cell, asset_dir, f"r{source_row}_{field}"))
    return text, [block for block in blocks if block.kind != "table" or block.rows]


def _content_warning(blocks: list[ContentBlock]) -> str:
    table_count = sum(1 for block in blocks if block.kind == "table")
    image_count = sum(1 for block in blocks if block.kind == "image")
    parts = []
    if table_count:
        parts.append(f"{table_count}个内嵌表格")
    if image_count:
        parts.append(f"{image_count}张图片")
    return "、".join(parts)


def parse_revision_docx(docx_path: Path, source_filename: str = "", converted_path: str = "", asset_dir: Path | None = None) -> ParsedRevisionDocument:
    """解析修订表、元数据、嵌套内容与附录；输入安全校验由上层入口负责。"""
    doc = Document(docx_path)
    paragraph_texts = [p.text for p in doc.paragraphs if p.text.strip()]
    table_texts: list[str] = []
    for table in doc.tables:
        for row in table.rows[:5]:
            table_texts.extend(_unique_cell_texts(row))

    title, manual_name, revision_version = _parse_title([*paragraph_texts, *table_texts, source_filename])
    table_index, table, header_index, header_map = _locate_revision_table(doc)
    meta_texts = [*paragraph_texts]
    for row in table.rows[:header_index]:
        meta_texts.extend(_unique_cell_texts(row))
    metadata = _extract_metadata(meta_texts, title, manual_name, revision_version, source_filename)

    rows: list[RevisionRow] = []
    skipped = 0
    warnings: list[str] = []
    if not metadata.manual_name:
        warnings.append("未能自动识别手册名称，请生成前确认。")
    if not metadata.revision_version:
        warnings.append("未能自动识别版本号，请生成前确认。")

    for source_row, row in enumerate(table.rows[header_index + 1 :], start=header_index + 2):
        before_text, before_blocks = _row_content(doc, row, header_map, "before_text", asset_dir, source_row)
        after_text, after_blocks = _row_content(doc, row, header_map, "after_text", asset_dir, source_row)
        values = {
            "before_text": before_text,
            "after_text": after_text,
            "chapter": _row_value(row, header_map, "chapter"),
            "department": _row_value(row, header_map, "department"),
            "basis": _row_value(row, header_map, "basis"),
            "remark": _row_value(row, header_map, "remark"),
        }
        if not any(values.values()) and not before_blocks and not after_blocks:
            skipped += 1
            continue
        if _normalize_header(values["before_text"]) == "修订前内容" or _normalize_header(values["after_text"]) == "修订后内容":
            skipped += 1
            continue
        if not values["before_text"] and not values["after_text"] and not before_blocks and not after_blocks:
            skipped += 1
            continue
        row_warnings: list[str] = []
        if values["before_text"] == values["after_text"] and before_blocks == after_blocks:
            row_warnings.append("修订前后内容一致，请确认是否需要保留。")
        rich_warning = _content_warning([*before_blocks, *after_blocks])
        if rich_warning:
            row_warnings.append(f"检测到{rich_warning}，已随解析结果保留。")
        candidate = RevisionRow(
            index=len(rows) + 1,
            before_text=values["before_text"],
            after_text=values["after_text"],
            chapter=values["chapter"],
            department=values["department"],
            basis=values["basis"],
            remark=values["remark"],
            source_row=source_row,
            warnings=row_warnings,
            before_blocks=before_blocks,
            after_blocks=after_blocks,
        )
        candidate.revision_type = classify_revision_row(candidate)
        rows.append(
            RevisionRow(
                index=candidate.index,
                before_text=candidate.before_text,
                after_text=candidate.after_text,
                chapter=candidate.chapter,
                department=candidate.department,
                basis=candidate.basis,
                remark=candidate.remark,
                revision_type=candidate.revision_type,
                source_row=candidate.source_row,
                warnings=candidate.warnings,
                before_blocks=candidate.before_blocks,
                after_blocks=candidate.after_blocks,
            )
        )

    if not rows:
        warnings.append("未抽取到有效修订记录。")

    return ParsedRevisionDocument(
        metadata=metadata,
        rows=rows,
        skipped_rows=skipped,
        warnings=warnings,
        source_path=str(docx_path),
        converted_path=converted_path,
        appendix_blocks=_extract_appendix_blocks(doc, table),
    )
