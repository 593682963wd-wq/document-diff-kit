"""作者：Wangyuan；日期：2026-10-29；职责：通用文档内容树、差异及渲染提取模块。"""
from __future__ import annotations

import re
import unicodedata
from difflib import SequenceMatcher

from manual_revision_models import ContentBlock, DiffRun, RenderedRevisionRow, RevisionRow


RED = "FF0000"
EMPTY_VALUES = {"", "/", "无", "无。", "无内容", "不涉及"}
DELETE_VALUES = {"删除", "删减", "取消", "废止"}
ADD_VALUES = {"新增", "增加", "新建"}


def _normalize_for_compare(text: str) -> str:
    text = unicodedata.normalize("NFKC", text or "")
    text = re.sub(r"\s+", "", text)
    text = text.replace("“", "\"").replace("”", "\"").replace("‘", "'").replace("’", "'")
    return text.strip().lower()


def _compact_value(text: str) -> str:
    return _normalize_for_compare(text).strip("。.;；")


def is_emptyish(text: str) -> bool:
    """判断文本是否为空或使用通用空项占位。"""
    return _compact_value(text) in {_compact_value(item) for item in EMPTY_VALUES}


def is_delete_keyword(text: str) -> bool:
    """判断修订后文本是否明确表示删除。"""
    return _compact_value(text) in {_compact_value(item) for item in DELETE_VALUES}


def is_add_keyword(text: str) -> bool:
    """判断修订前文本是否明确表示新增条款。"""
    return _compact_value(text) in {_compact_value(item) for item in ADD_VALUES}


def classify_revision(before: str, after: str) -> str:
    """按前后正文判定新增、删除、修改或无变化。"""
    before_empty = is_emptyish(before)
    before_add = is_add_keyword(before)
    after_empty = is_emptyish(after)
    after_delete = is_delete_keyword(after)
    if (before_empty or before_add) and not after_empty and not after_delete:
        return "add"
    if not before_empty and (after_empty or after_delete):
        return "delete"
    if _normalize_for_compare(before) == _normalize_for_compare(after):
        return "no_change"
    if _leading_chapter(before) and _leading_chapter(after) and _leading_chapter(before) != _leading_chapter(after):
        return "chapter_change"
    return "modify"


def _block_summary(blocks: list[ContentBlock]) -> str:
    table_count = sum(1 for block in blocks if block.kind == "table")
    image_count = sum(1 for block in blocks if block.kind == "image")
    parts = []
    if table_count:
        parts.append(f"{table_count}个内嵌表格")
    if image_count:
        parts.append(f"{image_count}张图片")
    return "、".join(parts)


def _content_empty(text: str, blocks: list[ContentBlock]) -> bool:
    return is_emptyish(text) and not blocks


def _blocks_signature(blocks: list[ContentBlock]) -> tuple[tuple[str, str], ...]:
    return tuple((block.kind, block.digest or repr(block.rows) or block.filename) for block in blocks)


def classify_revision_row(row: RevisionRow) -> str:
    """合并正文及内容块指纹判定修订类型。"""
    before_empty = _content_empty(row.before_text, row.before_blocks)
    before_add = is_add_keyword(row.before_text) and not row.before_blocks
    after_empty = _content_empty(row.after_text, row.after_blocks)
    after_delete = is_delete_keyword(row.after_text) and not row.after_blocks

    if (before_empty or before_add) and not after_empty and not after_delete:
        return "add"
    if not before_empty and (after_empty or after_delete):
        return "delete"
    if _normalize_for_compare(row.before_text) == _normalize_for_compare(row.after_text) and _blocks_signature(row.before_blocks) == _blocks_signature(row.after_blocks):
        return "no_change"
    if _leading_chapter(row.before_text) and _leading_chapter(row.after_text) and _leading_chapter(row.before_text) != _leading_chapter(row.after_text):
        return "chapter_change"
    return "modify"


def _leading_chapter(text: str) -> str:
    match = re.match(r"\s*(\d+(?:\.\d+)+)", text or "")
    return match.group(1) if match else ""


def _tokenize(text: str) -> list[str]:
    # Character-level tokens are deliberate for Chinese manual text: they avoid
    # marking an entire long sentence when only a date, number, or short phrase changed.
    return list(text or "")


def _run(text: str, role: str = "normal") -> DiffRun:
    if not text:
        return DiffRun(text="")
    if role == "normal":
        return DiffRun(text=text)
    return DiffRun(text=text, role=role, bold=True, color=RED)


def _merge_runs(runs: list[DiffRun]) -> list[DiffRun]:
    merged: list[DiffRun] = []
    for run in runs:
        if not run.text:
            continue
        if merged and merged[-1].role == run.role and merged[-1].bold == run.bold and merged[-1].color == run.color:
            merged[-1].text += run.text
        else:
            merged.append(run)
    return merged or [DiffRun(text="")]


def _diff_runs(before: str, after: str, policy: str) -> tuple[list[DiffRun], list[DiffRun]]:
    before_tokens = _tokenize(before)
    after_tokens = _tokenize(after)
    matcher = SequenceMatcher(None, before_tokens, after_tokens, autojunk=False)
    before_runs: list[DiffRun] = []
    after_runs: list[DiffRun] = []

    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        before_text = "".join(before_tokens[i1:i2])
        after_text = "".join(after_tokens[j1:j2])
        if tag == "equal":
            before_runs.append(_run(before_text))
            after_runs.append(_run(after_text))
        elif tag == "delete":
            before_runs.append(_run(before_text, "strike" if policy == "semantic_track" else "underline"))
        elif tag == "insert":
            after_runs.append(_run(after_text, "underline"))
        elif tag == "replace":
            before_runs.append(_run(before_text, "strike" if policy == "semantic_track" else "underline"))
            after_runs.append(_run(after_text, "underline"))

    return _merge_runs(before_runs), _merge_runs(after_runs)


def build_rendered_row(row: RevisionRow, policy: str = "semantic_track") -> RenderedRevisionRow:
    """生成可追踪的差异片段并保留表格图像内容块。"""
    revision_type = row.revision_type if row.revision_type and row.revision_type != "need_review" else classify_revision_row(row)
    warnings = list(row.warnings)
    rich_summary = _block_summary([*row.before_blocks, *row.after_blocks])
    if rich_summary:
        warnings.append(f"包含{rich_summary}，已保留到导出文件；表格和图片内部不做逐字差异标注。")

    if revision_type == "add":
        before_text = row.before_text if is_add_keyword(row.before_text) else "无。"
        before_runs = [_run(before_text, "underline") if is_add_keyword(before_text) else _run(before_text)]
        after_runs = [_run(row.after_text or _block_summary(row.after_blocks) or "新增内容", "underline")]
    elif revision_type == "delete":
        compare_after = row.after_text if row.after_text.strip() else "删除"
        before_runs, after_runs = _diff_runs(row.before_text or _block_summary(row.before_blocks) or "删除内容", compare_after, "semantic_track")
        if is_delete_keyword(compare_after):
            after_runs = [_run(compare_after, "underline")]
    elif revision_type == "no_change":
        before_runs = [_run(row.before_text)]
        after_runs = [_run(row.after_text)]
        warnings.append("修订前后内容一致，已按普通文本输出。")
    else:
        before_runs, after_runs = _diff_runs(row.before_text, row.after_text, policy)

    chapter_runs = [_run(row.chapter)]
    if revision_type == "chapter_change":
        warnings.append("检测到文本开头章节号变化，请人工确认涉及章节列是否同步。")

    return RenderedRevisionRow(
        index=row.index,
        before_runs=before_runs,
        after_runs=after_runs,
        chapter_runs=chapter_runs,
        department=row.department,
        basis=row.basis,
        remark=row.remark,
        revision_type=revision_type,
        warnings=warnings,
        before_blocks=row.before_blocks,
        after_blocks=row.after_blocks,
    )
