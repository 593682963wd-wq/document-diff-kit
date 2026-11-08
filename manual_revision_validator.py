"""作者：Wangyuan；日期：2026-10-29；职责：通用文档内容树、差异及渲染提取模块。"""
from __future__ import annotations

from pathlib import Path

from docx import Document

from manual_revision_models import ValidationIssue, ValidationResult


STANDARD_HEADERS = ["序号", "修订前内容", "修订后内容", "章节名称", "涉及部门（科室）", "依据", "备注"]
LEGACY_HEADERS = {"涉及章节"}


def _run_color(run) -> str | None:
    if run.font.color and run.font.color.rgb:
        return str(run.font.color.rgb)
    return None


def _is_red(run) -> bool:
    return (_run_color(run) or "").upper() == "FF0000"


def validate_output_docx(path: Path) -> ValidationResult:
    """重新打开输出文档，检查列头及差异标记，返回具体问题。"""
    issues: list[ValidationIssue] = []
    red_run_count = 0
    underline_count = 0
    strike_count = 0

    try:
        doc = Document(path)
    except Exception as exc:  # noqa: BLE001
        return ValidationResult(
            ok=False,
            issues=[ValidationIssue("error", "DOCX_OPEN_FAILED", f"输出文件无法反读：{exc}")],
        )

    if not doc.tables:
        issues.append(ValidationIssue("error", "NO_TABLE", "输出文档没有差异表。"))
        return ValidationResult(ok=False, issues=issues)

    table = doc.tables[0]
    header_row_index = None
    for idx, row in enumerate(table.rows[:5]):
        values = [cell.text.strip().replace("\n", "") for cell in row.cells]
        if "修订前内容" in values and "修订后内容" in values:
            header_row_index = idx
            for expected in STANDARD_HEADERS:
                if expected not in values:
                    issues.append(ValidationIssue("error", "HEADER_MISSING", f"缺少标准表头：{expected}", row_index=idx + 1))
            for legacy in LEGACY_HEADERS:
                if legacy in values:
                    issues.append(ValidationIssue("error", "LEGACY_HEADER", f"输出仍包含旧表头：{legacy}", row_index=idx + 1))
            break
    if header_row_index is None:
        issues.append(ValidationIssue("error", "HEADER_NOT_FOUND", "未找到标准表头行。"))

    for row_index, row in enumerate(table.rows, start=1):
        for col_index, cell in enumerate(row.cells, start=1):
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    if not run.text or not _is_red(run):
                        continue
                    red_run_count += 1
                    underlined = bool(run.underline or run.font.underline)
                    struck = bool(run.font.strike)
                    if underlined:
                        underline_count += 1
                    if struck:
                        strike_count += 1
                    if not bool(run.bold or run.font.bold):
                        issues.append(ValidationIssue("error", "RED_NOT_BOLD", "红色差异片段未加粗。", row_index, str(col_index)))
                    if not underlined and not struck:
                        issues.append(ValidationIssue("error", "RED_WITHOUT_MARK", "红色差异片段缺少下划线或删除线。", row_index, str(col_index)))
                    if underlined and struck:
                        issues.append(ValidationIssue("error", "DOUBLE_MARK", "同一差异片段同时使用下划线和删除线。", row_index, str(col_index)))

    if red_run_count == 0:
        issues.append(ValidationIssue("warning", "NO_RED_RUN", "输出未发现红色差异标注，请确认输入是否确实有变化。"))

    ok = not any(issue.severity == "error" for issue in issues)
    return ValidationResult(
        ok=ok,
        issues=issues,
        red_run_count=red_run_count,
        underline_count=underline_count,
        strike_count=strike_count,
    )
