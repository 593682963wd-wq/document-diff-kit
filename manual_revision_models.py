"""作者：Wangyuan；日期：2026-10-29；职责：通用文档内容树、差异及渲染提取模块。"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Literal


RevisionType = Literal["add", "delete", "modify", "chapter_change", "no_change", "need_review"]
DiffRole = Literal["normal", "underline", "strike"]
ContentBlockKind = Literal["table", "image"]
AppendixBlockKind = Literal["paragraph", "table"]


@dataclass
class ContentBlock:
    """表示嵌套表或图像内容块，保留可编辑内容及显式媒体引用。"""
    kind: str
    rows: list[list[str]] = field(default_factory=list)
    column_widths_mm: list[float] = field(default_factory=list)
    image_path: str = ""
    filename: str = ""
    content_type: str = ""
    digest: str = ""
    data_url: str = ""
    width_mm: float | None = None
    height_mm: float | None = None
    drawing_xml: str = ""

    def to_dict(self) -> dict[str, Any]:
        """导出当前对象的数据字段。"""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any] | None) -> "ContentBlock":
        """从显式字典构建对象并规范化内容字段。"""
        data = data or {}
        rows = data.get("rows") or []
        normalized_rows = [[str(cell or "") for cell in row] for row in rows if isinstance(row, list)]
        widths = []
        for value in data.get("column_widths_mm") or []:
            try:
                widths.append(float(value))
            except (TypeError, ValueError):
                continue

        def optional_float(key: str) -> float | None:
            """把可选尺寸转换为浮点数；无效值保留缺失。"""
            try:
                value = data.get(key)
                return float(value) if value not in (None, "") else None
            except (TypeError, ValueError):
                return None

        return cls(
            kind=str(data.get("kind") or ""),
            rows=normalized_rows,
            column_widths_mm=widths,
            image_path=str(data.get("image_path") or ""),
            filename=str(data.get("filename") or ""),
            content_type=str(data.get("content_type") or ""),
            digest=str(data.get("digest") or ""),
            data_url=str(data.get("data_url") or ""),
            width_mm=optional_float("width_mm"),
            height_mm=optional_float("height_mm"),
            drawing_xml=str(data.get("drawing_xml") or ""),
        )


@dataclass
class AppendixBlock:
    """表示修订表后的正文或表格附录。"""
    kind: str
    text: str = ""
    rows: list[list[str]] = field(default_factory=list)
    column_widths_mm: list[float] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        """导出当前对象的数据字段。"""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any] | None) -> "AppendixBlock":
        """从显式字典构建对象并规范化内容字段。"""
        data = data or {}
        rows = data.get("rows") or []
        widths = []
        for value in data.get("column_widths_mm") or []:
            try:
                widths.append(float(value))
            except (TypeError, ValueError):
                continue
        return cls(
            kind=str(data.get("kind") or ""),
            text=str(data.get("text") or ""),
            rows=[[str(cell or "") for cell in row] for row in rows if isinstance(row, list)],
            column_widths_mm=widths,
        )


@dataclass
class RevisionMetadata:
    """保存调用方文档元数据，不加载外部默认配置。"""
    title: str = ""
    manual_name: str = ""
    revision_version: str = ""
    writer: str = ""
    reviewer: str = ""
    revision_date: str = ""
    source_filename: str = ""

    def to_dict(self) -> dict[str, Any]:
        """导出当前对象的数据字段。"""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any] | None) -> "RevisionMetadata":
        """从显式字典构建对象并规范化内容字段。"""
        data = data or {}
        return cls(**{field_name: str(data.get(field_name) or "") for field_name in cls.__dataclass_fields__})


@dataclass
class RevisionRow:
    """保存修订前后正文、依据与结构化内容块。"""
    index: int
    before_text: str = ""
    after_text: str = ""
    chapter: str = ""
    department: str = ""
    basis: str = ""
    remark: str = ""
    revision_type: str = "need_review"
    source_row: int | None = None
    warnings: list[str] = field(default_factory=list)
    before_blocks: list[ContentBlock] = field(default_factory=list)
    after_blocks: list[ContentBlock] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        """导出当前对象的数据字段。"""
        data = asdict(self)
        data["before_blocks"] = [block.to_dict() for block in self.before_blocks]
        data["after_blocks"] = [block.to_dict() for block in self.after_blocks]
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "RevisionRow":
        """从显式字典构建对象并规范化内容字段。"""
        payload = dict(data)
        try:
            payload["index"] = int(payload.get("index") or 0)
        except (TypeError, ValueError):
            payload["index"] = 0
        payload["warnings"] = list(payload.get("warnings") or [])
        payload["before_blocks"] = [ContentBlock.from_dict(item) for item in payload.get("before_blocks", [])]
        payload["after_blocks"] = [ContentBlock.from_dict(item) for item in payload.get("after_blocks", [])]
        allowed = set(cls.__dataclass_fields__)
        return cls(**{key: value for key, value in payload.items() if key in allowed})


@dataclass
class DiffRun:
    """表示单个差异文本片段及其显示角色。"""
    text: str
    role: str = "normal"
    bold: bool = False
    color: str | None = None

    def to_dict(self) -> dict[str, Any]:
        """导出当前对象的数据字段。"""
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "DiffRun":
        """从显式字典构建对象并规范化内容字段。"""
        return cls(
            text=str(data.get("text") or ""),
            role=str(data.get("role") or "normal"),
            bold=bool(data.get("bold")),
            color=data.get("color"),
        )


@dataclass
class RenderedRevisionRow:
    """保存已分类和标记的修订行。"""
    index: int
    before_runs: list[DiffRun]
    after_runs: list[DiffRun]
    chapter_runs: list[DiffRun]
    department: str = ""
    basis: str = ""
    remark: str = ""
    revision_type: str = "need_review"
    warnings: list[str] = field(default_factory=list)
    before_blocks: list[ContentBlock] = field(default_factory=list)
    after_blocks: list[ContentBlock] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        """导出当前对象的数据字段。"""
        data = asdict(self)
        data["before_runs"] = [run.to_dict() for run in self.before_runs]
        data["after_runs"] = [run.to_dict() for run in self.after_runs]
        data["chapter_runs"] = [run.to_dict() for run in self.chapter_runs]
        data["before_blocks"] = [block.to_dict() for block in self.before_blocks]
        data["after_blocks"] = [block.to_dict() for block in self.after_blocks]
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "RenderedRevisionRow":
        """从显式字典构建对象并规范化内容字段。"""
        return cls(
            index=int(data.get("index") or 0),
            before_runs=[DiffRun.from_dict(item) for item in data.get("before_runs", [])],
            after_runs=[DiffRun.from_dict(item) for item in data.get("after_runs", [])],
            chapter_runs=[DiffRun.from_dict(item) for item in data.get("chapter_runs", [])],
            department=str(data.get("department") or ""),
            basis=str(data.get("basis") or ""),
            remark=str(data.get("remark") or ""),
            revision_type=str(data.get("revision_type") or "need_review"),
            warnings=list(data.get("warnings") or []),
            before_blocks=[ContentBlock.from_dict(item) for item in data.get("before_blocks", [])],
            after_blocks=[ContentBlock.from_dict(item) for item in data.get("after_blocks", [])],
        )


@dataclass
class ParsedRevisionDocument:
    """保存解析结果、来源索引与需要人工确认的警告。"""
    metadata: RevisionMetadata
    rows: list[RevisionRow]
    skipped_rows: int = 0
    warnings: list[str] = field(default_factory=list)
    source_path: str = ""
    converted_path: str = ""
    appendix_blocks: list[AppendixBlock] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        """导出当前对象的数据字段。"""
        return {
            "metadata": self.metadata.to_dict(),
            "rows": [row.to_dict() for row in self.rows],
            "skipped_rows": self.skipped_rows,
            "warnings": self.warnings,
            "source_path": self.source_path,
            "converted_path": self.converted_path,
            "appendix_blocks": [block.to_dict() for block in self.appendix_blocks],
        }


@dataclass
class ValidationIssue:
    """记录输出反读校验的问题及其位置。"""
    severity: str
    code: str
    message: str
    row_index: int | None = None
    column: str | None = None

    def to_dict(self) -> dict[str, Any]:
        """导出当前对象的数据字段。"""
        return asdict(self)


@dataclass
class ValidationResult:
    """记录可编辑文档结构与差异格式的检查结果。"""
    ok: bool
    issues: list[ValidationIssue] = field(default_factory=list)
    red_run_count: int = 0
    underline_count: int = 0
    strike_count: int = 0

    def to_dict(self) -> dict[str, Any]:
        """导出当前对象的数据字段。"""
        errors = [item for item in self.issues if item.severity == "error"]
        warnings = [item for item in self.issues if item.severity == "warning"]
        return {
            "ok": self.ok,
            "errors": len(errors),
            "warnings": len(warnings),
            "red_run_count": self.red_run_count,
            "underline_count": self.underline_count,
            "strike_count": self.strike_count,
            "issues": [item.to_dict() for item in self.issues],
        }


def path_to_str(path: str | Path | None) -> str:
    """将可选路径转换为可序列化字符串。"""
    return str(path) if path else ""
