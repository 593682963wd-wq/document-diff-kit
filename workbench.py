"""作者：Wangyuan；日期：2026-10-29；职责：文档差异、可配置职责条款和证据风险表。"""
from __future__ import annotations

import hashlib
import re
import zipfile
from dataclasses import asdict
from pathlib import Path
from typing import Any
from xml.etree import ElementTree

from docx import Document
from manual_revision_parser import parse_revision_docx
from manual_revision_diff import build_rendered_row,classify_revision_row
from manual_revision_renderer import render_revision_docx
from manual_revision_validator import validate_output_docx


def inspect_docx(path: str | Path, budget: int = 20_000_000) -> None:
    """拒绝外部关系、宏、路径穿越及超预算压缩包；不启动外部转换程序。"""
    target = Path(path)
    if target.suffix.lower() != ".docx" or target.stat().st_size > budget: raise ValueError("只接受预算内DOCX")
    with zipfile.ZipFile(target) as archive:
        entries = archive.infolist()
        if len(entries) > 3000 or sum(x.file_size for x in entries) > budget * 5: raise ValueError("压缩包展开预算超限")
        if len({x.filename for x in entries}) != len(entries): raise ValueError("重复压缩项")
        for item in entries:
            if ".." in item.filename.replace("\\", "/").split("/") or item.filename.startswith("/") or "vbaproject" in item.filename.lower(): raise ValueError("不安全压缩项")
            if item.file_size > max(item.compress_size, 1) * 200: raise ValueError("压缩比超限")
            if item.filename.endswith((".rels", ".xml")):
                payload = archive.read(item)
                if b"<!DOCTYPE" in payload.upper() or b"<!ENTITY" in payload.upper(): raise ValueError("禁止XML实体")
                if item.filename.endswith(".rels"):
                    tree = ElementTree.fromstring(payload)
                    if any(node.attrib.get("TargetMode") == "External" for node in tree): raise ValueError("禁止外部文档关系")


def revision(input_path: str | Path, output_path: str | Path, *, edits: dict[str, dict] | None = None,
             policy: str = "semantic_track") -> dict:
    """解析修订表并重建差异文档；保留嵌套表/图片/附录及明确人工修改。"""
    if policy not in {"semantic_track", "underline"}: raise ValueError("差异策略无效")
    source, target = Path(input_path), Path(output_path)
    if source.resolve() == target.resolve() or target.exists() and source.samefile(target): raise ValueError("输出不能覆盖输入")
    inspect_docx(source)
    assets = target.parent / (target.stem + "_assets"); assets.mkdir(parents=True, exist_ok=True)
    parsed = parse_revision_docx(source, asset_dir=assets)
    known = {str(x.index): x for x in parsed.rows}
    for key, changes in (edits or {}).items():
        if key not in known or set(changes) - {"before_text", "after_text", "chapter", "department", "basis", "remark"}: raise ValueError("修改引用或字段无效")
        for field, value in changes.items():
            if not isinstance(value, str): raise ValueError("修改值必须为文本")
            setattr(known[key], field, value)
        known[key].revision_type=classify_revision_row(known[key])
    rows = [build_rendered_row(row, policy) for row in parsed.rows]
    render_revision_docx(parsed.metadata, rows, target, {}, parsed.appendix_blocks)
    validation = validate_output_docx(target)
    return {"output": target.name, "source_sha256": hashlib.sha256(source.read_bytes()).hexdigest(),
            "rows": len(rows), "types": [x.revision_type for x in rows], "validation": asdict(validation),
            "warnings": parsed.warnings, "state": "candidate"}


def role_clauses(paragraphs: list[dict], aliases: list[str], replacements: dict[str, str]) -> dict:
    """按显式角色别名匹配并继承最近标题，改写保留原段落证据。"""
    if not aliases or any(not x for x in aliases) or any(not key for key in replacements): raise ValueError("别名或替换词不能为空")
    if len({x["id"] for x in paragraphs}) != len(paragraphs): raise ValueError("段落身份重复")
    result = []; heading = "工作要求"
    for paragraph in paragraphs:
        text = paragraph["text"]
        if not isinstance(text, str): raise ValueError("段落必须为文本")
        if paragraph.get("heading_level"):
            heading = text; continue
        if any(alias in text for alias in aliases):
            revised = text
            for before, after in replacements.items(): revised = revised.replace(before, after)
            result.append({"heading": heading, "text": revised, "source_id": paragraph["id"], "original": text, "state": "needs_review"})
    return {"clauses": result, "warnings": [] if result else ["未匹配到条款，请调整显式关键词或人工补充"]}


def read_paragraphs(path: str | Path) -> list[dict]:
    """读取文字及标题层级，未解析的表格不虚称已包含。"""
    inspect_docx(path); document = Document(path); result = []
    for index, paragraph in enumerate(document.paragraphs, 1):
        style = paragraph.style.name if paragraph.style else ""
        match = re.search(r"(?:Heading|标题)\s*(\d+)", style, re.I)
        if paragraph.text.strip(): result.append({"id": f"p{index}", "text": paragraph.text.strip(), "heading_level": int(match.group(1)) if match else None})
    return result


def analyze_risks(text: str, rules: list[dict]) -> list[dict]:
    """以调用方规则检索原文句证据，未匹配风险不生成事实结论。"""
    if len(text) > 200_000 or len(rules) > 1000: raise ValueError("输入超出预算")
    if len({x["id"] for x in rules}) != len(rules): raise ValueError("规则身份重复")
    sentences = [(m.start(), m.group()) for m in re.finditer(r"[^。！？\n]+[。！？]?", text)]
    output = []
    for rule in rules:
        keywords = rule.get("keywords", [])
        if not keywords or any(not isinstance(x, str) or not x for x in keywords): raise ValueError("规则关键词无效")
        evidence = [{"start": start, "end": start + len(sentence), "quote": sentence} for start, sentence in sentences if any(key in sentence for key in keywords)]
        if evidence:
            output.append({"id": rule["id"], "title": rule["title"], "dimension": rule.get("dimension"), "severity": rule.get("severity"), "evidence": evidence, "control_draft": rule.get("control_draft", ""), "human_comment": "", "state": "candidate"})
    return output


def build_dimension_summary(rows: list[dict], dimensions: dict[str, str], severity_order: list[str]) -> list[dict]:
    """按调用方维度及从低到高的级别汇总候选，不自带业务规则。"""
    if not severity_order or len(set(severity_order)) != len(severity_order):
        raise ValueError("严重度顺序须为非空且无重复的列表")
    if any(row.get('dimension') not in dimensions or row.get('severity') not in severity_order for row in rows):
        raise ValueError("候选维度或级别未配置")
    result = []
    for dimension, label in dimensions.items():
        selected = [row for row in rows if row['dimension'] == dimension]
        result.append({'id':dimension, 'label':label, 'count':len(selected), 'involved':bool(selected),
                       'severity': max((row['severity'] for row in selected), key=severity_order.index, default=None),
                       'summary': '；'.join(str(row.get('title','')) for row in selected[:2]), 'state':'candidate'})
    return result


def export_document(title: str, clauses: list[dict], risks: list[dict], output_path: str | Path) -> dict:
    """输出可编辑职责/风险草稿并反读确认行数，不载入真实模板。"""
    document = Document(); document.add_heading(title, 0); document.add_paragraph("通用配置生成的待复核草稿")
    for clause in clauses:
        document.add_heading(clause["heading"], 2); document.add_paragraph(clause["text"])
        document.add_paragraph("来源段落：" + clause["source_id"])
    table = document.add_table(rows=1, cols=4); headers = ["风险候选", "证据", "措施草稿", "人工意见"]
    for cell, value in zip(table.rows[0].cells, headers): cell.text = value
    for risk in risks:
        values = [risk["title"], "\n".join(x["quote"] for x in risk["evidence"]), risk["control_draft"], risk.get("human_comment", "")]
        for cell, value in zip(table.add_row().cells, values): cell.text = value
    document.save(output_path); reopened = Document(output_path)
    return {"output": Path(output_path).name, "risk_rows": len(reopened.tables[0].rows) - 1, "clauses": len(clauses), "state": "candidate"}
