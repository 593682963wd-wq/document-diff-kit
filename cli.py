"""作者：Wangyuan；日期：2026-10-29；职责：独立离线JSON命令行入口。"""
from __future__ import annotations
import argparse
import json
import sys
import tempfile
from pathlib import Path
from typing import Any
from workbench import revision, role_clauses, read_paragraphs, analyze_risks, export_document, build_dimension_summary
from pdf_text import extract_pdf
from section_diff import compare_group, diff_files, DiffRules
from dataclasses import asdict
def run(payload: dict, output: Path) -> Any:
    """运行文档差异、角色条款筛选或风险证据表。"""
    action = payload["action"]; args = payload.get("args", {})
    if action == "revision": return revision(output_path=output / "revision.docx", **args)
    if action == "clauses": return role_clauses(**args)
    if action == "paragraphs": return read_paragraphs(**args)
    if action == "risks": return analyze_risks(**args)
    if action == "risk_summary": return build_dimension_summary(**args)
    if action == "pdf": return extract_pdf(**args)
    if action == "pdf_group_diff": return compare_group(**args)
    if action == "pdf_pair_diff":
        params=dict(args); params["rules"]=DiffRules(**params.get("rules",{}))
        return asdict(diff_files(**params))
    if action == "diff_excel":
        from section_export import export_excel
        path=output/"section-diff.xlsx";export_excel(path,**args)
        return {"output":str(path),"state":"review_required"}
    if action == "export": return export_document(output_path=output / "analysis.docx", **args)
    raise ValueError("未知动作")

def demo(output: Path) -> dict:
    """从零创建修订表，验证文本差异与嵌套表/附录保留。"""
    from docx import Document
    source = output / "synthetic-input.docx"; doc = Document()
    doc.add_heading("《示例文档》R01-01修订差异记录表", 0)
    table = doc.add_table(rows=2, cols=7)
    for cell, text in zip(table.rows[0].cells, ["序号", "修订前内容", "修订后内容", "涉及章节", "涉及部门", "依据", "备注"]): cell.text = text
    for cell, text in zip(table.rows[1].cells, ["1", "资料组每周核对一次。", "资料组每周核对两次。", "1.1", "资料组", "示例变更", ""]): cell.text = text
    nested = table.rows[1].cells[2].add_table(rows=1, cols=2); nested.cell(0,0).text = "字段"; nested.cell(0,1).text = "合成值"
    doc.add_paragraph("附录：此处为合成材料。"); doc.save(source)
    diff = revision(source, output / "revision.docx")
    clauses = role_clauses([{"id": "p1", "text": "资料组核对清单。", "heading_level": None}], ["资料组"], {"资料组": "审阅角色"})
    risks = analyze_risks("示例材料缺少检查记录。", [{"id": "rule-1", "title": "记录缺口候选", "keywords": ["缺少"], "control_draft": "人工核对记录"}])
    export = export_document("示例分析", clauses["clauses"], risks, output / "analysis.docx")
    return {"revision": diff, "clauses": clauses, "risks": risks, "export": export}

def main() -> int:
    """运行合成演示或显式输入；错误返回非零，不访问外部服务。"""
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=["demo", "run"])
    parser.add_argument("input", nargs="?")
    parser.add_argument("--output", help="显式输出目录；默认临时目录")
    args = parser.parse_args()
    try:
        with tempfile.TemporaryDirectory() as temporary:
            output = Path(args.output or temporary)
            output.mkdir(parents=True, exist_ok=True)
            if args.command == "demo":
                result = demo(output)
            else:
                if not args.input:
                    raise ValueError("run必须提供JSON文件")
                source = Path(args.input)
                if source.stat().st_size > 10_000_000:
                    raise ValueError("JSON文件超出预算")
                result = run(json.loads(source.read_text(encoding="utf-8")), output)
            print(json.dumps(result, ensure_ascii=False, indent=2, allow_nan=False))
        return 0
    except (ValueError, TypeError, KeyError, OSError) as exc:
        print(json.dumps({"error": type(exc).__name__, "message": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 2
if __name__ == "__main__":
    raise SystemExit(main())
