"""作者：Wangyuan；日期：2026-10-29；职责：仅以从零合成输入验证核心与异常路径。"""

import tempfile, unittest, zipfile
from pathlib import Path
from docx import Document
from workbench import *
from manual_revision_diff import build_rendered_row
from manual_revision_models import RevisionRow, ContentBlock
class DocumentTests(unittest.TestCase):
    def test_diff_insert_delete_and_nochange(self):
        self.assertEqual(build_rendered_row(RevisionRow(index=1,before_text='',after_text='新增条款')).revision_type,'add')
        self.assertEqual(build_rendered_row(RevisionRow(index=2,before_text='旧条款',after_text='删除')).revision_type,'delete')
        self.assertEqual(build_rendered_row(RevisionRow(index=3,before_text='相同',after_text='相同')).revision_type,'no_change')
    def test_revision_nested_table_appendix_and_validation(self):
        from cli import demo
        with tempfile.TemporaryDirectory() as tmp:
            result=demo(Path(tmp)); self.assertTrue(result['revision']['validation']['ok']); self.assertEqual(result['revision']['rows'],1)
            doc=Document(Path(tmp)/'revision.docx')
            self.assertTrue(any(cell.tables for row in doc.tables[0].rows for cell in row.cells))
            self.assertIn('合成材料','\n'.join(p.text for p in doc.paragraphs))
    def test_revision_manual_edits_and_input_unchanged(self):
        from cli import demo
        with tempfile.TemporaryDirectory() as tmp:
            root=Path(tmp); demo(root); source=root/'synthetic-input.docx'; original=source.read_bytes()
            result=revision(source,root/'edited.docx',edits={'1':{'after_text':'人工修改后的条款'}})
            self.assertTrue(result['validation']['ok']); self.assertEqual(source.read_bytes(),original)
            with self.assertRaises(ValueError): revision(source,source)
    def test_zip_external_relationship_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'external.docx'
            with zipfile.ZipFile(path,'w') as z:
                z.writestr('_rels/.rels','<Relationships><Relationship TargetMode="External" Target="https://example.invalid"/></Relationships>')
                z.writestr('word/document.xml','<x/>')
            with self.assertRaises(ValueError): inspect_docx(path)
    def test_role_clauses_preserve_evidence_and_headings(self):
        result=role_clauses([{'id':'p1','text':'记录核对','heading_level':2},{'id':'p2','text':'资料组应核对记录。'}],['资料组'],{'资料组':'审阅角色'})
        row=result['clauses'][0]; self.assertEqual(row['heading'],'记录核对'); self.assertEqual(row['text'],'审阅角色应核对记录。'); self.assertIn('资料组',row['original']); self.assertEqual(row['state'],'needs_review')
    def test_risk_candidate_has_exact_anchor(self):
        text='第一句。记录缺少日期。最后一句。'; rows=analyze_risks(text,[{'id':'r','title':'缺项','keywords':['缺少']}]); self.assertEqual(len(rows),1)
        quote=rows[0]['evidence'][0]; self.assertEqual(text[quote['start']:quote['end']],quote['quote']); self.assertEqual(rows[0]['state'],'candidate')
        self.assertEqual(analyze_risks('没有相关字样。',[{'id':'r','title':'缺项','keywords':['缺少']}]),[])
    def test_editable_export_roundtrip(self):
        with tempfile.TemporaryDirectory() as tmp:
            path=Path(tmp)/'report.docx'; result=export_document('示例分析',[],[{'title':'候选','evidence':[{'quote':'原句'}],'control_draft':'核对','human_comment':'人工备注'}],path)
            self.assertEqual(result['risk_rows'],1); self.assertEqual(Document(path).tables[0].rows[1].cells[3].text,'人工备注')

if __name__ == "__main__":
    unittest.main()
