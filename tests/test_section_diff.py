"""作者：Wangyuan；日期：2026-10-29。职责：章节归属、合成PDF、质量缺口及Excel安全回归。"""
import tempfile
import unittest
from pathlib import Path
from openpyxl import load_workbook
from section_diff import diff_lines, diff_files, compare_group, DiffRules
from section_export import export_excel
from test_extraction import synthetic_pdf


class SectionTests(unittest.TestCase):
    def test_deletion_uses_old_section_and_empty_new_is_valid_text_diff(self):
        result=diff_lines(['## Before','字段名称','1 old'], [])
        self.assertIn('Before',result.sections)
        self.assertEqual(result.hunks[0]['kind'],'delete')
        result=diff_lines(['## A','字段名称','1 old','## B','另一个字段','2 same'],['## A','字段名称','## B','另一个字段','2 same'])
        self.assertEqual(result.hunks[0]['section'],'A')
        self.assertEqual(result.hunks[0]['field'],'字段名称')
        result=diff_lines(['## A','old field','## B','removed field'],['## A'])
        self.assertEqual(set(result.sections),{'A','B'})
        self.assertEqual(result.hunks[0]['affected_sections'],['A','B'])
        result=diff_lines(['## A','旧的字段'],['## A','旧的字段','## B','新字段'])
        self.assertEqual(result.sections['B'],[])
        result=diff_lines(['## A','甲字段','1 oldA','## B','1 before'],
                          ['## A','甲字段','1 oldA','## B','1 after'])
        self.assertEqual(result.sections,{'B':[]})
        self.assertEqual(result.hunks[0]['field'],'')

    def test_explicit_noise_and_labels(self):
        rules=DiffRules(label_pattern=r'[A-Za-z ]+',noise_patterns=[r'^Noise'])
        result=diff_lines(['## Example','Field label','1 before','Noise old'],['## Example','Field label','1 after','Noise new'],rules)
        self.assertEqual(len(result.hunks),1)
        self.assertEqual(result.hunks[0]['field'],'Field label')

    def test_pdf_pair_and_quality_errors(self):
        with tempfile.TemporaryDirectory() as directory:
            old=Path(directory)/'old';new=Path(directory)/'new';old.mkdir();new.mkdir()
            (old/'sample.pdf').write_bytes(synthetic_pdf('Before'))
            (new/'sample.pdf').write_bytes(synthetic_pdf('After'))
            self.assertTrue(diff_files(old/'sample.pdf',new/'sample.pdf').hunks)
            (new/'extra.pdf').write_bytes(synthetic_pdf('Added'))
            result=compare_group('sample',old,new)
            self.assertEqual(result['only_in_new'],['extra.pdf'])
            (new/'sample.pdf').write_bytes(synthetic_pdf(''))
            result=compare_group('sample',old,new)
            self.assertEqual(result['state'],'incomplete')
            self.assertTrue(result['errors'])

    def test_export_literal_and_no_overwrite(self):
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'sample.xlsx'
            export_excel(path,'示例',[('=1+1','=2+2','=3+3')])
            workbook=load_workbook(path,data_only=False)
            self.assertEqual(workbook.active['A7'].value,'=1+1')
            self.assertEqual(workbook.active['A7'].data_type,'s');workbook.close()
            with self.assertRaises(FileExistsError):export_excel(path,'示例',[])


if __name__=='__main__':unittest.main()
