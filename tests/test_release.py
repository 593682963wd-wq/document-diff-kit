"""作者：Wangyuan；日期：2026-11-08。职责：通过合成输入复验已确认缺陷。"""
import unittest,tempfile,os
from pathlib import Path
from docx import Document
from docx.shared import Inches
import base64
PNG=base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jXioAAAAASUVORK5CYII=")
from workbench import revision
from manual_revision_parser import parse_revision_docx
class ReleaseTests(unittest.TestCase):
 def source(self,p,before='',after='new'):
  d=Document();t=d.add_table(rows=2,cols=2);t.cell(0,0).text='修订前内容';t.cell(0,1).text='修订后内容';t.cell(1,0).text=before;t.cell(1,1).text=after;d.save(p);return d
 def test_hardlink_never_overwrites_source(self) -> None:
  with tempfile.TemporaryDirectory() as temp:
   p=Path(temp)/'source.docx';q=Path(temp)/'out.docx';self.source(p);os.link(p,q);before=p.read_bytes()
   with self.assertRaises(ValueError):revision(p,q)
   self.assertEqual(p.read_bytes(),before)
 def test_manual_edit_reclassifies_add_to_modify(self) -> None:
  with tempfile.TemporaryDirectory() as temp:
   p=Path(temp)/'source.docx';self.source(p)
   r=revision(p,Path(temp)/'out.docx',edits={'1':{'before_text':'old'}});self.assertEqual(r['types'],['modify'])
 def test_repeated_image_occurrences_keep_each_size(self) -> None:
  with tempfile.TemporaryDirectory() as temp:
   root=Path(temp);p=root/'source.docx';d=self.source(p);image=root/'sample.png';image.write_bytes(PNG)
   run=d.tables[0].cell(1,1).paragraphs[0].add_run();run.add_picture(str(image),width=Inches(1));run.add_picture(str(image),width=Inches(2));d.save(p)
   blocks=parse_revision_docx(p,asset_dir=root/'assets').rows[0].after_blocks
   self.assertEqual(len(blocks),2);self.assertAlmostEqual(blocks[1].width_mm,blocks[0].width_mm*2,places=1)
 def test_deeper_nested_content_explicitly_rejected(self) -> None:
  with tempfile.TemporaryDirectory() as temp:
   p=Path(temp)/'source.docx';d=self.source(p);t=d.tables[0].cell(1,1).add_table(rows=1,cols=1);t.cell(0,0).add_table(rows=1,cols=1).cell(0,0).text='deep';d.save(p)
   with self.assertRaisesRegex(ValueError,'更深表格'):parse_revision_docx(p)
 def test_appendix_image_only_explicitly_rejected(self) -> None:
  with tempfile.TemporaryDirectory() as temp:
   root=Path(temp);p=root/'source.docx';d=self.source(p);image=root/'sample.png';image.write_bytes(PNG);d.add_picture(str(image));d.save(p)
   with self.assertRaisesRegex(ValueError,'附录'):parse_revision_docx(p)
