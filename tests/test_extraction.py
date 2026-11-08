"""作者：Wangyuan；日期：2026-10-29。职责：合成PDF文字层与风险汇总回归。"""
import tempfile
import unittest
from pathlib import Path
from pdf_text import extract_pdf
from workbench import build_dimension_summary


def synthetic_pdf(text: str = 'Synthetic sample') -> bytes:
    """在内存构造简单单页PDF，不使用真实文档或附件。"""
    content = f'BT /F1 12 Tf 20 100 Td ({text}) Tj ET'.encode()
    objects = [b'<< /Type /Catalog /Pages 2 0 R >>',
               b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
               b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>',
               b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
               b'<< /Length '+str(len(content)).encode()+b' >>\nstream\n'+content+b'\nendstream']
    output = b'%PDF-1.4\n'; offsets = []
    for index, obj in enumerate(objects, 1):
        offsets.append(len(output)); output += f'{index} 0 obj\n'.encode()+obj+b'\nendobj\n'
    start = len(output)
    output += b'xref\n0 6\n0000000000 65535 f \n'
    output += b''.join(f'{offset:010} 00000 n \n'.encode() for offset in offsets)
    return output + f'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{start}\n%%EOF'.encode()


class ExtractionTests(unittest.TestCase):
    def test_pdf_text_and_budget(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory)/'sample.pdf'; path.write_bytes(synthetic_pdf())
            result = extract_pdf(path)
            self.assertEqual(result['pages'], [{'page':1,'text':'Synthetic sample'}])
            self.assertEqual(result['empty_text_pages'], [])
            self.assertFalse(result['ocr_performed'])
            with self.assertRaises(ValueError): extract_pdf(path, max_chars=2)
            with self.assertRaises(ValueError): extract_pdf(path, max_bytes=10)
            path.write_bytes(synthetic_pdf(''))
            self.assertEqual(extract_pdf(path)['empty_text_pages'], [1])

    def test_dimension_and_severity_are_explicit(self):
        rows = [{'dimension':'sample','severity':'low','title':'合成低级别'},
                {'dimension':'sample','severity':'high','title':'合成高级别'}]
        result = build_dimension_summary(rows, {'sample':'示例','empty':'空组'}, ['low','high'])
        self.assertEqual(result[0]['severity'],'high')
        self.assertEqual(result[0]['count'],2)
        self.assertIsNone(result[1]['severity'])
        with self.assertRaises(ValueError): build_dimension_summary(rows, {'sample':'示例'}, ['low'])


if __name__ == '__main__': unittest.main()
