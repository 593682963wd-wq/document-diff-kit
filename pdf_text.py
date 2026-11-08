"""作者：Wangyuan；日期：2026-10-29。职责：在有超时的子进程中提取PDF文字层及页码缺口。"""
from __future__ import annotations
import io
import json
import subprocess
import sys
from pathlib import Path


def extract_pdf(path: str | Path, *, max_bytes: int = 5_000_000,
                max_pages: int = 100, max_chars: int = 200_000, timeout: int = 10) -> dict:
    """仅读取显式文件；保留页码和空文字页，不将扫描页标成已提取。"""
    if any(type(value) is not int or value < 1 for value in (max_bytes, max_pages, max_chars, timeout)):
        raise ValueError('解析预算须为正整数')
    with Path(path).open('rb') as stream:
        data = stream.read(max_bytes + 1)
    if len(data) > max_bytes or not data.startswith(b'%PDF-'):
        raise ValueError('PDF超过字节预算或文件头无效')
    try:
        result = subprocess.run([sys.executable, '-B', str(Path(__file__).resolve()),
                                 str(max_pages), str(max_chars), str(timeout)],
                                input=data, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                timeout=timeout, check=False)
    except subprocess.TimeoutExpired as exc:
        raise ValueError('PDF解析超时') from exc
    if result.returncode:
        raise ValueError('PDF解析被拒绝：文件无效、加密、超过预算或缺少依赖')
    return json.loads(result.stdout)


def _parse(data: bytes, max_pages: int, max_chars: int) -> dict:
    """复用逐页文字层算法；无文字页明确留缺口，不执行OCR或活动内容。"""
    import pdfplumber
    pages, empty, count = [], [], 0
    with pdfplumber.open(io.BytesIO(data)) as document:
        for index, page in enumerate(document.pages, 1):
            if index > max_pages: raise ValueError('超过页数预算')
            text = page.extract_text(x_tolerance=1.5, y_tolerance=3) or ''
            count += len(text)
            if count > max_chars: raise ValueError('超过字符预算')
            pages.append({'page': index, 'text': text})
            if not text.strip(): empty.append(index)
            page.close()
    return {'pages': pages, 'empty_text_pages': empty, 'ocr_performed': False,
            'text': '\n\n'.join(f"===== PAGE {p['page']} =====\n{p['text']}" for p in pages),
            'state': 'review_required', 'coverage': 'text_layer_only'}


if __name__ == '__main__':
    # 子进程超时限制复杂解析；Linux另限制地址空间，macOS不虚称内存已硬隔离。
    import resource
    resource.setrlimit(resource.RLIMIT_CPU, (int(sys.argv[3]), int(sys.argv[3]) + 1))
    if sys.platform.startswith('linux'):
        resource.setrlimit(resource.RLIMIT_AS, (1_073_741_824, 1_073_741_824))
    try:
        print(json.dumps(_parse(sys.stdin.buffer.read(), int(sys.argv[1]), int(sys.argv[2])), ensure_ascii=False))
    except Exception:
        raise SystemExit(2)
