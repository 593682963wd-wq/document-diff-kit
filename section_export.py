"""作者：Wangyuan；日期：2026-10-29。职责：通用差异Excel导出，剥离全部原业务检查项。"""
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
TABLE_HEADER=('文档组','文字差异候选','人工复核意见')

def export_excel(out_path: Path, sheet_name: str, rows: list[tuple[str,str,str]],
                 before_version: str = '', after_version: str = '', maker: str = '', checker: str = '',
                 check_items: tuple = ()) -> None:
    """输出显式文本，不默认产生业务结论，不覆盖已有文件。"""
    out_path=Path(out_path)
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name or "对比"

    thin = Side(border_style="thin", color="888888")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    title_font = Font(bold=True)

    ws.append(["基本信息"])
    ws.append([f"旧版本：{before_version}"])
    ws.append([f"新版本：{after_version}"])
    ws.append([f"制作人：{maker}"])
    ws.append([f"校对人：{checker}"])
    for r in range(1, 6):
        ws.cell(row=r, column=1).font = title_font

    ws.append(list(TABLE_HEADER))
    header_row_idx = ws.max_row
    for col in range(1, 4):
        c = ws.cell(row=header_row_idx, column=col)
        c.font = title_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = border

    for code, change, conclusion in rows:
        ws.append([code, change, conclusion])
        r = ws.max_row
        for col in range(1, 4):
            cell = ws.cell(row=r, column=col)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = border

    ws.append(["补充检查事项"])
    ws.cell(row=ws.max_row, column=1).font = title_font

    for k, v in check_items:
        ws.append([k, v])
        cell = ws.cell(row=ws.max_row, column=2)
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    # 列宽
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 90
    ws.column_dimensions["C"].width = 28

    out_path.parent.mkdir(parents=True, exist_ok=True)
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value,str):cell.data_type='s'
    if out_path.suffix.lower() != '.xlsx':raise ValueError('输出须为XLSX')
    with out_path.open('xb') as stream:
        wb.save(stream)
    wb.close()
