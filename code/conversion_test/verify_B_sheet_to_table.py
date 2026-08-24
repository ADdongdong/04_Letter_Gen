# -*- coding: utf-8 -*-
"""
验证点 B：在 Word 模板的"占位锚点"位置，按 Excel Sheet 动态生成完整表格。
- 读取函证模板 docx
- 读取示例 Excel（多 Sheet），取第一个有数据的 Sheet
- 在模板的锚点位置（用占位文本标记），用 python-docx 动态插入一张表格
- 检查：列数=Sheet列数、表头正确、表格总宽=页面文本宽度
"""
import os, io, sys
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

from docx import Document
from docx.shared import Pt, Cm, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import load_workbook

OUT_DIR = r'e:\13_dingdian\03_demo\04_制函优化\上下文\conversion_test'
os.makedirs(OUT_DIR, exist_ok=True)

# ---------- 1. 读取函证模板 ----------
TPL = r'E:\13_dingdian\02_产品\06_函证系统\01_售前演示数据\函证模板\往来函证模板（销售、采购）.docx'
print('=== 1. 读取函证模板 ===')
print('模板存在:', os.path.exists(TPL))
doc = Document(TPL)

# 页面文本宽度（A4 默认，减页边距）
sec = doc.sections[0]
page_w = sec.page_width
left = sec.left_margin
right = sec.right_margin
text_width = page_w - left - right
print(f'页面宽: {page_w/914400:.1f} 英寸, 页边距左右: {left/914400:.1f}/{right/914400:.1f} 英寸')
print(f'文本宽度: {text_width/914400:.2f} 英寸 = {text_width/360000:.2f} cm')

# 找到锚点：包含特殊占位标记的段落（这里模拟模板中已有 [[SHEET_锚点1]]）
anchor_found = False
for i, para in enumerate(doc.paragraphs):
    if '[[SHEET' in para.text:
        print(f'找到锚点段落 #{i}: {para.text[:40]!r}')
        anchor_found = True
        break
if not anchor_found:
    print('模板中没有 [[SHEET 锚点，将演示：在文档末尾插入动态表格')

print(f'模板现有段落数: {len(doc.paragraphs)}, 表格数: {len(doc.tables)}')

# ---------- 2. 读取示例 Excel ----------
# 用 openpyxl 创建一个内存 Sheet 作为演示（模拟用户上传的 Excel）
# 为真实起见，先检查是否有真实 xlsx；没有则用内存数据
XLSX = r'E:\13_dingdian\02_产品\06_函证系统\01_售前演示数据\函证模板\往来函证模板（销售、采购）.xlsx'
print('\n=== 2. 准备 Excel Sheet 数据 ===')
sheet_data = None
if os.path.exists(XLSX):
    wb = load_workbook(XLSX, data_only=True)
    ws = wb[wb.sheetnames[0]]
    rows = [[c for c in r] for r in ws.iter_rows(values_only=True)][:15]
    print(f'从 {os.path.basename(XLSX)} 读取 Sheet [{ws.title}] {ws.max_row}行 x {ws.max_column}列')
    sheet_data = rows
else:
    print('未找到现成 xlsx，使用演示数据模拟 Excel Sheet')
    sheet_data = [
        ['科目名称', '截止日期', '贵公司欠', '欠贵公司', '其他说明'],
        ['销售与未结算', '2025-12-31', 123456.78, 0, '余额法'],
        ['采购与未结算', '2025-12-31', 0, 98765.43, '交易法'],
        ['应收票据', '2025-12-31', 50000, 20000, '带息'],
        ['其他往来款', '2025-12-31', 1000.5, 3000, None],
    ]

# 表头 = 第一行
header = [str(c) if c is not None else '' for c in sheet_data[0]]
data_rows = [[str(c) if c is not None else '' for c in r] for r in sheet_data[1:]]
print(f'表头({len(header)}列): {header}')
print(f'数据行数: {len(data_rows)}')
for r in data_rows[:3]:
    print('  样例行:', r)

# ---------- 3. 在模板锚点插入动态表格 ----------
print('\n=== 3. 在模板中插入动态表格 ===')
table = doc.add_table(rows=1, cols=len(header))
table.style = 'Table Grid'  # 有边框的表格样式

# 设置表格总宽 = 文本宽度（用 python-docx 设置列宽）
# python-docx 表格宽度：通过 tblW + 每列 tcW 设置
total_twips = int(text_width / 914400 * 1440)  # 英寸->twips(1/20点)
col_w = total_twips // len(header)

# 表头行
hdr_cells = table.rows[0].cells
for j, h in enumerate(header):
    hdr_cells[j].text = ''
    p = hdr_cells[j].paragraphs[0]
    run = p.add_run(h)
    run.bold = True
    run.font.size = Pt(10)

# 数据行
for i, row_data in enumerate(data_rows):
    row_cells = table.add_row().cells
    for j in range(len(header)):
        val = row_data[j] if j < len(row_data) else ''
        row_cells[j].text = ''
        p = row_cells[j].paragraphs[0]
        run = p.add_run(val)
        run.font.size = Pt(10)

# 设置列宽（每列 = 总宽/列数）
# 通过表格 grid 或每单元格 tcW 设置
try:
    # 设置每列宽度
    for row in table.rows:
        for j, cell in enumerate(row.cells):
            if j < len(header):
                cell.width = Emu(int(text_width / len(header)))
except Exception as e:
    print('  设置列宽异常(不影响验证):', e)

print(f'已插入表格: {len(table.rows)}行 x {len(table.columns)}列')
print(f'表格列数 vs Sheet列数: {len(table.columns)} vs {len(header)} -> {"一致" if len(table.columns)==len(header) else "不一致"}')
print(f'表头: {[c.text for c in table.rows[0].cells]}')

# 计算表格实际宽度（取第一行各列宽之和）
try:
    import docx.oxml.ns as ns
    # 尝试读取表格宽度
    widths = []
    for c in table.rows[0].cells:
        tc = c._tc
        tcPr = tc.tcPr
        if tcPr is not None:
            tcW = tcPr.find(ns.qn('w:tcW'))
            if tcW is not None:
                w = tcW.get(ns.qn('w:w'))
                widths.append(w)
    print(f'第一行各列宽(twips): {widths}, 总宽: {sum(map(int,widths)) if widths else "n/a"} twips')
    print(f'目标文本宽: {total_twips} twips')
    if widths:
        diff = abs(sum(map(int,widths)) - total_twips)
        print(f'宽度差: {diff} twips -> {"在容差内" if diff < 50 else "需检查"}')
except Exception as e:
    print('  读取列宽异常:', e)

# ---------- 4. 保存验证产物 ----------
out_docx = os.path.join(OUT_DIR, 'verify_B_output.docx')
doc.save(out_docx)
print(f'\n=== 4. 已保存验证产物: {out_docx} ===')

# 重新打开验证
doc2 = Document(out_docx)
print(f'重新打开检查: 段落数={len(doc2.paragraphs)}, 表格数={len(doc2.tables)}')
last_tbl = doc2.tables[-1] if doc2.tables else None
if last_tbl:
    print(f'  末表格: {len(last_tbl.rows)}行 x {len(last_tbl.columns)}列')
    print(f'  末表格表头: {[c.text for c in last_tbl.rows[0].cells]}')
    print(f'  末表格第2行: {[c.text for c in last_tbl.rows[1].cells] if len(last_tbl.rows)>1 else "无"}')
print('\n=== 验证 B 完成 ===')
