# -*- coding: utf-8 -*-
"""
制函渲染核心引擎（验证用 Python 实现，后续可移植到 Java/poi-tl）

核心能力：在 Word 模板的"锚点位置"，按 Excel Sheet 动态生成完整表格。
- 列数 = Sheet 列数
- 表头 = Sheet 第一行
- 数据 = Sheet 其余行
- 表格总宽锁定为模板文本宽度
"""
import io, os
from docx import Document
from docx.shared import Pt, Cm, Emu
from docx.oxml.ns import qn
from openpyxl import load_workbook


def get_text_width(doc):
    """获取页面文本宽度（EMU）"""
    sec = doc.sections[0]
    return sec.page_width - sec.left_margin - sec.right_margin


def read_excel_sheets(xlsx_path):
    """读取 Excel 所有 Sheet，返回 [{name, header, rows}]"""
    wb = load_workbook(xlsx_path, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if not rows:
            continue
        # 表头 = 第一行，数据 = 其余行
        header = [str(c) if c is not None else '' for c in rows[0]]
        data = [[str(c) if c is not None else '' for c in r] for r in rows[1:]]
        # 过滤全空数据行
        data = [r for r in data if any(v.strip() for v in r)]
        sheets.append({
            'name': ws.title,
            'header': header,
            'rows': data,
            'max_row': ws.max_row,
            'max_col': ws.max_column,
        })
    return sheets


def _set_cell_text(cell, text, bold=False, size=10, align='center'):
    """填充单元格文本并设置基本样式"""
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = 1  # center
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)


def _table_width_twips(doc):
    """文本宽度转 twips（1 twip = 1/20 点）"""
    return int(get_text_width(doc) / 914400 * 1440)


def inject_sheet_table(doc, sheet, anchor_mode='end', anchor_idx=None, col_widths_ratio=None):
    """
    在 doc 中注入一个 Sheet 表格。

    参数：
        doc: python-docx Document
        sheet: read_excel_sheets 返回的单个 sheet dict
        anchor_mode: 'end'(文档末尾) / 'after_table'(指定表格后) / 'after_para'(指定段落后)
        anchor_idx: anchor_mode 对应的索引（表格序号/段落序号）
        col_widths_ratio: 可选列宽比例列表，None 则均分

    返回：注入的表格
    """
    header = sheet['header']
    rows = sheet['rows']
    ncols = len(header)

    # 创建表格（1 行表头）
    table = doc.add_table(rows=1, cols=ncols)
    table.style = 'Table Grid'

    # 计算列宽（twips），默认均分
    total_twips = _table_width_twips(doc)
    if col_widths_ratio and len(col_widths_ratio) == ncols and sum(col_widths_ratio) > 0:
        twips = [int(total_twips * r / sum(col_widths_ratio)) for r in col_widths_ratio]
    else:
        twips = [total_twips // ncols] * ncols

    # 表头行
    for j, h in enumerate(header):
        _set_cell_text(table.rows[0].cells[j], h, bold=True, size=10)

    # 数据行
    for row_data in rows:
        cells = table.add_row().cells
        for j in range(ncols):
            val = row_data[j] if j < len(row_data) else ''
            _set_cell_text(cells[j], val, bold=False, size=10)

    # 设置列宽（每单元格 tcW）
    for row in table.rows:
        for j, cell in enumerate(row.cells):
            if j < ncols:
                cell.width = Emu(int(get_text_width(doc) / ncols))
                _set_tc_width(cell, twips[j])

    # 处理锚点定位：python-docx add_table 默认加到末尾，需要移动到指定位置
    if anchor_mode == 'after_table' and anchor_idx is not None:
        _move_table_after(doc, table, anchor_idx)
    elif anchor_mode == 'after_para' and anchor_idx is not None:
        _move_table_after_para(doc, table, anchor_idx)

    return table


def _set_tc_width(cell, twips):
    """设置单元格宽度（twips）"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcW = tcPr.find(qn('w:tcW'))
    if tcW is None:
        tcW = tcPr.makeelement(qn('w:tcW'), {})
        tcPr.append(tcW)
    tcW.set(qn('w:w'), str(int(twips)))
    tcW.set(qn('w:type'), 'dxa')


def _move_table_after(doc, table, anchor_table_idx):
    """把 table 移动到第 anchor_table_idx 个已有表格之后"""
    existing = doc.tables
    if anchor_table_idx >= len(existing) or anchor_table_idx < 0:
        return
    anchor_tbl = existing[anchor_table_idx]
    # 找到 anchor 表格后的段落（body 末尾元素）
    anchor_tbl_el = anchor_tbl._tbl
    body = doc.element.body
    # 在 anchor 表格之后插入新表格
    anchor_tbl_el.addnext(table._tbl)


def _move_table_after_para(doc, table, para_idx):
    """把 table 移动到第 para_idx 个段落之后"""
    paragraphs = doc.paragraphs
    if para_idx >= len(paragraphs) or para_idx < 0:
        return
    target_p = paragraphs[para_idx]._p
    target_p.addnext(table._tbl)


def extract_anchors(doc):
    """
    从模板文档提取"位置锚点"。
    规则：标记为锚点的段落 = 文本含「（此处」或含书签(bookmarkStart)的段落。
    返回 [{idx, text}]，idx 为段落序号（对应 after_para 锚点索引）。
    """
    anchors = []
    for i, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if not t:
            continue
        # 占位锚点：形如（此处...）或【插入...】
        if ('（此处' in t) or ('此处' in t and '）' in t) or ('【' in t and '插入' in t):
            anchors.append({'idx': i, 'text': t[:40]})
            continue
        # 书签锚点
        if p._p.find(qn('w:bookmarkStart')) is not None:
            anchors.append({'idx': i, 'text': t[:40] or f'(书签段 {i})'})
    return anchors


def ensure_anchor_count(doc, n):
    """
    确保模板至少有 n 个位置锚点。若现有锚点不足 n 个，则在文档末尾前追加
    占位锚点段落「（此处将动态插入 Sheet 表格）」，使位置数量可随 Excel 的 Sheet 数变化。
    返回全部锚点段落索引列表（升序）。
    """
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    anchors = extract_anchors(doc)
    while len(anchors) < n:
        # 在最后一个现有段落之前插入一个新锚点
        last = doc.paragraphs[-1]._p
        new_p = OxmlElement('w:p')
        new_p.append(OxmlElement('w:pPr'))  # 保持段落结构合法
        r = OxmlElement('w:r')
        t = OxmlElement('w:t')
        t.text = '（此处将动态插入 Sheet 表格）'
        r.append(t)
        new_p.append(r)
        last.addnext(new_p)
        anchors = extract_anchors(doc)  # 重新提取，idx 已更新
    return [a['idx'] for a in anchors]


def resolve_anchor_idx(doc, pos_index):
    """
    把"位置序号 pos_index"（0 基，与上传 Excel 的 Sheet 顺序一一对应）解析为
    模板中真实的锚点段落索引。若锚点不足，先 ensure 到 pos_index+1。
    """
    anchors = extract_anchors(doc)
    if len(anchors) <= pos_index:
        ensure_anchor_count(doc, pos_index + 1)
        anchors = extract_anchors(doc)
    return anchors[pos_index]['idx'] if pos_index < len(anchors) else None


def annotate_bindings(tpl_path, bindings, out_path, excel_path=None, sheet_count=None):
    """
    绑定预览（编辑器左侧显示）：在每个绑定锚点位置写入纯文本占位符
    「【Sheet「xxx」表格将在此处展示】」，让用户一眼看到"哪个位置放哪个 Sheet"。
    纯文本 → OnlyOffice 100% 兼容渲染。
    位置序号 pos_index（0 基）解析为模板真实锚点；不足时自动补占位锚点。
    """
    from docx.shared import RGBColor, Pt
    from docx.text.paragraph import Paragraph
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    doc = Document(tpl_path)

    # 确保锚点数量不少于绑定所需
    need = max([b.get('pos_index', b.get('anchor_idx', 0)) + 1 for b in bindings] + ([0] if not bindings else []))
    if sheet_count:
        need = max(need, sheet_count)
    if need > 0:
        ensure_anchor_count(doc, need)

    # 按位置序号升序处理
    ordered = sorted(bindings, key=lambda b: b.get('pos_index', b.get('anchor_idx') or 0))
    offset = 0
    for b in ordered:
        pos_index = b.get('pos_index', b.get('anchor_idx') or 0)
        anchor_idx = resolve_anchor_idx(doc, pos_index)
        if anchor_idx is None:
            continue
        idx = anchor_idx + offset
        label = b.get('pos_label') or ('位置' + str(pos_index + 1))
        sheet_name = b.get('sheet_name') or b.get('sheet') or ''
        paras = doc.paragraphs
        if idx < 0 or idx >= len(paras):
            continue

        # 在锚点段落后插入纯文本占位符段落
        placeholder_text = '【Sheet\u300c%s\u300d表格将在此处展示】' % sheet_name
        new_p = OxmlElement('w:p')
        paras[idx]._p.addnext(new_p)
        wp = Paragraph(new_p, doc)
        run = wp.add_run(placeholder_text)
        run.bold = True
        run.font.color.rgb = RGBColor(0x53, 0x4A, 0xB7)   # 蓝色醒目
        run.font.size = Pt(10)
        offset += 1

    doc.save(out_path)
    return out_path


# 占位符标记前缀（render 时用于识别并替换）
_PLACEHOLDER_PREFIX = '【Sheet'


def render_template(tpl_path, xlsx_path, bindings, out_path):
    """
    渲染：读取模板，按绑定关系注入多个 Sheet 表格，保存结果。

    逻辑：
    1. 先 ensure 锚点数量
    2. 对每个绑定：在锚点位置注入真实表格（替换 annotate 阶段的占位符段落）
    3. 若无占位符（直接 render），则在锚点段落后插入表格

    参数：
        tpl_path: Word 模板路径
        xlsx_path: Excel 路径
        bindings: [{sheet_name, pos_index, anchor_mode}] 列表
        out_path: 输出 docx 路径

    返回：统计信息
    """
    from docx.oxml.ns import qn

    doc = Document(tpl_path)
    text_width = get_text_width(doc)

    sheets = read_excel_sheets(xlsx_path)
    sheet_map = {s['name']: s for s in sheets}

    stats = {
        'template_para_count': len(doc.paragraphs),
        'template_table_count': len(doc.tables),
        'text_width_inch': round(text_width / 914400, 2),
        'sheets': sheets,
        'injected': [],
        'skipped_empty': [],
        'not_found': [],
    }

    # 确保锚点数量
    need = len(sheets)
    if bindings:
        need = max(need, max(b.get('pos_index', b.get('anchor_idx', 0)) + 1 for b in bindings))
    if need > 0:
        ensure_anchor_count(doc, need)

    for binding in bindings:
        sheet_name = binding['sheet_name']
        if sheet_name not in sheet_map:
            stats['not_found'].append(sheet_name)
            continue
        sheet = sheet_map[sheet_name]
        if not sheet['rows']:
            stats['skipped_empty'].append(sheet_name)
            continue

        pos_index = binding.get('pos_index', binding.get('anchor_idx'))
        anchor_idx = resolve_anchor_idx(doc, pos_index) if pos_index is not None else None

        # 策略：先找占位符段落（annotate 阶段写入的），有则替换；无则追加
        placeholder_para_idx = _find_placeholder_para(doc, sheet_name)

        if placeholder_para_idx is not None:
            # 找到占位符 → 在占位符位置注入表格，然后删除占位符段落
            table = inject_sheet_table(
                doc, sheet,
                anchor_mode='after_para',
                anchor_idx=placeholder_para_idx,
            )
            # 删除占位符段落（表格已插入在其后，需重新定位）
            _remove_paragraph_at(doc, placeholder_para_idx)
        elif anchor_idx is not None:
            # 无占位符 → 直接在锚点段落后插入
            table = inject_sheet_table(
                doc, sheet,
                anchor_mode='after_para',
                anchor_idx=anchor_idx,
            )
        else:
            table = inject_sheet_table(doc, sheet, anchor_mode='end')

        stats['injected'].append({
            'sheet': sheet_name,
            'cols': len(sheet['header']),
            'rows': len(sheet['rows']) + 1,
        })

    doc.save(out_path)
    stats['output_table_count'] = len(Document(out_path).tables)
    return stats


def _find_placeholder_para(doc, sheet_name):
    """查找 annotate 写入的占位符段落「【Sheet「xxx」表格将在此处展示】」，返回段落索引或 None"""
    marker = 'Sheet\u300c%s\u300d' % sheet_name   # 「xxx」
    for i, p in enumerate(doc.paragraphs):
        if _PLACEHOLDER_PREFIX in p.text and marker in p.text:
            return i
    return None


def _remove_paragraph_at(doc, para_idx):
    """删除指定索引的段落（从 body 中移除其 XML 元素）"""
    paras = doc.paragraphs
    if 0 <= para_idx < len(paras):
        p_el = paras[para_idx]._p
        p_el.getparent().remove(p_el)
