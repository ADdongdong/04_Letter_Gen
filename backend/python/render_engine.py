# -*- coding: utf-8 -*-
"""
制函渲染核心引擎（验证用 Python 实现，后续可移植到 Java/poi-tl）

核心能力：在 Word 模板的"锚点位置"，按 Excel Sheet 动态生成完整表格。
- 列数 = Sheet 列数
- 表头 = Sheet 第一行
- 数据 = Sheet 其余行
- 表格总宽锁定为模板文本宽度
"""
import io, os, re, unicodedata
from typing import TypedDict
from docx import Document
from docx.shared import Pt, Cm, Emu
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from openpyxl import load_workbook


# ===== 表格样式配置（模板配置页可改，缺省逐字段回落默认值=原硬编码行为）=====
class TableStyle(TypedDict, total=False):
    font_size: float    # 磅
    font_name: str      # 空=继承模板 Normal 字体
    h_align: str        # center | left | right
    v_align: str        # top | center
    row_height_mode: str  # auto | atLeast（最小值，内容多自动加高不截断）
    row_height_pt: float  # atLeast 时的最小行高（磅）
    col_widths_mode: str  # equal（均分）| auto（按内容自适应）| ratio（按 Sheet 自定义比例）
    col_widths: dict     # ratio 模式：{'Sheet名': '2:1:1'}；未配置的 Sheet 均分
    col_widths_ratio: str  # [兼容旧字段] 全局比例串，col_widths 缺失时回落生效


DEFAULT_TABLE_STYLE: TableStyle = {
    'font_size': 10,
    'font_name': '',
    'h_align': 'center',
    'v_align': 'top',
    'row_height_mode': 'auto',
    'row_height_pt': 20,
    'col_widths_mode': 'equal',
    'col_widths': {},
    'col_widths_ratio': '',
}

_H_ALIGN_MAP = {
    'left': WD_ALIGN_PARAGRAPH.LEFT,
    'center': WD_ALIGN_PARAGRAPH.CENTER,
    'right': WD_ALIGN_PARAGRAPH.RIGHT,
}
_V_ALIGN_MAP = {
    'top': WD_CELL_VERTICAL_ALIGNMENT.TOP,
    'center': WD_CELL_VERTICAL_ALIGNMENT.CENTER,
}


def _resolve_style(style) -> TableStyle:
    """用户样式与默认值合并：逐字段类型/枚举校验，非法值回落默认（防脏数据）"""
    s: TableStyle = dict(DEFAULT_TABLE_STYLE)
    if not isinstance(style, dict):
        return s
    try:
        v = style.get('font_size')
        if v is not None and 5 <= float(v) <= 72:
            s['font_size'] = float(v)
    except (TypeError, ValueError):
        pass
    v = style.get('font_name')
    if isinstance(v, str):
        s['font_name'] = v.strip()
    v = style.get('h_align')
    if v in _H_ALIGN_MAP:
        s['h_align'] = v
    v = style.get('v_align')
    if v in _V_ALIGN_MAP:
        s['v_align'] = v
    v = style.get('row_height_mode')
    if v in ('auto', 'atLeast'):
        s['row_height_mode'] = v
    try:
        v = style.get('row_height_pt')
        if v is not None and 5 <= float(v) <= 200:
            s['row_height_pt'] = float(v)
    except (TypeError, ValueError):
        pass
    v = style.get('col_widths_ratio')
    if isinstance(v, str):
        s['col_widths_ratio'] = v.strip()
    v = style.get('col_widths_mode')
    if v in ('equal', 'auto', 'ratio'):
        s['col_widths_mode'] = v
    v = style.get('col_widths')
    if isinstance(v, dict):
        s['col_widths'] = {str(k): str(val) for k, val in v.items()}
    return s


def _parse_col_ratio(ratio_str: str, ncols: int):
    """'2:1:1' -> [2.0,1.0,1.0]；为空/格式非法/列数不匹配返回 None（回落均分）"""
    if not ratio_str:
        return None
    try:
        parts = [float(x) for x in str(ratio_str).split(':') if x.strip()]
    except ValueError:
        return None
    if len(parts) != ncols or sum(parts) <= 0 or any(p < 0 for p in parts):
        return None
    return parts


def _char_w(ch):
    """单字符显示宽度：CJK/全角按 2，其余按 1"""
    return 2 if unicodedata.east_asian_width(ch) in ('W', 'F') else 1


def _str_w(s):
    """字符串显示宽度（CJK 加权）"""
    return sum(_char_w(c) for c in str(s))


def _col_weights_auto(sheet):
    """按内容自适应列宽权重：每列 = max(表头宽, 该列数据最大宽)，中文按 2 倍宽，下限 1。
    列数为 0 返回 None（回落均分）。"""
    header = sheet.get('header') or []
    ncols = len(header)
    if ncols == 0:
        return None
    weights = [_str_w(header[j]) for j in range(ncols)]
    for row in sheet.get('rows') or []:
        for j in range(ncols):
            v = row[j] if j < len(row) else ''
            w = _str_w(v)
            if w > weights[j]:
                weights[j] = w
    return [max(w, 1) for w in weights]


def get_text_width(doc):
    """获取页面文本宽度（EMU）"""
    sec = doc.sections[0]
    return sec.page_width - sec.left_margin - sec.right_margin


def read_excel_sheets(xlsx_path):
    """读取 Excel 所有 Sheet，返回 [{name, header, rows, row_nums, ...}]
    row_nums: 与 rows 一一对应的数据行在 Excel 中的真实行号（1 基，数据从第 2 行起，第 1 行为表头）
    """
    wb = load_workbook(xlsx_path, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
        if not rows:
            continue
        # 表头 = 第一行，数据 = 其余行
        header = [str(c) if c is not None else '' for c in rows[0]]
        # 过滤全空数据行，同步保留 Excel 真实行号（供批量制函报错定位）
        data = []
        row_nums = []
        for excel_row, r in enumerate(rows[1:], start=2):
            cells = [str(c) if c is not None else '' for c in r]
            if any(v.strip() for v in cells):
                data.append(cells)
                row_nums.append(excel_row)
        sheets.append({
            'name': ws.title,
            'header': header,
            'rows': data,
            'row_nums': row_nums,
            'max_row': ws.max_row,
            'max_col': ws.max_column,
        })
    return sheets


# 批量制函：分组列表头（Sheet 第 1 行前两列精确匹配才算分组 Sheet）
_GROUP_COL_AUDIT = '被审计单位'
_GROUP_COL_CONFIRM = '被询证单位'


def _is_group_header(header):
    """判断表头前两列是否为「被审计单位」「被询证单位」→ 是则为批量分组 Sheet"""
    if len(header) < 2:
        return False
    return header[0].strip() == _GROUP_COL_AUDIT and header[1].strip() == _GROUP_COL_CONFIRM


def read_excel_grouped(xlsx_path):
    """
    读取批量 Excel（多封函证数据写在同一个 Excel 中）：
    - 仅处理表头前两列为「被审计单位」「被询证单位」的 Sheet（其余跳过）
    - 按 (被审计单位, 被询证单位) 组合把每个分组 Sheet 的行分组，同一组合的数据可分散在多个 Sheet

    返回：
    {
        'sheets':  [sheet_name, ...],                # 分组 Sheet 名（保序）
        'groups':  [(audit, confirm), ...],          # 全部组合并集（保序）＝ N 封函证
        'empty_cells': [                             # 分组 Sheet 中前两列任一为空的数据行定位
            {'sheet': sheet_name, 'row': Excel行号, 'col': '被审计单位'|'被询证单位'},
        ],                                           # 非空时上层应终止制函并要求用户补充
        'virtual_sheets': {                          # 预构造的"虚拟 sheet"（供 render_template 直接消费）
            sheet_name: {
                (audit, confirm): {'name', 'header', 'rows', ...},  # header/rows 均去掉前两列（从第 3 列起）
            },                                       # 组合在该 Sheet 无行时 rows=[]（渲染时跳过并删占位段）
        },
    }
    """
    sheets = read_excel_sheets(xlsx_path)
    grouped_rows = {}   # sheet_name -> {key: [原始行,...]}
    sheet_headers = {}  # sheet_name -> 第 3 列起的表头
    sheet_order = []
    groups = []
    seen = set()
    empty_cells = []
    for sheet in sheets:
        if not _is_group_header(sheet['header']):
            continue
        buckets = {}
        for row, excel_row in zip(sheet['rows'], sheet.get('row_nums') or []):
            audit = row[0].strip() if len(row) > 0 else ''
            confirm = row[1].strip() if len(row) > 1 else ''
            # 前两列任一为空：记录精确位置（上层据此终止制函，要求用户补充）
            if not audit:
                empty_cells.append({'sheet': sheet['name'], 'row': excel_row, 'col': _GROUP_COL_AUDIT})
            if not confirm:
                empty_cells.append({'sheet': sheet['name'], 'row': excel_row, 'col': _GROUP_COL_CONFIRM})
            key = (audit, confirm)
            if key not in buckets:
                buckets[key] = []
            buckets[key].append(row)
            if key not in seen:
                seen.add(key)
                groups.append(key)
        grouped_rows[sheet['name']] = buckets
        sheet_headers[sheet['name']] = sheet['header'][2:]
        sheet_order.append(sheet['name'])

    # 为每个分组 Sheet × 每个全局组合构造虚拟 sheet（无数据的组合 rows=[]）
    virtual_sheets = {}
    for name in sheet_order:
        vmap = {}
        for key in groups:
            rows = [r[2:] for r in grouped_rows[name].get(key, [])]
            vmap[key] = {
                'name': name,
                'header': sheet_headers[name],
                'rows': rows,
                'max_row': len(rows) + 1,
                'max_col': len(sheet_headers[name]),
            }
        virtual_sheets[name] = vmap

    return {'sheets': sheet_order, 'groups': groups, 'empty_cells': empty_cells, 'virtual_sheets': virtual_sheets}


def _set_cell_text(cell, text, bold=False, size=10, align='center', font_name='', v_align=None):
    """填充单元格文本并设置样式：水平对齐/字号/加粗/字体（中西文双设置）/垂直对齐"""
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = _H_ALIGN_MAP.get(align, WD_ALIGN_PARAGRAPH.CENTER)
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(size)
    if font_name:
        # 中文字体必须同时设置 rFonts@eastAsia，否则 run.font.name 只对西文生效
        run.font.name = font_name
        rpr = run._element.get_or_add_rPr()
        rfonts = rpr.find(qn('w:rFonts'))
        if rfonts is None:
            rfonts = OxmlElement('w:rFonts')
            rpr.append(rfonts)
        rfonts.set(qn('w:eastAsia'), font_name)
    if v_align is not None:
        cell.vertical_alignment = v_align


def _table_width_twips(doc):
    """文本宽度转 twips（1 twip = 1/20 点）"""
    return int(get_text_width(doc) / 914400 * 1440)


def _set_row_min_height(row, pt: float):
    """设置行最小高度（w:trHeight hRule=atLeast：内容少时保持设定值，内容多自动加高不截断）"""
    trPr = row._tr.get_or_add_trPr()
    tr_height = trPr.find(qn('w:trHeight'))
    if tr_height is None:
        tr_height = OxmlElement('w:trHeight')
        trPr.append(tr_height)
    tr_height.set(qn('w:val'), str(int(pt * 20)))  # 磅 -> twips
    tr_height.set(qn('w:hRule'), 'atLeast')


def inject_sheet_table(doc, sheet, anchor_mode='end', anchor_idx=None, col_widths_ratio=None, style=None):
    """
    在 doc 中注入一个 Sheet 表格。

    参数：
        doc: python-docx Document
        sheet: read_excel_sheets 返回的单个 sheet dict
        anchor_mode: 'end'(文档末尾) / 'after_table'(指定表格后) / 'after_para'(指定段落后)
        anchor_idx: anchor_mode 对应的索引（表格序号/段落序号）
        col_widths_ratio: 可选列宽比例列表，None 则按 style 配置或均分
        style: 表格样式 dict（None/_resolve_style 回落 DEFAULT_TABLE_STYLE）

    返回：注入的表格
    """
    st = _resolve_style(style)
    header = sheet['header']
    rows = sheet['rows']
    ncols = len(header)
    if ncols == 0:
        # 分组 Sheet 去掉前两列后无数据列（或空 Sheet 结构）：无可注入内容，跳过
        return None

    # 创建表格（1 行表头）
    table = doc.add_table(rows=1, cols=ncols)
    table.style = 'Table Grid'

    # 计算列宽（twips）：调用方显式比例 > 三档模式（equal 均分 / auto 内容自适应 / ratio 按 Sheet 比例）
    # > 旧全局串兼容（存量模板） > 均分
    ratio = col_widths_ratio
    if not ratio and st['col_widths_mode'] == 'auto':
        ratio = _col_weights_auto(sheet)
    if not ratio and st['col_widths_mode'] == 'ratio':
        ratio = _parse_col_ratio(st['col_widths'].get(sheet['name'], ''), ncols)
    if not ratio:
        # 兼容存量模板：旧全局 col_widths_ratio 非空时仍生效
        ratio = _parse_col_ratio(st['col_widths_ratio'], ncols)
    total_twips = _table_width_twips(doc)
    if ratio and len(ratio) == ncols and sum(ratio) > 0:
        twips = [int(total_twips * r / sum(ratio)) for r in ratio]
    else:
        twips = [total_twips // ncols] * ncols

    v_align = _V_ALIGN_MAP.get(st['v_align'])

    # 表头行（表头固定加粗；字号/对齐/字体随样式配置）
    for j, h in enumerate(header):
        _set_cell_text(table.rows[0].cells[j], h, bold=True, size=st['font_size'],
                       align=st['h_align'], font_name=st['font_name'], v_align=v_align)

    # 数据行
    for row_data in rows:
        cells = table.add_row().cells
        for j in range(ncols):
            val = row_data[j] if j < len(row_data) else ''
            _set_cell_text(cells[j], val, bold=False, size=st['font_size'],
                           align=st['h_align'], font_name=st['font_name'], v_align=v_align)

    # 行高：最小值模式时对全部行设置 atLeast（内容多自动加高不截断）
    if st['row_height_mode'] == 'atLeast':
        for row in table.rows:
            _set_row_min_height(row, st['row_height_pt'])

    # 设置列宽（每单元格 tcW）
    for row in table.rows:
        for j, cell in enumerate(row.cells):
            if j < ncols:
                cell.width = Emu(int(twips[j] * 635))  # twips -> EMU（1 twip = 635 EMU）
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


def render_template(tpl_path, xlsx_path, bindings, out_path, group_key=None, style=None):
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
        group_key: 可选 (被审计单位, 被询证单位) 元组。传入时启用批量模式：
            按 read_excel_grouped 分组，仅渲染该组合的行；虚拟 sheet 的表头与数据
            均从第 3 列起（不含前两列分组列）；该组合在某个 Sheet 无行时跳过注入
            并删除其占位段。
        style: 可选表格样式 dict（模板配置），None 时用 DEFAULT_TABLE_STYLE

    返回：统计信息
    """
    from docx.oxml.ns import qn

    doc = Document(tpl_path)
    text_width = get_text_width(doc)

    if group_key is not None:
        grouped = read_excel_grouped(xlsx_path)
        sheet_map = {name: vmap[group_key] for name, vmap in grouped['virtual_sheets'].items()}
        sheets = [sheet_map[n] for n in grouped['sheets'] if n in sheet_map]
    else:
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

    # 注意：不再预建锚点（ensure_anchor_count）——占位段驱动的制函（模板即定义）若预建锚点，
    # 未被使用的锚点段落会残留在文档末尾（「（此处将动态插入 Sheet 表格）」×N）。
    # 锚点仅在旧流程 fallback（binding 带 pos_index 且模板无占位段）时由 resolve_anchor_idx 按需补建。
    occurrence_count = {}  # sheet_name -> 已替换/删除的占位段出现次数（支持同一 Sheet 多处插入）

    for binding in bindings:
        sheet_name = binding['sheet_name']
        if sheet_name not in sheet_map:
            stats['not_found'].append(sheet_name)
            # 该 Sheet 不在 Excel 中（无表头不参与分组 / Sheet 缺失 / 无数据）：
            # 同样按出现次序删除其占位段，避免输出残留「【Sheet「xxx」表格将在此处展示】」提示词
            placeholder_para_idx = _find_placeholder_para(doc, sheet_name, occurrence_count.get(sheet_name, 0))
            if placeholder_para_idx is not None:
                _remove_paragraph_at(doc, placeholder_para_idx)
                occurrence_count[sheet_name] = occurrence_count.get(sheet_name, 0) + 1
            continue
        sheet = sheet_map[sheet_name]
        if not sheet['rows']:
            stats['skipped_empty'].append(sheet_name)
            # 该 Sheet 无数据行：同时删除模板中它的占位段（按出现次序），避免输出残留「【Sheet「xxx」...】」
            placeholder_para_idx = _find_placeholder_para(doc, sheet_name, occurrence_count.get(sheet_name, 0))
            if placeholder_para_idx is not None:
                _remove_paragraph_at(doc, placeholder_para_idx)
                occurrence_count[sheet_name] = occurrence_count.get(sheet_name, 0) + 1
            continue

        # 策略：占位段优先——先找占位符段落（annotate 阶段写入的），命中即替换；
        # 同一 Sheet 插入多处时按出现次序依次替换（occurrence 计数）。
        # 占位段命中的路径不触碰锚点（不 ensure），避免产生残留的锚点段落。
        placeholder_para_idx = _find_placeholder_para(doc, sheet_name, occurrence_count.get(sheet_name, 0))

        if placeholder_para_idx is not None:
            # 找到占位符 → 在占位符位置注入表格，然后删除占位符段落
            table = inject_sheet_table(
                doc, sheet,
                anchor_mode='after_para',
                anchor_idx=placeholder_para_idx,
                style=style,
            )
            # 删除占位符段落（表格已插入在其后，需重新定位）
            _remove_paragraph_at(doc, placeholder_para_idx)
            occurrence_count[sheet_name] = occurrence_count.get(sheet_name, 0) + 1
        else:
            # 无占位符 → 旧流程 fallback：按 pos_index 解析锚点（内部按需 ensure），无 pos_index 则文档末尾
            pos_index = binding.get('pos_index', binding.get('anchor_idx'))
            anchor_idx = resolve_anchor_idx(doc, pos_index) if pos_index is not None else None
            if anchor_idx is not None:
                table = inject_sheet_table(
                    doc, sheet,
                    anchor_mode='after_para',
                    anchor_idx=anchor_idx,
                    style=style,
                )
            else:
                table = inject_sheet_table(doc, sheet, anchor_mode='end', style=style)

        stats['injected'].append({
            'sheet': sheet_name,
            'cols': len(sheet['header']),
            'rows': len(sheet['rows']) + 1,
        })

    # 终扫兜底：保存前删除文档中所有残留占位段（宽松匹配「【Sheet」+「将在此处展示」），
    # 覆盖占位文字被 OO 拆 run、手工书写等导致 _find_placeholder_para 精确匹配失败的场景；
    # 成功注入表格的占位段在注入时已被删除，不会被误删
    leftover = [i for i, p in enumerate(doc.paragraphs)
                if _PLACEHOLDER_PREFIX in p.text and '将在此处展示' in p.text]
    for i in reversed(leftover):  # 倒序删除，避免索引位移
        _remove_paragraph_at(doc, i)
    if leftover:
        print(f"[RENDER] 终扫删除残留占位段 {len(leftover)} 处", flush=True)

    doc.save(out_path)
    stats['output_table_count'] = len(Document(out_path).tables)
    return stats


def _find_placeholder_para(doc, sheet_name, occurrence=0):
    """查找 annotate 写入的占位符段落「【Sheet「xxx」表格将在此处展示】」。
    occurrence：同一 Sheet 插入多处时，取第 N 次出现的占位段（0 基）。
    返回段落索引或 None"""
    marker = 'Sheet\u300c%s\u300d' % sheet_name   # 「xxx」
    seen = -1
    for i, p in enumerate(doc.paragraphs):
        if _PLACEHOLDER_PREFIX in p.text and marker in p.text:
            seen += 1
            if seen >= occurrence:
                return i
    return None


def _remove_paragraph_at(doc, para_idx):
    """删除指定索引的段落（从 body 中移除其 XML 元素）"""
    paras = doc.paragraphs
    if 0 <= para_idx < len(paras):
        p_el = paras[para_idx]._p
        p_el.getparent().remove(p_el)


def extract_placeholder_bindings(docx_path):
    """
    扫描模板 docx 中所有「【Sheet「xxx」...】」占位段（按文档顺序），自动生成 bindings 列表。
    批量制函接口据此从模板自身提取插入关系——前端无需传 bindings，模板即全部定义。
    同一 Sheet 多处插入会生成多条 binding（render 时按出现次序依次替换）。
    """
    doc = Document(docx_path)
    bindings = []
    for p in doc.paragraphs:
        if _PLACEHOLDER_PREFIX not in p.text:
            continue
        m = re.search(r'Sheet\u300c(.+?)\u300d', p.text)
        if not m:
            continue
        name = m.group(1)
        bindings.append({
            'sheet_name': name,
            'pos_index': len(bindings),
            'anchor_mode': 'after_para',
            'pos_label': '占位-%s' % name,
        })
    return bindings
