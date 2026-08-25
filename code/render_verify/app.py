# -*- coding: utf-8 -*-
"""
制函渲染验证服务 - Flask Web 应用
启动: python app.py  → http://127.0.0.1:5002
"""
import io, os, uuid, json, re, shutil, socket
from datetime import datetime
from flask import Flask, request, render_template_string, jsonify, send_file, url_for

from render_engine import (
    read_excel_sheets, render_template, get_text_width, _table_width_twips,
    extract_anchors, annotate_bindings, extract_placeholders, render_by_template,
)
from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'outputs')
TEMPLATES_DIR = os.path.join(BASE_DIR, '..', '..', 'data', 'templates')
TEMPLATES_JSON = os.path.join(TEMPLATES_DIR, 'templates.json')
LEDGER_JSON = os.path.join(BASE_DIR, '..', '..', 'data', 'ledger.json')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(TEMPLATES_DIR, exist_ok=True)

DEFAULT_TPL = os.path.join(BASE_DIR, '..', 'conversion_test', 'demo_template.docx')
BLANK_TPL = os.path.join(TEMPLATES_DIR, 'blank.docx')

# 确保空白模板存在（新建模板时左侧直接加载空白 Word）
if not os.path.exists(BLANK_TPL):
    Document().save(BLANK_TPL)


def _resolve_source_path(source):
    """解析 source 为实际 docx 文件路径：default/blank 或模板 id"""
    if source == 'default':
        return DEFAULT_TPL
    if source == 'blank':
        if not os.path.exists(BLANK_TPL):
            Document().save(BLANK_TPL)
        return BLANK_TPL
    items = _load_templates_json()
    t = next((x for x in items if x.get('id') == source), None)
    return t.get('path') if t else None


# OnlyOffice 通过 http 加载 docx，必须用本机可被 OnlyOffice 访问的内网 IP
try:
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.connect(('8.8.8.8', 80))
    LOCAL_IP = s.getsockname()[0]
    s.close()
except Exception:
    LOCAL_IP = '127.0.0.1'

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB


# ============ 页面 ============
@app.route('/')
def index():
    return render_template_string(HTML)

# OnlyOffice 演示 Demo 页面（客户演示用）
@app.route('/demo')
def demo():
    return send_file(
        os.path.join(BASE_DIR, 'onlyoffice_demo.html'),
        mimetype='text/html'
    )


# ============ 上传模板 ============
@app.route('/api/upload_template', methods=['POST'])
def upload_template():
    f = request.files.get('file')
    if not f or not f.filename.lower().endswith('.docx'):
        return jsonify({'error': '请上传 .docx 模板'}), 400
    uid = uuid.uuid4().hex[:8]
    path = os.path.join(UPLOAD_DIR, f'tpl_{uid}.docx')
    f.save(path)

    doc = Document(path)
    paras = [p.text for p in doc.paragraphs]
    tables = []
    for i, t in enumerate(doc.tables):
        tables.append({
            'idx': i,
            'rows': len(t.rows),
            'cols': len(t.columns),
            'first_cell': t.rows[0].cells[0].text[:20] if len(t.rows) > 0 else '',
        })
    text_w = get_text_width(doc) / 914400

    return jsonify({
        'path': path,
        'uid': uid,
        'para_count': len(paras),
        'table_count': len(tables),
        'text_width_inch': round(text_w, 2),
        'paras': [{'idx': i, 'text': p[:30]} for i, p in enumerate(paras) if p.strip()][:200],
        'tables': tables,
    })


# ============ 上传 Excel ============
@app.route('/api/upload_excel', methods=['POST'])
def upload_excel():
    f = request.files.get('file')
    if not f or not f.filename.lower().endswith(('.xlsx', '.xlsm')):
        return jsonify({'error': '请上传 .xlsx 文件'}), 400
    uid = uuid.uuid4().hex[:8]
    path = os.path.join(UPLOAD_DIR, f'excel_{uid}.xlsx')
    f.save(path)

    sheets = read_excel_sheets(path)
    return jsonify({
        'path': path,
        'uid': uid,
        'sheets': [{
            'name': s['name'],
            'header': s['header'],
            'row_count': len(s['rows']),
            'preview': s['rows'][:5],
        } for s in sheets],
    })


# ============ 读取模板位置锚点（动态） ============
@app.route('/api/template_anchors', methods=['GET'])
def template_anchors():
    """解析默认模板文档中的占位锚点，返回动态位置列表。"""
    tpl = DEFAULT_TPL
    if not os.path.exists(tpl):
        return jsonify({'error': '默认模板不存在: ' + tpl}), 404
    try:
        anchors = extract_anchors(Document(tpl))
        return jsonify({'template': tpl, 'anchors': anchors})
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ============ 绑定回写（在模板上显示绑定关系） ============
@app.route('/api/bind_annotate', methods=['POST'])
def bind_annotate():
    """
    把前端收集的 bindings 回写到模板文档，生成一份带 📌 标注的 docx，
    供 OnlyOffice 重新加载后在左侧模板上直接看到"该位置绑了哪个 Sheet"。
    bindings: [{anchor_idx, pos_label, sheet_name}]
    返回新 docx 的下载路径（替换 onlyoffice_demo 的 docUrl 指向它）。
    """
    data = request.get_json()
    bindings = data.get('bindings', [])
    tpl = data.get('tpl_path') or DEFAULT_TPL
    excel_path = data.get('excel_path')  # 有则预览真实表格
    sheet_count = data.get('sheet_count')  # 位置数应 = Sheet 数，不足时自动补锚点
    if not os.path.exists(tpl):
        return jsonify({'error': '模板不存在'}), 400
    uid = uuid.uuid4().hex[:8]
    out_path = os.path.join(OUTPUT_DIR, f'annotated_{uid}.docx')
    try:
        annotate_bindings(tpl, bindings, out_path, excel_path=excel_path, sheet_count=sheet_count)
        # 复制到静态目录 8899，便于 OnlyOffice 通过 http 加载
        static_dir = os.path.join(BASE_DIR, '..', 'conversion_test')
        static_name = f'annotated_{uid}.docx'
        shutil.copy(out_path, os.path.join(static_dir, static_name))
        return jsonify({
            'path': out_path,
            'static_url': f'http://{LOCAL_IP}:8899/{static_name}',
            'download_url': f'/api/download/{os.path.basename(out_path)}',
            'bindings': bindings,
        })
    except Exception as e:
        return jsonify({'error': str(e)}), 500


# ============ 渲染 ============
@app.route('/api/render', methods=['POST'])
def render():
    data = request.get_json()
    tpl_path = data.get('tpl_path') or DEFAULT_TPL
    excel_path = data.get('excel_path')
    bindings = data.get('bindings', [])

    if not os.path.exists(tpl_path):
        return jsonify({'error': '模板不存在: ' + tpl_path}), 400
    if not excel_path or not os.path.exists(excel_path):
        return jsonify({'error': '请先上传 Excel'}), 400
    if not bindings:
        return jsonify({'error': '请至少添加一个绑定'}), 400

    uid = uuid.uuid4().hex[:8]
    out_path = os.path.join(OUTPUT_DIR, f'result_{uid}.docx')
    try:
        stats = render_template(tpl_path, excel_path, bindings, out_path)
        stats['download_url'] = f'/api/download/{os.path.basename(out_path)}'
        return jsonify(stats)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/download/<fname>')
def download(fname):
    path = os.path.join(OUTPUT_DIR, fname)
    if not os.path.exists(path):
        return jsonify({'error': '文件不存在'}), 404
    return send_file(path, as_attachment=True, download_name='渲染结果.docx')


# ============ 模板化批量制函（2026-08-24 新增）============

def _load_templates_json():
    """读取模板清单，不存在则返回 []"""
    if not os.path.exists(TEMPLATES_JSON):
        return []
    try:
        with open(TEMPLATES_JSON, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []


def _save_templates_json(items):
    """写入模板清单"""
    with open(TEMPLATES_JSON, 'w', encoding='utf-8') as f:
        json.dump(items, f, ensure_ascii=False, indent=2)


def _load_ledger():
    """读取函证台账（demo 模拟数据），不存在则返回 []"""
    if not os.path.exists(LEDGER_JSON):
        return []
    try:
        with open(LEDGER_JSON, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return []


def _parse_letter_no(filename):
    """从文件名解析函证编号：取首个 '_' 前段；无 '_' 则原样返回文件名（去扩展名）"""
    base = os.path.splitext(os.path.basename(filename))[0]
    if '_' in base:
        return base.split('_', 1)[0]
    return base


@app.route('/api/excel_sheets', methods=['POST'])
def excel_sheets():
    """解析上传 Excel 的 Sheet 列表，供快速导入制函选择位置。返回 {sheets: [{name}]}"""
    f = request.files.get('file')
    if not f or not f.filename.lower().endswith(('.xlsx', '.xlsm')):
        return jsonify({'error': '请上传 Excel 文件（.xlsx）'}), 400
    tmp = os.path.join(UPLOAD_DIR, f'q_sheets_{uuid.uuid4().hex[:8]}.xlsx')
    try:
        f.save(tmp)
        wb = openpyxl.load_workbook(tmp, read_only=True, data_only=True)
        sheets = [{'name': ws.title} for ws in wb.worksheets]
        wb.close()
    except Exception as e:
        return jsonify({'error': 'Excel 解析失败: ' + str(e)}), 400
    finally:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass
    return jsonify({'sheets': sheets})


def _match_ledger(letter_no):
    """按函证编号查台账，返回台账行或 None"""
    ledger = _load_ledger()
    for row in ledger:
        if row.get('函证编号') == letter_no:
            return row
    return None


@app.route('/api/template_save', methods=['POST'])
def template_save():
    """
    保存模板：前端提交 bindings（占位符绑定）→ 基于源模板生成带占位符 docx → 扫描提取绑定规则 → 存 data/templates/。
    JSON 字段:
        id: 可选；提供时表示「修改」该模板（原地更新，id 不变）
        name: 模板名称
        bindings: [{pos_index, sheet_name, pos_label?}] 占位符绑定（按插入顺序）
        source: 'default' / 'blank' 或已存模板 id（决定基于哪个 docx 生成）
    返回: 模板 id、绑定规则列表
    """
    data = request.get_json(silent=True) or {}
    name = (data.get('name') or '').strip()
    bindings = data.get('bindings', [])
    source = data.get('source') or 'default'
    excel_name = (data.get('excel_name') or '').strip()
    template_id = data.get('id')  # 修改模式：原地更新该 id

    if not bindings:
        return jsonify({'error': '请至少提供一个占位符绑定'}), 400
    if not name:
        return jsonify({'error': '请填写模板名称'}), 400

    # 确定源模板 docx
    src_tpl = _resolve_source_path(source)
    if not src_tpl or not os.path.exists(src_tpl):
        return jsonify({'error': '源模板不存在: ' + str(source)}), 400

    # 防止模板名含非法路径字符
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', name)
    items = _load_templates_json()

    # =================== 修改模式：id 已存在，原地更新 ===================
    if template_id:
        existing = next((t for t in items if t.get('id') == template_id), None)
        if not existing:
            return jsonify({'error': '模板不存在'}), 404
        # 若改名，不得与其他模板重名
        conflict = next((t for t in items if t.get('name') == name and t.get('id') != template_id), None)
        if conflict:
            return jsonify({'error': f'模板名称「{name}」已存在，请使用其他名称'}), 409

        # 基于源 docx 生成新文件
        uid = uuid.uuid4().hex[:8]
        tpl_path = os.path.join(TEMPLATES_DIR, f'{uid}_{safe_name}.docx')
        try:
            annotate_bindings(src_tpl, bindings, tpl_path)
            doc = Document(tpl_path)
            placeholders = extract_placeholders(doc)
            tpl_bindings = [{'sheet_name': ph['sheet_name'], 'para_index': ph['para_index']} for ph in placeholders]
        except Exception as e:
            return jsonify({'error': '模板生成失败: ' + str(e)}), 400

        old_path = existing.get('path')
        existing.update({
            'name': name,
            'path': tpl_path,
            'bindings': tpl_bindings,
            'excel_name': excel_name,
            'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        })
        _save_templates_json(items)

        # 删除旧 docx（容错）
        try:
            if old_path and os.path.exists(old_path) and old_path != tpl_path:
                os.remove(old_path)
        except Exception:
            pass

        return jsonify({
            'id': template_id,
            'name': name,
            'path': tpl_path,
            'bindings': tpl_bindings,
            'excel_name': excel_name,
            'placeholder_count': len(placeholders),
        })

    # =================== 新建模式：无 id，拒绝同名 ===================
    if any(t.get('name') == name for t in items):
        return jsonify({'error': f'模板名称「{name}」已存在，请使用其他名称'}), 409

    uid = uuid.uuid4().hex[:8]
    tpl_path = os.path.join(TEMPLATES_DIR, f'{uid}_{safe_name}.docx')
    try:
        annotate_bindings(src_tpl, bindings, tpl_path)
        doc = Document(tpl_path)
        placeholders = extract_placeholders(doc)
        tpl_bindings = [{'sheet_name': ph['sheet_name'], 'para_index': ph['para_index']} for ph in placeholders]
    except Exception as e:
        return jsonify({'error': '模板生成失败: ' + str(e)}), 400

    items.append({
        'id': uid,
        'name': name,
        'path': tpl_path,
        'bindings': tpl_bindings,
        'excel_name': excel_name,
        'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    })
    _save_templates_json(items)

    return jsonify({
        'id': uid,
        'name': name,
        'path': tpl_path,
        'bindings': tpl_bindings,
        'excel_name': excel_name,
        'placeholder_count': len(placeholders),
    })


@app.route('/api/templates')
def templates():
    """列出所有已配置模板及绑定规则"""
    items = _load_templates_json()
    return jsonify({'templates': items})


@app.route('/api/match_excel', methods=['POST'])
def match_excel():
    """
    批量制函第一步（同页）：上传多个 Excel → 解析文件名取函证编号 → 查台账 → 返回三态匹配结果。
    所有上传文件均列入列表（不丢弃）：
      status: 'ok'     自动匹配（文件名解析出的编号在台账中）
      status: 'pending' 待补编号（解析失败 / 台账暂无）
      status: 'fail'   补后仍无（前端已手填但仍不在台账，由前端以 manual_no 重传触发）
    请求: form 字段 file（可多个 excel），可选 manual_no_map（JSON: {原文件名: 手动函证编号}）
    返回: list: [{file, letter_no, status, ledger(台账行或null)}]
    """
    files = request.files.getlist('file')
    manual_map = {}
    mraw = request.form.get('manual_no_map')
    if mraw:
        try:
            manual_map = json.loads(mraw)
        except Exception:
            manual_map = {}

    if not files:
        return jsonify({'error': '请上传至少一个 Excel'}), 400

    results = []
    for f in files:
        if not f or not f.filename.lower().endswith(('.xlsx', '.xlsm')):
            continue
        fname = f.filename
        # 手动补编号优先：前端传了则用，否则从文件名解析
        letter_no = manual_map.get(fname)
        if not letter_no:
            letter_no = _parse_letter_no(fname)
        ledger = _match_ledger(letter_no)
        status = 'ok' if ledger else 'pending'
        results.append({
            'file': fname,
            'letter_no': letter_no,
            'status': status,
            'ledger': ledger,
        })
    return jsonify({'matches': results})





@app.route('/api/template_delete', methods=['POST'])
def template_delete():
    """删除模板（按 id），同时删除 docx 文件"""
    data = request.get_json() or {}
    tid = data.get('id')
    items = _load_templates_json()
    target = next((t for t in items if t.get('id') == tid), None)
    if not target:
        return jsonify({'error': '模板不存在'}), 404
    items = [t for t in items if t.get('id') != tid]
    _save_templates_json(items)
    try:
        if os.path.exists(target.get('path', '')):
            os.remove(target['path'])
    except Exception:
        pass
    return jsonify({'ok': True})


@app.route('/api/oo/template/<tid>')
def oo_get_saved_template(tid):
    """OnlyOffice 通过此 URL 下载「模板 docx」：
    - tid='blank' → 返回空白模板（新建模板初始画布）
    - tid='default' → 返回默认模板 demo_template.docx
    - 其他 → 返回已保存模板（配置模板时重新打开）
    """
    if tid == 'blank':
        if not os.path.exists(BLANK_TPL):
            Document().save(BLANK_TPL)
        return send_file(
            BLANK_TPL,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            download_name='blank.docx',
            as_attachment=False,
        )
    if tid == 'default':
        if not os.path.exists(DEFAULT_TPL):
            return jsonify({'error': '默认模板不存在'}), 404
        return send_file(
            DEFAULT_TPL,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            download_name='demo_template.docx',
            as_attachment=False,
        )
    items = _load_templates_json()
    tpl = next((t for t in items if t.get('id') == tid), None)
    if not tpl or not os.path.exists(tpl.get('path', '')):
        return jsonify({'error': '模板不存在'}), 404
    return send_file(
        tpl['path'],
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        download_name=os.path.basename(tpl['path']),
        as_attachment=False,
    )


@app.route('/api/upload_template_oo', methods=['POST'])
def upload_template_oo():
    """
    上传用户自己的 Word 模板（.docx）→ 保存到 templates 目录 → 写入清单（bindings 后置）→
    返回 OnlyOffice 可打开的 URL（/api/oo/template/<id>），供前端加载到编辑器。
    """
    f = request.files.get('file')
    if not f:
        return jsonify({'error': '未收到文件'}), 400
    if not f.filename.lower().endswith('.docx'):
        return jsonify({'error': '仅支持 .docx 模板文件'}), 400
    name = (request.form.get('name') or '').strip() or os.path.splitext(f.filename)[0]
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', name)

    items = _load_templates_json()
    # 模板名称全局唯一：上传新模板时若名称已存在，直接拒绝
    if any(t.get('name') == name for t in items):
        return jsonify({'error': f'模板名称「{name}」已存在，请修改名称后重新上传'}), 409

    uid = uuid.uuid4().hex[:8]
    tpl_path = os.path.join(TEMPLATES_DIR, f'{uid}_{safe_name}.docx')
    try:
        f.save(tpl_path)
    except Exception as e:
        return jsonify({'error': '保存失败: ' + str(e)}), 500

    # 上传的模板占位符绑定待用户在 OnlyOffice 中插入/保存时补全；此处先登记为「未配置绑定」
    items.append({
        'id': uid,
        'name': name,
        'path': tpl_path,
        'bindings': [],
        'excel_name': '',
        'updated_at': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
    })
    _save_templates_json(items)
    return jsonify({
        'ok': True,
        'id': uid,
        'name': name,
        'url': f'/api/oo/template/{uid}',
    })


@app.route('/api/upload_template_replace', methods=['POST'])
def upload_template_replace():
    """
    修改模板时替换当前 Word 文件：不新建模板清单条目，只更新原条目的 path。
    请求: file + replace_id
    返回: {ok, id, name, url}
    """
    f = request.files.get('file')
    if not f:
        return jsonify({'error': '未收到文件'}), 400
    if not f.filename.lower().endswith('.docx'):
        return jsonify({'error': '仅支持 .docx 模板文件'}), 400
    replace_id = request.form.get('replace_id')
    if not replace_id:
        return jsonify({'error': '缺少 replace_id'}), 400

    items = _load_templates_json()
    target = next((t for t in items if t.get('id') == replace_id), None)
    if not target:
        return jsonify({'error': '模板不存在'}), 404

    name = target.get('name') or os.path.splitext(f.filename or '')[0]
    safe_name = re.sub(r'[\\/:*?"<>|]', '_', name)
    new_uid = uuid.uuid4().hex[:8]
    new_path = os.path.join(TEMPLATES_DIR, f'{new_uid}_{safe_name}.docx')
    try:
        f.save(new_path)
    except Exception as e:
        return jsonify({'error': '保存失败: ' + str(e)}), 500

    old_path = target.get('path')
    target['path'] = new_path
    target['updated_at'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    _save_templates_json(items)

    # 删除旧 docx（容错）
    try:
        if old_path and os.path.exists(old_path) and old_path != new_path:
            os.remove(old_path)
    except Exception:
        pass

    return jsonify({
        'ok': True,
        'id': replace_id,
        'name': name,
        'url': f'/api/oo/template/{replace_id}',
    })


@app.route('/api/batch_render', methods=['POST'])
def batch_render():
    """
    批量制函：选模板 + 上传已匹配的 Excel → 逐个按绑定规则注入表格 → 按函证编号命名 → 打包 zip。
    只制函前端传入的可制函项（✓ 自动匹配 与 ✎ 补到并匹配上的）；✗ 无匹配项不传入。
    请求: form 字段 template_id; file 字段（可多个 excel）; letter_no_map（JSON: {原文件名: 函证编号}）
    返回: zip 下载 url + 每个文件的状态
    """
    template_id = request.form.get('template_id')
    files = request.files.getlist('file')
    letter_no_map = {}
    ln_raw = request.form.get('letter_no_map')
    if ln_raw:
        try:
            letter_no_map = json.loads(ln_raw)
        except Exception:
            letter_no_map = {}

    if not template_id:
        return jsonify({'error': '请选择模板'}), 400
    if not files:
        return jsonify({'error': '请上传至少一个 Excel'}), 400

    items = _load_templates_json()
    tpl = next((t for t in items if t.get('id') == template_id), None)
    if not tpl:
        return jsonify({'error': '模板不存在'}), 404
    tpl_path = tpl.get('path')
    if not os.path.exists(tpl_path):
        return jsonify({'error': '模板文件不存在: ' + tpl_path}), 404

    uid = uuid.uuid4().hex[:8]
    batch_dir = os.path.join(OUTPUT_DIR, f'batch_{uid}')
    os.makedirs(batch_dir, exist_ok=True)

    results = []
    for f in files:
        if not f or not f.filename.lower().endswith(('.xlsx', '.xlsm')):
            results.append({'file': f.filename if f else '?', 'status': 'error', 'msg': '非 Excel 文件，跳过'})
            continue
        fname = f.filename
        # 输出按函证编号命名（与文档一致：输出命名=函证编号，而非 Excel 文件名）
        letter_no = letter_no_map.get(fname) or _parse_letter_no(fname)
        xlsx_tmp = os.path.join(UPLOAD_DIR, f'batch_{uuid.uuid4().hex[:8]}.xlsx')
        f.save(xlsx_tmp)
        out_path = os.path.join(batch_dir, f'{letter_no}.docx')
        try:
            stats = render_by_template(tpl_path, xlsx_tmp, out_path)
            results.append({
                'file': fname,
                'letter_no': letter_no,
                'status': 'ok',
                'injected': stats['injected'],
                'not_found': stats['not_found'],
                'skipped_empty': stats['skipped_empty'],
            })
        except Exception as e:
            results.append({'file': fname, 'status': 'error', 'msg': str(e)})
        finally:
            try:
                os.remove(xlsx_tmp)
            except Exception:
                pass

    # 打包 zip（重名自动加序号）
    zip_path = os.path.join(OUTPUT_DIR, f'batch_{uid}.zip')
    import zipfile
    used = set()
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fname in os.listdir(batch_dir):
            if not fname.endswith('.docx'):
                continue
            base, ext = os.path.splitext(fname)
            final_name = fname
            n = 1
            while final_name in used:
                final_name = f'{base}({n}){ext}'
                n += 1
            used.add(final_name)
            zf.write(os.path.join(batch_dir, fname), final_name)

    # 清理临时目录
    shutil.rmtree(batch_dir, ignore_errors=True)

    return jsonify({
        'results': results,
        'zip_url': f'/api/download/{os.path.basename(zip_path)}',
        'count': len([r for r in results if r['status'] == 'ok']),
    })


@app.route('/api/quick_render', methods=['POST'])
def quick_render():
    """
    快速导入制函：一步到位（导入 Word 模板 + 导入 Excel + 选 Sheet 位置 → 直接制函）。
    请求 form：
        word: 用户 Word 模板 .docx
        excel: 一个 Excel .xlsx（文件名含函证编号 或 前端传 letter_no）
        bindings: JSON 字符串，[{pos_index, sheet_name}] 用户选择的 Sheet 与位置
        letter_no: 可选，函证编号（默认从 Excel 文件名解析）
    流程：保存 word → annotate_bindings 写入占位符 → 扫描提取绑定 → render_by_template 注入 → 返回 docx 下载
    """
    word = request.files.get('word')
    excel = request.files.get('excel')
    bindings_raw = request.form.get('bindings') or '[]'
    letter_no = (request.form.get('letter_no') or '').strip()

    if not word or not word.filename.lower().endswith('.docx'):
        return jsonify({'error': '请上传 Word 模板（.docx）'}), 400
    if not excel or not excel.filename.lower().endswith(('.xlsx', '.xlsm')):
        return jsonify({'error': '请上传 Excel（.xlsx）'}), 400

    try:
        bindings = json.loads(bindings_raw)
    except Exception:
        bindings = []
    if not isinstance(bindings, list) or len(bindings) == 0:
        return jsonify({'error': '请至少选择一个 Sheet 位置'}), 400

    uid = uuid.uuid4().hex[:8]
    word_path = os.path.join(UPLOAD_DIR, f'quick_{uid}_tpl.docx')
    excel_path = os.path.join(UPLOAD_DIR, f'quick_{uid}.xlsx')
    try:
        word.save(word_path)
        excel.save(excel_path)
    except Exception as e:
        return jsonify({'error': '保存上传文件失败: ' + str(e)}), 500

    # 解析函证编号
    if not letter_no:
        letter_no = _parse_letter_no(excel.filename)

    # 生成带占位符的临时模板
    annotated_path = os.path.join(UPLOAD_DIR, f'quick_{uid}_annotated.docx')
    try:
        annotate_bindings(word_path, bindings, annotated_path)
        doc = Document(annotated_path)
        placeholders = extract_placeholders(doc)
        tpl_bindings = [{'sheet_name': ph['sheet_name'], 'para_index': ph['para_index']} for ph in placeholders]
    except Exception as e:
        return jsonify({'error': '模板处理失败: ' + str(e)}), 400

    # 渲染制函
    out_path = os.path.join(OUTPUT_DIR, f'quick_{uid}_{letter_no}.docx')
    try:
        stats = render_by_template(annotated_path, excel_path, out_path)
    except Exception as e:
        return jsonify({'error': '制函失败: ' + str(e)}), 400
    finally:
        # 清理临时文件
        for p in (word_path, excel_path, annotated_path):
            try:
                if os.path.exists(p):
                    os.remove(p)
            except Exception:
                pass

    dl_name = f'{letter_no}.docx'
    return jsonify({
        'ok': True,
        'letter_no': letter_no,
        'docx_url': f'/api/download/{os.path.basename(out_path)}?name={dl_name}',
        'bindings': tpl_bindings,
        'injected': stats.get('injected', []),
        'not_found': stats.get('not_found', []),
        'skipped_empty': stats.get('skipped_empty', []),
    })


# ============ OnlyOffice 文档下载接口（hanzheng 模式：docUrl 指向后端 API）============
@app.route('/api/oo/getTemplate')
def oo_get_template():
    """OnlyOffice Document Server 通过此 URL 下载模板文档。
    与 hanzheng 项目 ConfirmOnlyOfficeController.getFile 模式一致：
    后端动态返回文件流，OO 服务端从此接口下载文档。"""
    import datetime
    print(f"[OO-DL][{datetime.datetime.now()}] OnlyOffice 请求模板下载 来自: {request.remote_addr}  UA: {request.headers.get('User-Agent','')[:60]}", flush=True)
    tpl = DEFAULT_TPL
    if not os.path.exists(tpl):
        return jsonify({'error': '模板不存在: ' + tpl}), 404
    return send_file(
        tpl,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        download_name='demo_template.docx',
        as_attachment=False,  # OO 需要内联下载，不是附件
    )


# ============ OnlyOffice 静态托管（模板/标注 docx）============
# 前端 OnlyOffice 通过 /static-docx/<file> 加载模板，Vite 代理到本路由（本地同源）。
STATIC_DOCX_DIR = os.path.join(BASE_DIR, '..', 'conversion_test')

@app.route('/static-docx/<fname>')
def static_docx(fname):
    # 防目录穿越
    fname = os.path.basename(fname)
    path = os.path.join(STATIC_DOCX_DIR, fname)
    if not os.path.exists(path):
        return jsonify({'error': '文档不存在: ' + fname}), 404
    return send_file(path, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document')


HTML = """<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<title>制函渲染验证服务（Excel 直填模式 Demo）</title>
<style>
  * { box-sizing: border-box; }
  body { font-family: "Microsoft YaHei", Arial, sans-serif; margin: 0; background: #f5f7fa; color: #333; }
  header { background: #534AB7; color: #fff; padding: 16px 24px; }
  header h1 { margin: 0; font-size: 20px; }
  header p { margin: 4px 0 0; font-size: 13px; opacity: .9; }
  .container { max-width: 1200px; margin: 20px auto; padding: 0 16px; display: grid; grid-template-columns: 1fr 1fr; gap: 16px; }
  .card { background: #fff; border-radius: 8px; padding: 16px; box-shadow: 0 1px 3px rgba(0,0,0,.08); }
  .card h2 { font-size: 15px; margin: 0 0 12px; border-bottom: 2px solid #534AB7; padding-bottom: 8px; }
  .card h3 { font-size: 13px; color: #534AB7; margin: 14px 0 8px; }
  .btn { background: #534AB7; color: #fff; border: none; padding: 8px 16px; border-radius: 5px; cursor: pointer; font-size: 14px; }
  .btn:hover { background: #433A9A; }
  .btn.secondary { background: #e0e0e0; color: #333; }
  .btn.danger { background: #d9534f; }
  .file-box { border: 2px dashed #ccc; border-radius: 6px; padding: 20px; text-align: center; margin-bottom: 12px; }
  .file-box input { display: none; }
  .file-box label { cursor: pointer; color: #534AB7; font-weight: bold; }
  .info { background: #f7f7fb; border-radius: 5px; padding: 10px; font-size: 13px; margin-top: 10px; }
  .info div { margin: 2px 0; }
  .badge { display: inline-block; background: #534AB7; color: #fff; padding: 2px 8px; border-radius: 10px; font-size: 12px; margin-right: 4px; }
  .binding-item { background: #f0f0f5; border-radius: 5px; padding: 8px; margin: 6px 0; font-size: 13px; display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
  .binding-item select { padding: 3px; border: 1px solid #ccc; border-radius: 4px; }
  .result-box { margin-top: 12px; }
  .result-item { background: #e8f5e9; padding: 8px; border-radius: 5px; margin: 4px 0; font-size: 13px; }
  .result-item.err { background: #fdecea; color: #c62828; }
  a.download { display: inline-block; background: #2e7d32; color: #fff; padding: 10px 20px; border-radius: 5px; text-decoration: none; margin-top: 12px; font-weight: bold; }
  select, input[type=number] { padding: 4px; border: 1px solid #ccc; border-radius: 4px; }
  .row { display: flex; gap: 8px; flex-wrap: wrap; margin: 6px 0; }
</style>
</head>
<body>
<header>
  <h1>制函渲染验证服务（Excel 直填模式 Demo）</h1>
  <p>在 Word 模板锚点位置，按 Excel Sheet 动态生成完整表格。仅本地操作，不连接 OnlyOffice，对生产零影响。</p>
  <p style="margin-top:8px;"><a href="/demo" style="color:#fff;background:#2e7d32;padding:6px 16px;border-radius:5px;text-decoration:none;font-weight:bold;">🚀 打开 OnlyOffice 演示 Demo →</a></p>
</header>

<div class="container">
  <!-- 左列：模板 + Excel 上传 -->
  <div>
    <div class="card">
      <h2>1. 上传 Word 模板</h2>
      <div class="file-box">
        <label for="tpl-file">📄 点击选择 Word 模板 (.docx)</label>
        <input type="file" id="tpl-file" accept=".docx">
        <div id="tpl-name" style="margin-top:8px;color:#666;font-size:12px;"></div>
      </div>
      <div class="info" id="tpl-info">尚未上传模板</div>
      <h3>模板锚点位置</h3>
      <div id="tpl-anchors" style="font-size:12px;color:#999;">上传模板后可选择锚点</div>
    </div>

    <div class="card" style="margin-top:16px;">
      <h2>2. 上传 Excel 数据（多 Sheet）</h2>
      <div class="file-box">
        <label for="excel-file">📊 点击选择 Excel (.xlsx)</label>
        <input type="file" id="excel-file" accept=".xlsx,.xlsm">
        <div id="excel-name" style="margin-top:8px;color:#666;font-size:12px;"></div>
      </div>
      <div class="info" id="excel-info">尚未上传 Excel</div>
    </div>
  </div>

  <!-- 右列：绑定配置 + 渲染 -->
  <div>
    <div class="card">
      <h2>3. 配置绑定并渲染</h2>
      <div id="binding-list"></div>
      <div class="row">
        <button class="btn secondary" onclick="addBinding()">+ 添加绑定</button>
      </div>
      <div class="row" style="margin-top:12px;">
        <button class="btn" onclick="render()">🚀 渲染</button>
        <span id="render-msg" style="font-size:13px;color:#666;align-self:center;"></span>
      </div>
      <div class="result-box" id="result-box"></div>
    </div>
  </div>
</div>

<script>
let tpl = null, excel = null;
let bindings = [];

// ---- 上传模板 ----
document.getElementById('tpl-file').addEventListener('change', async function(e){
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('tpl-name').textContent = file.name;
  const fd = new FormData();
  fd.append('file', file);
  const res = await fetch('/api/upload_template', {method:'POST', body: fd});
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  tpl = data;
  document.getElementById('tpl-info').innerHTML =
    `<div><b>模板已加载</b>：段落 ${data.para_count}，表格 ${data.table_count}，文本宽 ${data.text_width_inch} 英寸</div>` +
    `<div><b>已有表格：</b>${data.tables.map(t=>`表${t.idx}(${t.rows}行×${t.cols}列)`).join('、') || '无'}</div>`;
  renderAnchors();
  updateBindings();
});

// ---- 上传 Excel ----
document.getElementById('excel-file').addEventListener('change', async function(e){
  const file = e.target.files[0];
  if (!file) return;
  document.getElementById('excel-name').textContent = file.name;
  const fd = new FormData();
  fd.append('file', file);
  const res = await fetch('/api/upload_excel', {method:'POST', body: fd});
  const data = await res.json();
  if (data.error) { alert(data.error); return; }
  excel = data;
  document.getElementById('excel-info').innerHTML =
    `<div><b>Excel 已加载</b>，共 ${data.sheets.length} 个 Sheet：</div>` +
    data.sheets.map(s=>`<div style="margin-top:4px;"><span class="badge">${s.name}</span> ${s.row_count}行数据，表头: ${s.header.join('/')}</div>`).join('');
  updateBindings();
});

// ---- 锚点选项 ----
function renderAnchors() {
  if (!tpl) { document.getElementById('tpl-anchors').textContent = '上传模板后可选择锚点'; return; }
  let html = '可用锚点：';
  html += `<div class="row">`;
  html += `<select id="anchor-type"><option value="end">文档末尾</option><option value="after_table">指定表格后</option><option value="after_para">指定段落后</option></select>`;
  html += `<input type="number" id="anchor-idx" value="0" min="0" style="width:80px;" placeholder="索引">`;
  html += `</div>`;
  html += `<div style="font-size:11px;color:#999;margin-top:4px;">表格索引: ${tpl.tables.map(t=>t.idx).join(',')} | 最大段落: ${tpl.para_count-1}</div>`;
  document.getElementById('tpl-anchors').innerHTML = html;
}

// ---- 绑定列表 ----
function addBinding() {
  if (!excel || !excel.sheets.length) { alert('请先上传 Excel'); return; }
  if (!tpl) { alert('请先上传模板'); return; }
  bindings.push({sheet_name: excel.sheets[0].name, anchor_mode: 'end', anchor_idx: 0});
  updateBindings();
}

function updateBindings() {
  const box = document.getElementById('binding-list');
  if (!bindings.length) {
    box.innerHTML = '<div style="color:#999;font-size:13px;">尚未添加绑定，点击"添加绑定"。</div>';
    return;
  }
  let html = '';
  bindings.forEach((b, i) => {
    html += `<div class="binding-item">
      <span>Sheet:</span>
      <select onchange="setBinding(${i},'sheet',this.value)">${excel.sheets.map(s=>`<option ${s.name===b.sheet_name?'selected':''}>${s.name}</option>`).join('')}</select>
      <span>位置:</span>
      <select onchange="setBinding(${i},'mode',this.value)">
        <option value="end" ${b.anchor_mode==='end'?'selected':''}>文档末尾</option>
        <option value="after_table" ${b.anchor_mode==='after_table'?'selected':''}>指定表格后</option>
        <option value="after_para" ${b.anchor_mode==='after_para'?'selected':''}>指定段落后</option>
      </select>
      <input type="number" value="${b.anchor_idx}" min="0" style="width:70px;" onchange="setBinding(${i},'idx',parseInt(this.value)||0)">
      <button class="btn danger" style="padding:2px 8px;font-size:12px;" onclick="removeBinding(${i})">删</button>
    </div>`;
  });
  box.innerHTML = html;
}

function setBinding(i, key, val) {
  if (key === 'sheet') bindings[i].sheet_name = val;
  else if (key === 'mode') { bindings[i].anchor_mode = val; }
  else if (key === 'idx') bindings[i].anchor_idx = val;
}

function removeBinding(i) {
  bindings.splice(i, 1);
  updateBindings();
}

// ---- 渲染 ----
async function render() {
  const msg = document.getElementById('render-msg');
  const box = document.getElementById('result-box');
  if (!tpl || !excel) { alert('请先上传模板和 Excel'); return; }
  if (!bindings.length) { alert('请至少添加一个绑定'); return; }

  msg.textContent = '渲染中...';
  box.innerHTML = '';
  try {
    const res = await fetch('/api/render', {
      method: 'POST',
      headers: {'Content-Type':'application/json'},
      body: JSON.stringify({tpl_path: tpl.path, excel_path: excel.path, bindings})
    });
    const data = await res.json();
    msg.textContent = '';
    if (data.error) {
      box.innerHTML = `<div class="result-item err">渲染失败：${data.error}</div>`;
      return;
    }
    let html = `<div class="result-item"><b>✅ 渲染成功</b> | 文本宽 ${data.text_width_inch} 英寸</div>`;
    html += `<div class="result-item">模板原有：段落 ${data.template_para_count}，表格 ${data.template_table_count}</div>`;
    html += `<div class="result-item">渲染后表格总数：${data.output_table_count}（新增 ${data.injected.length} 个）</div>`;
    data.injected.forEach(inj => {
      html += `<div class="result-item">📊 Sheet「${inj.sheet}」→ ${inj.cols}列 × ${inj.rows}行</div>`;
    });
    if (data.skipped_empty.length) html += `<div class="result-item err">已跳过空 Sheet：${data.skipped_empty.join(',')}</div>`;
    if (data.not_found.length) html += `<div class="result-item err">未找到 Sheet：${data.not_found.join(',')}</div>`;
    html += `<a class="download" href="${data.download_url}">⬇️ 下载渲染结果 (.docx)</a>`;
    html += `<div style="font-size:12px;color:#666;margin-top:8px;">下载后用 Word/OnlyOffice 打开，查看表格是否正确生成（列数=Sheet列数、表头正确、总宽=文本宽）</div>`;
    box.innerHTML = html;
  } catch (e) {
    msg.textContent = '';
    box.innerHTML = `<div class="result-item err">异常：${e.message}</div>`;
  }
}
</script>
</body>
</html>
"""

if __name__ == '__main__':
    print("=" * 50)
    print("制函渲染验证服务已启动")
    print("打开浏览器访问: http://127.0.0.1:5002")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5002, debug=False)
