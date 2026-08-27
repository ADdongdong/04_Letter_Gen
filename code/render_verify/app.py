# -*- coding: utf-8 -*-
"""
制函渲染验证服务 - Flask Web 应用
启动: python app.py  → http://127.0.0.1:5002
"""
import io, os, uuid, json, shutil, socket
from flask import Flask, request, render_template_string, jsonify, send_file, url_for

from render_engine import (
    read_excel_sheets, render_template, get_text_width, _table_width_twips,
    extract_anchors, annotate_bindings,
)
from docx import Document

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, 'uploads')
OUTPUT_DIR = os.path.join(BASE_DIR, 'outputs')
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

DEFAULT_TPL = os.path.join(BASE_DIR, '..', 'conversion_test', 'demo_template.docx')
# 用户上传的 Word 模板（覆盖内置默认模板）
TEMPLATE_UPLOAD_DIR = os.path.join(UPLOAD_DIR, 'templates')
os.makedirs(TEMPLATE_UPLOAD_DIR, exist_ok=True)
CURRENT_TEMPLATE = os.path.join(TEMPLATE_UPLOAD_DIR, 'current.docx')


def get_current_template():
    """返回当前生效的模板路径：用户上传过的优先，否则用内置默认模板。"""
    if os.path.exists(CURRENT_TEMPLATE) and os.path.getsize(CURRENT_TEMPLATE) > 0:
        return CURRENT_TEMPLATE
    return DEFAULT_TPL

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


# ============ 临时调试：OO 内部 API 探测结果接收（开发期使用，可删） ============
@app.route('/api/probe', methods=['POST'])
def probe_receive():
    data = request.get_json(silent=True) or {}
    probe_file = os.path.join(BASE_DIR, 'probe_result.json')
    with open(probe_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return jsonify({'ok': True})


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

    # 同步为"当前模板"：此后 /api/oo/getTemplate 与 /api/render 均使用它
    shutil.copyfile(path, CURRENT_TEMPLATE)
    print(f"[TPL-UP] 模板已上传并生效: {f.filename} -> {CURRENT_TEMPLATE}", flush=True)

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
        'name': f.filename,
        'size_kb': round(os.path.getsize(path) / 1024, 1),
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
    tpl = get_current_template()
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
    tpl = data.get('tpl_path') or get_current_template()
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
    tpl_path = data.get('tpl_path') or get_current_template()
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


# ============ OnlyOffice 文档下载接口（hanzheng 模式：docUrl 指向后端 API）============
@app.route('/api/oo/getTemplate')
def oo_get_template():
    """OnlyOffice Document Server 通过此 URL 下载模板文档。
    与 hanzheng 项目 ConfirmOnlyOfficeController.getFile 模式一致：
    后端动态返回文件流，OO 服务端从此接口下载文档。"""
    import datetime
    print(f"[OO-DL][{datetime.datetime.now()}] OnlyOffice 请求模板下载 来自: {request.remote_addr}  UA: {request.headers.get('User-Agent','')[:60]}", flush=True)
    tpl = get_current_template()
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
