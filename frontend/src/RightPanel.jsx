import React, { useState, useEffect, useRef } from 'react';
import SelectField from './SelectField.jsx';

/**
 * 右侧操作面板 —— 精简版（与 HanZheng 一致的交互）
 *
 * 流程：
 *   1. 上传 Excel → 解析 Sheet 列表
 *   2. 点击某个 Sheet → 直接在左侧 OnlyOffice 当前光标处插入标注
 *   3. 点"生成正式函证" → 后端将标注文字替换为真实 Excel 表格
 *
 * 所有状态变化封装在此组件内，不影响左侧 OnlyOffice。
 */
// 表格样式默认值（与后端 render_engine.DEFAULT_TABLE_STYLE 一致）
const DEFAULT_TABLE_STYLE = {
  font_size: 10, font_name: '', h_align: 'center', v_align: 'top',
  row_height_mode: 'auto', row_height_pt: 20,
  col_widths_mode: 'equal', col_widths: {}, col_widths_ratio: '',
};

export default function RightPanel({ ready, insertText, forceSave, updateStatus, onTemplateUploaded, editingId, editingName, templateStyle, templateSheets, onTemplateSaved }) {
  const [sheets, setSheets] = useState([]);
  const [excelPath, setExcelPath] = useState(null);
  const [excelInfo, setExcelInfo] = useState('尚未上传 Excel');
  const [insertedSheets, setInsertedSheets] = useState(new Set()); // 记录已插入的 sheet
  const [tplInfo, setTplInfo] = useState('未上传（使用内置默认模板）');
  const [tplName, setTplName] = useState('');   // 模板配置名称
  const [saving, setSaving] = useState(false);  // 保存模板配置进行中
  // 表格样式（高级可选）：默认收起；跟随模板配置保存，制函渲染时生效
  const [styleCfg, setStyleCfg] = useState(DEFAULT_TABLE_STYLE);
  const [styleOpen, setStyleOpen] = useState(false);
  const upStyle = (k, v) => setStyleCfg(s => ({ ...s, [k]: v }));

  // 编辑模式回显已存样式；新建模式重置默认
  useEffect(() => {
    if (!editingId) { setStyleCfg(DEFAULT_TABLE_STYLE); return; }
    if (templateStyle && typeof templateStyle === 'object') {
      setStyleCfg({ ...DEFAULT_TABLE_STYLE, ...templateStyle });
    }
  }, [editingId, templateStyle]);

  // 模板名称自动带出：编辑模式填入当前模板名（可直接修改后保存）；新建模式为空
  useEffect(() => {
    setTplName(editingName || '');
  }, [editingName]);

  // 列宽按 Sheet 配置的 Sheet 名单：上传的 Excel ∪ 编辑回显的模板 sheets ∪ 已存配置键（去重保序）
  const sheetNamesForWidth = (() => {
    const names = [];
    const push = (n) => { if (n && !names.includes(n)) names.push(n); };
    (sheets || []).forEach(s => push(s.name));
    (templateSheets || []).forEach(s => push(typeof s === 'string' ? s : (s && s.name)));
    Object.keys(styleCfg.col_widths || {}).forEach(push);
    return names;
  })();
  const colCountOf = (name) => {
    const s = (sheets || []).find(x => x.name === name);
    return s && Array.isArray(s.header) ? s.header.length : null;
  };

  // ---- 上传 Word 模板 → 后端保存 → 通知 App 换 key 重载编辑器 ----
  const onTemplateChange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setTplInfo('上传中...');
    const fd = new FormData();
    fd.append('file', file);
    const res = await fetch('/api/upload_template', { method: 'POST', body: fd });
    const data = await res.json();
    if (data.error) { alert(data.error); setTplInfo('上传失败，仍使用原模板'); return; }
    setTplInfo(`✓ 已上传「${data.name}」（${data.size_kb}KB），编辑器重载中...`);
    updateStatus('✓ 模板已上传，左侧编辑器重载中...');
    if (onTemplateUploaded) onTemplateUploaded();
  };

  // ---- 上传 Excel ----
  const onExcelChange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setExcelInfo('上传中...');
    const fd = new FormData();
    fd.append('file', file);
    const res = await fetch('/api/upload_excel', { method: 'POST', body: fd });
    const data = await res.json();
    if (data.error) { alert(data.error); setExcelInfo('上传失败'); return; }
    setExcelPath(data.path);
    setSheets(data.sheets);
    setInsertedSheets(new Set());
    const groupedCount = data.sheets.filter(s => s.is_grouped).length;
    setExcelInfo(`已加载：${data.sheets.length} 个 Sheet（${file.name}），其中 ${groupedCount} 个为分组结构 Sheet`);
    updateStatus(`✓ Excel 已加载：${data.sheets.length} 个 Sheet（${groupedCount} 个分组 Sheet）`);
  };

  // ---- 点击 Sheet → 在左侧 OnlyOffice 光标处插入标注 ----
  const onSheetClick = (sheet) => {
    if (!ready) { alert('编辑器尚未就绪，请稍候'); return; }
    const text = `【Sheet「${sheet.name}」表格将在此处展示】`;
    const ok = insertText(text);
    if (!ok) {
      updateStatus('⚠ 插入失败：请在左侧模板中点一下再重试', true);
      return;
    }
    updateStatus(`✓ 已在光标处插入「${sheet.name}」标注`);
    setInsertedSheets((prev) => new Set([...prev, sheet.name]));
  };

  // ---- 保存模板配置：forceSave 回写 current.docx → 复制到模板配置目录 ----
  const saveTemplate = async () => {
    const name = tplName.trim();
    if (!name) { alert('请填写模板名称'); return; }
    // 新建模式必须上传 Excel（定义 Sheet 结构）；编辑模式可选——未重新上传则沿用该模板已存的 Excel
    if (!editingId && !excelPath) { alert('请先上传 Excel（用于定义 Sheet 结构，可以没有数据）'); return; }
    if (saving) return;
    setSaving(true);

    try {
      // 1. forceSave 把 OO 当前内容（含占位段标注）回写到 current.docx
      updateStatus('正在保存模板（含标注）到后端...');
      let beforeTs = 0;
      try {
        const br = await fetch('/api/oo/save_status');
        const bd = await br.json();
        beforeTs = bd.saved_ts || 0;
      } catch (e) {}
      let fsRes = null;
      try { fsRes = forceSave ? await forceSave() : null; } catch (e) {}
      if (!fsRes || fsRes.ok === false) {
        alert('保存失败：无法触发强制保存（编辑器未就绪或 key 缺失）');
        updateStatus('保存失败', true);
        return;
      }

      // no_changes：文档自上次回写后无新修改，current.docx 已是最新，跳过轮询
      if (!fsRes.no_changes) {
        // 2. 轮询 save_status 等回写完成（最多 20 秒）
        let saved = false;
        for (let i = 0; i < 40; i++) {
          await new Promise(r => setTimeout(r, 500));
          try {
            const sr = await fetch('/api/oo/save_status');
            const sd = await sr.json();
            if (sd.saved_ts > beforeTs) { saved = true; break; }
          } catch (e) {}
        }
        if (!saved) {
          alert('模板保存超时，请确认编辑器已就绪后重试');
          updateStatus('模板保存超时', true);
          return;
        }
      }

      // 3. 落盘为模板配置（id 空=新建；编辑模式带 editingId；excel_path 缺省=沿用已存 Excel）
      const res = await fetch('/api/templates/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: editingId || undefined, name, excel_path: excelPath || undefined, style: styleCfg }),
      });
      const data = await res.json();
      if (data.error) { alert(data.error); updateStatus('保存失败', true); return; }
      updateStatus(`✓ 模板配置「${data.name}」已保存（含 ${data.sheets.length} 个 Sheet），可到制函页使用`);
      if (onTemplateSaved) onTemplateSaved();  // 刷新配置页顶部模板列表
      alert(`模板配置「${data.name}」已保存`);
    } catch (e) {
      alert('保存失败：' + e.message);
      updateStatus('保存失败', true);
    } finally {
      setSaving(false);
    }
  };

  const insertedCount = insertedSheets.size;

  return (
    <div className="right">
      <h2>1. 上传 Word 模板（可选）</h2>
      <div className="file-box">
        <label htmlFor="tpl-file">📝 点击选择 Word 模板 (.docx)</label>
        <input id="tpl-file" type="file" accept=".docx" onChange={onTemplateChange} />
      </div>
      <div className="info">
        {tplInfo}
        {!tplInfo && (
          <div style={{ marginTop: '4px' }}>
            没有现成模板？<a href="/api/templates/sample/word" download>📥 下载示例 Word 模板</a>，修改内容后上传即可
          </div>
        )}
      </div>

      <h2 style={{ marginTop: '16px' }}>2. 上传 Excel（多 Sheet）</h2>
      <div className="file-box">
        <label htmlFor="excel-file">📊 点击选择 Excel (.xlsx)</label>
        <input id="excel-file" type="file" accept=".xlsx,.xlsm" onChange={onExcelChange} />
      </div>
      <div className="info">{excelInfo}</div>

      {sheets.length > 0 && (
        <>
          <h2>3. 可用 Excel Sheet（点击插入表格）</h2>
          <div className="hint" style={{ marginBottom: '10px' }}>
            操作步骤：
            <br />① 先<strong>在左侧模板中点一下</strong>定位插入位置（只需一次）
            <br />② 点击下方某个 <strong>Sheet</strong> → 自动插入标注
            <br />③ 可在不同位置重复点击插入多张表格
          </div>

          <div className="sheet-list">
            {sheets.map((s, i) => {
              const isInserted = insertedSheets.has(s.name);
              return (
                <div
                  key={s.name}
                  className={`sheet-item clickable ${isInserted ? 'inserted' : ''}`}
                  onClick={() => onSheetClick(s)}
                  title={isInserted ? '已插入（可再次点击重复插入）' : `点击插入「${s.name}」表格`}
                >
                  <span className="badge">{s.name}</span>
                  <span>{s.row_count} 行数据</span>
                  <span style={{ color: '#888', fontSize: '11px' }}>
                    表头: {(s.header || []).slice(0, 4).join('/')}
                    {(s.header || []).length > 4 ? '...' : ''}
                  </span>
                  <span style={{ marginLeft: 'auto', fontSize: '11px' }}>
                    {isInserted
                      ? <span style={{ color: '#2e7d32' }}>✅ 已插入</span>
                      : <span style={{ color: 'var(--primary)' }}>👆 点击插入</span>
                    }
                  </span>
                </div>
              );
            })}
          </div>

          <div className="hint" style={{ marginTop: '12px' }}>
            💡 提示：每次点击 Sheet，文字会插入到模板中<strong>当前光标所在位置</strong>。
            若插入失败，请先在模板中点一下再重试。制函时所有
            <code style={{ background: '#eee', padding: '1px 4px' }}>【Sheet「...」】</code>
            标注会被替换为真实表格。
          </div>
        </>
      )}

      {/* 保存模板配置：始终可见（Excel 可以没数据，但必须有 Sheet 名称） */}
      <div style={{ marginTop: '20px', padding: '14px', background: 'var(--primary-bg-light)', borderRadius: '8px', border: '1px solid var(--primary-border)' }}>
        <h3 style={{ margin: '0 0 10px' }}>
          {editingId
            ? '修改模板配置' + (editingName ? '：「' + editingName + '」' : '')
            : '保存为模板配置'}
        </h3>
        <input
          type="text"
          placeholder="模板名称（如：往来非标询证函）"
          value={tplName}
          onChange={e => setTplName(e.target.value)}
          style={{ width: '100%', padding: '8px', boxSizing: 'border-box', marginBottom: '10px' }}
        />
        {/* 表格样式（高级，可选）：默认收起；全部带默认值，不改则与现状渲染一致。
            容器无边框（去框线层级），标题行用分隔线与内容区分隔 */}
        <div style={{ marginBottom: '10px' }}>
          <div
            onClick={() => setStyleOpen(o => !o)}
            style={{ padding: '6px 0', cursor: 'pointer', fontSize: '13px', fontWeight: 600, userSelect: 'none', display: 'flex', justifyContent: 'space-between', alignItems: 'center', borderBottom: '1px solid var(--primary-border)' }}
          >
            <span>表格样式（高级，可选）</span>
            <span className={styleOpen ? 'sfield-arrow open' : 'sfield-arrow'} style={{ fontSize: '13px' }}>▾</span>
          </div>
          {styleOpen && (
            <div style={{ padding: '10px 0 0', display: 'grid', gridTemplateColumns: '1fr 1fr', gap: '8px' }}>
              <label style={{ fontSize: '12px', color: '#555' }}>
                字号（磅，默认 10）
                <input
                  type="number" min={5} max={72} step={0.5}
                  value={styleCfg.font_size}
                  onChange={e => {
                    const v = parseFloat(e.target.value);
                    if (!isNaN(v)) upStyle('font_size', Math.min(72, Math.max(5, v)));
                  }}
                  style={{ width: '100%', marginTop: '2px' }}
                />
              </label>
              <label style={{ fontSize: '12px', color: '#555' }}>
                字体（默认继承模板）
                <SelectField
                  value={styleCfg.font_name}
                  onChange={v => upStyle('font_name', v)}
                  options={[
                    { value: '', label: '继承模板' },
                    { value: '宋体', label: '宋体' },
                    { value: '仿宋', label: '仿宋' },
                    { value: '黑体', label: '黑体' },
                    { value: '楷体', label: '楷体' },
                    { value: '微软雅黑', label: '微软雅黑' },
                    { value: '等线', label: '等线' },
                    { value: 'Arial', label: 'Arial' },
                  ]}
                />
              </label>
              <label style={{ fontSize: '12px', color: '#555' }}>
                水平对齐（默认居中）
                <SelectField
                  value={styleCfg.h_align}
                  onChange={v => upStyle('h_align', v)}
                  options={[
                    { value: 'center', label: '居中' },
                    { value: 'left', label: '左对齐' },
                    { value: 'right', label: '右对齐' },
                  ]}
                />
              </label>
              <label style={{ fontSize: '12px', color: '#555' }}>
                垂直对齐（默认顶端）
                <SelectField
                  value={styleCfg.v_align}
                  onChange={v => upStyle('v_align', v)}
                  options={[
                    { value: 'top', label: '顶端' },
                    { value: 'center', label: '居中' },
                  ]}
                />
              </label>
              <label style={{ fontSize: '12px', color: '#555' }}>
                行高（默认自动）
                <SelectField
                  value={styleCfg.row_height_mode}
                  onChange={v => upStyle('row_height_mode', v)}
                  options={[
                    { value: 'auto', label: '自动' },
                    { value: 'atLeast', label: '最小值' },
                  ]}
                />
              </label>
              <label style={{ fontSize: '12px', color: '#555' }}>
                {styleCfg.row_height_mode === 'atLeast' ? '最小行高（磅）' : '行高磅数（自动时不生效）'}
                <input
                  type="number" min={5} max={200}
                  value={styleCfg.row_height_pt}
                  onChange={e => upStyle('row_height_pt', parseFloat(e.target.value) || 20)}
                  disabled={styleCfg.row_height_mode !== 'atLeast'}
                  style={{ width: '100%', marginTop: '2px', opacity: styleCfg.row_height_mode === 'atLeast' ? 1 : 0.5 }}
                />
              </label>
              <label style={{ fontSize: '12px', color: '#555', gridColumn: '1 / -1' }}>
                列宽（默认均分）
                <SelectField
                  value={styleCfg.col_widths_mode}
                  onChange={v => upStyle('col_widths_mode', v)}
                  options={[
                    { value: 'equal', label: '均分' },
                    { value: 'auto', label: '按内容自适应' },
                    { value: 'ratio', label: '自定义比例（按 Sheet）' },
                  ]}
                />
              </label>
              {styleCfg.col_widths_mode === 'auto' && (
                <div style={{ fontSize: '12px', color: '#888', gridColumn: '1 / -1' }}>
                  按每列的表头与数据内容长度自动分配列宽（中文按 2 倍宽计），无需配置。
                </div>
              )}
              {styleCfg.col_widths_mode === 'ratio' && (
                <div style={{ gridColumn: '1 / -1', display: 'grid', gap: '6px' }}>
                  <div style={{ fontSize: '12px', color: '#888' }}>
                    按 Sheet 配置列宽比例（冒号分隔，比例段数需与该 Sheet 列数一致；留空的 Sheet 按均分）：
                  </div>
                  {sheetNamesForWidth.length === 0 && (
                    <div style={{ fontSize: '12px', color: '#888' }}>请先上传 Excel 以获取 Sheet 列表</div>
                  )}
                  {sheetNamesForWidth.map(name => {
                    const n = colCountOf(name);
                    const fullLabel = name + (n ? `（${n}列）` : '');
                    return (
                      <label key={name} title={fullLabel} style={{ fontSize: '12px', color: '#555', display: 'flex', alignItems: 'center', gap: '8px' }}>
                        <span style={{ flex: '0 0 150px', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>
                          {fullLabel}
                        </span>
                        <input
                          type="text" placeholder="如 2:1:1（留空=均分）"
                          value={(styleCfg.col_widths || {})[name] || ''}
                          onChange={e => upStyle('col_widths', { ...(styleCfg.col_widths || {}), [name]: e.target.value })}
                          style={{ flex: 1, minWidth: 0, marginTop: 0 }}
                        />
                      </label>
                    );
                  })}
                </div>
              )}
            </div>
          )}
        </div>
        <button
          className="btn primary"
          onClick={saveTemplate}
          disabled={saving}
          style={{ width: '100%', padding: '10px' }}
        >
          {saving ? '保存中...' : (editingId ? '保存修改' : '保存模板配置')}
        </button>
        <div className="hint" style={{ marginTop: '10px' }}>
          {editingId ? '正在编辑已有模板：保存将覆盖该模板的 Word；未重新上传 Excel 则沿用原 Sheet 结构。' : '保存后可到「Excel 制函」页选择此模板批量制函。'}
          保存时自动把左侧编辑器当前内容（含占位段标注）回写为模板 Word 文件；上传的 Excel 仅用于定义 Sheet 结构，可以没有数据。
        </div>
      </div>
    </div>
  );
}
