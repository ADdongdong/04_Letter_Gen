import React, { useState, useEffect, useRef } from 'react';

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
export default function RightPanel({ ready, insertText, forceSave, updateStatus, onTemplateUploaded, editingId, editingName, onTemplateSaved }) {
  const [sheets, setSheets] = useState([]);
  const [excelPath, setExcelPath] = useState(null);
  const [excelInfo, setExcelInfo] = useState('尚未上传 Excel');
  const [insertedSheets, setInsertedSheets] = useState(new Set()); // 记录已插入的 sheet
  const [tplInfo, setTplInfo] = useState('未上传（使用内置默认模板）');
  const [tplName, setTplName] = useState('');   // 模板配置名称
  const [saving, setSaving] = useState(false);  // 保存模板配置进行中

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
    if (!excelPath) { alert('请先上传 Excel（用于定义 Sheet 结构，可以没有数据）'); return; }
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

      // 3. 落盘为模板配置（id 空=新建；编辑模式带 editingId）
      const res = await fetch('/api/templates/save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: editingId || undefined, name, excel_path: excelPath }),
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
      <h2>0. 上传 Word 模板（可选）</h2>
      <div className="file-box">
        <label htmlFor="tpl-file">📝 点击选择 Word 模板 (.docx)</label>
        <input id="tpl-file" type="file" accept=".docx" onChange={onTemplateChange} />
      </div>
      <div className="info">{tplInfo}</div>

      <h2 style={{ marginTop: '16px' }}>1. 上传 Excel（多 Sheet）</h2>
      <div className="file-box">
        <label htmlFor="excel-file">📊 点击选择 Excel (.xlsx)</label>
        <input id="excel-file" type="file" accept=".xlsx,.xlsm" onChange={onExcelChange} />
      </div>
      <div className="info">{excelInfo}</div>

      {sheets.length > 0 && (
        <>
          <h2>2. 可用 Excel Sheet（点击插入表格）</h2>
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
                      : <span style={{ color: '#534AB7' }}>👆 点击插入</span>
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
      <div style={{ marginTop: '20px', padding: '14px', background: '#f7f7fb', borderRadius: '8px', border: '1px solid #e3e0f5' }}>
        <h3 style={{ margin: '0 0 10px' }}>
          {editingId
            ? '✏️ 修改模板配置' + (editingName ? '：「' + editingName + '」' : '')
            : '💾 保存为模板配置'}
        </h3>
        <input
          type="text"
          placeholder="模板名称（如：往来非标询证函）"
          value={tplName}
          onChange={e => setTplName(e.target.value)}
          style={{ width: '100%', padding: '8px', boxSizing: 'border-box', marginBottom: '10px' }}
        />
        <button
          className="btn primary"
          onClick={saveTemplate}
          disabled={saving}
          style={{ width: '100%', padding: '10px' }}
        >
          {saving ? '⏳ 保存中...' : (editingId ? '💾 保存修改' : '💾 保存模板配置')}
        </button>
        <div className="hint" style={{ marginTop: '10px' }}>
          {editingId ? '正在编辑已有模板，保存后将覆盖该模板的 Word 与 Excel。' : '保存后可到「Excel 制函」页选择此模板批量制函。'}
          保存时自动把左侧编辑器当前内容（含占位段标注）回写为模板 Word 文件；上传的 Excel 仅用于定义 Sheet 结构，可以没有数据。
        </div>
      </div>
    </div>
  );
}
