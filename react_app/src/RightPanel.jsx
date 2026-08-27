import React, { useState, useEffect } from 'react';

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
export default function RightPanel({ ready, insertText, updateStatus, onTemplateUploaded }) {
  const [sheets, setSheets] = useState([]);
  const [excelPath, setExcelPath] = useState(null);
  const [excelInfo, setExcelInfo] = useState('尚未上传 Excel');
  const [insertedSheets, setInsertedSheets] = useState(new Set()); // 记录已插入的 sheet
  const [tplInfo, setTplInfo] = useState('未上传（使用内置默认模板）');

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
    setExcelInfo(`已加载：${data.sheets.length} 个 Sheet（${file.name}）`);
    updateStatus(`✓ Excel 已加载：${data.sheets.length} 个 Sheet`);
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

  // ---- 生成正式函证 ----
  const renderDoc = async () => {
    if (!excelPath) { alert('请先上传 Excel'); return; }
    if (insertedSheets.size === 0) { alert('请先点击至少一个 Sheet，在模板中插入标注'); return; }

    const bindings = Array.from(insertedSheets).map((sheetName, idx) => ({
      pos_index: idx,
      anchor_mode: 'after_para',
      sheet_name: sheetName,
      pos_label: `手动插入-${sheetName}`,
    }));

    updateStatus(`渲染中...（${bindings.length} 个表格）`);
    const res = await fetch('/api/render', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        tpl_path: null,
        excel_path: excelPath,
        bindings: bindings,
      }),
    });
    const data = await res.json();
    if (data.error) { alert(data.error); updateStatus('渲染失败'); return; }
    updateStatus(`✓ 渲染成功！共 ${data.output_table_count} 个表格`);
    window.open(data.download_url, '_blank');
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

          <div className="actions" style={{ marginTop: '16px' }}>
            <button
              className="btn primary"
              disabled={insertedCount === 0}
              onClick={renderDoc}
              style={{ opacity: insertedCount === 0 ? 0.5 : 1 }}
            >
              🚀 生成正式函证{insertedCount > 0 ? `（${insertedCount} 个表格）` : ''}
            </button>
          </div>

          <div className="hint" style={{ marginTop: '12px' }}>
            💡 提示：每次点击 Sheet，文字会插入到模板中<strong>当前光标所在位置</strong>。
            若插入失败，请先在模板中点一下再重试。生成函证时所有
            <code style={{ background: '#eee', padding: '1px 4px' }}>【Sheet「...」】</code>
            标注会被替换为真实表格。
          </div>
        </>
      )}
    </div>
  );
}
