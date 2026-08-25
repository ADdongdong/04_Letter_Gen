import React, { useState, useRef } from 'react';
import './styles.css';

/**
 * 快速导入Excel制函向导：
 *  Step 1: 导入 Word 模板（.docx）
 *  Step 2: 导入 Excel（.xlsx）→ 选择 Sheet 位置
 *  Step 3: 直接制函 → 下载 docx
 */
export default function QuickImportModal({ onClose }) {
  const [step, setStep] = useState(1);
  const [wordFile, setWordFile] = useState(null);
  const [excelFile, setExcelFile] = useState(null);
  const [sheets, setSheets] = useState([]);       // [{name}]
  const [selected, setSelected] = useState(new Set()); // 选中的 sheet 名集合
  const [letterNo, setLetterNo] = useState('');    // 函证编号（可编辑）
  const [renderStatus, setRenderStatus] = useState('');
  const [result, setResult] = useState(null);      // {docx_url, letter_no}
  const [busy, setBusy] = useState(false);
  const excelInputRef = useRef(null);

  // Step 1: 上传 Word
  const onWordChange = async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    if (!f.name.toLowerCase().endswith('.docx')) {
      alert('请选择 Word 模板文件（.docx）');
      return;
    }
    setWordFile(f);
    setStep(2);
    e.target.value = '';
  };

  // Step 2: 上传 Excel → 解析 sheet 列表
  const onExcelChange = async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    if (!f.name.toLowerCase().endswith(('.xlsx', '.xlsm'))) {
      alert('请选择 Excel 文件（.xlsx）');
      return;
    }
    setExcelFile(f);
    setBusy(true);
    setRenderStatus('正在解析 Excel 中的 Sheet…');
    try {
      const fd = new FormData();
      fd.append('file', f);
      const res = await fetch('/api/excel_sheets', { method: 'POST', body: fd });
      const data = await res.json();
      if (data.error) { alert('解析失败：' + data.error); return; }
      setSheets(data.sheets || []);
      setSelected(new Set());
      // 默认函证编号从文件名解析
      const m = f.name.match(/([A-Za-z0-9_-]+)/);
      setLetterNo(m ? m[1] : '');
    } catch (err) {
      alert('解析失败：' + err.message);
    } finally {
      setBusy(false);
      setRenderStatus('');
      e.target.value = '';
    }
  };

  const toggleSheet = (name) => {
    setSelected((prev) => {
      const n = new Set(prev);
      if (n.has(name)) n.delete(name);
      else n.add(name);
      return n;
    });
  };

  // Step 3: 制函
  const onRender = async () => {
    if (!wordFile || !excelFile) { alert('请先导入 Word 模板和 Excel'); return; }
    if (selected.size === 0) { alert('请至少选择一个 Sheet 位置'); return; }
    setBusy(true);
    setRenderStatus('正在制函…');
    try {
      const fd = new FormData();
      fd.append('word', wordFile);
      fd.append('excel', excelFile);
      fd.append('letter_no', letterNo);
      const bindings = Array.from(selected).map((name, i) => ({ pos_index: i, sheet_name: name }));
      fd.append('bindings', JSON.stringify(bindings));
      const res = await fetch('/api/quick_render', { method: 'POST', body: fd });
      const data = await res.json();
      if (data.error) { alert('制函失败：' + data.error); return; }
      setResult({ docx_url: data.docx_url, letter_no: data.letter_no });
      setStep(3);
    } catch (err) {
      alert('制函失败：' + err.message);
    } finally {
      setBusy(false);
      setRenderStatus('');
    }
  };

  return (
    <div className="modal-backdrop" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="quick-modal">
        <div className="quick-header">
          <span className="quick-title">⚡ 快速导入Excel制函</span>
          <button className="modal-close" onClick={onClose}>×</button>
        </div>

        {/* 步骤指示 */}
        <div className="quick-steps">
          <div className={'quick-step ' + (step >= 1 ? 'active' : '')}><span>1</span> 导入 Word 模板</div>
          <div className={'quick-step ' + (step >= 2 ? 'active' : '')}><span>2</span> 导入 Excel 选 Sheet</div>
          <div className={'quick-step ' + (step >= 3 ? 'active' : '')}><span>3</span> 制函</div>
        </div>

        <div className="quick-body">
          {/* Step 1 */}
          <div className="quick-block">
            <h3>① 导入 Word 模板</h3>
            <div className="file-box">
              <label htmlFor="quick-word">📄 点击选择 Word 模板 (.docx)</label>
              <input id="quick-word" type="file" accept=".docx" onChange={onWordChange} />
            </div>
            {wordFile && <div className="ok-tip">✓ 已加载：{wordFile.name}</div>}
          </div>

          {/* Step 2 */}
          <div className="quick-block">
            <h3>② 导入 Excel 并选择 Sheet 位置</h3>
            <div className="file-box">
              <label htmlFor="quick-excel">📊 点击选择 Excel (.xlsx)</label>
              <input id="quick-excel" type="file" accept=".xlsx,.xlsm" onChange={onExcelChange} ref={excelInputRef} />
            </div>
            {excelFile && <div className="ok-tip">✓ 已加载：{excelFile.name}</div>}

            {sheets.length > 0 && (
              <div className="sheet-picker">
                <div className="sheet-picker-hint">选择要插入的 Sheet：</div>
                {sheets.map((s) => (
                  <label key={s.name} className="sheet-check">
                    <input
                      type="checkbox"
                      checked={selected.has(s.name)}
                      onChange={() => toggleSheet(s.name)}
                    />
                    {s.name}
                  </label>
                ))}
              </div>
            )}

            {excelFile && (
              <div className="letter-no-row">
                <label>函证编号：</label>
                <input
                  value={letterNo}
                  onChange={(e) => setLetterNo(e.target.value)}
                  placeholder="默认从文件名解析"
                />
              </div>
            )}
          </div>

          {/* Step 3 */}
          {step >= 3 && result && (
            <div className="quick-block">
              <h3>③ 制函完成</h3>
              <a className="download-btn" href={result.docx_url} download>📥 下载制函结果（{result.letter_no}.docx）</a>
            </div>
          )}
        </div>

        <div className="quick-footer">
          {step < 3 && (
            <button
              className="btn primary"
              disabled={busy || !wordFile || !excelFile || selected.size === 0}
              onClick={onRender}
            >
              {busy ? '处理中…' : '直接制函'}
            </button>
          )}
          {step >= 3 && (
            <>
              <button className="btn" onClick={() => { setStep(1); setWordFile(null); setExcelFile(null); setSheets([]); setSelected(new Set()); setResult(null); }}>再制一份</button>
              <button className="btn primary" onClick={onClose}>完成</button>
            </>
          )}
        </div>

        {renderStatus && <div className="quick-status">{renderStatus}</div>}
      </div>
    </div>
  );
}
