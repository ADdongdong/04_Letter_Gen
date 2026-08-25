import React, { useState, useEffect } from 'react';

/**
 * Excel 批量制函对话框（复刻 docs/proto/batch_wizard_step2_demo.html）。
 * 纯对话框：① 选模板 + ② 上传匹配同页，无 OnlyOffice。
 * 三态匹配：✓ 自动匹配 / ✎ 待补编号 / ✗ 无匹配；支持手动补编号重匹配、单条删除；
 * 开始制函仅制函已匹配项，底部显示「可制函 N 份」。
 */
export default function BatchWizardModal({ onClose }) {
  const [templates, setTemplates] = useState([]);
  const [selTemplate, setSelTemplate] = useState('');
  // matches: [{key, file, letter_no, status('ok'|'pending'|'fail'), ledger}]
  const [matches, setMatches] = useState([]);
  const [matching, setMatching] = useState(false);
  const [batchResults, setBatchResults] = useState(null);
  const [rendering, setRendering] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const res = await fetch('/api/templates');
        const data = await res.json();
        setTemplates(data.templates || []);
      } catch (e) { /* ignore */ }
    })();
  }, []);

  const onBatchFilesChange = async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setMatching(true);
    setBatchResults(null);
    try {
      const fd = new FormData();
      files.forEach((f) => fd.append('file', f));
      const res = await fetch('/api/match_excel', { method: 'POST', body: fd });
      const data = await res.json();
      if (data.error) { alert('匹配失败：' + data.error); return; }
      const keyed = data.matches.map((m) => {
        const f = files.find((x) => x.name === m.file);
        return { ...m, key: m.file + '|' + Math.random().toString(36).slice(2, 8), file: f };
      });
      setMatches((prev) => [...prev, ...keyed]);
    } catch (err) {
      alert('匹配失败：' + err.message);
    } finally {
      setMatching(false);
      e.target.value = '';
    }
  };

  const rematchOne = async (key, manualNo) => {
    const item = matches.find((m) => m.key === key);
    if (!item) return;
    if (!manualNo || !manualNo.trim()) { alert('请输入函证编号'); return; }
    setMatching(true);
    try {
      const fd = new FormData();
      if (item.file) fd.append('file', item.file);
      fd.append('manual_no_map', JSON.stringify({ [item.file ? item.file.name : item.file]: manualNo.trim() }));
      const res = await fetch('/api/match_excel', { method: 'POST', body: fd });
      const data = await res.json();
      if (data.error) { alert('匹配失败：' + data.error); return; }
      const row = data.matches[0];
      setMatches((prev) => prev.map((m) =>
        m.key === key ? { ...m, letter_no: manualNo.trim(), status: row.status, ledger: row.ledger } : m
      ));
    } catch (err) {
      alert('匹配失败：' + err.message);
    } finally {
      setMatching(false);
    }
  };

  const deleteOne = (key) => setMatches((prev) => prev.filter((m) => m.key !== key));

  const startBatchRender = async () => {
    if (!selTemplate) { alert('请先选择模板'); return; }
    const makable = matches.filter((m) => m.status === 'ok' && m.file);
    if (makable.length === 0) { alert('没有可制函的文件（全部未匹配或已删除）'); return; }
    setRendering(true);
    setBatchResults(null);
    try {
      const fd = new FormData();
      fd.append('template_id', selTemplate);
      const letterNoMap = {};
      makable.forEach((m) => {
        fd.append('file', m.file);
        letterNoMap[m.file.name] = m.letter_no;
      });
      fd.append('letter_no_map', JSON.stringify(letterNoMap));
      const res = await fetch('/api/batch_render', { method: 'POST', body: fd });
      const data = await res.json();
      if (data.error) { alert('制函失败：' + data.error); return; }
      setBatchResults(data);
      if (data.zip_url) window.open(data.zip_url, '_blank');
    } catch (err) {
      alert('制函失败：' + err.message);
    } finally {
      setRendering(false);
    }
  };

  const okCnt = matches.filter((m) => m.status === 'ok').length;
  const pendCnt = matches.filter((m) => m.status === 'pending').length;
  const failCnt = matches.filter((m) => m.status === 'fail').length;
  const makable = matches.filter((m) => m.status === 'ok').length;

  return (
    <div className="wizard-backdrop" onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="wizard">
        <div className="dialog-title">
          <h1>Excel 批量制函</h1>
          <span className="close" onClick={onClose}>×</span>
        </div>

        <div className="body">
          {/* ① 选模板 */}
          <div className="tpl-bar">
            <div className="field">
              <label>① 选择函证模板</label>
              <select value={selTemplate} onChange={(e) => { setSelTemplate(e.target.value); setBatchResults(null); }}>
                <option value="">请选择已配置模板...</option>
                {templates.map((t) => (
                  <option key={t.id} value={t.id}>{t.name}</option>
                ))}
              </select>
            </div>
          </div>

          {/* ② 上传匹配 */}
          <div className="sec-title">② 上传 Excel 并匹配函证台账</div>
          <p className="upload-hint">
            文件名 = 函证编号（可带描述），如 <span className="code">wlfz002001.xlsx</span> 或 <span className="code">wlfz002001_中信证券.xlsx</span>。系统取首个 <span className="code">_</span> 前段作为函证编号去台账匹配。
          </p>
          <div className="upload">
            <label className="big" htmlFor="bw-excel">＋ 点击或拖拽上传多个 .xlsx</label>
            <input id="bw-excel" type="file" accept=".xlsx,.xlsm" multiple style={{ display: 'none' }} onChange={onBatchFilesChange} />
            <div className="small">支持多选；所有文件都会列入下方台账列表，匹配不上也不会被丢弃</div>
          </div>

          {matches.length > 0 && (
            <>
              <div className="summary">
                <div className="item">共上传 <b>{matches.length}</b> 个</div>
                <div className="item ok">✓ 已匹配 <b>{okCnt}</b></div>
                <div className="item pend">✎ 待补编号 <b>{pendCnt}</b></div>
                <div className="item fail">✗ 无匹配 <b>{failCnt}</b></div>
              </div>
              {failCnt > 0 && (
                <div className="alert">⚠️ 有 {failCnt} 个未匹配（已置灰不参与制函），可手动补填函证编号救回，或删除该行。</div>
              )}

              <table className="bw-table">
                <thead>
                  <tr>
                    <th style={{ width: '22%' }}>文件</th>
                    <th style={{ width: '16%' }}>函证编号</th>
                    <th style={{ width: '16%' }}>客户名称</th>
                    <th style={{ width: '14%' }}>被询证单位</th>
                    <th style={{ width: '12%' }}>状态</th>
                    <th style={{ width: '20%' }}>操作</th>
                  </tr>
                </thead>
                <tbody>
                  {matches.map((m) => (
                    <tr key={m.key} className={m.status === 'pending' ? 'row-pend' : m.status === 'fail' ? 'row-fail' : ''}>
                      <td className="code">{m.file ? m.file.name : m.file}</td>
                      {m.status === 'pending' ? (
                        <>
                          <td colSpan="3">
                            <ManualInput onMatch={(v) => rematchOne(m.key, v)} disabled={matching} />
                          </td>
                          <td><span className="tag pend">✎ 待补编号</span></td>
                          <td><button className="btn danger sm" onClick={() => deleteOne(m.key)}>删除</button></td>
                        </>
                      ) : (
                        <>
                          <td className="code">{m.letter_no}</td>
                          <td>{m.ledger ? m.ledger['客户名称'] : <span className="muted">—</span>}</td>
                          <td>{m.ledger ? m.ledger['被询证单位'] : <span className="muted">—</span>}</td>
                          <td>
                            {m.status === 'ok' && <span className="tag ok">✓ 自动匹配</span>}
                            {m.status === 'fail' && <span className="tag fail">✗ 无匹配</span>}
                          </td>
                          <td><button className="btn danger sm" onClick={() => deleteOne(m.key)}>删除</button></td>
                        </>
                      )}
                    </tr>
                  ))}
                </tbody>
              </table>

              {batchResults && (
                <div className="bw-result">
                  ✅ 制函完成：成功 {batchResults.count} 份，失败 {batchResults.results.length - batchResults.count} 份
                </div>
              )}
            </>
          )}
        </div>

        {/* 底部操作 */}
        <div className="footer">
          <div className="left-info">可制函 <b className="primary">{makable}</b> 份（✓ {okCnt} + 补到的计入）；✗ 不计入</div>
          <div className="actions">
            <button className="btn ghost" onClick={onClose}>取消</button>
            <button className="btn-start" disabled={rendering || !selTemplate || makable === 0} onClick={startBatchRender}>
              {rendering ? '制函中...' : `🚀 开始制函（${makable}）`}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function ManualInput({ onMatch, disabled }) {
  const [val, setVal] = useState('');
  return (
    <div className="match-cell">
      <input value={val} placeholder="手动输入函证编号" disabled={disabled}
        onChange={(e) => setVal(e.target.value)}
        onKeyDown={(e) => { if (e.key === 'Enter' && val.trim()) onMatch(val); }} />
      <button className="btn ghost sm" disabled={disabled || !val.trim()} onClick={() => onMatch(val)}>匹配</button>
    </div>
  );
}
