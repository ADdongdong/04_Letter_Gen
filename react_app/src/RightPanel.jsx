import React, { useState, useRef } from 'react';

/**
 * 待补编号输入框（行内）：用户输入函证编号后点「匹配」回传
 */
function ManualNoInput({ onMatch, disabled }) {
  const [val, setVal] = useState('');
  return (
    <span className="match-cell">
      <input
        className="manual-no"
        placeholder="手动输入函证编号"
        value={val}
        disabled={disabled}
        onChange={(e) => setVal(e.target.value)}
        onKeyDown={(e) => { if (e.key === 'Enter' && val.trim()) onMatch(val); }}
      />
      <button className="btn ghost sm" disabled={disabled || !val.trim()} onClick={() => onMatch(val)}>
        匹配
      </button>
    </span>
  );
}

/**
 * 右侧操作面板 —— 模板化制函（统一单一主线）
 *
 * Tab1「配置模板」（一次性）：
 *   1. 选择模板来源（默认模板 / 已存模板）→ OnlyOffice 加载
 *   2. （可选）加载参考 Excel → 点击 Sheet → 在 OnlyOffice 光标处写入占位符
 *   3. 输入模板名称 → 保存为模板（后端扫描占位符提取绑定规则）
 *
 * Tab2「批量制函」（每次，单页）：
 *   1. ① 选择模板（同页，不展示绑定 Sheet 列表）
 *   2. ② 上传 1 个或多个 Excel → 自动匹配台账（全部列入列表，三态：✓/✎/✗）
 *   3. ✎ 待补编号可手填重匹配救回；任意行可删除
 *   4. 点「开始制函」→ 仅制函已匹配项 → 按函证编号命名 → 打包 zip 下载
 */
export default function RightPanel({
  ready, insertText, updateStatus,
  templates, onLoadTemplates, onOpenTemplate,
  defaultTab = 'config', showTabs = true,
  hideSourceSelect = false,
  onUploadTemplate,
  sourceTemplateId = 'default',
  mode = 'new',              // 'new' | 'edit'：新建模板 / 修改模板
  templateName = '',         // 修改模式下当前模板名称
  templateId = null,         // 修改模式下当前模板 id（保存时原地更新）
}) {
  const [tab, setTab] = useState(defaultTab);       // config | batch
  const [refSheets, setRefSheets] = useState([]);  // 参考 Excel 的 sheet 列表
  const [refExcelName, setRefExcelName] = useState('尚未加载参考 Excel');
  const [excelFileName, setExcelFileName] = useState(''); // 上传的参考 Excel 文件名
  const [placeholderSheets, setPlaceholderSheets] = useState([]); // 已插入占位符的 sheet（有序）
  const [tplName, setTplName] = useState(mode === 'edit' ? templateName : '');
  const [saving, setSaving] = useState(false);

  const isEdit = mode === 'edit';
  const finalTemplateName = (isEdit ? templateName : tplName).trim();
  const excelStep = hideSourceSelect ? '1' : '2';
  const saveStep = hideSourceSelect ? '2' : '3';

  // ---- 批量制函状态 ----
  const [selTemplate, setSelTemplate] = useState('');
  // matches: [{key, file, letter_no, status('ok'|'pending'|'fail'), ledger}]
  const [matches, setMatches] = useState([]);
  const [matching, setMatching] = useState(false);
  const [batchResults, setBatchResults] = useState(null);
  const [rendering, setRendering] = useState(false);

  // ---- 配置模板：加载参考 Excel（可选）----
  const onRefExcelChange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const fd = new FormData();
    fd.append('file', file);
    try {
      const res = await fetch('/api/upload_excel', { method: 'POST', body: fd });
      const data = await res.json();
      if (data.error) { alert(data.error); return; }
      setRefSheets(data.sheets);
      setRefExcelName(`已加载：${data.sheets.length} 个 Sheet（${file.name}）`);
      setExcelFileName(file.name);
      updateStatus(`✓ 参考 Excel 已加载，共 ${data.sheets.length} 个 Sheet`);
    } catch (err) {
      alert('参考 Excel 加载失败：' + err.message);
    }
  };

  // ---- 配置模板：点击 Sheet → 在 OnlyOffice 光标处写入占位符 ----
  const onSheetClick = (sheet) => {
    if (!ready) { alert('编辑器尚未就绪，请稍候'); return; }
    const text = `【Sheet「${sheet.name}」表格将在此处展示】`;
    const ok = insertText(text);
    if (!ok) {
      updateStatus('⚠ 插入失败：请在左侧模板中点一下再重试', true);
      return;
    }
    updateStatus(`✓ 已写入「${sheet.name}」占位符`);
    setPlaceholderSheets((prev) => [...prev, sheet.name]);
  };

  // ---- 配置模板：保存为模板 ----
  const saveTemplate = async () => {
    if (placeholderSheets.length === 0) {
      alert('请先在左侧模板中插入至少一个 Sheet 占位符');
      return;
    }
    const name = finalTemplateName || ('模板_' + new Date().toISOString().slice(0, 10));
    if (isEdit) {
      if (!window.confirm(`保存到当前模板「${name}」？\n绑定的 Excel：${excelFileName || '未绑定'}\n占位 Sheet：${placeholderSheets.join('、')}`)) return;
    } else {
      if (!window.confirm(`保存为模板「${name}」？\n绑定的 Excel：${excelFileName || '未绑定'}\n占位 Sheet：${placeholderSheets.join('、')}`)) return;
    }

    setSaving(true);
    try {
      // bindings：按插入顺序，pos_index 0..n-1（占位符在模板中的位置）
      const bindings = placeholderSheets.map((sn, i) => ({
        pos_index: i,
        sheet_name: sn,
        pos_label: `位置${i + 1}`,
      }));
      const res = await fetch('/api/template_save', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          id: templateId || undefined,       // 修改模式：原地更新该模板；新建模式：不传 id
          name: name,
          bindings: bindings,
          source: sourceTemplateId || 'default', // 基于当前 OnlyOffice 里的 docx 生成
          excel_name: excelFileName,
        }),
      });
      const data = await res.json();
      if (data.error) {
        const msg = res.status === 409 ? '模板名称重复：' + data.error : '保存失败：' + data.error;
        alert(msg);
        return;
      }
      updateStatus(`✓ 模板「${data.name}」已保存（Excel：${data.excel_name || '未绑定'}）`);
      alert(`模板「${data.name}」保存成功！\n绑定的 Excel：${data.excel_name || '未绑定'}\n占位规则：\n` +
        data.bindings.map((b) => `  · ${b.sheet_name} → 位置${b.para_index}`).join('\n'));
      onLoadTemplates(); // 刷新模板列表
      setPlaceholderSheets([]);
      setExcelFileName('');
      setRefExcelName('尚未加载参考 Excel');
      setTplName('');
    } catch (err) {
      alert('保存失败：' + err.message);
    } finally {
      setSaving(false);
    }
  };

  // ---- 配置模板：打开模板来源 ----
  const onSourceChange = (e) => {
    const v = e.target.value;
    onOpenTemplate(v); // '' = 默认模板
    setPlaceholderSheets([]);
  };

  // ---- 批量制函：选择模板时清空上次结果 ----
  const onSelTemplateChange = (e) => {
    setSelTemplate(e.target.value);
    setBatchResults(null);
  };

  // ---- 批量制函：多选 Excel → 调匹配接口，全部列入列表（不丢弃）----
  const onBatchFilesChange = async (e) => {
    const files = Array.from(e.target.files || []);
    if (files.length === 0) return;
    setMatching(true);
    setBatchResults(null);
    updateStatus(`匹配台账中...（${files.length} 个 Excel）`);
    try {
      const fd = new FormData();
      files.forEach((f) => fd.append('file', f));
      const res = await fetch('/api/match_excel', { method: 'POST', body: fd });
      const data = await res.json();
      if (data.error) { alert('匹配失败：' + data.error); return; }
      // 关联本地 File 引用，便于后续制函
      const keyed = data.matches.map((m) => {
        const f = files.find((x) => x.name === m.file);
        return { ...m, key: m.file + '|' + Math.random().toString(36).slice(2, 8), file: f };
      });
      setMatches((prev) => [...prev, ...keyed]);
      const okCnt = keyed.filter((k) => k.status === 'ok').length;
      const pendCnt = keyed.filter((k) => k.status === 'pending').length;
      updateStatus(`✓ 匹配完成：已匹配 ${okCnt}，待补编号 ${pendCnt}（共 ${keyed.length}）`);
    } catch (err) {
      alert('匹配失败：' + err.message);
    } finally {
      setMatching(false);
      e.target.value = '';
    }
  };

  // ---- 批量制函：手填函证编号 → 对该文件重新匹配 ----
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
      updateStatus(row.status === 'ok' ? `✓ ${manualNo} 已匹配台账` : `⚠ ${manualNo} 台账仍无，已置灰`);
    } catch (err) {
      alert('匹配失败：' + err.message);
    } finally {
      setMatching(false);
    }
  };

  // ---- 批量制函：单条删除 ----
  const deleteOne = (key) => {
    setMatches((prev) => prev.filter((m) => m.key !== key));
  };

  // ---- 批量制函：开始制函（仅制函 ok 项）----
  const startBatchRender = async () => {
    if (!selTemplate) { alert('请先选择模板'); return; }
    const makable = matches.filter((m) => m.status === 'ok' && m.file);
    if (makable.length === 0) { alert('没有可制函的文件（全部未匹配或已删除）'); return; }
    setRendering(true);
    setBatchResults(null);
    updateStatus(`制函中...（${makable.length} 个可制函）`);
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
      const okCount = data.count;
      const failCount = data.results.filter((r) => r.status === 'error').length;
      updateStatus(`✓ 制函完成：成功 ${okCount}，失败 ${failCount}`);
      if (data.zip_url) {
        window.open(data.zip_url, '_blank');
      }
    } catch (err) {
      alert('制函失败：' + err.message);
    } finally {
      setRendering(false);
    }
  };

  return (
    <div className={`right ${!showTabs ? 'right-modal' : ''}`}>
      {/* Tab 切换（仅在 showTabs 时显示）*/}
      {showTabs && (
        <div className="tabs">
          <div className={`tab ${tab === 'config' ? 'active' : ''}`} onClick={() => setTab('config')}>
            🧩 配置模板
          </div>
          <div className={`tab ${tab === 'batch' ? 'active' : ''}`} onClick={() => setTab('batch')}>
            📦 批量制函
          </div>
        </div>
      )}

      {/* ============ Tab1：配置模板 ============ */}
      {tab === 'config' && (
        <>
          {!hideSourceSelect && (
            <>
              <h3>1. 模板来源</h3>
              <select className="tpl-select" value="" onChange={onSourceChange}>
                <option value="">默认模板（demo_template）</option>
                {templates.map((t) => (
                  <option key={t.id} value={t.id}>{t.name}</option>
                ))}
              </select>
              <div className="hint">
                选择后会重新加载到左侧 OnlyOffice。已有模板可直接继续编辑再保存。
              </div>
            </>
          )}

          {hideSourceSelect && onUploadTemplate && (
            <div className="upload-tpl-block">
              <h3>{isEdit ? '替换 Word 模板' : '上传 Word 模板'}</h3>
              <div className="file-box">
                <label htmlFor="tpl-word">📄 点击选择 Word 模板 (.docx)</label>
                <input id="tpl-word" type="file" accept=".docx" onChange={onUploadTemplate} />
              </div>
              <div className="hint">{isEdit ? '替换当前模板 Word 文件，左侧 OnlyOffice 会重新加载。' : '上传后会在左侧 OnlyOffice 打开该模板，可继续编辑占位符并保存。'}</div>
            </div>
          )}

          <h3>{excelStep}.（可选）加载参考 Excel</h3>
          <div className="file-box">
            <label htmlFor="ref-excel">📊 点击选择参考 Excel (.xlsx)</label>
            <input id="ref-excel" type="file" accept=".xlsx,.xlsm" onChange={onRefExcelChange} />
          </div>
          <div className="info">{refExcelName}</div>

          {refSheets.length > 0 && (
            <>
              <div className="hint" style={{ marginTop: '10px' }}>
                操作：先在<strong>左侧模板中点一下</strong>定位光标，再点击下方 Sheet → 自动写入占位符。
                <br />（也可手写 <code>【Sheet「xxx」表格将在此处展示】</code>）
              </div>
              <div className="sheet-list">
                {refSheets.map((s) => {
                  const inserted = placeholderSheets.includes(s.name);
                  return (
                    <div
                      key={s.name}
                      className={`sheet-item clickable ${inserted ? 'inserted' : ''}`}
                      onClick={() => onSheetClick(s)}
                    >
                      <span className="badge">{s.name}</span>
                      <span>{s.row_count} 行数据</span>
                      <span style={{ marginLeft: 'auto', fontSize: '11px' }}>
                        {inserted ? '✅ 已写入' : '👆 点击写入'}
                      </span>
                    </div>
                  );
                })}
              </div>
              <div className="info" style={{ marginTop: '8px' }}>
                已写入占位符（{placeholderSheets.length}）：{placeholderSheets.join('、') || '无'}
              </div>
            </>
          )}

          {isEdit ? (
            <>
              <h3>{saveStep}. 保存当前模板</h3>
              <div className="info" style={{ marginBottom: '8px' }}>
                当前模板名称：<strong>{templateName}</strong>
              </div>
              <div className="actions">
                <button className="btn primary" disabled={saving || placeholderSheets.length === 0} onClick={saveTemplate}>
                  {saving ? '保存中...' : '💾 保存'}
                </button>
              </div>
            </>
          ) : (
            <>
              <h3>{saveStep}. 保存为模板</h3>
              <input
                className="tpl-name-input"
                placeholder="模板名称（如：银行询证函-销售采购）"
                value={tplName}
                onChange={(e) => setTplName(e.target.value)}
              />
              <div className="actions">
                <button className="btn primary" disabled={saving || placeholderSheets.length === 0} onClick={saveTemplate}>
                  {saving ? '保存中...' : '💾 保存为模板'}
                </button>
              </div>
            </>
          )}
        </>
      )}

      {/* ============ Tab2：批量制函（单页：选模板 + 上传匹配同页）============ */}
      {tab === 'batch' && (
        <>
          <h2>批量制函</h2>

          {/* ① 选择模板（同页，不展示绑定 Sheet 列表）*/}
          <div className="batch-tpl">
            <label className="sec-label">① 选择函证模板</label>
            <select
              className="tpl-select"
              value={selTemplate}
              onChange={onSelTemplateChange}
            >
              <option value="">请选择已配置模板...</option>
              {templates.map((t) => (
                <option key={t.id} value={t.id}>
                  {t.name}（{t.bindings ? t.bindings.length : 0} 个绑定）
                </option>
              ))}
            </select>
          </div>

          {/* ② 上传 Excel + 匹配台账 */}
          <div className="sec-title">② 上传 Excel 并匹配函证台账</div>
          <div className="hint" style={{ marginBottom: '8px' }}>
            文件名 = 函证编号（可带描述），如 <code>wlfz002001.xlsx</code> 或 <code>wlfz002001_中信证券.xlsx</code>。
            系统取首个 <code>_</code> 前段作为函证编号去台账匹配。
          </div>
          <div className="file-box">
            <label htmlFor="batch-excel">📊 点击或拖拽选择 1 个或多个 Excel (.xlsx)</label>
            <input id="batch-excel" type="file" accept=".xlsx,.xlsm" multiple onChange={onBatchFilesChange} />
          </div>

          {matches.length > 0 && (() => {
            const okCnt = matches.filter((m) => m.status === 'ok').length;
            const pendCnt = matches.filter((m) => m.status === 'pending').length;
            const failCnt = matches.filter((m) => m.status === 'fail').length;
            const makable = matches.filter((m) => m.status === 'ok').length;
            return (
              <>
                {/* 汇总条 */}
                <div className="match-summary">
                  <span>共上传 <b>{matches.length}</b></span>
                  <span className="ok">✓ 已匹配 <b>{okCnt}</b></span>
                  <span className="pend">✎ 待补编号 <b>{pendCnt}</b></span>
                  <span className="fail">✗ 无匹配 <b>{failCnt}</b></span>
                </div>
                {failCnt > 0 && (
                  <div className="alert">
                    ⚠️ 有 {failCnt} 个未匹配（已置灰不参与制函），可手动补填函证编号救回，或删除该行。
                  </div>
                )}

                {/* 台账列表 */}
                <table className="match-table">
                  <thead>
                    <tr>
                      <th>文件</th>
                      <th>函证编号</th>
                      <th>客户名称</th>
                      <th>被询证单位</th>
                      <th>状态</th>
                      <th>操作</th>
                    </tr>
                  </thead>
                  <tbody>
                    {matches.map((m) => (
                      <tr key={m.key} className={m.status === 'pending' ? 'row-pend' : m.status === 'fail' ? 'row-fail' : ''}>
                        <td className="code">{m.file ? m.file.name : m.file}</td>
                        <td className="code">{m.letter_no}</td>
                        <td>{m.ledger ? m.ledger['客户名称'] : <span className="muted">—</span>}</td>
                        <td>{m.ledger ? m.ledger['被询证单位'] : <span className="muted">—</span>}</td>
                        <td>
                          {m.status === 'ok' && <span className="tag ok">✓ 自动匹配</span>}
                          {m.status === 'pending' && <span className="tag pend">✎ 待补编号</span>}
                          {m.status === 'fail' && <span className="tag fail">✗ 无匹配</span>}
                        </td>
                        <td>
                          {m.status === 'pending' && (
                            <ManualNoInput onMatch={(v) => rematchOne(m.key, v)} disabled={matching} />
                          )}
                          <button className="btn danger sm" onClick={() => deleteOne(m.key)}>删除</button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                {/* 开始制函 */}
                <div className="actions" style={{ marginTop: '12px' }}>
                  <button
                    className="btn primary"
                    disabled={rendering || !selTemplate || makable === 0}
                    onClick={startBatchRender}
                  >
                    {rendering ? '制函中...' : `🚀 开始制函（${makable}）`}
                  </button>
                </div>
              </>
            );
          })()}

          {batchResults && (
            <div className="result-box">
              <div className="info" style={{ background: '#e8f5e9' }}>
                ✅ 制函完成：成功 {batchResults.count} 份，失败 {batchResults.results.length - batchResults.count} 份
              </div>
              {batchResults.results.map((r, i) => (
                <div key={i} className={`result-item ${r.status === 'ok' ? '' : 'err'}`}>
                  <b>{r.file}</b>（{r.letter_no}）
                  {r.status === 'ok'
                    ? ` → 注入 ${r.injected.map((x) => `「${x.sheet}」${x.cols}列×${x.rows}行`).join('、') || '无表格'}`
                    : ` → 失败：${r.msg}`}
                  {r.not_found && r.not_found.length > 0 && (
                    <div style={{ color: '#e65100' }}>缺 Sheet（已跳过）：{r.not_found.join('、')}</div>
                  )}
                </div>
              ))}
              {batchResults.zip_url && (
                <a className="download" href={batchResults.zip_url}>
                  ⬇️ 下载全部（zip）
                </a>
              )}
            </div>
          )}
        </>
      )}
    </div>
  );
}
