import React, { useState, useEffect, useCallback, useRef } from 'react';
import SelectField from './SelectField.jsx';

/**
 * 制函页（#/generate）：下拉选择模板配置 → 上传 Excel → 一键制函出 ZIP。
 * 无 OnlyOffice：模板（含占位段标注的 word + sheet 结构 excel）在配置页维护，
 * bindings 由后端从模板 word 自动提取（模板即定义）。
 * 上传 Excel 后展示与函证登记数据的匹配结果（未匹配红色，制函时跳过）。
 */
export default function GeneratePage() {
  const [templates, setTemplates] = useState([]);
  const [tplId, setTplId] = useState('');
  const [excelPath, setExcelPath] = useState(null);
  const [excelInfo, setExcelInfo] = useState('尚未上传 Excel');
  const [matchResults, setMatchResults] = useState([]);  // 函证编号匹配结果：[{audit, confirm, letter_no, matched}]
  const [status, setStatus] = useState('请选择模板并上传 Excel');
  const [generating, setGenerating] = useState(false);
  const generatingRef = useRef(false);

  // 加载模板配置清单；常驻挂载：切到本页（#/generate）时自动刷新（配置页新增/修改模板后数据最新）
  const loadTemplates = useCallback(() => {
    fetch('/api/templates')
      .then(r => r.json())
      .then(d => setTemplates(d.templates || []))
      .catch(() => {});
  }, []);

  useEffect(() => {
    loadTemplates();
    const onHash = () => {
      if ((window.location.hash || '').startsWith('#/generate')) loadTemplates();
    };
    window.addEventListener('hashchange', onHash);
    return () => window.removeEventListener('hashchange', onHash);
  }, [loadTemplates]);

  const onPick = (id) => {
    setTplId(id);
    const t = templates.find(x => x.id === id);
    if (t) {
      const names = (t.sheets || []).join('、') || '（无）';
      setStatus(`已选模板「${t.name}」：包含 Sheet ${names}`);
    }
  };

  // 上传 Excel：复用 /api/upload_excel 做即时校验（批量格式 + 空分组行拦截 + 函证编号匹配）
  const onExcelChange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    if (!tplId) { alert('请先选择制函模板'); e.target.value = ''; return; }
    setExcelInfo('上传校验中...');
    const fd = new FormData();
    fd.append('file', file);
    const res = await fetch('/api/upload_excel', { method: 'POST', body: fd });
    const data = await res.json();
    if (data.error) {
      alert(data.error);
      setExcelInfo('上传失败：' + data.error);
      setExcelPath(null);
      setMatchResults([]);
      e.target.value = '';
      return;
    }
    if (!data.batch_mode) {
      alert('请上传批量格式 Excel：各 Sheet 表头前两列需为「被审计单位」「被询证单位」');
      setExcelInfo('格式不支持（非批量格式）');
      setExcelPath(null);
      setMatchResults([]);
      e.target.value = '';
      return;
    }
    setExcelPath(data.path);
    setMatchResults(data.match_results || []);
    const unmatchedCount = (data.match_results || []).filter(m => !m.matched).length;
    setExcelInfo(`已加载：${file.name} · ${(data.groups || []).length} 封函证`);
    if (unmatchedCount > 0) {
      setStatus(`⚠ Excel 已加载：${(data.groups || []).length} 封函证，其中 ${unmatchedCount} 封在系统中未登记（见红色行），制函时将跳过`);
    } else {
      setStatus(`✓ Excel 已就绪：${(data.groups || []).length} 封函证全部匹配，可点击制函`);
    }
  };

  const generate = async () => {
    if (!tplId) { alert('请先选择制函模板'); return; }
    if (!excelPath) { alert('请先上传 Excel'); return; }
    if (generatingRef.current) return;
    generatingRef.current = true;
    setGenerating(true);
    try {
      setStatus('制函中，每封函证独立渲染，请耐心等待...');
      const res = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ template_id: tplId, excel_path: excelPath }),
      });
      const data = await res.json();
      if (data.error) { alert(data.error); setStatus('制函失败'); return; }
      // sheet_missing 不再弹窗阻断：无数据/缺失的 Sheet 位置不生成表格，仅在成功提示尾部附带说明
      const missingNote = (data.sheet_missing && data.sheet_missing.length)
        ? `（未生成表格的 Sheet：${data.sheet_missing.join('、')}）`
        : '';
      if (data.unmatched && data.unmatched.length) {
        setStatus(`✓ 制函成功！共 ${data.count} 封函证（已跳过 ${data.unmatched.length} 封未匹配函证），ZIP 已开始下载${missingNote}`);
      } else {
        setStatus(`✓ 制函成功！共 ${data.count} 封函证，ZIP 已开始下载${missingNote}`);
      }
      // 隐藏 <a download> 触发下载（不开新标签，避免页面闪烁）
      const a = document.createElement('a');
      a.href = data.download_url;
      a.download = '批量函证.zip';
      document.body.appendChild(a);
      a.click();
      a.remove();
    } catch (e) {
      alert('制函失败：' + e.message);
      setStatus('制函失败');
    } finally {
      generatingRef.current = false;
      setGenerating(false);
    }
  };

  const allUnmatched = matchResults.length > 0 && matchResults.every(m => !m.matched);

  return (
    <div className="page">
      <h2 style={{ margin: '8px 0 16px' }}>Excel 批量制函</h2>

      <div className="card">
        <h3>1. 选择制函模板</h3>
        <SelectField
          value={tplId}
          onChange={onPick}
          placeholder="— 请选择模板 —"
          options={[
            { value: '', label: '— 请选择模板 —' },
            ...templates.map(t => ({ value: t.id, label: t.name })),
          ]}
        />
        {tplId && (
          <div className="hint" style={{ marginTop: '8px' }}>
            模板包含 Sheet：{(templates.find(x => x.id === tplId)?.sheets || []).join('、') || '（无）'}
          </div>
        )}
      </div>

      <div className="card" style={{ marginTop: '12px' }}>
        <h3>2. 上传 Excel</h3>
        <div className="file-box">
          <label htmlFor="gen-excel">📊 点击选择 Excel (.xlsx)</label>
          <input id="gen-excel" type="file" accept=".xlsx,.xlsm" onChange={onExcelChange} disabled={!tplId} />
        </div>
        <div className="info">
          {excelInfo}
          <div style={{ marginTop: '4px' }}>
            <a href="/api/templates/sample/excel" download>📥 下载数据导入模板（含示例数据）</a>
            ，按模板填入实际数据后上传
          </div>
        </div>
        {/* 函证编号匹配结果：未匹配行红色，制函时跳过 */}
        {matchResults.length > 0 && (
          <div style={{ marginTop: '10px' }}>
            <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: '13px' }}>
              <thead>
                <tr style={{ background: 'var(--primary-bg)' }}>
                  <th style={{ border: '1px solid #ddd', padding: '6px 8px' }}>被审计单位</th>
                  <th style={{ border: '1px solid #ddd', padding: '6px 8px' }}>被询证单位</th>
                  <th style={{ border: '1px solid #ddd', padding: '6px 8px' }}>函证编号</th>
                  <th style={{ border: '1px solid #ddd', padding: '6px 8px' }}>状态</th>
                </tr>
              </thead>
              <tbody>
                {matchResults.map((m, i) => (
                  <tr key={i} style={{ background: m.matched ? '#fff' : '#fdecea' }}>
                    <td style={{ border: '1px solid #ddd', padding: '6px 8px' }}>{m.audit}</td>
                    <td style={{ border: '1px solid #ddd', padding: '6px 8px' }}>{m.confirm}</td>
                    <td className={m.matched ? '' : 'warn-red'} style={{ border: '1px solid #ddd', padding: '6px 8px', fontWeight: m.matched ? 400 : 700 }}>
                      {m.letter_no || '未匹配'}
                    </td>
                    <td className={m.matched ? '' : 'warn-red'} style={{ border: '1px solid #ddd', padding: '6px 8px' }}>
                      {m.matched ? '✓ 已匹配' : '⚠ 系统没有对应的函证，制函将跳过'}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            {matchResults.some(m => !m.matched) && (
              <div className="warn-red" style={{ marginTop: '8px', fontSize: '13px' }}>
                共 {matchResults.filter(m => !m.matched).length} 封未匹配（红色行），制函时将跳过这些函证的数据。
              </div>
            )}
          </div>
        )}
        <div className="hint">
          要求：各 Sheet 表头前两列为「被审计单位」「被询证单位」，这两列的组合唯一确定一封函证；
          制函时前两列不展示（仅用于分组与文件命名）。
        </div>
      </div>

      <div className="card" style={{ marginTop: '12px' }}>
        <h3>3. 制函</h3>
        <button
          className="btn primary"
          onClick={generate}
          disabled={!tplId || !excelPath || generating || allUnmatched}
          style={{ opacity: !tplId || !excelPath || allUnmatched ? 0.5 : 1, padding: '10px 24px' }}
        >
          {generating ? '制函中...' : `制函${matchResults.length > 0 ? `（${matchResults.filter(m => m.matched).length} 封）` : ''}`}
        </button>
        {allUnmatched && (
          <div className="warn-red" style={{ marginTop: '8px', fontSize: '13px' }}>
            Excel 中所有函证在系统中均未登记，无法制函。请先在函证系统完成登记后重新上传。
          </div>
        )}
      </div>

      <div className="status-line">{status}</div>
    </div>
  );
}
