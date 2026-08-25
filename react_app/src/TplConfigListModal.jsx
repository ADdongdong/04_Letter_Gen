import React, { useState, useEffect } from 'react';

/**
 * 模板配置列表弹窗（复刻 docs/proto/模板配置列表弹窗.html）。
 * 行操作：修改模板（进入 OnlyOffice 编辑页）、删除（调用 /api/template_delete）。
 * 顶部「新建模板」基于默认模板进入编辑页。
 */
export default function TplConfigListModal({ onClose, onEditTemplate }) {
  const [templates, setTemplates] = useState([]);
  const [deleting, setDeleting] = useState(null); // 正在删除的模板 id

  const load = async () => {
    try {
      const res = await fetch('/api/templates');
      const data = await res.json();
      setTemplates(data.templates || []);
    } catch (e) { /* ignore */ }
  };
  useEffect(() => { load(); }, []);

  const handleDelete = async (t) => {
    if (!window.confirm(`确认删除模板「${t.name}」？\n删除后不可恢复。`)) return;
    setDeleting(t.id);
    try {
      const res = await fetch('/api/template_delete', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ id: t.id }),
      });
      const data = await res.json();
      if (data.error) { alert('删除失败：' + data.error); return; }
      await load();
    } catch (err) {
      alert('删除失败：' + err.message);
    } finally {
      setDeleting(null);
    }
  };

  return (
    <div className="tpl-overlay" onClick={(e) => { if (e.target === e.currentTarget) onClose(); }}>
      <div className="tpl-modal">
        <div className="tpl-modal-head">
          <div className="t">🧩 模板配置</div>
          <div className="x" onClick={onClose}>✕</div>
        </div>

        <div className="tpl-modal-body">
          <div className="tpl-head-row">
            <div className="hint">共 {templates.length} 个模板 · 每个模板 = 一份 Word + 绑定的 Excel 名称</div>
            <button className="btn-new" onClick={() => onEditTemplate(null)}>+ 新建模板</button>
          </div>

          <table className="tpl-list">
            <thead>
              <tr>
                <th style={{ width: '26%' }}>模板名称</th>
                <th>绑定的 Excel 名称</th>
                <th style={{ width: '14%' }}>更新时间</th>
                <th style={{ width: '20%' }}>操作</th>
              </tr>
            </thead>
            <tbody>
              {templates.length === 0 && (
                <tr><td colSpan="4" style={{ textAlign: 'center', color: '#aaa' }}>暂无模板</td></tr>
              )}
              {templates.map((t) => (
                <tr key={t.id}>
                  <td>
                    <span className="tpl-name">
                      {t.name} <span className="v">{t.version || 'v1'}</span>
                    </span>
                  </td>
                  <td>
                    <span className="bind-chips">
                      <span className="chip">{t.excel_name || '未绑定 Excel'}</span>
                    </span>
                  </td>
                  <td>{t.updated_at || '—'}</td>
                  <td>
                    <div className="op">
                      <button className="edit" onClick={() => onEditTemplate(t)}>修改模板</button>
                      <button className="del" disabled={deleting === t.id} onClick={() => handleDelete(t)}>
                        {deleting === t.id ? '删除中...' : '删除'}
                      </button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>

          <div className="tip">
            💡 点击「<strong>修改模板</strong>」进入 OnlyOffice 编辑页，编辑该模板的占位符位置；<br />
            点击「<strong>新建模板</strong>」可上传自己的 Word 模板或基于默认模板创建新模板。
          </div>
        </div>

        <div className="tpl-modal-foot">
          <button className="btn btn-default" onClick={onClose}>关闭</button>
        </div>
      </div>
    </div>
  );
}
