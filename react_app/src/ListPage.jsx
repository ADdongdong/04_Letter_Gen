import React, { useState, useEffect, useCallback } from 'react';

/**
 * 模板列表页（#/templates，默认页）：
 * 展示所有模板配置；点【+ 新增配置】或某条的【修改】进入配置编辑页（#/config）。
 * 监听 hashchange：切回本页（#/templates）时自动刷新数据（保存/删除后列表最新）。
 */
export default function ListPage() {
  const [templates, setTemplates] = useState([]);

  const loadTemplates = useCallback(() => {
    fetch('/api/templates')
      .then(r => r.json())
      .then(d => setTemplates(d.templates || []))
      .catch(() => {});
  }, []);

  useEffect(() => {
    loadTemplates();
    // 常驻挂载：切回本页（hash 变为 #/templates）时刷新，保证保存/删除后数据最新
    const onHash = () => {
      if ((window.location.hash || '').startsWith('#/templates')) loadTemplates();
    };
    window.addEventListener('hashchange', onHash);
    return () => window.removeEventListener('hashchange', onHash);
  }, [loadTemplates]);

  const del = (t) => {
    if (!window.confirm('确认删除模板配置「' + t.name + '」？该操作不可恢复。')) return;
    fetch('/api/templates/' + t.id, { method: 'DELETE' })
      .then(r => r.json())
      .then(d => { if (d.error) { alert(d.error); return; } loadTemplates(); })
      .catch(e => alert('删除失败：' + e.message));
  };

  const card = { background: '#fff', border: '1px solid #e0e0e0', borderRadius: '8px', padding: '16px' };

  return (
    <div style={{ maxWidth: '760px', margin: '0 auto', padding: '16px' }}>
      <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', margin: '8px 0 16px' }}>
        <h2 style={{ margin: 0 }}>模板配置管理</h2>
        <button
          className="btn primary"
          onClick={() => { window.location.hash = '#/config'; }}
          style={{ padding: '8px 20px' }}
        >
          ＋ 新增模板
        </button>
      </div>

      {templates.length === 0 ? (
        <div style={card}>
          <div style={{ textAlign: 'center', padding: '32px 0', color: '#888' }}>
            <div style={{ fontSize: '32px', marginBottom: '8px' }}>📄</div>
            还没有模板配置
            <div style={{ marginTop: '8px', fontSize: '13px' }}>
              点击右上角「＋ 新增配置」上传 Word 模板并标注 Sheet 插入位置
            </div>
          </div>
        </div>
      ) : (
        templates.map(t => (
          <div key={t.id} style={{ ...card, marginBottom: '12px' }}>
            <div style={{ display: 'flex', alignItems: 'center', gap: '12px', flexWrap: 'wrap' }}>
              <strong style={{ fontSize: '15px' }}>{t.name}</strong>
              <span style={{ color: '#888', fontSize: '12px' }}>更新时间：{t.updated_at || '-'}</span>
              <span style={{ marginLeft: 'auto' }}>
                <button
                  className="btn"
                  onClick={() => { window.location.hash = '#/config?id=' + t.id; }}
                  style={{ padding: '6px 14px', marginRight: '8px' }}
                >
                  ✏️ 修改
                </button>
                <button className="btn" onClick={() => del(t)} style={{ padding: '6px 14px', color: '#c62828' }}>
                  🗑 删除
                </button>
              </span>
            </div>
            <div style={{ marginTop: '10px', fontSize: '13px', color: '#555' }}>
              包含 Sheet（{(t.sheets || []).length}）：{(t.sheets || []).join('、') || '（无）'}
            </div>
          </div>
        ))
      )}

      <div style={{ marginTop: '16px', padding: '10px', background: '#f5f5f5', borderRadius: '6px', fontSize: '13px', color: '#666' }}>
        模板配置 = Word 模板（含 Sheet 插入位置标注）+ Excel 的 Sheet 结构定义。保存后即可在「Excel 制函」页选择使用。
      </div>
    </div>
  );
}
