import React, { useRef, useState, useCallback, useEffect } from 'react';
import OnlyOfficeEditor from './OnlyOfficeEditor.jsx';
import RightPanel from './RightPanel.jsx';

// OnlyOffice 在本地 Docker 中运行（端口8080），文档通过 Flask 后端 API 返回（与 hanzheng 项目模式一致：docUrl 指向后端接口）
// OO 容器内通过 host.docker.internal 访问宿主机 Flask（已验证连通）
const FLASK_BASE = 'http://host.docker.internal:5002';
const TEMPLATE_URL = FLASK_BASE + '/api/oo/getTemplate';

// 生成全局唯一的 document.key（时间戳+随机段，永不复用）。
// 关键：OO 服务器按 document.key 缓存文档——同 key 直接返回缓存不重新下载。
// 唯一 key 强制 OO 重新下载文档（旧方案 key 复用导致"上传新模板加载的还是旧文档"）。
const genDocKey = () =>
  'zhihanyou-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);

export default function App() {
  const editorRef = useRef(null);
  const [ready, setReady] = useState(false);
  const [docKey, setDocKey] = useState(genDocKey);  // 当前编辑器的唯一 document.key

  // 修改模式：editingId 非空 = 编辑已有模板（OO 加载该模板已保存的 word）；null = 新增模式。
  // 配置页常驻挂载：editingId 跟随 hashchange 同步（列表页点【修改】→ hash=#/config?id=x → 本页自动重载目标模板）
  const [editingId, setEditingId] = useState(() => {
    const h = window.location.hash || '';
    const m = h.match(/[?&]id=(tpl_[0-9a-f]+)/);
    return h.startsWith('#/config') && m ? m[1] : null;
  });
  const [editingName, setEditingName] = useState('');  // 编辑中模板名（标题展示，从配置清单查询）
  // 新增模式：wordLoaded=false 时 OO 显示空白占位（不加载残留的 current.docx）；上传 Word 后才加载
  const [wordLoaded, setWordLoaded] = useState(!!editingId);
  // docUrl：修改模式指向模板配置的 word；新增模式未上传 Word 前为 null（OO 空态占位）
  const docUrl = editingId
    ? FLASK_BASE + '/api/templates/' + editingId + '/word'
    : (wordLoaded ? TEMPLATE_URL : null);

  // 用 ref 存储状态信息，避免触发 App 重渲染导致 OnlyOffice 刷新
  const statusRef = useRef('加载中...');
  // 与 editingId 同步的 ref：hashchange 同步逻辑判断"是否真的变化"，避免 OO 无谓重建
  const editingRef = useRef(editingId);

  // ---- 插入文本的核心方法（由 RightPanel 调用）----
  // 优先走 Click2Insert 插件（OO 8.x 官方方式，Hanzheng 同款），fallback 到旧的内部 insertText
  const insertText = useCallback((text) => {
    if (!editorRef.current) return false;
    if (typeof editorRef.current.insertTextViaPlugin === 'function') {
      const ok = editorRef.current.insertTextViaPlugin(text);
      if (ok) return true;
    }
    return editorRef.current.insertText(text);
  }, []);

  // ---- 触发 OO 强制保存（把含占位段的文档回写后端）----
  const forceSave = useCallback(async () => {
    if (!editorRef.current || typeof editorRef.current.forceSave !== 'function') return false;
    return editorRef.current.forceSave();
  }, []);

  // ---- 模板上传成功：换新 docKey 重载（竞态保护：预注册 active key）----
  // 旧编辑器销毁时 OO 会对旧 key 延迟触发 forcesave 回写旧文档，
  // 若不提前切换 active key，旧文档会覆盖刚上传的新模板。
  const handleTemplateUploaded = useCallback(() => {
    setReady(false);
    setWordLoaded(true);  // 新增模式：上传 Word 后开始加载上传的内容
    const nextKey = genDocKey();
    fetch('/api/oo/active_key', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ key: nextKey }),
    }).catch(() => {});
    setDocKey(nextKey);
  }, []);

  // ---- 回调：更新状态栏文字（不触发重渲染）----
  const updateStatus = useCallback((msg, highlight = false) => {
    statusRef.current = msg;
    const el = document.getElementById('oo-status');
    if (el) {
      el.textContent = msg;
      el.className = 'status' + (highlight ? ' status-alert' : '');
    }
  }, []);

  // ---- editingId 跟随 hashchange 同步（配置页常驻挂载，靠此机制响应列表页的进入动作）----
  // 列表页点【新增】→ hash=#/config（editingId=null，新增模式）
  // 列表页点【修改】→ hash=#/config?id=x（editingId=x，OO 以新 docKey 重载该模板 word）
  // 与当前 editingId 相同则不做任何操作（避免 OO 无谓重建）
  useEffect(() => {
    const sync = () => {
      const h = window.location.hash || '';
      if (!h.startsWith('#/config')) return;
      const m = h.match(/[?&]id=(tpl_[0-9a-f]+)/);
      const newId = m ? m[1] : null;
      if (editingRef.current === newId) return;
      editingRef.current = newId;
      setEditingId(newId);
      setWordLoaded(!!newId);  // 切换编辑目标：编辑模式直接加载已存 word，新增模式回到空白占位
      setReady(false);
      const nextKey = genDocKey();
      fetch('/api/oo/active_key', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ key: nextKey }),
      }).catch(() => {});
      setDocKey(nextKey);
      // 查询编辑中模板的名称（标题展示）
      if (newId) {
        fetch('/api/templates')
          .then(r => r.json())
          .then(d => {
            const t = (d.templates || []).find(x => x.id === newId);
            setEditingName(t ? t.name : '');
          })
          .catch(() => {});
      } else {
        setEditingName('');
      }
    };
    sync();
    window.addEventListener('hashchange', sync);
    return () => window.removeEventListener('hashchange', sync);
  }, []);

  // 保存成功（RightPanel onTemplateSaved）→ 跳回模板列表页（列表页监听 hashchange 自动刷新）
  const onTemplateSaved = useCallback(() => {
    window.location.hash = '#/templates';
  }, []);

  return (
    <div style={{ height: '100%', display: 'flex', flexDirection: 'column' }}>
      {/* 顶部：返回列表 + 模式标题 + 流程提示 */}
      <div style={{ display: 'flex', alignItems: 'center', gap: '12px', padding: '6px 16px', background: '#fff', borderBottom: '1px solid #e0e0e0' }}>
        <a
          href="#/templates"
          style={{ color: 'var(--primary)', cursor: 'pointer', fontWeight: 600, textDecoration: 'none', whiteSpace: 'nowrap' }}
        >
          ← 返回模板列表
        </a>
        <strong style={{ fontSize: '14px', whiteSpace: 'nowrap' }}>
          {editingId ? '✏️ 编辑模板' + (editingName ? '「' + editingName + '」' : '') : '📄 新增模板配置'}
        </strong>
        <span style={{ color: '#888', fontSize: '12px' }}>
          上传 Word 与 Excel → 在左侧光标处插入 Sheet 占位段 → 保存模板配置
        </span>
      </div>
      <div className="layout">
        <div className="left">
          <div className="toolbar">
            {/* 状态栏内容由 updateStatus 直接操作 DOM（脱离 React 管理），避免 textContent 与 React diff 冲突 */}
            <span id="oo-status" className="status" />
          </div>
          <div className="editor-area">
            {/* OnlyOfficeEditor 被 React.memo 包裹，props 不变时绝不重渲染 */}
            <OnlyOfficeEditor
              docKey={docKey}
              ref={editorRef}
              docUrl={docUrl}
              onReady={() => {
                setReady(true);
                updateStatus('模板已加载 ✓ 点击右侧 Sheet 即可在光标处插入表格标注');
              }}
            />
          </div>
        </div>

        {/* 右侧面板：所有状态变化都封装在此组件内，不会影响左侧 OnlyOffice */}
        <RightPanel
          ready={ready}
          insertText={insertText}
          forceSave={forceSave}
          updateStatus={updateStatus}
          onTemplateUploaded={handleTemplateUploaded}
          editingId={editingId}
          editingName={editingName}
          onTemplateSaved={onTemplateSaved}
        />
      </div>
    </div>
  );
}
