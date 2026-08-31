import React, { useRef, useState, useCallback } from 'react';
import OnlyOfficeEditor from './OnlyOfficeEditor.jsx';
import RightPanel from './RightPanel.jsx';

// OnlyOffice 在本地 Docker 中运行（端口8080），文档通过 Flask 后端 API 返回（与 hanzheng 项目模式一致：docUrl 指向后端接口）
// OO 容器内通过 host.docker.internal 访问宿主机 Flask（已验证连通）
const TEMPLATE_URL = 'http://host.docker.internal:5002/api/oo/getTemplate';

// 生成全局唯一的 document.key（时间戳+随机段，永不复用）。
// 关键：OO 服务器按 document.key 缓存文档——同 key 直接返回缓存不重新下载。
// 旧方案 key=日期+版本号，刷新页面后版本号重置 → key 复用 → OO 返回缓存的旧编辑文档，
// 导致"上传新模板后加载的还是上次改过的模板"。唯一 key 强制 OO 重新下载 current.docx。
const genDocKey = () =>
  'zhihanyou-' + Date.now().toString(36) + '-' + Math.random().toString(36).slice(2, 7);

export default function App() {
  const editorRef = useRef(null);
  const [ready, setReady] = useState(false);
  // 模板版本号：每上传一次新 Word 模板 +1，用于换 OO document.key 强制重载新模板
  const [docKey, setDocKey] = useState(genDocKey);  // 当前编辑器的唯一 document.key

  // 用 ref 存储状态信息，避免触发 App 重渲染导致 OnlyOffice 刷新
  const statusRef = useRef('加载中...');

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

  // ---- 模板上传成功：docKeySeed 变化 → OnlyOfficeEditor 内部 useEffect 自动销毁旧实例 + 创建新实例 ----
  // (自己管理生命周期：await destroyEditor + 手动清空 div + 重新 new DocEditor)
  const handleTemplateUploaded = useCallback(() => {
    setReady(false);
    // 预注册新 document.key（先于旧编辑器的 disconnect forcesave 到达）：
    // 旧编辑器销毁时 OO 会对旧 key 延迟触发 forcesave 回写旧文档，
    // 若不提前切换 active key，旧文档会覆盖刚上传的新模板（竞态保护）
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

  return (
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
            docUrl={TEMPLATE_URL}
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
      />
    </div>
  );
}
