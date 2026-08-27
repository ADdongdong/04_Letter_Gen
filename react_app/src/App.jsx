import React, { useRef, useState, useCallback } from 'react';
import OnlyOfficeEditor from './OnlyOfficeEditor.jsx';
import RightPanel from './RightPanel.jsx';

// OnlyOffice 在本地 Docker 中运行（端口8080），文档通过 Flask 后端 API 返回（与 hanzheng 项目模式一致：docUrl 指向后端接口）
// OO 容器内通过 host.docker.internal 访问宿主机 Flask（已验证连通）
const TEMPLATE_URL = 'http://host.docker.internal:5002/api/oo/getTemplate';

export default function App() {
  const editorRef = useRef(null);
  const [ready, setReady] = useState(false);
  // 模板版本号：每上传一次新 Word 模板 +1，用于换 OO document.key 强制重载新模板
  const [templateVersion, setTemplateVersion] = useState(0);

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

  // ---- 模板上传成功：换 key 强制重建 OnlyOffice 编辑器，加载新模板 ----
  const handleTemplateUploaded = useCallback(() => {
    setReady(false);
    setTemplateVersion((v) => v + 1);
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
          <span className="title">银行询证函模板（OnlyOffice）</span>
          {/* 状态栏内容由 updateStatus 直接操作 DOM（脱离 React 管理），避免 textContent 与 React diff 冲突 */}
          <span id="oo-status" className="status" />
        </div>
        <div className="editor-area">
          {/* OnlyOfficeEditor 被 React.memo 包裹，props 不变时绝不重渲染 */}
          <OnlyOfficeEditor
            key={'oo-' + templateVersion}
            docKeySeed={'v' + templateVersion}
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
        updateStatus={updateStatus}
        onTemplateUploaded={handleTemplateUploaded}
      />
    </div>
  );
}
