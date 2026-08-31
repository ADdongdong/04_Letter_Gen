import React, { forwardRef, useImperativeHandle, useRef, useEffect, memo } from 'react';

/**
 * OnlyOffice 编辑器（自管理生命周期，绕过 @onlyoffice/document-editor-react 库）
 *
 * 设计原因：
 *   @onlyoffice/react 组件的 destroyEditor() 是异步的，且内部会操作 onlyoffice-editor div 的 DOM
 *   （移除 iframe/div），与 React 的卸载/重渲染产生 removeChild 冲突，反复导致白屏。
 *   这里完全自管：docKeySeed 变化 → await destroyEditor + 手动清空 div + 重新 new DocEditor。
 *   不经过 React 卸载/重挂，避免 removeChild 冲突。
 *
 * insertText 通过 Click2Insert 插件（8.x 官方方式，Hanzheng 同款）：
 *   找 src 含 'click2insert' 的插件 iframe → 调其 window.insertText(text) / postMessage
 */

// 加载 OnlyOffice api.js（带去重 + 状态机）
let apiJsLoading = null;
const ensureDocsApi = (documentServerUrl) => {
  if (window.DocsAPI) return Promise.resolve();
  if (apiJsLoading) return apiJsLoading;
  apiJsLoading = new Promise((resolve, reject) => {
    const script = document.createElement('script');
    script.src = documentServerUrl.replace(/\/$/, '') + '/web-apps/apps/api/documents/api.js';
    script.async = true;
    script.onload = () => { apiJsLoading = null; resolve(); };
    script.onerror = (e) => { apiJsLoading = null; reject(e); };
    document.body.appendChild(script);
  });
  return apiJsLoading;
};

const OnlyOfficeEditor = memo(forwardRef(({ docUrl, docKey, onReady }, ref) => {
  const containerRef = useRef(null);
  const editorRef = useRef(null);
  const isUnmountedRef = useRef(false);
  const docKeyRef = useRef(null);  // 当前 document.key（forceSave 时传给后端 CommandService）

  useEffect(() => {
    isUnmountedRef.current = false;
    const container = containerRef.current;
    const documentServerUrl = (typeof window !== 'undefined' && window.location)
      ? window.location.origin + '/onlyoffice'
      : '/onlyoffice';

    let cancelled = false;

    const init = async () => {
      try {
        // 1. 销毁旧实例（await 等 OO SDK 8.2 异步清理完成）
        if (editorRef.current) {
          try { await editorRef.current.destroyEditor(); } catch (e) { /* ignore */ }
          editorRef.current = null;
        }
        if (window.DocEditor && window.DocEditor.instances) {
          window.DocEditor.instances['onlyoffice-editor'] = undefined;
        }
        // 2. 手动清空 div 内的所有节点（destroyEditor 8.2 未必移除 iframe DOM，必须手动清理）
        if (container) container.innerHTML = '';
        if (cancelled) return;

        // 3. 加载 api.js
        await ensureDocsApi(documentServerUrl);
        if (cancelled) return;

        if (!window.DocsAPI || !window.DocsAPI.DocEditor) {
          console.error('[OO] DocsAPI.DocEditor 不可用');
          return;
        }

        // 4. 创建新编辑器（同一 div id，唯一 document.key 由 App 层生成并传入）
        const currentDocKey = docKey;
        docKeyRef.current = currentDocKey;
        // callbackUrl：OO forcesave 时 POST 当前文档到此（与 getTemplate 同路径，OO 容器经 host.docker.internal 访问）
        const callbackUrl = docUrl.replace('/api/oo/getTemplate', '/api/oo/callback');
        const config = {
          document: {
            fileType: 'docx',
            // key 拼接 docKeySeed：上传新模板时切到 v1 → OO 重新下载转换
            key: currentDocKey,
            title: '银行询证函模板',
            url: docUrl,
          },
          documentType: 'word',
          editorConfig: {
            mode: 'edit',
            lang: 'zh-CN',  // 编辑器界面语言（默认按浏览器语言，未识别时回落英文）
            callbackUrl: callbackUrl,
            customization: {
              forcesave: true,  // 启用强制保存：用户点保存按钮或前端调 server_forceSave 触发回写后端
            },
          },
          height: '100%',
          width: '100%',
          events: {
            onAppReady: () => { if (!cancelled && onReady) onReady('appReady'); },
            onDocumentReady: () => { if (!cancelled && onReady) onReady('documentReady'); },
            onError: (e) => { console.error('[OO] 错误:', e); },
          },
        };

        const editor = new window.DocsAPI.DocEditor('onlyoffice-editor', config);
        editorRef.current = editor;
        // 注册当前活跃 key（callback 回写校验用；初始加载兜底，上传新模板时 App 会预注册）
        fetch('/api/oo/active_key', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ key: currentDocKey }),
        }).catch(() => {});
        if (window.DocEditor) {
          if (!window.DocEditor.instances) window.DocEditor.instances = {};
          window.DocEditor.instances['onlyoffice-editor'] = editor;
        }
      } catch (e) {
        console.error('[OO] 初始化异常:', e);
      }
    };

    init();

    return () => {
      cancelled = true;
      isUnmountedRef.current = true;
      // 卸载时同步销毁（不 await：React 卸载是同步的，不等异步）
      if (editorRef.current) {
        try { editorRef.current.destroyEditor(); } catch (e) {}
        editorRef.current = null;
      }
      if (container) container.innerHTML = '';
    };
  }, [docKey, docUrl]);

  // 暴露给父组件的插入文本方法（兼容旧 API）
  // 优先走 Click2Insert 插件，fallback 到旧内部 API
  useImperativeHandle(ref, () => ({
    insertText: (value) => {
      if (!value) return false;
      console.log('[OO] insertText called:', value.substring(0, 30));
      // 策略1: 旧内部 API（仅 7.x 可用，8.x 已移除）
      try {
        const frameEditor = document.querySelector('iframe[name=frameEditor]');
        if (frameEditor) {
          const win = frameEditor.contentWindow;
          if (win && typeof win.insertText === 'function') {
            win.insertText(value);
            return true;
          }
        }
      } catch (e) { /* ignore */ }
      // 策略2: execCommand（兼容性后备）
      try {
        const frameEditor = document.querySelector('iframe[name=frameEditor]');
        if (frameEditor && frameEditor.contentDocument) {
          frameEditor.contentDocument.execCommand('insertText', false, value);
          return true;
        }
      } catch (e) { /* ignore */ }
      return false;
    },
    // 通过 Click2Insert 插件在光标处插入（8.x 推荐方式）
    insertTextViaPlugin: (value) => {
      if (!value) return false;
      try {
        // 递归找 src 含 'click2insert' 的 iframe
        const findPluginIframe = () => {
          let found = null;
          const visit = (win, depth) => {
            if (depth > 8 || found) return;
            try {
              const ifs = win.document.querySelectorAll('iframe');
              for (let i = 0; i < ifs.length; i++) {
                const f = ifs[i];
                if (f.src && f.src.indexOf('click2insert') !== -1) { found = f; return; }
                try { if (f.contentWindow) visit(f.contentWindow, depth + 1); } catch (e) {}
              }
            } catch (e) {}
          };
          visit(window, 0);
          return found;
        };
        const iframe = findPluginIframe();
        if (!iframe) {
          console.warn('[OO] insertTextViaPlugin: 找不到 click2insert 插件 iframe');
          return false;
        }
        const win = iframe.contentWindow;
        if (win && typeof win.insertText === 'function') {
          win.insertText(value);
          console.log('[OO] ✅ insertTextViaPlugin 成功');
          return true;
        }
        // 兜底：postMessage
        win.postMessage({ type: 'insertText', text: value }, '*');
        return true;
      } catch (e) {
        console.error('[OO] insertTextViaPlugin 异常:', e);
        return false;
      }
    },
    // 触发 forcesave：OO 8.x SDK 无公开 server_forceSave 方法，
    // 改走后端 /api/oo/force_save（向 OO CommandService 发 {"c":"forceSave","key":...} 命令），
    // OO 收到命令后把当前文档（含占位段）POST(status=4) 到 callbackUrl 回写后端。
    forceSave: async () => {
      try {
        const key = docKeyRef.current;
        if (!key) { console.warn('[OO] forceSave 失败：document key 未就绪'); return false; }
        const res = await fetch('/api/oo/force_save', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ key }),
        });
        const data = await res.json();
        if (data.error) { console.error('[OO] forceSave 失败:', data); return false; }
        if (data.no_changes) {
          // OO 报告文档自上次回写后无新修改（如连续二次点击生成）→ 后端已是最新，直接放行渲染
          console.log('[OO] forceSave: 文档无新修改，后端已是最新');
          return { ok: true, no_changes: true };
        }
        console.log('[OO] ✅ forceSave 命令已发送');
        return { ok: true };
      } catch (e) { console.error('[OO] forceSave 异常:', e); return false; }
    },
  }));

  return (
    <div
      ref={containerRef}
      id="onlyoffice-editor"
      style={{ width: '100%', height: '100%' }}
    />
  );
}));

export default OnlyOfficeEditor;
