import React, { forwardRef, useImperativeHandle, useRef, memo, useEffect, useState, useCallback } from 'react';

/**
 * OnlyOffice 编辑器（直接调 DocsAPI，跳过 @onlyoffice/document-editor-react）。
 *
 * 关键原理：
 *   DocsAPI.DocEditor(id, config) 创建 iframe 时会把 placeholder div 从 DOM 移除
 *   （replaceWith iframe），且 React 卸载/重建 SDK 用过的 div 会触发 removeChild 崩溃。
 *
 * 因此本组件【完全不让 React 管理 SDK 的 placeholder div】：
 *   - React 只渲染一个空容器 <div ref={containerRef}>（固定，从不变化）。
 *   - docUrl 变化 → instanceId+1 → useEffect 里手动操作容器：
 *       1. container.innerHTML = ''  清空旧 iframe（不经过 React，安全）
 *       2. 手动 createElement('div') 设 id，append 到容器
 *       3. DocsAPI.DocEditor(editorId, config) 创建全新实例
 *   - SDK 移除 div/iframe 时 React 不知情、不追踪、不崩溃。
 *
 * document.key 与 docUrl 关联（不同模板不同 key，避免 OO 缓存误判）。
 *
 * insertText：找 OnlyOffice 内部 iframe_asc.* 编辑区，调用 contentWindow.insertText。
 */

const DOCS_API_URL = '/onlyoffice/web-apps/apps/api/documents/api.js';

function loadDocsApi() {
  return new Promise((resolve, reject) => {
    if (window.DocsAPI && window.DocsAPI.DocEditor) return resolve();
    const existing = document.querySelector('script[data-onlyoffice-api]');
    if (existing) {
      if (window.DocsAPI && window.DocsAPI.DocEditor) return resolve();
      existing.addEventListener('load', () => resolve());
      existing.addEventListener('error', reject);
      return;
    }
    const s = document.createElement('script');
    s.src = DOCS_API_URL;
    s.setAttribute('data-onlyoffice-api', 'true');
    s.onload = () => resolve();
    s.onerror = () => reject(new Error('DocsAPI 脚本加载失败: ' + DOCS_API_URL));
    document.head.appendChild(s);
  });
}

function deriveKey(docUrlStr) {
  if (docUrlStr.indexOf('/api/oo/template/') !== -1) {
    const id = docUrlStr.split('/api/oo/template/')[1].split(/[?#]/)[0];
    return 'zhihanyou-tpl-' + id;
  }
  return 'zhihanyou-tpl-default';
}

/**
 * 把传给 OnlyOffice 的 document.url 转成【容器内可访问的绝对 URL】。
 * OnlyOffice 运行在 Docker bridge 网络中，容器内 127.0.0.1 指向容器自身，
 * 必须用 host.docker.internal 才能访问宿主机 Flask(5002)。
 * - 相对路径（如 /api/oo/template/<id>）→ 拼上 host.docker.internal:5002
 * - 绝对 http(s) 但 host 为 127.0.0.1/localhost → 替换为 host.docker.internal
 */
function toAbsoluteDocUrl(docUrl) {
  if (!docUrl) return docUrl;
  if (/^https?:\/\//i.test(docUrl)) {
    return docUrl.replace(
      /^https?:\/\/(127\.0\.0\.1|localhost)(:\d+)?/i,
      (m, host, port) => 'http://host.docker.internal' + (port || ':5002')
    );
  }
  return 'http://host.docker.internal:5002' + (docUrl.startsWith('/') ? docUrl : '/' + docUrl);
}

const OnlyOfficeEditor = memo(forwardRef(({ docUrl, onReady }, ref) => {
  const containerRef = useRef(null);
  const [instanceId, setInstanceId] = useState(0);
  const readyRef = useRef(false);
  const onReadyRef = useRef(onReady);
  onReadyRef.current = onReady;
  const firstRunRef = useRef(true);

  const editorId = 'onlyoffice-editor-' + instanceId;

  // ---- docUrl 变化 → instanceId+1（首次不触发，初始 0 已创建）----
  useEffect(() => {
    if (firstRunRef.current) {
      firstRunRef.current = false;
      return;
    }
    setInstanceId((n) => n + 1);
  }, [docUrl]);

  // ---- 每个 instanceId 手动构建 div + 创建 OO 实例（容器手动管理，React 不追踪）----
  useEffect(() => {
    let cancelled = false;
    readyRef.current = false;
    loadDocsApi().then(() => {
      if (cancelled) return;
      const container = containerRef.current;
      if (!container) return;
      // 清空旧内容（旧 iframe/div 全移除，不经过 React，安全）
      container.innerHTML = '';
      // 手动创建 placeholder div
      const div = document.createElement('div');
      div.id = editorId;
      div.style.width = '100%';
      div.style.height = '100%';
      container.appendChild(div);
      // 创建 OO 实例
      const config = {
        document: {
          fileType: 'docx',
          key: deriveKey(docUrl),
          title: '银行询证函模板',
          url: toAbsoluteDocUrl(docUrl),
        },
        documentType: 'word',
        editorConfig: {
          mode: 'edit',
          lang: 'zh-CN',
          customization: { autosave: false, toolbar: true },
        },
        events: {
          onAppReady: () => {
            console.log('[OO] onAppReady', editorId);
            if (!readyRef.current) {
              readyRef.current = true;
              onReadyRef.current?.();
            }
          },
          onDocumentReady: () => {
            console.log('[OO] onDocumentReady', editorId);
            if (!readyRef.current) {
              readyRef.current = true;
              onReadyRef.current?.();
            }
          },
        },
      };
      try {
        const absUrl = toAbsoluteDocUrl(docUrl);
        window.DocsAPI.DocEditor(editorId, config);
        console.log('[OO] DocEditor created', editorId, 'absUrl=', absUrl, 'key=', deriveKey(docUrl));
      } catch (e) {
        console.error('[OO] DocEditor 创建失败:', e);
      }
    }).catch((e) => console.error('[OO] loadDocsApi 失败:', e.message));
    return () => {
      cancelled = true;
      // 卸载时清空容器（若容器仍在）
      const container = containerRef.current;
      if (container) container.innerHTML = '';
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [instanceId]);

  // ---- insertText ----
  const findIframeByCondition = (startNode, condition, depth = 0) => {
    if (depth > 5) return null;
    const iframes = startNode.querySelectorAll
      ? startNode.querySelectorAll('iframe')
      : startNode.getElementsByTagName('iframe');
    for (let i = 0; i < iframes.length; i++) {
      const iframe = iframes[i];
      if (condition(iframe)) return iframe;
      try {
        const cw = iframe.contentWindow;
        if (cw && cw.document) {
          const nested = findIframeByCondition(cw.document, condition, depth + 1);
          if (nested) return nested;
        }
      } catch (e) { /* 跨域 */ }
    }
    return null;
  };

  const findAscIframe = () => {
    let r = findIframeByCondition(document, (iframe) => iframe.id && iframe.id.indexOf('iframe_asc.') !== -1);
    if (r) return r;
    return findIframeByCondition(document, (iframe) => iframe.name === 'frameEditor');
  };

  useImperativeHandle(ref, () => ({
    insertText: (value) => {
      if (!value) return false;
      console.log('[OO] insertText:', value.substring(0, 30));
      try {
        const iframe = findAscIframe();
        if (iframe && iframe.contentWindow && typeof iframe.contentWindow.insertText === 'function') {
          iframe.contentWindow.insertText(`${value}`);
          console.log('[OO] 策略1成功');
          return true;
        }
      } catch (e) { console.warn('[OO] 策略1异常:', e.message); }
      try {
        const iframe = findAscIframe();
        if (iframe && iframe.contentDocument && iframe.contentDocument.execCommand) {
          iframe.contentDocument.execCommand('insertText', false, value);
          console.log('[OO] 策略2成功');
          return true;
        }
      } catch (e) { console.warn('[OO] 策略2异常:', e.message); }
      try {
        const iframe = findAscIframe();
        const win = iframe && iframe.contentWindow;
        if (win && win.getSelection) {
          const sel = win.getSelection();
          if (sel && sel.rangeCount > 0) {
            const doc = iframe.contentDocument || win.document;
            const range = sel.getRangeAt(0);
            range.deleteContents();
            range.insertNode(doc.createTextNode(value));
            range.collapse(false);
            sel.removeAllRanges();
            sel.addRange(range);
            console.log('[OO] 策略3成功');
            return true;
          }
        }
      } catch (e) { console.warn('[OO] 策略3异常:', e.message); }
      console.log('[OO] ❌ insertText 全部失败');
      return false;
    },
  }));

  // React 只渲染固定容器；docUrl 为空时显示占位提示，不创建编辑器
  return (
    <div
      ref={containerRef}
      style={{ width: '100%', height: '100%', position: 'relative' }}
    >
      {!docUrl && (
        <div className="oo-placeholder">
          <p>暂无 Word 模板</p>
          <p>请使用右侧「上传 Word 模板」按钮加载文档</p>
        </div>
      )}
    </div>
  );
}));

export default OnlyOfficeEditor;