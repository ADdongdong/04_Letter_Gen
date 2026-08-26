import React, { forwardRef, useImperativeHandle, useRef, memo, useMemo } from 'react';
import { DocumentEditor } from '@onlyoffice/document-editor-react';

/**
 * OnlyOffice 编辑器组件（基于 @onlyoffice/document-editor-react 封装）。
 *
 * insertText 采用与 HanZheng 项目一致的方式：
 *   找到 OnlyOffice 内部 iframe_asc.* 编辑区，
 *   直接调用其 contentWindow.insertText(text) 在「当前光标处」插入文字。
 *   （OnlyOffice 会记忆上次光标位置，无需用户反复点模板）
 */
const OnlyOfficeEditor = memo(forwardRef(({ docUrl, onReady }, ref) => {
  // 递归查找满足条件的 iframe（兼容嵌套 iframe）
  const findIframeByCondition = (startNode, condition, depth = 0) => {
    if (depth > 5) return null;
    const iframes = startNode.querySelectorAll
      ? startNode.querySelectorAll('iframe')
      : startNode.getElementsByTagName('iframe');
    for (let i = 0; i < iframes.length; i++) {
      const iframe = iframes[i];
      if (condition(iframe)) return iframe;
      try {
        const iframeWindow = iframe.contentWindow;
        if (iframeWindow && iframeWindow.document) {
          const nested = findIframeByCondition(iframeWindow.document, condition, depth + 1);
          if (nested) return nested;
        }
      } catch (e) { /* 跨域忽略 */ }
    }
    return null;
  };

  // 找 OnlyOffice 编辑区 iframe
  // 可能是 iframe_asc.*（内部编辑区）或 frameEditor（外层容器）
  const findAscIframe = () => {
    // 先找 iframe_asc.*
    let result = findIframeByCondition(
      document,
      (iframe) => iframe.id && iframe.id.indexOf('iframe_asc.') !== -1
    );
    if (result) return result;

    // 兜底：找 frameEditor（OnlyOffice 外层 iframe）
    result = findIframeByCondition(
      document,
      (iframe) => iframe.name === 'frameEditor'
    );
    return result;
  };

  // 查找 Click2Insert 插件 iframe（部署在 /sdkjs-plugins/click2insert/）
  // 通过递归遍历所有 iframe，匹配 src 含 'click2insert' 的插件窗口
  const findPluginIframe = () => {
    let found = null;
    const visit = (win, depth) => {
      if (depth > 8 || found) return;
      try {
        const ifs = win.document.querySelectorAll('iframe');
        for (let i = 0; i < ifs.length; i++) {
          const f = ifs[i];
          if (f.src && f.src.indexOf('click2insert') !== -1) {
            found = f;
            return;
          }
          try { if (f.contentWindow) visit(f.contentWindow, depth + 1); } catch (e) {}
        }
      } catch (e) {}
    };
    visit(window, 0);
    return found;
  };

  useImperativeHandle(ref, () => ({
    insertText: (value) => {
      if (!value) return false;
      console.log('[OO] insertText called:', value.substring(0, 30));

      // === 策略1：HanZheng 方式 — ascWindow.insertText ===
      try {
        const ascIframe = findAscIframe();
        console.log('[OO] 策略1: findAscIframe=', !!ascIframe, ascIframe?.id);
        if (ascIframe && ascIframe.contentWindow) {
          const ascWindow = ascIframe.contentWindow;
          console.log('[OO] 策略1: has insertText?', typeof ascWindow.insertText);
          if (typeof ascWindow.insertText === 'function') {
            ascWindow.insertText(`${value}`);
            console.log('[OO] ✅ 策略1成功');
            return true;
          }
        }
      } catch (e) { console.warn('[OO] 策略1异常:', e.message); }

      // === 策略2：document.execCommand ===
      try {
        const ascIframe = findAscIframe();
        console.log('[OO] 策略2: iframe=', !!ascIframe, ascIframe?.id || ascIframe?.name);
        if (ascIframe && ascIframe.contentWindow) {
          const doc = ascIframe.contentDocument || ascIframe.contentWindow.document;
          console.log('[OO] 策略2: doc=', !!doc, 'execCommand=', !!doc?.execCommand);
          if (doc && doc.execCommand) {
            doc.execCommand('insertText', false, value);
            console.log('[OO] ✅ 策略2成功 (execCommand)');
            return true;
          }
        }
      } catch (e) { console.warn('[OO] 策略2异常:', e.message); }

      // === 策略3：DOM Selection/Range 操作 ===
      try {
        const ascIframe = findAscIframe();
        if (ascIframe && ascIframe.contentWindow) {
          const win = ascIframe.contentWindow;
          console.log('[OO] 策略3: win=', !!win, 'getSelection=', typeof win?.getSelection);
          if (win.getSelection) {
            const sel = win.getSelection();
            console.log('[OO] 策略3: sel=', !!sel, 'rangeCount=', sel?.rangeCount);
            if (sel && sel.rangeCount > 0) {
              const doc = ascIframe.contentDocument || win.document;
              const range = sel.getRangeAt(0);
              range.deleteContents();
              range.insertNode(doc.createTextNode(value));
              range.collapse(false);
              sel.removeAllRanges();
              sel.addRange(range);
              console.log('[OO] ✅ 策略3成功 (DOM range)');
              return true;
            }
          }
        }
      } catch (e) { console.warn('[OO] 策略3异常:', e.message); }

      // 诊断：列出页面上所有 iframe
      try {
        const all = document.getElementsByTagName('iframe');
        console.log(`[OO] ❌ 全部失败。页面共 ${all.length} 个 iframe:`);
        for (let i = 0; i < Math.min(all.length, 8); i++) {
          console.log(`  [${i}] id=${all[i].id} name=${all[i].name}`);
        }
      } catch (e) {}

      return false;
    },
    // 通过 Click2Insert 插件在光标处插入（8.x 推荐方式：Asc.plugin.callCommand + InsertContent([para], true)）
    insertTextViaPlugin: (value) => {
      if (!value) return false;
      const iframe = findPluginIframe();
      if (!iframe) {
        console.warn('[OO] insertTextViaPlugin: 找不到 click2insert 插件 iframe（请确认插件已部署到 OO 容器 sdkjs-plugins/click2insert/）');
        return false;
      }
      try {
        const win = iframe.contentWindow;
        if (win && typeof win.insertText === 'function') {
          win.insertText(value);
          console.log('[OO] ✅ insertTextViaPlugin 成功');
          return true;
        }
        // 兜底：postMessage（与 Hanzheng connector 一致）
        win.postMessage({ type: 'insertText', text: value }, '*');
        console.log('[OO] ✅ insertTextViaPlugin postMessage 已发送');
        return true;
      } catch (e) {
        console.error('[OO] insertTextViaPlugin 异常:', e);
        return false;
      }
    },
  }));

  // 用 useMemo 锁定 config 引用稳定，避免 React 重复渲染触发 OO SDK 重建 iframe
  // 关键：document.key 不能含 Date.now()，否则 OO SDK 会死循环卸载/重建
  const config = useMemo(() => ({
    document: {
      fileType: 'docx',
      // 容器内该 key 已有转换好的 Editor.bin 缓存（绕过 docservie 缺失 /coauthoring/convert 端点）
      key: 'zhihanyou-demo-v2-' + new Date().toISOString().slice(0, 10),  // 必须与容器内 cache 目录一致
      title: '银行询证函模板',
      url: docUrl,
    },
    documentType: 'word',
    editorConfig: {
      mode: 'edit',
      lang: 'zh-CN',
      customization: {
        autosave: false,
        toolbar: true,
      },
    },
    events: {
      onAppReady: () => {
        console.log('[OO] onAppReady');
        if (onReady) onReady();
      },
      onDocumentReady: () => {
        console.log('[OO] onDocumentReady');
        if (onReady) onReady();
      },
    },
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }), []);  // 故意空依赖：只挂载时计算一次，docUrl 变化用 key 重新挂载

  return (
    <DocumentEditor
      id="onlyoffice-editor"
      documentServerUrl="/onlyoffice"
      config={config}
      height="100%"
      width="100%"
    />
  );
}));

export default OnlyOfficeEditor;
