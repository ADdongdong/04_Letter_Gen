// CDP 真实时间探测：打开 oo-probe.html，等待编辑器完全加载后读取内部 API 状态
// 用法: node cdp-check.mjs [等待秒数]
const WAIT_MS = (parseInt(process.argv[2] || '50', 10)) * 1000;
const CDP_PORT = 9222;
const PAGE_URL = 'http://127.0.0.1:5173/oo-probe.html';

async function main() {
  // 1. 创建新 target
  const resp = await fetch(`http://127.0.0.1:${CDP_PORT}/json/new?${encodeURIComponent(PAGE_URL)}`, { method: 'PUT' });
  const target = await resp.json();
  console.log('target:', target.webSocketDebuggerUrl);

  const ws = new WebSocket(target.webSocketDebuggerUrl);
  let msgId = 0;
  const pending = new Map();

  ws.addEventListener('message', (ev) => {
    const data = JSON.parse(ev.data);
    if (data.id && pending.has(data.id)) {
      pending.get(data.id)(data);
      pending.delete(data.id);
    }
  });

  await new Promise((resolve, reject) => {
    ws.addEventListener('open', resolve, { once: true });
    ws.addEventListener('error', reject, { once: true });
  });
  console.log('ws connected, waiting', WAIT_MS, 'ms for editor init...');
  await new Promise((r) => setTimeout(r, WAIT_MS));

  const send = (method, params = {}) => new Promise((resolve) => {
    const id = ++msgId;
    pending.set(id, resolve);
    ws.send(JSON.stringify({ id, method, params }));
  });

  // 2. 读取探测页保存的状态
  const evalJS = async (expr) => {
    const r = await send('Runtime.evaluate', { expression: expr, returnByValue: true, awaitPromise: true });
    if (r.result && r.result.result && r.result.result.value !== undefined) return r.result.result.value;
    return JSON.stringify(r.result || r);
  };

  const last = await evalJS('window.__lastProbe ? JSON.stringify(window.__lastProbe) : "NOT_READY"');
  console.log('=== __lastProbe ===');
  console.log(last);

  // 3. 额外深入检查 frameEditor 内部编辑器实例
  const deep = await evalJS(`(() => {
    const out = {};
    const frames = [];
    (function walk(win, depth){
      if (depth > 8) return;
      const iframes = win.document.querySelectorAll('iframe');
      for (const f of iframes) {
        const name = f.id || f.name || 'iframe';
        frames.push({name, depth});
        try { if (f.contentWindow && f.contentWindow.document) walk(f.contentWindow, depth + 1); } catch(e){}
      }
    })(window, 0);
    out.frames = frames;
    // frameEditor
    const fe = document.querySelector('iframe[name=frameEditor]');
    if (fe && fe.contentWindow) {
      const cw = fe.contentWindow;
      out.fe = {};
      out.fe.insertText = typeof cw.insertText;
      out.fe.hasAscPlugin = !!(cw.Asc && cw.Asc.plugin);
      out.fe.hasAscPluginExecute = !!(cw.Asc && cw.Asc.plugin && cw.Asc.plugin.executeMethod);
      if (cw.Asc) { try { out.fe.ascKeys = Object.keys(cw.Asc).slice(0, 60); } catch(e){} }
      if (cw.Asc && cw.Asc.plugin) { try { out.fe.pluginKeys = Object.keys(cw.Asc.plugin).slice(0, 40); } catch(e){} }
      // 找 window 上所有含 insert/exec 的方法
      const all = Object.keys(cw);
      out.fe.insertish = all.filter(k => /insert|paste|exec|replace/i.test(k)).slice(0, 40);
      out.fe.funcCount = all.length;
    }
    return JSON.stringify(out);
  })()`);
  console.log('=== deep check ===');
  console.log(deep);

  const errs = await evalJS('window.__errs ? JSON.stringify(window.__errs) : "[]"');
  console.log('=== errs ===');
  console.log(errs);

  ws.close();
  process.exit(0);
}

main().catch((e) => { console.error('FAIL', e); process.exit(1); });
