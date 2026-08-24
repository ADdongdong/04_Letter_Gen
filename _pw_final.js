const puppeteer = require('puppeteer-core');

(async () => {
  const browser = await puppeteer.launch({
    executablePath: 'C:\\Users\\10355\\AppData\\Local\\ms-playwright\\chromium_headless_shell-1148\\chrome-win\\headless_shell.exe',
    headless: 'new',
    args: ['--no-sandbox', '--disable-gpu', '--disable-dev-shm-usage'],
  });
  const page = await browser.newPage();
  const consoleMsgs = [];
  const errors = [];
  page.on('console', m => consoleMsgs.push(`[${m.type()}] ${m.text()}`));
  page.on('pageerror', e => errors.push('PAGEERROR: ' + e.message));
  await page.setViewport({ width: 1366, height: 800 });
  await page.goto('http://127.0.0.1:5173/', { waitUntil: 'networkidle2', timeout: 30000 });
  await new Promise(r => setTimeout(r, 18000));

  // 深入嵌套 iframe 检查文档内容
  const state = await page.evaluate(() => {
    function findTexts(frameDoc, depth) {
      if (depth > 3) return [];
      if (!frameDoc || !frameDoc.body) return [];
      const result = [];
      // 找 canvas 是否渲染（OO 用 canvas 渲染文字）
      const canvases = frameDoc.querySelectorAll('canvas');
      const editors = frameDoc.querySelectorAll('.editor, .viewer, #editor_sdk, .asc-editor');
      const mainTxt = frameDoc.body ? frameDoc.body.innerText.slice(0, 200) : '';
      result.push({ depth, frameCount: frameDoc.querySelectorAll('iframe').length, canvasCount: canvases.length, editorSel: editors.length, text: mainTxt });
      const iframes = frameDoc.querySelectorAll('iframe');
      for (const f of iframes) {
        try {
          const inner = findTexts(f.contentDocument || f.contentWindow.document, depth + 1);
          result.push(...inner);
        } catch (e) { result.push({ depth: depth+1, err: 'cross-origin' }); }
      }
      return result;
    }
    const outer = document.querySelector('iframe#frameEditor');
    if (!outer) return { ok: false, msg: 'no frameEditor' };
    const results = findTexts(outer.contentDocument, 1);
    return { ok: true, results, toolbar: document.getElementById('oo-status')?.textContent };
  });
  console.log('=== STATE ===');
  console.log(JSON.stringify(state, null, 2));
  console.log('=== CONSOLE (onDocumentReady?) ===');
  const relevant = consoleMsgs.filter(m => /DocumentReady|onAppReady|onDocumentReady|error|Error|fail|Failed/i.test(m));
  relevant.slice(-15).forEach(m => console.log(m));
  console.log('=== ERRORS ===');
  errors.slice(-8).forEach(e => console.log(e));
  await page.screenshot({ path: 'e:\\13_dingdian\\03_demo\\04_Letter_Gen\\_final_shot.png', fullPage: true });
  console.log('screenshot saved');
  await browser.close();
})();
