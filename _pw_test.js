const puppeteer = require('puppeteer-core');
const fs = require('fs');

(async () => {
  const browser = await puppeteer.launch({
    executablePath: 'C:\\Users\\10355\\AppData\\Local\\ms-playwright\\chromium_headless_shell-1148\\chrome-win\\headless_shell.exe',
    headless: 'new',
    args: ['--no-sandbox', '--disable-gpu', '--disable-dev-shm-usage'],
  });
  const page = await browser.newPage();
  const errors = [];
  const consoleMsgs = [];
  page.on('pageerror', e => errors.push('PAGEERROR: ' + e.message));
  page.on('console', m => consoleMsgs.push(`[${m.type()}] ${m.text()}`));
  page.on('requestfailed', req => errors.push(`REQFAIL: ${req.url()} - ${req.failure()?.errorText}`));
  await page.setViewport({ width: 1366, height: 800 });
  await page.goto('http://127.0.0.1:5173/', { waitUntil: 'networkidle2', timeout: 30000 });
  // wait for OO to load
  await new Promise(r => setTimeout(r, 12000));
  // get editor area state
  const editorState = await page.evaluate(() => {
    const area = document.querySelector('.editor-area');
    const iframes = area ? area.querySelectorAll('iframe') : [];
    const result = {
      areaExists: !!area,
      areaHeight: area?.offsetHeight,
      iframeCount: iframes.length,
      iframeIds: Array.from(iframes).map(f => f.id || f.name),
      toolbarText: document.getElementById('oo-status')?.textContent,
      bodyText: document.body.innerText.slice(0, 300),
    };
    return result;
  });
  console.log('=== EDITOR STATE ===');
  console.log(JSON.stringify(editorState, null, 2));
  console.log('=== ERRORS (last 15) ===');
  errors.slice(-15).forEach(e => console.log(e));
  console.log('=== CONSOLE (last 15) ===');
  consoleMsgs.slice(-15).forEach(m => console.log(m));
  await page.screenshot({ path: 'e:\\13_dingdian\\03_demo\\04_Letter_Gen\\_test_full.png', fullPage: true });
  console.log('screenshot saved');
  await browser.close();
})();
