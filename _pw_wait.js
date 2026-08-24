const puppeteer = require('puppeteer-core');

(async () => {
  const browser = await puppeteer.launch({
    executablePath: 'C:\\Users\\10355\\AppData\\Local\\ms-playwright\\chromium_headless_shell-1148\\chrome-win\\headless_shell.exe',
    headless: 'new',
    args: ['--no-sandbox', '--disable-gpu', '--disable-dev-shm-usage'],
  });
  const page = await browser.newPage();
  const consoleMsgs = [];
  page.on('console', m => consoleMsgs.push(`[${m.type()}] ${m.text()}`));
  page.on('pageerror', e => consoleMsgs.push('PAGEERROR: ' + e.message));
  await page.setViewport({ width: 1366, height: 800 });
  await page.goto('http://127.0.0.1:5173/', { waitUntil: 'domcontentloaded', timeout: 30000 });
  // 轮询等待 iframe 出现
  for (let i = 0; i < 20; i++) {
    await new Promise(r => setTimeout(r, 1000));
    const hasFrame = await page.evaluate(() => {
      return !!(document.querySelector('iframe#frameEditor') || document.querySelector('iframe[name="frameEditor"]'));
    });
    if (hasFrame) { console.log(`iframe appeared at ${i+1}s`); break; }
    if (i === 19) console.log('iframe never appeared after 20s');
  }
  await new Promise(r => setTimeout(r, 8000));
  // 完整状态
  const state = await page.evaluate(() => {
    const iframes = document.querySelectorAll('iframe');
    const all = document.querySelectorAll('*');
    return {
      allIframes: Array.from(iframes).map(f => f.id || f.name),
      editorAreaHTML: (document.querySelector('.editor-area')?.innerHTML || '').slice(0, 300),
      toolbar: document.getElementById('oo-status')?.textContent,
    };
  });
  console.log('=== STATE ===');
  console.log(JSON.stringify(state, null, 2));
  console.log('=== CONSOLE (last 20) ===');
  consoleMsgs.slice(-20).forEach(m => console.log(m));
  await page.screenshot({ path: 'e:\\13_dingdian\\03_demo\\04_Letter_Gen\\_wait_shot.png', fullPage: true });
  console.log('screenshot saved');
  await browser.close();
})();
