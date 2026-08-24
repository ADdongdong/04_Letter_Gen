const puppeteer = require('puppeteer-core');

(async () => {
  const browser = await puppeteer.launch({
    executablePath: 'C:\\Users\\10355\\AppData\\Local\\ms-playwright\\chromium_headless_shell-1148\\chrome-win\\headless_shell.exe',
    headless: 'new',
    args: ['--no-sandbox', '--disable-gpu', '--disable-dev-shm-usage'],
  });
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', e => errors.push('PAGEERROR: ' + e.message));
  page.on('requestfailed', req => errors.push(`REQFAIL: ${req.url()} - ${req.failure()?.errorText}`));
  await page.setViewport({ width: 1366, height: 800 });
  await page.goto('http://127.0.0.1:5173/', { waitUntil: 'networkidle2', timeout: 30000 });
  await new Promise(r => setTimeout(r, 15000));

  // 深入 iframe 内部检查文档内容
  const frameInfo = await page.evaluate(() => {
    const iframe = document.querySelector('iframe#frameEditor');
    return { hasIframe: !!iframe, id: iframe?.id };
  });
  console.log('=== OUTER IFRAME ===', JSON.stringify(frameInfo));

  // 尝试访问嵌套 iframe 内容
  const inner = await page.evaluate(() => {
    const outer = document.querySelector('iframe#frameEditor');
    if (!outer || !outer.contentWindow) return { ok: false, msg: 'no contentWindow' };
    const doc = outer.contentDocument || outer.contentWindow.document;
    const text = doc.body ? doc.body.innerText.slice(0, 500) : '';
    // 找内部编辑器 iframe
    const innerIframes = doc.querySelectorAll('iframe');
    const ids = Array.from(innerIframes).map(f => f.id || f.name);
    return { ok: true, outerText: text, innerIds: ids, innerCount: innerIframes.length };
  });
  console.log('=== INNER CONTENT ===');
  console.log(JSON.stringify(inner, null, 2));

  console.log('=== ERRORS ===');
  errors.slice(-10).forEach(e => console.log(e));

  await page.screenshot({ path: 'e:\\13_dingdian\\03_demo\\04_Letter_Gen\\_test_full2.png', fullPage: true });
  console.log('screenshot2 saved');
  await browser.close();
})();
