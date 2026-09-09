import { chromium } from 'playwright';
const b = await chromium.launch({ channel: 'msedge', headless: true });
const p = await b.newPage({ viewport: { width: 1280, height: 860 } });
const shot = (f) => p.screenshot({ path: 'docs/images/' + f });
console.log('bootstrap ok');

await p.goto('http://127.0.0.1:5173/#/templates');
await p.waitForTimeout(2200);
await shot('ui_01_模板列表页.png');
console.log('shot 1');

await p.goto('http://127.0.0.1:5173/#/config?id=tpl_c6e74f35');
await p.waitForTimeout(16000);
await shot('ui_02a_配置编辑页整体.png');
console.log('shot 2a');

await p.evaluate(() => {
  const f = document.querySelector('iframe[name=frameEditor]');
  if (f && f.contentDocument) {
    const el = Array.from(f.contentDocument.querySelectorAll('span, p')).find(e => (e.textContent || '').includes('表格将在此处展示'));
    if (el) el.scrollIntoView({ block: 'center' });
  }
});
await p.waitForTimeout(600);
await shot('ui_02b_占位段插入效果.png');
console.log('shot 2b');

await p.evaluate(() => {
  const t = Array.from(document.querySelectorAll('div')).find(e => e.textContent.trim() === '表格样式（高级，可选）');
  if (t) t.click();
});
await p.waitForTimeout(500);
await p.locator('input[placeholder*="模板名称"]').scrollIntoViewIfNeeded();
await p.waitForTimeout(300);
await shot('ui_02c_表格样式高级区.png');
console.log('shot 2c');

await p.goto('http://127.0.0.1:5173/#/generate');
await p.waitForTimeout(2200);
await shot('ui_03a_制函页布局.png');
console.log('shot 3a');

await p.click('.sfield-trigger');
await p.waitForTimeout(400);
await p.locator('.sfield-opt', { hasText: '往来函证配置1' }).click();
await p.waitForTimeout(500);
await p.evaluate(async () => {
  const buf = await (await fetch('/_shot_sample.xlsx')).arrayBuffer();
  const dt = new DataTransfer();
  dt.items.add(new File([buf], '批量数据.xlsx'));
  const inp = document.querySelector('#gen-excel');
  inp.files = dt.files;
  inp.dispatchEvent(new Event('change', { bubbles: true }));
});
await p.waitForTimeout(2500);
await shot('ui_03b_制函页匹配结果.png');
console.log('shot 3b');

await p.goto('http://127.0.0.1:5173/#/help');
await p.waitForTimeout(1500);
await p.screenshot({ path: 'docs/images/ui_04_使用说明页.png', fullPage: true });
console.log('shot 4');

await b.close();
console.log('ALL DONE');






