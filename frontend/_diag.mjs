import { chromium } from 'playwright';
const b = await chromium.launch({ channel: 'msedge', headless: true });
const p = await b.newPage({ viewport: { width: 1280, height: 860 } });
p.on('dialog', async d => { console.log('DIALOG:', d.message().substring(0, 80)); await d.dismiss(); });
await p.goto('http://127.0.0.1:5173/#/generate');
await p.waitForTimeout(2200);
const head = await p.evaluate(async () => {
  const r = await fetch('/_shot_sample.xlsx');
  const buf = await r.arrayBuffer();
  const bytes = new Uint8Array(buf.slice(0, 4));
  return r.status + ' len=' + buf.byteLength + ' magic=' + Array.from(bytes).join(',');
});
console.log('FETCH:', head);
