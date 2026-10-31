// 用 CDP 给看板截图（Node 18+ 自带 fetch / WebSocket，无需第三方依赖）
// 用法：node tools/screenshot.js <url> <out.png> [宽度=1440] [端口=9333] [最大高度=6000]
// 依赖：本机已启动带 --remote-debugging-port 的 Chromium/Edge
const fs = require('fs');

const [,, url, out, wStr, portStr, maxHStr] = process.argv;
const W = parseInt(wStr || '1440', 10);
const PORT = parseInt(portStr || '9333', 10);
const MAXH = parseInt(maxHStr || '6000', 10);
const SCALE = parseFloat(process.env.SCALE || '1');   // 1 = 1440px 宽，够 README 用且体积小
const BASE = `http://127.0.0.1:${PORT}`;

(async () => {
  // 1. 找一个 page target；没有就新建（Edge 的 /json/new 只接受 PUT）
  let list = await (await fetch(`${BASE}/json/list`)).json();
  let page = list.find(t => t.type === 'page');
  if (!page) {
    page = await (await fetch(`${BASE}/json/new?about:blank`, { method: 'PUT' })).json();
  }

  const ws = new WebSocket(page.webSocketDebuggerUrl);
  let id = 0;
  const pending = new Map();
  const send = (method, params = {}) => new Promise((res, rej) => {
    const i = ++id;
    pending.set(i, { res, rej });
    ws.send(JSON.stringify({ id: i, method, params }));
  });
  ws.onmessage = ev => {
    const m = JSON.parse(ev.data);
    if (m.id && pending.has(m.id)) {
      const p = pending.get(m.id); pending.delete(m.id);
      m.error ? p.rej(new Error(JSON.stringify(m.error))) : p.res(m.result);
    }
  };
  await new Promise(r => (ws.onopen = r));

  // 2. 先在窄视口导航，量出真实文档高度，再按高度重设视口整页截图
  await send('Page.enable');
  await send('Emulation.setDeviceMetricsOverride', { width: W, height: 1000, deviceScaleFactor: SCALE, mobile: false });
  await send('Page.navigate', { url });
  await new Promise(r => setTimeout(r, 2500));

  const h = await send('Runtime.evaluate', {
    expression: 'Math.min(Math.max(document.documentElement.scrollHeight, document.body.scrollHeight), ' + MAXH + ')',
    returnByValue: true,
  });
  const H = Math.ceil(h.result.value || 2000);

  const probe = await send('Runtime.evaluate', {
    expression: 'JSON.stringify({svg:document.querySelectorAll("svg").length, err:(window.__e||null)})',
    returnByValue: true,
  });
  console.log('dom:', probe.result.value, 'height:', H);

  await send('Emulation.setDeviceMetricsOverride', { width: W, height: H, deviceScaleFactor: SCALE, mobile: false });
  await new Promise(r => setTimeout(r, 800));
  const { data } = await send('Page.captureScreenshot', { format: 'png', captureBeyondViewport: true });
  fs.writeFileSync(out, Buffer.from(data, 'base64'));
  console.log('saved', out, `${(fs.statSync(out).size / 1024).toFixed(0)} KB`);
  process.exit(0);
})().catch(e => { console.error('ERR', e.message); process.exit(1); });
