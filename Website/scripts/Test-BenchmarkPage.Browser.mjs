import assert from 'node:assert/strict';

const debugPort = Number(process.argv[2]);
const siteOrigin = process.argv[3];
if (!Number.isInteger(debugPort) || !siteOrigin) {
  throw new Error('Usage: node Test-BenchmarkPage.Browser.mjs <debug-port> <site-origin>');
}

const targets = await fetch(`http://127.0.0.1:${debugPort}/json/list`).then(response => response.json());
const page = targets.find(target => target.type === 'page' && target.webSocketDebuggerUrl);
assert.ok(page, 'Chromium did not expose a page target.');

const socket = new WebSocket(page.webSocketDebuggerUrl);
await new Promise((resolve, reject) => {
  socket.addEventListener('open', resolve, { once: true });
  socket.addEventListener('error', reject, { once: true });
});

let nextId = 0;
const pending = new Map();
socket.addEventListener('message', event => {
  const message = JSON.parse(String(event.data));
  if (!message.id || !pending.has(message.id)) return;
  const { resolve, reject } = pending.get(message.id);
  pending.delete(message.id);
  if (message.error) reject(new Error(message.error.message));
  else resolve(message.result);
});

function send(method, params = {}) {
  const id = ++nextId;
  socket.send(JSON.stringify({ id, method, params }));
  return new Promise((resolve, reject) => pending.set(id, { resolve, reject }));
}

async function evaluate(expression) {
  const response = await send('Runtime.evaluate', {
    expression,
    returnByValue: true,
    awaitPromise: true
  });
  if (response.exceptionDetails) {
    throw new Error(response.exceptionDetails.text || 'Browser evaluation failed.');
  }
  return response.result.value;
}

async function waitFor(expression, description) {
  const deadline = Date.now() + 15000;
  while (Date.now() < deadline) {
    if (await evaluate(expression)) return;
    await new Promise(resolve => setTimeout(resolve, 100));
  }
  throw new Error(`Timed out waiting for ${description}.`);
}

const pageUrl = new URL('/benchmarks/', siteOrigin);
pageUrl.searchParams.set('benchmark-workload', 'pdf-structured-generation-net10.0');
pageUrl.searchParams.set('benchmark-os', 'windows');
pageUrl.searchParams.set('benchmark-mode', 'full');
pageUrl.searchParams.set('benchmark-cpu', '0xffff');

await send('Page.enable');
await send('Runtime.enable');
await send('Page.navigate', { url: pageUrl.toString() });

await waitFor(`(() => {
  const selected = document.querySelector('[data-library-comparison-affinity="0xffff"]');
  const table = document.querySelector('[data-library-comparison-table]');
  const rows = document.querySelector('[data-library-comparison-rows]')?.textContent || '';
  const meta = document.querySelector('[data-library-comparison-meta]')?.textContent || '';
  return selected?.getAttribute('aria-pressed') === 'true' &&
    table?.hidden === false &&
    rows.includes('QuestPDF 2026.5.0') &&
    rows.includes('OfficeIMO') &&
    meta.includes('QuestPDF 2026.5.0') && meta.includes('CPU affinity 0xFFFF');
})()`, 'the first CPU-domain benchmark evidence');

assert.equal(await evaluate(`(() => {
  document.querySelector('[data-library-comparison-affinity="0xffff0000"]')?.click();
  return true;
})()`), true);

await waitFor(`(() => {
  const selected = document.querySelector('[data-library-comparison-affinity="0xffff0000"]');
  const rows = document.querySelector('[data-library-comparison-rows]')?.textContent || '';
  const meta = document.querySelector('[data-library-comparison-meta]')?.textContent || '';
  return selected?.getAttribute('aria-pressed') === 'true' &&
    new URL(location.href).searchParams.get('benchmark-cpu') === '0xffff0000' &&
    rows.includes('QuestPDF 2026.5.0') &&
    rows.includes('OfficeIMO') &&
    meta.includes('QuestPDF 2026.5.0') && meta.includes('CPU affinity 0xFFFF0000');
})()`, 'the clicked second CPU-domain benchmark evidence');

await send('Browser.close');
socket.close();
console.log('Rendered benchmark dependency version and CPU-domain selection behavior verified.');
