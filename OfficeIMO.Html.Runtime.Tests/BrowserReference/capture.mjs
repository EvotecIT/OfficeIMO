import { createHash } from 'node:crypto';
import { readFile, mkdir, writeFile } from 'node:fs/promises';
import { createServer } from 'node:http';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright';

const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, '..', 'Fixtures');
const output = process.argv[2] && resolve(process.argv[2]);
if (!output) throw new Error('Pass an output directory for the browser reference.');

const preactHtml = `<!doctype html><html lang="en"><head><meta charset="utf-8"><title>Preact report</title></head><body>
<div id="app"></div><script src="/preact.umd.js"></script><script src="/hooks.umd.js"></script>
<script src="/report.js"></script><script>localStorage.setItem('report:name','Monthly');mountReport()</script>
</body></html>`;
const preactData = '[{"name":"North","value":24},{"name":"South","value":18}]';
const specs = {
  vanilla: { folder: 'StandaloneApplication', files: ['app.js', 'view.js', 'app.css', 'theme.css',
    'data.json', 'review.js', 'health.txt'] },
  'react-build': { folder: 'ReactBuild', files: ['style.css', 'data.json', 'app.js',
    'chunk-chunk-YDHNFGZD.js', 'chunk-review-7VOCLQDE.js'] },
  preact: { folder: 'Preact', files: ['preact.umd.js', 'hooks.umd.js', 'report.js'] },
  legacy: { folder: 'LegacyApplication', files: [] }
};
const expectedStates = {
  vanilla: { route: '/review', heading: 'Review report', selection: 'Quarterly / South', total: 'Total: 25' },
  'react-build': { route: '/review', heading: 'Review report', selection: 'Quarterly / South', total: 'Total: 21' },
  preact: { route: '/index.html', heading: 'Application report', selection: 'Report: Quarterly', total: 'Total: 18' },
  legacy: { route: '/index.html', heading: 'Approval register', state: 'Approved' }
};

function mediaType(name) {
  if (name.endsWith('.js')) return 'text/javascript';
  if (name.endsWith('.css')) return 'text/css';
  if (name.endsWith('.json')) return 'application/json';
  if (name.endsWith('.html')) return 'text/html; charset=utf-8';
  return 'text/plain';
}

async function serve(spec, kind, request, response, served) {
  const path = new URL(request.url, 'http://localhost').pathname;
  let file;
  let content;
  let status = 200;
  if (path === '/' || path === '/index.html' || path === '/review') {
    if (kind === 'preact') content = Buffer.from(preactHtml);
    else {
      file = kind === 'vanilla' && path === '/review' ? 'review.html' : 'index.html';
      content = await readFile(join(fixtures, spec.folder, file));
    }
  } else if (kind === 'preact' && path === '/data.json') {
    content = Buffer.from(preactData);
    file = 'data.json';
  } else {
    file = path.slice(1);
    if (!spec.files.includes(file)) {
      response.writeHead(404).end();
      return;
    }
    const filePath = kind === 'react-build' && file.endsWith('.js')
      ? join(fixtures, spec.folder, 'dist', file) : join(fixtures, spec.folder, file);
    content = await readFile(filePath);
    if (kind === 'vanilla' && file === 'health.txt') status = 503;
  }
  const identity = file ?? 'generated-index.html';
  served.push({ path, status, sha256: createHash('sha256').update(content).digest('hex') });
  response.writeHead(status, { 'Content-Type': mediaType(identity), 'Content-Length': content.length });
  response.end(content);
}

async function act(kind, page) {
  if (kind === 'vanilla') {
    await page.waitForFunction(() => window.applicationReady === true);
    await page.getByLabel('Report title').fill('Quarterly');
    await page.getByLabel('Region').selectOption('South');
    await page.getByRole('button', { name: 'Prepare review' }).click();
    await page.waitForFunction(() => window.reviewReady === true);
    return { route: new URL(page.url()).pathname, heading: await page.locator('h1').innerText(),
      selection: await page.locator('#selection').innerText(), total: await page.locator('#review-total').innerText() };
  }
  if (kind === 'react-build') {
    await page.waitForFunction(() => document.querySelector('#total')?.textContent === 'Total: 42');
    await page.getByLabel('Region').selectOption('South');
    await page.getByLabel('Report title').fill('Quarterly');
    await page.getByRole('button', { name: 'Add adjustment' }).click();
    await page.getByRole('button', { name: 'Prepare review' }).click();
    await page.getByRole('heading', { name: 'Review report' }).waitFor();
    return { route: new URL(page.url()).pathname, heading: await page.locator('h1').innerText(),
      selection: await page.locator('#report-heading').innerText(), total: await page.locator('#total').innerText() };
  }
  if (kind === 'preact') {
    await page.waitForFunction(() => document.querySelector('#total')?.textContent === 'Total: 42');
    await page.getByLabel('Region').selectOption('South');
    await page.getByLabel('Report name').fill('Quarterly');
    await page.getByText('Report: Quarterly').waitFor();
    return { route: new URL(page.url()).pathname, heading: await page.locator('h1').innerText(),
      selection: await page.locator('#report-name').innerText(), total: await page.locator('#total').innerText() };
  }
  await page.getByRole('button', { name: 'Approve' }).click();
  await page.getByText('Approved').waitFor();
  return { route: new URL(page.url()).pathname, heading: await page.locator('h1').innerText(),
    state: await page.locator('#state').innerText() };
}

await mkdir(output, { recursive: true });
const browser = await chromium.launch({ headless: true });
const manifest = { browser: `Chromium ${browser.version()}`, playwright: '1.62.1',
  viewport: { width: 816, height: 900, deviceScaleFactor: 1 }, cases: {} };
try {
  for (const [kind, spec] of Object.entries(specs)) {
    const served = [];
    const errors = [];
    const server = createServer((request, response) => {
      serve(spec, kind, request, response, served).catch(error => {
        errors.push(String(error));
        if (!response.headersSent) response.writeHead(500);
        response.end();
      });
    });
    await new Promise(done => server.listen(0, '127.0.0.1', done));
    const address = server.address();
    const context = await browser.newContext({ viewport: { width: 816, height: 900 }, deviceScaleFactor: 1 });
    const page = await context.newPage();
    page.on('pageerror', error => errors.push(String(error)));
    page.on('requestfailed', request => errors.push(`${request.url()}: ${request.failure()?.errorText}`));
    try {
      await page.goto(`http://127.0.0.1:${address.port}/index.html`);
      const state = await act(kind, page);
      if (JSON.stringify(state) !== JSON.stringify(expectedStates[kind])) {
        throw new Error(`${kind} reached unexpected state: ${JSON.stringify(state)}`);
      }
      const folder = join(output, kind);
      await mkdir(folder, { recursive: true });
      await page.screenshot({ path: join(folder, 'chromium-screen.png'), fullPage: true });
      await page.pdf({ path: join(folder, 'chromium-print.pdf'), format: 'A4', printBackground: true });
      manifest.cases[kind] = { state, served, errors };
      if (errors.length) throw new Error(`${kind} browser errors: ${errors.join('; ')}`);
    } finally {
      await context.close();
      await new Promise((done, reject) => server.close(error => error ? reject(error) : done()));
    }
  }
} finally {
  await browser.close();
  await writeFile(join(output, 'chromium-manifest.json'), JSON.stringify(manifest, null, 2) + '\n');
}
