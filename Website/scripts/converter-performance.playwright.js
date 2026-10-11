async (page) => {
  const consoleErrors = [];
  const requests = [];
  page.on('console', message => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });
  page.on('pageerror', error => consoleErrors.push(error.message));
  page.on('request', request => requests.push(request.url()));
  page.on('worker', worker => worker.on('console', message => {
    if (message.type() === 'error') consoleErrors.push('worker: ' + message.text());
  }));
  await page.addInitScript(() => {
    const tools = Object.create(null);
    const objectUrlBlobs = new Map();
    const createObjectURL = URL.createObjectURL.bind(URL);
    const revokeObjectURL = URL.revokeObjectURL.bind(URL);
    Object.defineProperty(window, '__officeImoWebMcpTools', { value: tools, configurable: true });
    Object.defineProperty(window, '__officeImoObjectUrlBlobs', { value: objectUrlBlobs, configurable: true });
    URL.createObjectURL = blob => {
      const url = createObjectURL(blob);
      objectUrlBlobs.set(url, blob);
      return url;
    };
    URL.revokeObjectURL = url => {
      objectUrlBlobs.delete(url);
      revokeObjectURL(url);
    };
    Object.defineProperty(document, 'modelContext', {
      configurable: true,
      value: {
        registerTool: async (tool, options) => {
          tools[tool.name] = tool;
          options?.signal?.addEventListener('abort', () => { delete tools[tool.name]; }, { once: true });
        }
      }
    });
  });
  await page.setViewportSize({ width: 1440, height: 1000 });
  // run-code has no URL global; the harness opens the site's /convert/ page first.
  const origin = page.url().match(/^https?:\/\/[^/]+/)[0];
  const tools = [
    { routeId: 'docx-pdf', toolId: 'word-to-pdf' },
    { routeId: 'xlsx-pdf', toolId: 'excel-to-pdf' },
    { routeId: 'pptx-pdf', toolId: 'powerpoint-to-pdf' }
  ];

  // The directory is plain HTML: it must not start the engine or expose the Website Tool.
  await page.goto(`${origin}/convert/`, { waitUntil: 'domcontentloaded' });
  await page.locator('.imo-browser-tools').waitFor({ state: 'visible', timeout: 60000 });
  await page.waitForTimeout(500);
  const directoryMilliseconds = await page.evaluate(() => performance.now());
  const prematureRuntime = requests.find(name => /\/apps\/officeimo-converter\/|\/_framework\/|\.wasm(?:\?|$)/.test(name));
  if (prematureRuntime || await page.locator('iframe').count()) {
    throw new Error(`The directory started the document runtime before a tool was opened: ${prematureRuntime || 'iframe'}`);
  }
  const removedOutsideConverter = await page.evaluate(() => !window.__officeImoWebMcpTools?.convert_selected_document);

  // Opening a tool shows the page immediately; the engine finishes downloading in a worker.
  requests.length = 0;
  await page.locator('.imo-browser-tools a[data-tool="word-to-pdf"]').click();
  await page.locator('[data-browser-tool="word-to-pdf"]').waitFor({ state: 'visible', timeout: 60000 });
  const interactiveMilliseconds = await page.evaluate(() => performance.now());
  await page.locator('[data-bt-engine-status][data-state="ready"]').waitFor({ state: 'attached', timeout: 120000 });
  const startupMilliseconds = await page.evaluate(() => performance.now());
  // The Excel, PowerPoint and Visio engines themselves (OfficeIMO.Excel.<fingerprint>.wasm); small *.Pdf bridges are shared.
  const unrelatedOfficeEngine = requests.find(url => /\/OfficeIMO\.(?:Excel|PowerPoint|Visio)\.[a-z0-9]+\.wasm$/.test(url));
  if (unrelatedOfficeEngine) throw new Error(`Word conversion downloaded an unrelated engine: ${unrelatedOfficeEngine}`);
  await page.waitForFunction(() => Boolean(window.__officeImoWebMcpTools?.convert_selected_document), null, { timeout: 60000 });
  const restoredWithConverter = await page.evaluate(() => Boolean(window.__officeImoWebMcpTools?.convert_selected_document));

  const readMemory = async () => page.evaluate(async () => {
    const memory = performance.memory;
    if (!memory || !Number.isFinite(memory.usedJSHeapSize) || memory.totalJSHeapSize <= 0) return null;
    const engine = window.OfficeIMOBrowserTool ? await window.OfficeIMOBrowserTool.engineMemory() : 0;
    return { used: memory.usedJSHeapSize, total: memory.totalJSHeapSize, engine: Number(engine) || 0 };
  });

  const results = [];
  let maximumBrowserHeapBytes = 0;
  let maximumEngineMemoryBytes = 0;
  let webMcp = null;
  for (let index = 0; index < tools.length; index++) {
    const { routeId, toolId } = tools[index];
    if (index > 0) await page.goto(`${origin}/browser/${toolId}/`, { waitUntil: 'domcontentloaded' });
    await page.locator('[data-bt-engine-status][data-state="ready"]').waitFor({ state: 'attached', timeout: 120000 });
    await page.locator('[data-bt-sample]').click();
    await page.locator('[data-bt-files] li').waitFor({ state: 'visible', timeout: 60000 });
    await page.locator('[data-bt-run]:not([disabled])').waitFor({ state: 'visible', timeout: 60000 });

    const measureConversion = async (previousDownloadUrl, useWebMcp) => {
      let sampling = true;
      let memorySamples = 0;
      let peakBrowserHeapBytes = 0;
      let peakBrowserUsedHeapBytes = 0;
      let peakEngineMemoryBytes = 0;
      const sample = async () => {
        const memory = await readMemory();
        if (!memory) throw new Error('Chromium performance.memory is unavailable; peak-memory evidence is required.');
        memorySamples++;
        peakBrowserHeapBytes = Math.max(peakBrowserHeapBytes, memory.total);
        peakBrowserUsedHeapBytes = Math.max(peakBrowserUsedHeapBytes, memory.used);
        peakEngineMemoryBytes = Math.max(peakEngineMemoryBytes, memory.engine);
      };
      await sample();
      const sampler = (async () => {
        while (sampling) {
          await page.waitForTimeout(20);
          await sample();
        }
      })();

      let webMcpOutput = null;
      if (useWebMcp) {
        webMcpOutput = await page.evaluate(async () => {
          const tool = window.__officeImoWebMcpTools.convert_selected_document;
          const cancelledSignal = new AbortController();
          cancelledSignal.abort();
          const cancelled = await tool.execute({}, { signal: cancelledSignal.signal });
          const output = await tool.execute({}, { signal: new AbortController().signal });
          return {
            registeredTools: Object.keys(window.__officeImoWebMcpTools).sort(),
            schema: tool.inputSchema,
            annotations: tool.annotations,
            cancelled,
            output,
            outputCharacters: JSON.stringify(output).length
          };
        });
      } else {
        await page.locator('[data-bt-run]').click();
      }
      await page.waitForFunction(previousUrl => {
        const output = document.querySelector('[data-bt-output]');
        const link = document.querySelector('[data-bt-download="primary"]');
        return output?.getAttribute('data-result-state') === 'ok' && Boolean(link?.href?.startsWith('blob:') && link.href !== previousUrl);
      }, previousDownloadUrl, { timeout: 120000 });
      sampling = false;
      await sampler;
      await sample();

      const downloadUrl = await page.locator('[data-bt-download="primary"]').getAttribute('href');
      const pdfMagic = await page.evaluate(async url => {
        const blob = window.__officeImoObjectUrlBlobs?.get(url);
        if (!(blob instanceof Blob)) throw new Error(`Generated output blob is unavailable for ${url}.`);
        const bytes = new Uint8Array(await blob.arrayBuffer());
        return String.fromCharCode(...bytes.slice(0, 4));
      }, downloadUrl);
      const metrics = await page.locator('[data-bt-output]').evaluate(element => ({
        conversionMilliseconds: Number(element.getAttribute('data-conversion-ms') || '0'),
        peakRetainedBytes: Number(element.getAttribute('data-peak-retained-bytes') || '0'),
        resultBytes: Number(element.getAttribute('data-result-bytes') || '0')
      }));
      return { downloadUrl, pdfMagic, memorySamples, peakBrowserHeapBytes, peakBrowserUsedHeapBytes, peakEngineMemoryBytes, webMcpOutput, ...metrics };
    };

    const first = await measureConversion('', index === 0);
    const repeat = await measureConversion(first.downloadUrl, false);
    if (first.webMcpOutput) webMcp = first.webMcpOutput;
    maximumBrowserHeapBytes = Math.max(maximumBrowserHeapBytes, first.peakBrowserHeapBytes, repeat.peakBrowserHeapBytes);
    maximumEngineMemoryBytes = Math.max(maximumEngineMemoryBytes, first.peakEngineMemoryBytes, repeat.peakEngineMemoryBytes);
    results.push({
      routeId,
      toolId,
      memorySamples: first.memorySamples,
      peakBrowserHeapBytes: first.peakBrowserHeapBytes,
      peakBrowserUsedHeapBytes: first.peakBrowserUsedHeapBytes,
      peakEngineMemoryBytes: first.peakEngineMemoryBytes,
      conversionMilliseconds: first.conversionMilliseconds,
      peakRetainedBytes: first.peakRetainedBytes,
      resultBytes: first.resultBytes,
      pdfMagic: first.pdfMagic,
      repeatMemorySamples: repeat.memorySamples,
      repeatPeakBrowserHeapBytes: repeat.peakBrowserHeapBytes,
      repeatPeakBrowserUsedHeapBytes: repeat.peakBrowserUsedHeapBytes,
      repeatConversionMilliseconds: repeat.conversionMilliseconds,
      repeatPeakRetainedBytes: repeat.peakRetainedBytes,
      repeatResultBytes: repeat.resultBytes,
      repeatPdfMagic: repeat.pdfMagic
    });
  }

  const chooseFile = async (name, bytesOrSample) => {
    await page.locator('[data-bt-file-input]').evaluate(async (input, args) => {
      const bytes = Array.isArray(args.bytes)
        ? new Uint8Array(args.bytes)
        : new Uint8Array(await (await fetch(args.bytes)).arrayBuffer());
      const file = new File([bytes], args.name, { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
      const transfer = new DataTransfer();
      transfer.items.add(file);
      input.files = transfer.files;
      input.dispatchEvent(new Event('change', { bubbles: true }));
    }, { name, bytes: bytesOrSample });
    await page.locator('[data-bt-files] li').filter({ hasText: name.slice(0, 20) }).waitFor({ state: 'visible', timeout: 60000 });
    // The filename appears before asynchronous staging completes. Exercise conversion
    // only after the same readiness gate as the visible Run button has settled.
    await page.waitForFunction(() => !document.querySelector('[data-bt-run]')?.disabled, null, { timeout: 60000 });
  };

  await page.goto(`${origin}/browser/word-to-pdf/`, { waitUntil: 'domcontentloaded' });
  await page.locator('[data-bt-engine-status][data-state="ready"]').waitFor({ state: 'attached', timeout: 120000 });
  await page.waitForFunction(() => Boolean(window.__officeImoWebMcpTools?.convert_selected_document), null, { timeout: 60000 });
  await chooseFile(`${'a'.repeat(179)}🚀.docx`, '/apps/officeimo-converter/samples/basic.docx');
  const longNameWebMcp = await page.evaluate(async () => {
    const tool = window.__officeImoWebMcpTools.convert_selected_document;
    const output = await tool.execute({}, { signal: new AbortController().signal });
    const fileName = String(output.outputFileName || '');
    const hasUnpairedSurrogate = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(fileName);
    return { output, outputCharacters: JSON.stringify(output).length, outputFileNameCharacters: fileName.length, hasUnpairedSurrogate };
  });

  await chooseFile(`${'malformed-'.repeat(18)}document.docx`, [110, 111, 116, 45, 97, 110, 45, 111, 112, 101, 110, 45, 120, 109, 108, 45, 112, 97, 99, 107, 97, 103, 101]);
  const malformedWebMcp = await page.evaluate(async () => {
    const tool = window.__officeImoWebMcpTools.convert_selected_document;
    const output = await tool.execute({}, { signal: new AbortController().signal });
    return { output, outputCharacters: JSON.stringify(output).length };
  });
  malformedWebMcp.visibleState = await page.locator('[data-bt-output]').getAttribute('data-result-state');
  malformedWebMcp.visibleDiagnostics = await page.locator('.bt-verdict').allTextContents().then(items => items.join(' '));

  return JSON.stringify({
    directoryMilliseconds,
    interactiveMilliseconds,
    startupMilliseconds,
    maximumBrowserHeapBytes,
    maximumEngineMemoryBytes,
    routes: results,
    webMcp,
    webMcpLifecycle: { removedOutsideConverter, restoredWithConverter },
    longNameWebMcp,
    malformedWebMcp,
    consoleErrors
  });
}
