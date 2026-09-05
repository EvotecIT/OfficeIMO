async (page) => {
  await page.addInitScript(() => {
    const tools = Object.create(null);
    Object.defineProperty(window, '__officeImoWebMcpTools', { value: tools, configurable: true });
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
  await page.reload({ waitUntil: 'domcontentloaded' });
  const routeIds = ['docx-pdf', 'xlsx-pdf', 'pptx-pdf'];
  const consoleErrors = [];
  page.on('console', message => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });
  page.on('pageerror', error => consoleErrors.push(error.message));

  const baseUrl = page.url().split('?')[0];
  const results = [];
  await page.locator('[data-converter-ready="true"]').waitFor({ state: 'visible', timeout: 60000 });
  const startupMilliseconds = await page.evaluate(() => performance.now());
  const initialResourceErrors = await page.evaluate(() => performance.getEntriesByType('resource')
    .filter(entry => Number.isFinite(entry.responseStatus) && entry.responseStatus >= 400)
    .map(entry => `${entry.responseStatus} ${entry.name}`));
  consoleErrors.push(...initialResourceErrors);
  if (await page.locator('#blazor-error-ui').isVisible()) {
    consoleErrors.push('Blazor error UI became visible during startup.');
  }
  await page.waitForFunction(() => Boolean(window.__officeImoWebMcpTools?.convert_selected_document), null, { timeout: 60000 });
  await page.getByRole('button', { name: 'PDF tools', exact: true }).click();
  await page.waitForFunction(() => !window.__officeImoWebMcpTools?.convert_selected_document, null, { timeout: 60000 });
  const removedOutsideConverter = await page.evaluate(() => !window.__officeImoWebMcpTools?.convert_selected_document);
  await page.getByRole('button', { name: 'Convert', exact: true }).click();
  await page.waitForFunction(() => Boolean(window.__officeImoWebMcpTools?.convert_selected_document), null, { timeout: 60000 });
  const restoredWithConverter = await page.evaluate(() => Boolean(window.__officeImoWebMcpTools?.convert_selected_document));
  let maximumBrowserHeapBytes = 0;
  let webMcp = null;

  const readBrowserHeap = async () => page.evaluate(() => {
    const memory = performance.memory;
    if (!memory || !Number.isFinite(memory.usedJSHeapSize) || !Number.isFinite(memory.totalJSHeapSize)) return null;
    if (memory.usedJSHeapSize <= 0 || memory.totalJSHeapSize <= 0) return null;
    return { used: memory.usedJSHeapSize, total: memory.totalJSHeapSize };
  });

  for (let index = 0; index < routeIds.length; index++) {
    const routeId = routeIds[index];
    if (index > 0) {
      await page.goto(`${baseUrl}?route=${encodeURIComponent(routeId)}`, { waitUntil: 'domcontentloaded' });
    }
    await page.locator('[data-converter-ready="true"]').waitFor({ state: 'visible', timeout: 60000 });
    await page.locator(`[data-active-route="${routeId}"]`).waitFor({ state: 'attached', timeout: 60000 });

    await page.locator(`[data-load-sample="${routeId}"]`).click();
    await page.locator('.ocx-diagnostic').filter({ hasText: 'Sample ready' }).waitFor({ state: 'visible', timeout: 60000 });
    const summary = page.locator(`[data-performance-result="true"][data-route="${routeId}"]`);
    const downloadLink = page.getByRole('link', { name: 'Download result', exact: true });
    const measureConversion = async (previousDownloadUrl, useWebMcp) => {
      let sampling = true;
      let memorySamples = 0;
      let peakBrowserHeapBytes = 0;
      let peakBrowserUsedHeapBytes = 0;
      const sampleHeap = async () => {
        const memory = await readBrowserHeap();
        if (!memory) throw new Error('Chromium performance.memory is unavailable; peak-memory evidence is required.');
        memorySamples++;
        peakBrowserHeapBytes = Math.max(peakBrowserHeapBytes, memory.total);
        peakBrowserUsedHeapBytes = Math.max(peakBrowserUsedHeapBytes, memory.used);
      };
      await sampleHeap();
      const sampler = (async () => {
        while (sampling) {
          await page.waitForTimeout(10);
          await sampleHeap();
        }
      })();

      let webMcpOutput = null;
      if (useWebMcp) {
        await page.waitForFunction(() => Boolean(window.__officeImoWebMcpTools?.convert_selected_document), null, { timeout: 60000 });
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
        await page.locator(`[data-convert-route="${routeId}"]`).click();
      }
      await page.waitForFunction(previousUrl => {
        const link = Array.from(document.querySelectorAll('a'))
          .find(element => element.textContent?.trim() === 'Download result');
        return Boolean(link?.href?.startsWith('blob:') && link.href !== previousUrl);
      }, previousDownloadUrl, { timeout: 120000 });
      await summary.waitFor({ state: 'visible', timeout: 120000 });
      sampling = false;
      await sampler;
      await sampleHeap();

      const downloadUrl = await downloadLink.getAttribute('href');
      const pdfMagic = await page.evaluate(async url => {
        const bytes = new Uint8Array(await (await fetch(url)).arrayBuffer());
        return String.fromCharCode(...bytes.slice(0, 4));
      }, downloadUrl);
      const metrics = await summary.evaluate(element => ({
        conversionMilliseconds: Number(element.getAttribute('data-conversion-ms') || '0'),
        peakRetainedBytes: Number(element.getAttribute('data-peak-retained-bytes') || '0'),
        resultBytes: Number(element.getAttribute('data-result-bytes') || '0')
      }));
      return { downloadUrl, pdfMagic, memorySamples, peakBrowserHeapBytes, peakBrowserUsedHeapBytes, webMcpOutput, ...metrics };
    };

    const first = await measureConversion('', index === 0);
    const repeat = await measureConversion(first.downloadUrl, false);
    if (first.webMcpOutput) webMcp = first.webMcpOutput;
    maximumBrowserHeapBytes = Math.max(
      maximumBrowserHeapBytes,
      first.peakBrowserHeapBytes,
      repeat.peakBrowserHeapBytes);
    results.push({
      routeId,
      memorySamples: first.memorySamples,
      peakBrowserHeapBytes: first.peakBrowserHeapBytes,
      peakBrowserUsedHeapBytes: first.peakBrowserUsedHeapBytes,
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

  await page.goto(`${baseUrl}?route=docx-pdf`, { waitUntil: 'domcontentloaded' });
  await page.locator('[data-converter-ready="true"]').waitFor({ state: 'visible', timeout: 60000 });
  await page.waitForFunction(() => Boolean(window.__officeImoWebMcpTools?.convert_selected_document), null, { timeout: 60000 });
  await page.getByLabel('Choose a DOCX file', { exact: true }).evaluate(async input => {
    const bytes = new Uint8Array(await (await fetch('samples/basic.docx')).arrayBuffer());
    const file = new File(
      [bytes],
      `${'a'.repeat(179)}🚀.docx`,
      { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
    );
    const transfer = new DataTransfer();
    transfer.items.add(file);
    input.files = transfer.files;
    input.dispatchEvent(new Event('change', { bubbles: true }));
  });
  await page.locator('.ocx-diagnostic').filter({ hasText: 'is loaded in this browser tab' })
    .waitFor({ state: 'visible', timeout: 60000 });
  const longNameWebMcp = await page.evaluate(async () => {
    const tool = window.__officeImoWebMcpTools.convert_selected_document;
    const output = await tool.execute({}, { signal: new AbortController().signal });
    const fileName = String(output.outputFileName || '');
    const hasUnpairedSurrogate = /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(fileName);
    return {
      output,
      outputCharacters: JSON.stringify(output).length,
      outputFileNameCharacters: fileName.length,
      hasUnpairedSurrogate
    };
  });
  await page.getByLabel('Choose a DOCX file', { exact: true }).evaluate(input => {
    const bytes = new Uint8Array([
      110, 111, 116, 45, 97, 110, 45, 111, 112, 101, 110,
      45, 120, 109, 108, 45, 112, 97, 99, 107, 97, 103, 101
    ]);
    const file = new File(
      [bytes],
      `${'malformed-'.repeat(18)}document.docx`,
      { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }
    );
    const transfer = new DataTransfer();
    transfer.items.add(file);
    input.files = transfer.files;
    input.dispatchEvent(new Event('change', { bubbles: true }));
  });
  await page.locator('.ocx-diagnostic').filter({ hasText: 'is loaded in this browser tab' })
    .waitFor({ state: 'visible', timeout: 60000 });
  const malformedWebMcp = await page.evaluate(async () => {
    const tool = window.__officeImoWebMcpTools.convert_selected_document;
    const output = await tool.execute({}, { signal: new AbortController().signal });
    return {
      output,
      outputCharacters: JSON.stringify(output).length,
      visibleDiagnostics: Array.from(document.querySelectorAll('.ocx-diagnostic')).map(element => element.textContent || '').join(' ')
    };
  });

  return JSON.stringify({
    startupMilliseconds,
    maximumBrowserHeapBytes,
    routes: results,
    webMcp,
    webMcpLifecycle: { removedOutsideConverter, restoredWithConverter },
    longNameWebMcp,
    malformedWebMcp,
    consoleErrors
  });
}
