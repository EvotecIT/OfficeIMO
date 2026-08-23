async (page) => {
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
  let maximumBrowserHeapBytes = 0;

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
    const measureConversion = async previousDownloadUrl => {
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

      await page.locator(`[data-convert-route="${routeId}"]`).click();
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
      return { downloadUrl, pdfMagic, memorySamples, peakBrowserHeapBytes, peakBrowserUsedHeapBytes, ...metrics };
    };

    const first = await measureConversion('');
    const repeat = await measureConversion(first.downloadUrl);
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

  return JSON.stringify({ startupMilliseconds, maximumBrowserHeapBytes, routes: results, consoleErrors });
}
