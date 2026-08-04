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
    const summary = page.locator(`[data-performance-result="true"][data-route="${routeId}"]`);
    await summary.waitFor({ state: 'visible', timeout: 120000 });
    sampling = false;
    await sampler;
    await sampleHeap();

    const metrics = await summary.evaluate(element => ({
      conversionMilliseconds: Number(element.getAttribute('data-conversion-ms') || '0'),
      peakRetainedBytes: Number(element.getAttribute('data-peak-retained-bytes') || '0'),
      resultBytes: Number(element.getAttribute('data-result-bytes') || '0')
    }));
    maximumBrowserHeapBytes = Math.max(maximumBrowserHeapBytes, peakBrowserHeapBytes);
    results.push({ routeId, memorySamples, peakBrowserHeapBytes, peakBrowserUsedHeapBytes, ...metrics });
  }

  return JSON.stringify({ startupMilliseconds, maximumBrowserHeapBytes, routes: results, consoleErrors });
}
