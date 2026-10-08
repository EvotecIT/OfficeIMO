/* Native capture semantics stay with CanopyX; document and sink ownership stays with OfficeIMO. */
function canopyAssert(condition, message) { if (!condition) throw new Error(message); }
async function canopyCollect(capture) { const rows = []; for await (const row of capture.rows()) rows.push(row); return rows; }
function canopyFonts(regular, bold) {
  const bytes = text => Uint8Array.from(atob(text), char => char.charCodeAt(0));
  return { regular: new OfficeIMO.PdfFont(bytes(regular)), bold: new OfficeIMO.PdfFont(bytes(bold)) };
}
function canopyOptions(fonts) {
  return { tones: { neutral: {}, warning: { background: 'FFF2CC', color: '9C6500' }, danger: { background: 'FFC7CE', color: '9C0006', bold: true },
    success: { background: 'E2F0D9' }, info: { background: 'DDEBF7', color: '1F4E78' } },
    columnOptions: { amount: { width: 16, format: '0.00' }, when: { width: 42, wrapText: true, format: 'yyyy-mm-dd hh:mm:ss.000' } },
    pdf: { fonts, fontSize: 9, pageNumbers: false }, xlsx: { dateMode: 'utc' } };
}
async function canopySave(name, capture, format, options, bytes) {
  const rows = await canopyCollect(capture);
  if (!bytes) bytes = new Uint8Array(await (await OfficeIMO.exportCanopy(capture, format, options)).arrayBuffer());
  let binary = ''; for (let i = 0; i < bytes.length; i += 32768) binary += String.fromCharCode(...bytes.subarray(i, i + 32768));
  await writeCanopyFixture(name + '.' + format, btoa(binary), JSON.stringify({ columns: capture.request.columns, values: capture.request.values, rows }));
}
function canopyInlineView() {
  return { protocol: 'canopyx/2', type: 'grid', id: 'officeimo-records', label: 'Inventory', source: { kind: 'inline' },
    rowHeight: 30, pageSize: 7, presentation: 'paged', pageLength: 5, selection: 'multiple', search: true,
    reporting: { selectionControls: true }, columns: [
      { id: 'name', title: 'Record', kind: 'text', width: 180, wrap: true, sortable: true, presenter: { kind: 'link', hrefColumn: 'url' } },
      { id: 'url', title: 'Address', kind: 'text', width: 180, hidden: true },
      { id: 'amount', title: 'Amount', kind: 'number', width: 100, sortable: true, presenter: { kind: 'progress', maximum: 100 } },
      { id: 'when', title: 'Updated', kind: 'datetime', width: 220 },
      { id: 'flag', title: 'Enabled', kind: 'text', width: 100, presenter: { kind: 'boolean' } }
    ], highlightRules: [ { column: 'amount', operator: 'gt', values: [10], tone: 'warning' },
      { column: 'amount', operator: 'gt', values: [10], target: 'cell', tone: 'info' } ],
    rows: Array.from({ length: 24 }, (_, i) => ({ id: 'r' + i, tone: i === 0 ? 'success' : undefined, cells: {
      name: { value: (i % 2 ? 'Archive ' : 'Current ') + i, text: (i === 1 ? '=literal Archive ' : (i % 2 ? 'Archive Łódź ' : 'Current Łódź ')) + i, title: 'Record _x0041_ 🧪' },
      url: 'https://example.com/report/' + i + '#details', amount: { value: i + .5, text: (i + .5).toFixed(2) + ' USD' },
      when: i === 2 ? '2026-10-08T12:34:56.1234567+02:00' : i === 3 ? '2026-10-08T12:34Z' : i === 4 ? '2026-10-08T12:34:56.123456789Z' : '2026-10-08T12:34:56.123Z', flag: i % 2 === 0
    } })) };
}
async function canopyMount(view, options = {}) {
  const container = document.querySelector('#fixture'); container.replaceChildren();
  const host = document.createElement('div'); host.className = 'cx-host'; host.setAttribute('data-canopyx-host', ''); host.setAttribute('data-cx-theme', 'light'); container.append(host);
  const api = CanopyX.createDataGrid(host, { ...options, view }); await api.ready; return api;
}

// A bounded host-owned row/byte bridge. The live native grid and source remain on the page.
async function canopyWorkerExport(capture, format, options, workerScript, regular, bold) {
  const source = workerScript + '\n' + canopyFonts.toString() + '\n' + `
let sequence=0, controller; const pending=new Map();
function ask(kind, fields={}) { const id=++sequence; return new Promise((resolve,reject)=>{pending.set(id,{resolve,reject});postMessage({kind,id,...fields});}); }
onmessage=async ({data})=>{
  if(data.reply) { const entry=pending.get(data.reply); if(entry){pending.delete(data.reply);data.error?entry.reject(new Error(data.error)):entry.resolve(data);} return; }
  if(data.cancel) { controller?.abort(); return; }
  controller=new AbortController();
  try {
    const capture={request:data.request, async *rows(){while(true){const page=await ask('rows'); for(const row of page.rows)yield row;if(page.done)return;}}};
    const options={...data.options,signal:controller.signal,pdf:{...data.options.pdf,fonts:canopyFonts(data.regular,data.bold)}};
    const result=await OfficeIMO.writeCanopyTo(capture,data.format,{async write(bytes){await ask('bytes',{bytes:bytes.slice()});}},options);
    postMessage({done:true,result});
  }catch(error){postMessage({error:error.stack||String(error)});}
};`;
  const url = URL.createObjectURL(new Blob([source], { type: 'text/javascript' })), worker = new Worker(url), chunks = [];
  const iterator = capture.rows()[Symbol.asyncIterator](); let maxBatch = 0, result;
  try {
    result = await new Promise((resolve, reject) => {
      worker.onerror = event => reject(new Error(event.message));
      worker.onmessage = async ({ data }) => {
        if (data.error) { reject(new Error(data.error)); return; }
        if (data.done) { resolve(data.result); return; }
        try {
          if (data.kind === 'bytes') { chunks.push(data.bytes); worker.postMessage({ reply: data.id }); }
          else {
            const rows = []; let done = false;
            while (rows.length < 4) { const next = await iterator.next(); if (next.done) { done = true; break; } rows.push(next.value); }
            maxBatch = Math.max(maxBatch, rows.length); worker.postMessage({ reply: data.id, rows, done });
          }
        } catch (error) { worker.postMessage({ reply: data.id, error: String(error) }); }
      };
      worker.postMessage({ request: capture.request, format, options: { ...options, pdf: { ...options.pdf, fonts: undefined } }, regular, bold });
    });
    const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0)); let offset = 0;
    for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
    canopyAssert(maxBatch <= 4 && result.rows === capture.request.recordCount, 'Worker source bridge lost its bound or row count.');
    return bytes;
  } finally { worker.terminate(); URL.revokeObjectURL(url); await iterator.return?.(); }
}

globalThis.runCanopyContracts = async function ({ regular, bold, workerScript }) {
  const fonts = canopyFonts(regular, bold), options = canopyOptions(fonts), formats = ['csv', 'xlsx', 'pdf'];
  let grid = await canopyMount(canopyInlineView());
  const all = grid.prepareExport({ scope: 'all', values: 'raw', locale: 'en-US', timeZone: 'utc' });
  const records = await canopyCollect(all);
  canopyAssert(all.request.presentation === 'semantic' && all.request.recordCount === 24 && records.length === 24, 'Native semantic record capture differs.');
  canopyAssert(records[11].tone === 'warning' && records[11].cells.amount.tone === 'info', 'Native highlight precedence differs.');
  const fineDate = records.find(row => row.id === 'r2');
  canopyAssert(fineDate?.cells.when.value === '2026-10-08T12:34:56.1234567+02:00', 'Native timestamp precision was lost: ' + JSON.stringify(fineDate));
  await captureCanopyScreen('records');
  for (const mode of ['classic', 'fallback', 'worker']) for (const format of formats) {
    const compression = globalThis.CompressionStream;
    try {
      if (mode === 'fallback') globalThis.CompressionStream = undefined;
      const bytes = mode === 'worker' ? await canopyWorkerExport(all, format, options, workerScript, regular, bold) : undefined;
      await canopySave(mode + '-all', all, format, options, bytes);
    } finally { globalThis.CompressionStream = compression; }
  }
  grid.setColumnOrder(['amount', 'name', 'when', 'flag', 'url']); grid.setColumnVisible('flag', false);
  await grid.setQuery({ search: 'Archive', sort: [{ column: 'amount', direction: 'desc' }] });
  const filtered = grid.prepareExport({ scope: 'filtered', values: 'raw', locale: 'en-US', timeZone: 'utc' });
  grid.selectAllMatching(); grid.setRowSelected('r23', false);
  const selected = grid.prepareExport({ scope: 'selected', values: 'display', columns: ['name', 'amount', 'when'], locale: 'en-US', timeZone: 'utc' });
  await grid.setQuery({ search: 'Current' }); grid.clearSelection(); grid.setColumnOrder(['name', 'amount', 'when', 'flag', 'url']);
  canopyAssert(filtered.request.columns.map(c => c.id).join() === 'amount,name,when', 'Native visible column capture changed.');
  canopyAssert((await canopyCollect(filtered)).map(r => r.id).join() === Array.from({ length: 12 }, (_, i) => 'r' + (23 - i * 2)).join(), 'Captured filter/order changed with the UI.');
  canopyAssert(selected.request.recordCount === 11 && !(await canopyCollect(selected)).some(r => r.id === 'r23'), 'Selected-query exclusion was lost.');
  for (const format of formats) { await canopySave('filtered', filtered, format, options); await canopySave('selected-display', selected, format, options); }
  grid.destroy();

  const cursorView = canopyInlineView(); cursorView.id = 'officeimo-cursor'; cursorView.source = { kind: 'host', navigation: 'cursor' }; delete cursorView.rows;
  let held = false, changed = false, sourceSignal;
  const source = { navigation: 'cursor', async query(query, { signal }) {
    sourceSignal = signal; const offset = query.cursor ? Number(query.cursor) : 0;
    if (held && offset) await new Promise((_, reject) => signal.addEventListener('abort', () => reject(signal.reason), { once: true }));
    const rows = canopyInlineView().rows.slice(offset, offset + query.limit);
    return { protocol: 'canopyx/2', type: 'gridCursorPage', view: query.view, revision: changed ? 'r2' : 'r1', offset, total: null, matched: null, items: rows,
      ...(offset + rows.length < 24 ? { nextCursor: String(offset + rows.length) } : {}) };
  } };
  const delivered = [], hostProgress = [];
  grid = await canopyMount(cursorView, { source, async onExport(event) {
    canopyAssert(event.capture.request === event.request, 'Host export did not receive the original capture.');
    const progress = [], chunks = [], result = await OfficeIMO.writeCanopyTo(event.capture, event.format, { write(bytes) { chunks.push(bytes.slice()); } }, {
      ...options, signal: event.signal, onProgress(progressEvent) {
        progress.push(progressEvent.rows); event.reportProgress?.('generating', progressEvent.rows);
      }
    });
    canopyAssert(progress.at(-1) === result.rows && progress.every((count, index) => !index || count >= progress[index - 1]), 'Native host progress lost its completion or monotonic count.');
    hostProgress.push(progress.length);
    const bytes = new Uint8Array(chunks.reduce((n, c) => n + c.length, 0)); let offset = 0; for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.length; }
    delivered.push(result); await canopySave('host-' + event.format, event.capture, event.format, options, bytes);
  } });
  const cursor = grid.prepareExport({ scope: 'all', values: 'raw', locale: 'en-US', timeZone: 'utc' });
  canopyAssert(cursor.request.recordCount === null && (await canopyCollect(cursor)).length === 24, 'Unknown cursor counts were invented or truncated.');
  for (const format of formats) { await canopySave('cursor', cursor, format, options); await grid.exportToHost(format, { scope: 'all', values: 'raw', locale: 'en-US', timeZone: 'utc' }); }
  held = true; const controller = new AbortController();
  const cancelled = OfficeIMO.exportCanopy(cursor, 'xlsx', { ...options, signal: controller.signal });
  setTimeout(() => controller.abort(), 20);
  const error = await cancelled.then(() => null, e => e);
  canopyAssert(error?.name === 'AbortError' && sourceSignal.aborted, 'Writer cancellation did not reach native page I/O.');
  held = false; changed = true;
  const revisionError = await OfficeIMO.exportCanopy(cursor, 'csv', options).then(() => null, e => e);
  canopyAssert(revisionError && String(revisionError).includes('snapshot changed'), 'Native revision mismatch was ignored.');
  changed = false; held = true;
  const closing = OfficeIMO.exportCanopy(cursor, 'xlsx', options);
  setTimeout(() => grid.destroy(), 20);
  const closeError = await closing.then(() => null, error => error);
  canopyAssert(closeError?.name === 'AbortError' && sourceSignal.aborted, 'Native grid disposal did not cancel in-flight export I/O.');
  held = false;

  const diagnosticView = canopyInlineView(); diagnosticView.rows = [{ id: 'custom', cells: { name: { value: 'Raw name', text: 'Captured text', className: 'custom' },
    url: '../offline-report.html', amount: 1, when: null, flag: false } }];
  grid = await canopyMount(diagnosticView, { renderers: { name(row, element) { element.textContent = 'DOM-only renderer'; return true; } },
    async onExport(event) {
      await canopySave('host-button', event.capture, event.format, { ...options, unsupportedPresentation: 'text', signal: event.signal });
    } });
  const diagnostic = grid.prepareExport({ scope: 'all', values: 'raw', locale: 'en-US', timeZone: 'utc' });
  const strictError = await OfficeIMO.exportCanopy(diagnostic, 'xlsx', options).then(() => null, e => e);
  canopyAssert(strictError && String(strictError).includes('custom-renderer'), 'Custom renderer was silently treated as portable.');
  const notices = [];
  for (const format of formats) await canopySave('text-fallback', diagnostic, format, { ...options, unsupportedPresentation: 'text', onDiagnostic: notice => notices.push(notice.code) });
  canopyAssert(notices.includes('custom-renderer') && notices.includes('relative-link'), 'Text fallback lost its diagnostics.');
  const button = document.createElement('button'); button.id = 'canopy-host-export'; button.textContent = 'Export Excel with OfficeIMO';
  button.addEventListener('click', () => grid.exportToHost('xlsx', { scope: 'all', values: 'raw', locale: 'en-US', timeZone: 'utc' })
    .then(() => { globalThis.canopyHostButtonComplete = true; }, error => { globalThis.canopyHostButtonError = String(error); }));
  document.querySelector('#fixture').prepend(button);
  return { files: 24, modes: ['classic', 'fallback', 'worker'], nativeCapture: 'semantic', cancellation: true, revisionMismatch: true,
    selectedQuery: true, unknownCursorCount: true, gridDisposal: true, hostCapture: delivered.length, hostProgress, diagnosticCodes: [...new Set(notices)] };
};

globalThis.runOrdinaryCanopyContracts = async function ({ regular, bold }) {
  const view = canopyInlineView(); delete view.reporting; delete view.highlightRules;
  for (const column of view.columns) delete column.presenter;
  const grid = await canopyMount(view);
  const capture = grid.prepareExport({ scope: 'all', values: 'raw' });
  canopyAssert(capture.request.presentation === 'text', 'Ordinary grid falsely claims semantic capture.');
  const options = canopyOptions(canopyFonts(regular, bold));
  const error = await OfficeIMO.exportCanopy(capture, 'xlsx', options).then(() => null, error => error);
  canopyAssert(String(error).includes('TEXT_ONLY_CAPTURE'), 'Ordinary capture silently claims presentation fidelity.');
  for (const format of ['csv', 'xlsx', 'pdf']) await canopySave('ordinary', capture, format, { ...options, unsupportedPresentation: 'text' });
  grid.destroy(); return { ordinaryTextCapture: true, explicitVisualFallback: true };
};
