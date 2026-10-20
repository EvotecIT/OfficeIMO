// Real selected grid sources exercise the shared example; file semantics are read independently in .NET.
async function runEditableAnalysisContracts(input) {
  const { moduleBase, exampleUrl, profile, classicScript, executionScript } = input;
  const offline = profile === 'offline-currentview';
  const api = profile === 'datatables' || offline ? globalThis.OfficeIMO : await import(moduleBase + '/index.js');
  const execution = offline ? api : await import(moduleBase + '/execution/index.js');
  const example = offline ? globalThis.serviceAnalysisExample : await import(exampleUrl);
  const sourceApi = offline ? undefined : await import(moduleBase + '/integrations/' + profile + '/index.js');
  const check = (value, message) => { if (!value) throw Error(message); };
  const ids = ['service', 'site', 'status', 'latency', 'availability', 'checked'];
  const headers = example.reportColumns.map(column => column.header), order = [1, 0, 2, 3, 4, 5];
  // Selected rows exceed the example's 128-row sizing sample so awaited output can
  // demonstrably stop a source that still has records, rather than observing a finished source.
  const records = [...example.operationalRows(256)].map((row, index) => index % 4 === 3 ? ['UNSELECTED-' + index, ...row.slice(1)] : row);
  // Formula-looking captured content stays a raw string, independent of the deliberate summary formulas.
  records[1][1] = '=literal Łódź';
  const excluded = records.length - 2; // First loaded accepted record after descending service order.
  const expected = records.map((row, index) => ({ row, index })).filter(item => item.index % 4 !== 3 && item.index !== excluded)
    .sort((a, b) => b.row[0].localeCompare(a.row[0]));
  let table, grid;
  const summaries = rows => Object.fromEntries(['Healthy', 'Attention', 'Offline'].map(status => [status, rows.filter(row => row[2] === status).length]));
  const contract = empty => ({ profile, headers: order.map(index => headers[index]),
    rows: (empty ? [] : expected).map(item => order.map(index => item.row[index] instanceof Date ? item.row[index].toISOString() : item.row[index])),
    counts: summaries(empty ? [] : expected.map(item => item.row)), sourceScope: 'selected filtered records in accepted order',
    summaryScope: 'all exported records; later Excel filters do not change COUNTIF' });
  const tone = value => ({ background: value === 'Offline' ? 'FEE2E2' : value === 'Attention' ? 'FEF3C7' : 'DCFCE7' });
  if (profile === 'datatables') {
    document.body.innerHTML = '<h1>Editable service analysis</h1><table id="analysis-table"></table>';
    table = new DataTable('#analysis-table', { data: records.map(row => [...row, 'INTERNAL-ONLY']),
      columns: [...headers.map(title => ({ title, render: (value, type) => type === 'export' ? value : value instanceof Date ? value.toISOString() : typeof value === 'number' ? 'DISPLAY-ONLY ' + value : value })), { title: 'Internal', visible: false }],
      select: true, colReorder: true, order: [[0, 'desc']], pageLength: 8 });
    table.search('WEB-').draw(); table.rows({ search: 'applied' }).select(); table.row(excluded).deselect();
    table.colReorder.order([1, 0, 2, 3, 4, 5, 6]);
  } else if (profile === 'canopyx') {
    document.body.innerHTML = '<h1>Editable service analysis</h1><div id="analysis-grid" style="height:500px"></div>';
    const view = { protocol: 'canopyx/2', type: 'grid', id: 'analysis-services', label: 'Service inventory', source: { kind: 'inline' },
      rowHeight: 30, pageSize: 7, presentation: 'paged', pageLength: 8, selection: 'multiple', search: true,
      reporting: { selectionControls: true }, columns: [...ids.map((id, index) => ({ id, title: headers[index], sortable: true,
        width: [160, 110, 120, 130, 130, 190][index], kind: index === 5 ? 'datetime' : [3, 4].includes(index) ? 'number' : 'text' })),
        { id: 'internal', title: 'Internal', kind: 'text', hidden: true }],
      rows: records.map((row, index) => ({ id: 'record-' + index, cells: { ...Object.fromEntries(ids.map((id, column) => [id,
        column === 2 ? { value: row[column], tone: row[column] === 'Offline' ? 'danger' : row[column] === 'Attention' ? 'warning' : 'success' }
          : { value: column === 5 ? row[column].toISOString() : row[column], text: 'DISPLAY-ONLY ' + String(row[column]) }])), internal: 'INTERNAL-ONLY' } })) };
    grid = CanopyX.createDataGrid(document.querySelector('#analysis-grid'), { view }); await grid.ready;
    await grid.setQuery({ search: 'WEB-', sort: [{ column: 'service', direction: 'desc' }] });
    grid.setColumnOrder([...order.map(index => ids[index]), 'internal']); grid.selectAllMatching(); grid.setRowSelected('record-' + excluded, false);
  } else {
    check(offline, 'Unknown analysis source profile.');
    document.body.innerHTML = '<h1>Offline current-view service analysis</h1><table id="analysis-currentview"><thead></thead><tbody></tbody></table>';
    const host = document.querySelector('table'), head = host.tHead.insertRow();
    for (const index of [...order, 6]) { const cell = document.createElement('th'); cell.textContent = headers[index] ?? 'Internal'; cell.dataset.source = index; cell.hidden = index === 6; head.append(cell); }
    for (const item of records.map((row, index) => ({ row, index })).sort((a, b) => b.row[0].localeCompare(a.row[0]))) {
      const row = host.tBodies[0].insertRow(); row.dataset.record = item.index;
      row.hidden = item.index % 4 === 3 || item.index === excluded;
      for (const index of [...order, 6]) { const cell = row.insertCell(); cell.textContent = index === 6 ? 'INTERNAL-ONLY' : String(item.row[index]); cell.dataset.source = index; cell.hidden = index === 6; }
    }
  }
  async function capture(signal, empty = false, withoutStatus = false) {
    if (profile === 'datatables') {
      const columnOptions = {};
      table.columns().indexes().toArray().forEach(index => { const original = table.column(index).dataSrc(); if (original < 6) columnOptions[index] = example.reportColumns[original]; });
      const source = sourceApi.createDataTablesExport(DataTable, table, { signal, headings: 'leaf', batchRows: 7, columnOptions,
        exportOptions: { orthogonal: 'export', columns: withoutStatus ? (index => table.column(index).visible() && table.column(index).dataSrc() !== 2) : ':visible', ...(empty ? { rows: [] } : {}) },
        project(value, context) { const original = table.column(context.sourceColumnIndex).dataSrc(); return original === 2 ? new api.ExportCell(value, { presentation: tone(value) }) : value; } });
      return { columns: source.columns, rows: source.rows };
    }
    if (profile === 'canopyx') {
      if (empty) grid.clearSelection();
      const native = grid.prepareExport({ scope: 'selected', values: 'raw', columns: order.filter(index => !withoutStatus || index !== 2).map(index => ids[index]), locale: 'en-US', timeZone: 'utc' });
      check(native.request.values === 'raw' && native.request.recordCount === (empty ? 0 : expected.length), 'Native selected raw scope differs.');
      const source = sourceApi.createCanopyExport(native, 'xlsx', { signal,
        tones: { success: tone('Healthy'), warning: tone('Attention'), danger: tone('Offline') },
        columnOptions: Object.fromEntries(ids.map((id, index) => [id, { width: example.reportColumns[index].width, ...(example.reportColumns[index].format ? { format: example.reportColumns[index].format } : {}) }])) });
      return { columns: source.columns, rows: source.rows() };
    }
    const host = document.querySelector('#analysis-currentview');
    const selected = [...host.tHead.rows[0].cells].filter(cell => !cell.hidden && (!withoutStatus || Number(cell.dataset.source) !== 2)).map(cell => Number(cell.dataset.source));
    return { columns: selected.map(index => example.reportColumns[index]), rows: (async function* () {
      if (empty) return;
      for (const row of host.tBodies[0].rows) if (!row.hidden) yield selected.map(index => records[Number(row.dataset.record)][index]);
    })() };
  }
  async function save(name, blob, empty = false) {
    const bytes = new Uint8Array(await blob.arrayBuffer()); let text = '';
    for (let first = 0; first < bytes.length; first += 8192) text += String.fromCharCode(...bytes.subarray(first, first + 8192));
    await writeEditableAnalysis(name + '.xlsx', btoa(text), JSON.stringify(contract(empty)));
    return bytes.length;
  }
  function worker() {
    const operation = 'async ({input,rows,sink,signal,onProgress}) => example.writeServiceAnalysis(api,{input,columns:input.columns,rows,sink,signal,onProgress})';
    const body = offline ? classicScript + '\n' + executionScript + '\nconst api=OfficeIMO,execution=OfficeIMO; const reportColumns=' + JSON.stringify(example.reportColumns) + ';const example={writeServiceAnalysis:' + example.writeServiceAnalysis.toString() + '};\n'
      : 'import * as api from ' + JSON.stringify(moduleBase + '/index.js') + ';\nimport * as execution from ' + JSON.stringify(moduleBase + '/execution/index.js') + ';\nimport * as example from ' + JSON.stringify(exampleUrl) + ';\n';
    const url = URL.createObjectURL(new Blob([body + 'execution.handleWorkerExports(self,' + operation + ');'], { type: 'text/javascript' }));
    const owned = new Worker(url, offline ? undefined : { type: 'module' }), listeners = new Set(), observed = { maxRows: 0, maxChunk: 0 };
    const endpoint = {
      postMessage(message) { if (message.kind === 'rows') observed.maxRows = Math.max(observed.maxRows, message.rows.length); owned.postMessage(message); },
      addEventListener(type, listener) {
        const wrapped = event => { if (type === 'message' && event.data.kind === 'output') observed.maxChunk = Math.max(observed.maxChunk, event.data.bytes.length); listener(event); };
        listeners.add({ type, listener, wrapped }); owned.addEventListener(type, wrapped);
      },
      removeEventListener(type, listener) { const entry = [...listeners].find(item => item.type === type && item.listener === listener); if (entry) { owned.removeEventListener(type, entry.wrapped); listeners.delete(entry); } }
    };
    return { endpoint, observed, listeners, release() { owned.terminate(); URL.revokeObjectURL(url); } };
  }
  const results = [];
  try {
    for (const delivery of ['blob', 'sink', 'worker']) {
      const source = await capture(), chunks = [], progress = []; let enumerations = 0, produced = 0, returned = 0, writes = 0, earlyWrite = false;
      async function* rows() { enumerations++; try { for await (const row of source.rows) { produced++; yield row; } } finally { returned++; } }
      const destination = { async write(chunk) {
        const stopped = produced; earlyWrite ||= stopped < expected.length;
        await new Promise(resolve => setTimeout(resolve, 1));
        check(produced === stopped, 'Rows advanced while a destination chunk awaited acknowledgement.'); chunks.push(chunk.slice()); writes++;
      } };
      let blob, completed, observed;
      if (delivery === 'blob') blob = await example.writeServiceAnalysis(api, { rows: rows(), columns: source.columns, onProgress: event => progress.push(event.rows) });
      else if (delivery === 'sink') {
        completed = await example.writeServiceAnalysis(api, { rows: rows(), columns: source.columns, sink: destination, onProgress: event => progress.push(event.rows) });
        blob = new Blob(chunks);
      } else {
        const owned = worker();
        try {
          completed = await execution.executeWorkerExport(owned.endpoint, rows(), destination, { input: { columns: source.columns },
            batchRows: 7, maxBatchCells: 42, maxChunkBytes: 4096, onProgress: event => progress.push(event.rows) });
          check(owned.listeners.size === 0 && owned.observed.maxRows <= 7 && owned.observed.maxChunk <= 4096, 'Worker bounds/listener lifetime differ.');
          observed = owned.observed; blob = new Blob(chunks);
        } finally { owned.release(); }
      }
      check(enumerations === 1 && produced === expected.length && returned === 1, 'Selected source was replayed, lost or retained.');
      check(progress.at(-1) === expected.length, 'Summary metadata changed report progress.');
      if (completed) check(completed.rows === expected.length && completed.columns === 6 && completed.bytes === blob.size, 'Sink/worker completion differs.');
      if (delivery !== 'blob') check(writes > 0 && earlyWrite, 'No awaited output occurred before the selected source completed.');
      const bytes = await save(profile + '-' + delivery, blob);
      results.push({ delivery, rows: produced, enumerations, returned, writes, bytes, ...(observed ? { worker: observed } : {}) });
    }
    // A missing analysis field must not be silently restored or consume its captured source.
    const missing = await capture(undefined, false, true); let consumed = false, missingError;
    async function* unused() { consumed = true; yield* missing.rows; }
    try { await example.writeServiceAnalysis(api, { rows: unused(), columns: missing.columns }); } catch (error) { missingError = error; }
    check(missingError instanceof TypeError && /status/.test(missingError.message) && !consumed, 'A hidden required Status field was backfilled or read.');
    const controller = new AbortController(), reason = Error('Selected analysis cancelled'), source = await capture(controller.signal);
    let produced = 0, returned = 0, release;
    const released = new Promise(resolve => { release = resolve; });
    async function* cancelledRows() { try { for await (const row of source.rows) { if (++produced === 20) controller.abort(reason); yield row; } } finally { returned++; release(); } }
    const sink = { write() {} }; let cancelled;
    try { await example.writeServiceAnalysis(api, { rows: cancelledRows(), columns: source.columns, sink, signal: controller.signal }); } catch (error) { cancelled = error; }
    check(cancelled === reason && produced === 20, 'Cancellation lost source bounds or its reason.');
    await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('Cancelled analysis source did not release.')), 10000); released.then(() => { clearTimeout(timer); resolve(); }); });
    check(returned === 1, 'Cancelled analysis source did not return once.');
    const empty = await capture(undefined, true);
    const emptyBlob = await example.writeServiceAnalysis(api, { rows: empty.rows, columns: empty.columns });
    const emptyBytes = await save(profile + '-empty', emptyBlob, true);
    return { profile, results, emptyBytes, cancelledAfterRows: produced, requiredStatusField: true,
      summaryScope: 'exported records; later native Excel filtering does not alter COUNTIF' };
  } finally { table?.destroy(); grid?.destroy(); }
}
