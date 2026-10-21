// Expected values come from the selected host records, without using the grouping or document engines.
function hierarchicalExpectedContract(records, columns, groupKeys, text, language = 'en', operations = { latency: 'average', availability: 'average' }) {
  const keys = columns.map(column => column.key);
  const totals = rows => keys.map((key, column) => {
    if (!operations[key]) return null;
    const numbers = rows.map(row => row[column]).filter(value => typeof value === 'number' && Number.isFinite(value));
    if (operations[key] === 'count') return numbers.length;
    if (operations[key] === 'sum') return numbers.reduce((sum, value) => sum + value, 0);
    return !numbers.length ? null : operations[key] === 'min' ? Math.min(...numbers) : operations[key] === 'max' ? Math.max(...numbers)
      : numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  });
  const rows = records.map(row => row.slice(0, columns.length)), events = [], active = [];
  function close() {
    const group = active.pop();
    events.push({ kind: 'groupEnd', level: group.level, label: group.label, path: group.path, rows: group.rows.length,
      totals: totals(group.rows), summaryLabel: text.subtotal + ' · ' + group.label + ' · ' + group.rows.length.toLocaleString(language) });
  }
  for (let index = 0; index < records.length; index++) {
    const groups = records[index].slice(columns.length), common = groups.findIndex((key, level) => !active[level] || active[level].key !== key);
    const first = common < 0 ? groups.length : common;
    while (active.length > first) close();
    for (let level = first; level < groups.length; level++) {
      const label = String(groups[level] ?? ''), path = [...active.map(group => group.label), label];
      active.push({ key: groups[level], label, level, path, rows: [] }); events.push({ kind: 'groupStart', level, label, path });
    }
    for (const group of active) group.rows.push(rows[index]);
    events.push({ kind: 'row', index });
  }
  while (active.length) close();
  return { language, keys, headers: columns.map(column => column.header), groupKeys, operations, rows, events, totals: totals(rows),
    grandLabel: text.total + ' · ' + rows.length.toLocaleString(language), text };
}
globalThis.hierarchicalExpectedContract = hierarchicalExpectedContract;

async function runHierarchicalReportContracts({ moduleBase, exampleUrl, inventoryUrl, profile, stack, regular, bold, chart, chartSummary }) {
  const api = profile === 'datatables' ? globalThis.OfficeIMO : await import(moduleBase + '/index.js');
  const example = await import(exampleUrl), inventory = await import(inventoryUrl), execution = await import(moduleBase + '/execution/index.js');
  const sourceApi = await import(moduleBase + '/integrations/' + profile + '/index.js');
  const check = (condition, message) => { if (!condition) throw Error(message); };
  const bytes = base64 => Uint8Array.from(atob(base64), value => value.charCodeAt(0));
  const ids = ['service', 'site', 'status', 'latency', 'availability', 'checked'], groupKeys = ['portfolio', 'environment', 'owner'];
  const titles = ['Endpoint custom', 'Tenant location', 'Health value', 'Response ms', 'Uptime ratio', 'Observed UTC'];
  const allKeys = [...ids, ...groupKeys], order = [1, 0, 2, 3, 4, 5], excluded = 126;
  const columns = inventory.reportColumns.map((column, index) => ({ ...column, header: titles[index] }));
  const records = [...inventory.operationalRows(128)].map((row, index) => [index % 4 === 3 ? 'UNSELECTED-' + index : row[0], row[1], row[2],
    index % 17 === 0 ? null : row[3], index % 23 === 0 ? null : row[4], row[5],
    index < 100 ? 'Portfolio A' : 'Portfolio B', index < 64 ? 'Production' : 'Staging', index < 48 || index >= 112 ? 'Owner 0' : 'Owner 1', 'INTERNAL-ONLY']);
  const selected = records.map((row, index) => ({ row, index })).filter(item => item.index % 4 !== 3 && item.index !== excluded);
  const compare = (left, right) => {
    for (const column of [6, 7, 8]) { if (left.row[column] < right.row[column]) return -1; if (left.row[column] > right.row[column]) return 1; }
    return left.row[0] < right.row[0] ? 1 : left.row[0] > right.row[0] ? -1 : 0;
  };
  selected.sort(compare);
  const status = { total: selected.length, Healthy: 0, Attention: 0, Offline: 0 };
  selected.forEach(item => status[item.row[2]]++);
  check(Object.keys(status).every(key => status[key] === chartSummary[key]), 'Canonical ChartForgeX figure does not describe selected records.');
  const visibleColumns = order.map(index => columns[index]), portable = selected.map(item => [...order.map(index => item.row[index]), ...item.row.slice(6, 9)]);
  const text = { ...example.analyticalReportText.en, scope: 'Selected filtered inventory; explicitly captured hidden dimensions; UTC values.' };
  const contract = hierarchicalExpectedContract(portable, visibleColumns, groupKeys, text, 'en', { latency: 'average', availability: 'count' });
  contract.figureCaption = 'Figure 1. Status of the 95 selected services: 84 healthy, 8 requiring attention and 3 offline.';
  const display = (value, column) => value === null ? '' : value instanceof Date ? value.toISOString().slice(0, 16).replace('T', ' ')
    : typeof value === 'number' ? new Intl.NumberFormat('en', { maximumFractionDigits: 3 }).format(value) : String(value);
  let table, grid;
  if (profile === 'datatables') {
    document.body.innerHTML = '<h1>Selected hierarchical inventory</h1><table id="analytical-table"></table>';
    table = new DataTable('#analytical-table', { data: records, select: true, colReorder: true, pageLength: 8,
      order: [[6, 'asc'], [7, 'asc'], [8, 'asc'], [0, 'desc']],
      columns: [...titles.map(title => ({ title, render: (value, type) => type === 'export' || type === 'sort' ? value : display(value) })),
        ...groupKeys.map(title => ({ title, visible: false })), { title: 'Internal', visible: false }] });
    table.search('WEB-').draw(); table.rows({ search: 'applied' }).select(); table.row(excluded).deselect();
    table.colReorder.order([1, 0, 2, 3, 4, 5, 6, 7, 8, 9]);
  } else {
    document.body.innerHTML = '<h1>Selected hierarchical inventory</h1><div id="analytical-grid" style="height:500px"></div>';
    const view = { protocol: 'canopyx/2', type: 'grid', id: 'analytical-inventory', label: 'Inventory', source: { kind: 'inline' },
      rowHeight: 30, pageSize: 7, presentation: 'paged', pageLength: 8, selection: 'multiple', search: true, reporting: { selectionControls: true },
      columns: [...allKeys.map((id, index) => ({ id, title: titles[index] ?? id, sortable: true, hidden: index >= 6,
        kind: index === 5 ? 'datetime' : index === 3 || index === 4 ? 'number' : 'text', width: 150 })), { id: 'internal', title: 'Internal', kind: 'text', hidden: true }],
      rows: records.map((row, index) => ({ id: 'record-' + index, cells: Object.fromEntries([...allKeys, 'internal'].map((id, column) => [id,
        { value: column === 5 ? row[column].toISOString() : row[column], text: display(row[column], column) }])) })) };
    grid = CanopyX.createDataGrid(document.querySelector('#analytical-grid'), { view }); await grid.ready;
    await grid.setQuery({ search: 'WEB-', sort: [...groupKeys.map(column => ({ column, direction: 'asc' })), { column: 'service', direction: 'desc' }] });
    grid.setColumnOrder([...order.map(index => ids[index]), ...groupKeys, 'internal']);
    grid.selectAllMatching();
    const excludedPosition = records.map((row, index) => ({ row, index })).filter(item => item.index % 4 !== 3).sort(compare).findIndex(item => item.index === excluded);
    // Native row selection is scoped to loaded accepted records, so visit its page before excluding it.
    await grid.goToPage(Math.floor(excludedPosition / 8));
    grid.setRowSelected('record-' + excluded, false);
  }
  async function capture(signal, format) {
    if (profile === 'datatables') {
      const columnOptions = {}, selectedColumns = [];
      table.columns().indexes().toArray().forEach(index => {
        const original = table.column(index).dataSrc();
        if (original < 9) { selectedColumns.push(index); columnOptions[index] = columns[original] ?? { key: allKeys[original], header: allKeys[original] }; }
      });
      const source = sourceApi.createDataTablesExport(DataTable, table, { signal, headings: 'leaf', batchRows: 7, maxBatchCells: 63,
        columnOptions, exportOptions: { orthogonal: 'export', columns: selectedColumns } });
      check(source.columns.map(column => column.key).join('|') === [...order.map(index => ids[index]), ...groupKeys].join('|'), 'DataTables hidden intake/current visible order differs.');
      return { columns: source.columns.slice(0, 6), rows: source.rows };
    }
    // Hidden dimensions are an explicit part of this immutable native record capture, never visible output columns.
    const native = grid.prepareExport({ scope: 'selected', values: 'raw', columns: [...order.map(index => ids[index]), ...groupKeys], locale: 'en-US', timeZone: 'utc' });
    check(native.request.recordCount === selected.length, 'Canopy selected scope differs.');
    check(native.request.columns.map(column => column.id).join('|') === [...order.map(index => ids[index]), ...groupKeys].join('|'), 'Canopy explicit hidden intake differs.');
    const source = sourceApi.createCanopyExport(native, format, { signal, columnOptions: Object.fromEntries(ids.map((id, index) => [id, { width: columns[index].width }])) });
    return { columns: source.columns.slice(0, 6), rows: { async *[Symbol.asyncIterator]() {
      for await (const row of source.rows()) yield row.map((value, index) => index < 6 ? value : value instanceof api.ExportCell ? value.value : value);
    } } };
  }
  const inputFor = format => ({ format, groupKeys, totals: contract.operations, compression: 'store', scope: text.scope,
    ...(format === 'pdf' ? { fonts: { regular: bytes(regular), bold: bytes(bold) }, figure: { image: { format: 'png', data: bytes(chart) }, width: 420,
      caption: contract.figureCaption, alternativeText: 'Selected service status: 84 healthy, 8 requiring attention and 3 offline; 95 services in total.' } } : {}) });
  function worker() {
    const body = 'import * as api from ' + JSON.stringify(moduleBase + '/index.js') + ';\nimport * as execution from ' + JSON.stringify(moduleBase + '/execution/index.js') + ';\nimport * as example from ' + JSON.stringify(exampleUrl) + ';\n' +
      'execution.handleWorkerExports(self,({input,rows,sink,signal,onProgress})=>example.writeAnalyticalReport(api,{input,columns:input.columns,rows,sink,signal,onProgress}));';
    const url = URL.createObjectURL(new Blob([body], { type: 'text/javascript' })), owned = new Worker(url, { type: 'module' }), listeners = new Set(), observed = { maxRows: 0, maxChunk: 0 };
    return { observed, listeners, endpoint: {
      postMessage(message) { if (message.kind === 'rows') { observed.maxRows = Math.max(observed.maxRows, message.rows.length); check(message.rows.every(row => row.length === 9), 'Worker metadata was not a bounded flat nine-cell row.'); } owned.postMessage(message); },
      addEventListener(type, listener) { const wrapped = event => { if (type === 'message' && event.data.kind === 'output') observed.maxChunk = Math.max(observed.maxChunk, event.data.bytes.length); listener(event); }; listeners.add({ type, listener, wrapped }); owned.addEventListener(type, wrapped); },
      removeEventListener(type, listener) { const entry = [...listeners].find(item => item.type === type && item.listener === listener); if (entry) { owned.removeEventListener(type, entry.wrapped); listeners.delete(entry); } },
    }, release() { owned.terminate(); URL.revokeObjectURL(url); } };
  }
  const results = [];
  try {
    for (const format of ['xlsx', 'pdf']) for (const delivery of stack === 'current' ? ['blob', 'sink', 'worker'] : ['blob', 'sink']) {
      const source = await capture(undefined, format), chunks = [], progress = []; let produced = 0, returned = 0, enumerations = 0, writes = 0, earlyWrite = false;
      async function* rows() { enumerations++; try { for await (const row of source.rows) { produced++; yield row; } } finally { returned++; } }
      const destination = { async write(chunk) { const before = produced; earlyWrite ||= before < selected.length; await new Promise(resolve => setTimeout(resolve, 1));
        check(before === produced, 'Analytical rows advanced during destination acknowledgement.'); chunks.push(chunk.slice()); writes++; } };
      let blob, completion, observed;
      if (delivery === 'blob') blob = await example.writeAnalyticalReport(api, { input: inputFor(format), rows: rows(), columns: source.columns, onProgress: event => progress.push(event.rows) });
      else if (delivery === 'sink') { completion = await example.writeAnalyticalReport(api, { input: inputFor(format), rows: rows(), columns: source.columns, sink: destination, onProgress: event => progress.push(event.rows) }); blob = new Blob(chunks); }
      else {
        const owned = worker();
        try {
          completion = await execution.executeWorkerExport(owned.endpoint, rows(), destination, { input: { ...inputFor(format), columns: source.columns }, batchRows: 7,
            maxBatchCells: 63, maxChunkBytes: 4096, onProgress: event => progress.push(event.rows) });
          check(owned.listeners.size === 0 && owned.observed.maxRows <= 7 && owned.observed.maxChunk <= 4096, 'Analytical worker bounds/listener lifetime differ.'); observed = owned.observed; blob = new Blob(chunks);
        } finally { owned.release(); }
      }
      check(enumerations === 1 && returned === 1 && produced === selected.length, 'Analytical source was replayed, truncated or retained.');
      check(progress.at(-1) === selected.length, 'Generated report sections changed selected-record progress.');
      if (completion) check(completion.rows === selected.length && completion.columns === 7 && completion.bytes === blob.size, 'Analytical completion differs from selected source/accepted bytes.');
      if (delivery !== 'blob') check(writes > 0 && earlyWrite, 'Analytical report did not stream before source completion.');
      const data = new Uint8Array(await blob.arrayBuffer()); let binary = '';
      for (let offset = 0; offset < data.length; offset += 8192) binary += String.fromCharCode(...data.subarray(offset, offset + 8192));
      await writeHierarchicalArtifact(profile + '-' + delivery + '.' + format, btoa(binary), JSON.stringify(contract));
      results.push({ format, delivery, produced, enumerations, returned, writes, bytes: data.length, ...(observed ? { worker: observed } : {}) });
    }
    for (const format of ['xlsx', 'pdf']) {
      const controller = new AbortController(), reason = Error('Hierarchical selected report cancelled'), source = await capture(controller.signal, format);
      let returned = 0, produced = 0, caught, released;
      const release = new Promise(resolve => { released = resolve; });
      async function* cancelled() { try { for await (const row of source.rows) { if (++produced === 20) controller.abort(reason); yield row; } } finally { returned++; released(); } }
      try { await example.writeAnalyticalReport(api, { input: inputFor(format), rows: cancelled(), columns: source.columns, sink: { write() {} }, signal: controller.signal }); } catch (error) { caught = error; }
      check(caught === reason && produced === 20, 'Analytical cancellation lost original reason/source bound.');
      await new Promise((resolve, reject) => { const timeout = setTimeout(() => reject(Error('Cancelled hierarchical source was not returned.')), 10000); release.then(() => { clearTimeout(timeout); resolve(); }); });
      check(returned === 1, 'Cancelled analytical source was not returned once.');
    }
    return { profile, stack, selectedRows: selected.length, levels: groupKeys.length, results, hiddenFields: groupKeys,
      selection: 'real native filter, selection exclusion, current visible column order, explicit hidden metadata intake', cancelledAfterRows: 20, chartSummary };
  } finally { table?.destroy(); grid?.destroy(); }
}
