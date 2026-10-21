// Real grid captures feed the shared operation. Expected rows/totals are computed separately here.
async function runGroupedReportContracts({ moduleBase, exampleUrl, profile, stack, regular, bold, emptyOnly = false }) {
  const api = profile === 'datatables' ? globalThis.OfficeIMO : await import(moduleBase + '/index.js');
  const execution = await import(moduleBase + '/execution/index.js'), example = await import(exampleUrl);
  const sourceApi = await import(moduleBase + '/integrations/' + profile + '/index.js');
  const check = (condition, message) => { if (!condition) throw Error(message); };
  const ids = ['service', 'site', 'status', 'latency', 'availability', 'checked'];
  const titles = ['Endpoint custom', 'Tenant location', 'Health value', 'Response ms', 'Uptime ratio', 'Observed UTC'];
  const columns = example.reportColumns.map((column, index) => ({ ...column, header: titles[index] }));
  const order = [1, 0, 2, 3, 4, 5], sites = ['AA · Łódź', 'BB · Praha', 'CC · Berlin', 'DD · Warszawa'];
  const records = [...example.operationalRows(256)].map((row, index) => [index % 5 === 4 ? 'UNSELECTED-' + index : row[0], sites[index % 4],
    row[2], index % 5 === 0 ? null : row[3], index % 7 === 0 ? null : row[4], row[5]]);
  const excluded = 252;
  const expected = records.map((row, index) => ({ row, index })).filter(item => item.index % 5 !== 4 && item.index !== excluded)
    .sort((left, right) => left.row[1] < right.row[1] ? -1 : left.row[1] > right.row[1] ? 1 : right.row[0].localeCompare(left.row[0]));
  const language = profile === 'canopyx' ? 'pl' : undefined;
  const labels = language === 'pl' ? { header: 'Grupa', group: 'Lokalizacja: ', subtotal: 'Suma grupy', total: 'Wszystkie wybrane usługi', records: 'usług', continued: ' (ciąg dalszy)' }
    : { header: 'Group', group: 'Site: ', subtotal: 'Group total', total: 'All selected services', records: 'services', continued: ' (continued)' };
  const translated = { service: 'Usługa', site: 'Lokalizacja', status: 'Stan', latency: 'Opóźnienie (ms)', availability: 'Dostępność', checked: 'Sprawdzone (UTC)' };
  const selectedHeaders = order.map(index => language ? translated[ids[index]] : titles[index]);
  const average = (rows, column) => { const values = rows.map(row => row[column]).filter(value => value !== null); return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : null; };
  const groupRows = empty => {
    const groups = [];
    for (const item of empty ? [] : expected) {
      if (groups.at(-1)?.key !== item.row[1]) groups.push({ key: item.row[1], rows: [] });
      groups.at(-1).rows.push(item.row);
    }
    return groups.map(group => ({ key: group.key, label: labels.group + group.key, rows: group.rows.length,
      subtotal: labels.subtotal + ' · ' + group.rows.length.toLocaleString(language ?? 'en') + ' ' + labels.records,
      latency: average(group.rows, 3), availability: average(group.rows, 4) }));
  };
  const decimal = new Intl.NumberFormat(language ?? 'en', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const percent = new Intl.NumberFormat(language ?? 'en', { style: 'percent', minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const summaryText = (latency, availability) => [latency === null ? '' : decimal.format(latency), availability === null ? '' : percent.format(availability)];
  const pageLabels = language === 'pl' ? { prefix: 'Strona ', separator: ' z ' } : { prefix: 'Page ', separator: ' of ' };
  const contract = (empty, includeGrandTotal = true) => ({ profile, stack, language: language ?? 'en', includeGrandTotal, keys: order.map(index => ids[index]), headers: selectedHeaders,
    rows: (empty ? [] : expected).map(item => order.map(index => item.row[index] instanceof Date ? item.row[index].toISOString() : item.row[index])),
    groups: groupRows(empty), totals: { latency: average(empty ? [] : expected.map(item => item.row), 3), availability: average(empty ? [] : expected.map(item => item.row), 4) },
    labelHeader: labels.header, totalLabel: labels.total + ' · ' + (empty ? 0 : expected.length).toLocaleString(language ?? 'en') + ' ' + labels.records,
    summaryRows: [...groupRows(empty).map(group => ({ label: group.subtotal, values: summaryText(group.latency, group.availability) })), ...(includeGrandTotal ? [
      { label: labels.total + ' · ' + (empty ? 0 : expected.length).toLocaleString(language ?? 'en') + ' ' + labels.records,
        values: summaryText(average(empty ? [] : expected.map(item => item.row), 3), average(empty ? [] : expected.map(item => item.row), 4)) }] : [])],
    pageLabels,
    continuationSuffix: labels.continued, sourceScope: 'selected filtered records; host ordered by raw site key and service; hidden fields omitted' });
  let table, grid;
  if (profile === 'datatables') {
    document.body.innerHTML = '<h1>Grouped service report capture</h1><table id="grouped-table"></table>';
    table = new DataTable('#grouped-table', { data: records.map(row => [...row, 'INTERNAL-ONLY']),
      columns: [...titles.map(title => ({ title, render: (value, type) => type === 'export' || type === 'sort' ? value
        : value instanceof Date ? value.toISOString() : typeof value === 'number' ? 'DISPLAY-ONLY ' + value : value })), { title: 'Internal', visible: false }],
      select: true, colReorder: true, order: [[1, 'asc'], [0, 'desc']], pageLength: 8 });
    table.search('WEB-').draw(); table.rows({ search: 'applied' }).select(); table.row(excluded).deselect();
    table.colReorder.order([1, 0, 2, 3, 4, 5, 6]);
  } else {
    document.body.innerHTML = '<h1>Grouped service report capture</h1><div id="grouped-grid" style="height:500px"></div>';
    const view = { protocol: 'canopyx/2', type: 'grid', id: 'grouped-services', label: 'Grouped inventory', source: { kind: 'inline' },
      rowHeight: 30, pageSize: 7, presentation: 'paged', pageLength: 8, selection: 'multiple', search: true, reporting: { selectionControls: true },
      columns: [...ids.map((id, index) => ({ id, title: titles[index], sortable: true, width: [160, 150, 120, 130, 130, 190][index],
        kind: index === 5 ? 'datetime' : [3, 4].includes(index) ? 'number' : 'text' })), { id: 'internal', title: 'Internal', kind: 'text', hidden: true }],
      rows: records.map((row, index) => ({ id: 'record-' + index, cells: { ...Object.fromEntries(ids.map((id, column) => [id,
        { value: column === 5 ? row[column].toISOString() : row[column], text: row[column] === null ? '' : column === 5 ? row[column].toISOString().slice(0, 16).replace('T', ' ')
          : column === 3 ? decimal.format(row[column]) : column === 4 ? percent.format(row[column]) : String(row[column]) }])), internal: 'INTERNAL-ONLY' } })) };
    grid = CanopyX.createDataGrid(document.querySelector('#grouped-grid'), { view }); await grid.ready;
    await grid.setQuery({ search: 'WEB-', sort: [{ column: 'site', direction: 'asc' }, { column: 'service', direction: 'desc' }] });
    grid.setColumnOrder([...order.map(index => ids[index]), 'internal']); grid.selectAllMatching(); grid.setRowSelected('record-' + excluded, false);
  }
  async function capture(signal, empty = false, format = 'xlsx') {
    if (profile === 'datatables') {
      const columnOptions = {};
      table.columns().indexes().toArray().forEach(index => { const original = table.column(index).dataSrc(); if (original < 6) columnOptions[index] = columns[original]; });
      const source = sourceApi.createDataTablesExport(DataTable, table, { signal, headings: 'leaf', batchRows: 7, columnOptions,
        exportOptions: { orthogonal: 'export', columns: ':visible', ...(empty ? { rows: [] } : {}) } });
      return { columns: source.columns, rows: source.rows };
    }
    if (empty) grid.clearSelection();
    const native = grid.prepareExport({ scope: 'selected', values: 'raw', columns: order.map(index => ids[index]), locale: 'pl-PL', timeZone: 'utc' });
    check(native.request.values === 'raw' && native.request.recordCount === (empty ? 0 : expected.length), 'Canopy native scope differs.');
    const source = sourceApi.createCanopyExport(native, format, { signal,
      columnOptions: Object.fromEntries(ids.map((id, index) => [id, { width: columns[index].width,
        ...(format === 'pdf' ? { wrapText: true } : {}), ...(columns[index].format ? { format: columns[index].format } : {}) }])) });
    return { columns: source.columns, rows: source.rows() };
  }
  const inputFor = format => ({ format, ...(language ? { language } : {}), compression: 'store',
    ...(format === 'pdf' ? { fonts: { regular: Uint8Array.from(atob(regular), item => item.charCodeAt(0)), bold: Uint8Array.from(atob(bold), item => item.charCodeAt(0)) } } : {}) });
  async function save(name, format, blob, empty = false, includeGrandTotal = true) {
    const bytes = new Uint8Array(await blob.arrayBuffer()); let binary = '';
    for (let offset = 0; offset < bytes.length; offset += 8192) binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
    await writeGroupedArtifact(name + '.' + format, btoa(binary), JSON.stringify(contract(empty, includeGrandTotal))); return bytes.length;
  }
  function worker() {
    const body = 'import * as api from ' + JSON.stringify(moduleBase + '/index.js') + ';\nimport * as execution from ' + JSON.stringify(moduleBase + '/execution/index.js') + ';\nimport * as example from ' + JSON.stringify(exampleUrl) + ';\n' +
      'execution.handleWorkerExports(self,async ({input,rows,sink,signal,onProgress})=>example.writeGroupedOperationalReport(api,{input,columns:input.columns,rows,sink,signal,onProgress}));';
    const url = URL.createObjectURL(new Blob([body], { type: 'text/javascript' })), owned = new Worker(url, { type: 'module' }), listeners = new Set(), observed = { maxRows: 0, maxChunk: 0 };
    return { observed, listeners, endpoint: {
      postMessage(message) { if (message.kind === 'rows') observed.maxRows = Math.max(observed.maxRows, message.rows.length); owned.postMessage(message); },
      addEventListener(type, listener) { const wrapped = event => { if (type === 'message' && event.data.kind === 'output') observed.maxChunk = Math.max(observed.maxChunk, event.data.bytes.length); listener(event); }; listeners.add({ type, listener, wrapped }); owned.addEventListener(type, wrapped); },
      removeEventListener(type, listener) { const entry = [...listeners].find(item => item.type === type && item.listener === listener); if (entry) { owned.removeEventListener(type, entry.wrapped); listeners.delete(entry); } }
    }, release() { owned.terminate(); URL.revokeObjectURL(url); } };
  }
  const results = [];
  try {
    const deliveries = emptyOnly ? [] : [['xlsx', 'blob'], ['xlsx', 'sink'], ['pdf', 'blob'], ...(stack === 'current' ? [['pdf', 'sink'], ['xlsx', 'worker']] : [])];
    for (const [format, delivery] of deliveries) {
      const source = await capture(undefined, false, format), chunks = [], progress = []; let produced = 0, returned = 0, enumerations = 0, writes = 0, earlyWrite = false;
      async function* rows() { enumerations++; try { for await (const row of source.rows) { produced++; yield row; } } finally { returned++; } }
      const destination = { async write(chunk) { const stopped = produced; earlyWrite ||= stopped < expected.length; await new Promise(resolve => setTimeout(resolve, 1));
        check(produced === stopped, 'Grouped source advanced during an awaited destination acknowledgement.'); chunks.push(chunk.slice()); writes++; } };
      let blob, completion, observed;
      if (delivery === 'blob') blob = await example.writeGroupedOperationalReport(api, { input: inputFor(format), columns: source.columns, rows: rows(), onProgress: event => progress.push(event.rows) });
      else if (delivery === 'sink') { completion = await example.writeGroupedOperationalReport(api, { input: inputFor(format), columns: source.columns, rows: rows(), sink: destination, onProgress: event => progress.push(event.rows) }); blob = new Blob(chunks); }
      else {
        const owned = worker();
        try { completion = await execution.executeWorkerExport(owned.endpoint, rows(), destination, { input: { ...inputFor(format), columns: source.columns }, batchRows: 7, maxBatchCells: 42, maxChunkBytes: 4096,
          onProgress: event => progress.push(event.rows) });
          check(owned.listeners.size === 0 && owned.observed.maxRows <= 7 && owned.observed.maxChunk <= 4096, 'Grouped worker bounds/listener lifetime differ.'); observed = owned.observed; blob = new Blob(chunks);
        } finally { owned.release(); }
      }
      check(enumerations === 1 && returned === 1 && produced === expected.length, 'Grouped capture was replayed, incomplete or retained.');
      check(progress.at(-1) === expected.length, 'Generated grouped summaries changed source progress.');
      if (completion) check(completion.rows === expected.length && completion.columns === 7 && completion.bytes === blob.size, 'Grouped completion differs from accepted output.');
      if (delivery !== 'blob') check(writes > 0 && earlyWrite, 'Grouped output never reached its destination before source completion.');
      results.push({ format, delivery, produced, enumerations, returned, writes, bytes: await save(profile + '-' + delivery, format, blob), ...(observed ? { worker: observed } : {}) });
    }
    for (const format of emptyOnly ? [] : ['xlsx', 'pdf']) {
      const controller = new AbortController(), reason = Error('Grouped selected export cancelled'), source = await capture(controller.signal, false, format); let produced = 0, returned = 0, caught, release;
      const released = new Promise(resolve => { release = resolve; });
      async function* cancelled() { try { for await (const row of source.rows) { if (++produced === 20) controller.abort(reason); yield row; } } finally { returned++; release(); } }
      try { await example.writeGroupedOperationalReport(api, { input: inputFor(format), columns: source.columns, rows: cancelled(), signal: controller.signal, sink: { write() {} } }); } catch (error) { caught = error; }
      check(caught === reason && produced === 20, 'Grouped cancellation reason/source bounds differ for ' + format + ': ' + JSON.stringify({ sameReason: caught === reason, produced, returned, error: String(caught) }));
      await new Promise((resolve, reject) => { const timer = setTimeout(() => reject(Error('Cancelled grouped source did not release.')), 10000); released.then(() => { clearTimeout(timer); resolve(); }); });
      check(returned === 1, 'Cancelled grouped source was not returned once.');
    }
    const recurrent = api.createGroupedReport([['A', 1], ['B', 2], ['A', 3]], { columns: [{ key: 'place', header: '任意の列' }, { key: 'value', header: 'Metric', type: 'number' }], groupBy: 'place', totals: { value: 'sum' } });
    let rejected; try { for await (const _ of recurrent.events) {} } catch (error) { rejected = error; }
    check(rejected instanceof TypeError && /contiguous/.test(rejected.message), 'A repeated closed group was silently split.');
    let emptyNoGrandColumns;
    if (stack === 'current') {
      for (const format of ['xlsx', 'pdf']) { const empty = await capture(undefined, true, format); const blob = await example.writeGroupedOperationalReport(api, { input: inputFor(format), columns: empty.columns, rows: empty.rows }); await save(profile + '-empty', format, blob, true); }
      const empty = await capture(undefined, true), fonts = { regular: await api.PdfFont.create(inputFor('pdf').fonts.regular), bold: await api.PdfFont.create(inputFor('pdf').fonts.bold) };
      const source = api.createGroupedReport([], { columns: empty.columns, groupBy: 'site', totals: { latency: 'average', availability: 'average' } });
      const chunks = [];
      const completion = await api.writeGroupedPdfTo(source, { write(chunk) { chunks.push(chunk.slice()); } }, { fonts, labelColumn: { header: labels.header, width: 26 },
        pageSize: 'A4', orientation: 'landscape', margins: 30, fontSize: 9, pageNumberLabels: { ...pageLabels, suffix: '' },
        subtotalLabel: () => labels.subtotal, grandTotalLabel: () => labels.total, includeGrandTotal: false });
      check(completion.rows === 0 && completion.columns === 7, 'Empty/no-grand grouped PDF lost its declared source column count.'); emptyNoGrandColumns = completion.columns;
      await save(profile + '-empty-no-grand', 'pdf', new Blob(chunks), true, false);
    }
    return { profile, stack, results, selectedRows: emptyOnly ? 0 : expected.length, groups: emptyOnly ? 0 : groupRows(false).length,
      ...(emptyOnly ? {} : { cancellationRows: 20 }), recurrentGroupRejected: true,
      headersSelectedByStableKeys: true, ...(emptyNoGrandColumns === undefined ? {} : { emptyNoGrandColumns }), emptyOnly,
      values: 'raw numbers, null metrics and UTC dates; host-resolved PDF display text; hidden fields and DataTables display substitutions excluded' };
  } finally { table?.destroy(); grid?.destroy(); }
}
