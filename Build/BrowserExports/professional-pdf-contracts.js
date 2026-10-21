// Real grid captures feed the shared executable example; no document/layout implementation lives here.
async function runProfessionalPdfContracts({ moduleBase, exampleUrl, profile, regular, bold, chart, chartSummary }) {
  const api = profile === 'datatables' ? globalThis.OfficeIMO : await import(moduleBase + '/index.js');
  const sourceApi = await import(moduleBase + '/integrations/' + profile + '/index.js');
  const example = await import(exampleUrl);
  const check = (condition, message) => { if (!condition) throw Error(message); };
  const bytes = text => Uint8Array.from(atob(text), character => character.charCodeAt(0));
  const headers = example.reportColumns.map(column => column.header);
  const records = [...example.operationalRows(128)].map((row, index) => index % 4 === 3 ? ['UNSELECTED-' + index, ...row.slice(1)] : row);
  const excludedRecord = 126; // The first loaded, accepted row after descending service sort.
  const expected = records.map((row, index) => ({ row, index })).filter(item => item.index % 4 !== 3 && item.index !== excludedRecord).reverse();
  const summary = { total: expected.length, Healthy: 0, Attention: 0, Offline: 0 };
  for (const item of expected) summary[item.row[2]]++;
  check(['total', 'Healthy', 'Attention', 'Offline'].every(key => summary[key] === chartSummary[key]), 'The independently produced figure does not describe the selected services.');
  const order = [1, 0, 2, 3, 4, 5], titles = order.map(index => headers[index]);
  const format = (value, index) => value instanceof Date ? value.toISOString().slice(0, 16).replace('T', ' ')
    : index === 3 ? new Intl.NumberFormat('en', { minimumFractionDigits: 1, maximumFractionDigits: 1 }).format(value)
    : index === 4 ? new Intl.NumberFormat('en', { style: 'percent', minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value) : String(value);
  const contract = { profile, headers: titles, rows: expected.map(item => order.map(index => format(item.row[index], index))),
    summary, figureCaption: 'Figure 1. Status of the 95 selected services: 84 healthy, 8 requiring attention and 3 offline.' };
  const input = { format: 'pdf', summary, scope: 'Selected production services · typed capture · UTC checks',
    fonts: { regular: bytes(regular), bold: bytes(bold) }, figure: { image: { format: 'png', data: bytes(chart) }, width: 480,
      caption: contract.figureCaption, alternativeText: 'Selected service status: 84 healthy, 8 requiring attention and 3 offline; 95 services in total.' } };
  const tone = status => ({ background: status === 'Offline' ? 'FEE2E2' : status === 'Attention' ? 'FEF3C7' : 'DCFCE7',
    color: status === 'Offline' ? '991B1B' : status === 'Attention' ? '92400E' : '166534' });
  let table, grid;
  async function capture(signal) {
    if (profile === 'datatables') {
      const columnOptions = {};
      table.columns().indexes().toArray().forEach(index => {
        const original = table.column(index).dataSrc();
        if (original < 6) columnOptions[index] = example.reportColumns[original];
      });
      const source = sourceApi.createDataTablesExport(DataTable, table, { signal, headings: 'leaf', batchRows: 7,
        exportOptions: { columns: ':visible' }, columnOptions,
        project(_value, context) {
          const original = table.column(context.sourceColumnIndex).dataSrc(), value = records[context.sourceRowIndex][original];
          return original === 2 ? new api.ExportCell(value, { presentation: tone(value) }) : value;
        } });
      return { columns: source.columns, rows: source.rows };
    }
    const native = grid.prepareExport({ scope: 'selected', values: 'raw', columns: order.map(index => ['service', 'site', 'status', 'latency', 'availability', 'checked'][index]), locale: 'en-US', timeZone: 'utc' });
    check(native.request.values === 'raw' && native.request.recordCount === expected.length, 'Native selection/raw capture differs.');
    const source = sourceApi.createCanopyExport(native, 'pdf', { signal,
      tones: { success: tone('Healthy'), warning: tone('Attention'), danger: tone('Offline') },
      columnOptions: Object.fromEntries(['service', 'site', 'status', 'latency', 'availability', 'checked'].map((id, index) => [id, { width: example.reportColumns[index].width }])) });
    return { columns: source.columns, rows: source.rows() };
  }
  if (profile === 'datatables') {
    document.body.innerHTML = '<h1>Selected service review</h1><table id="professional-table"></table>';
    table = new DataTable('#professional-table', { data: records.map(row => [...row, 'INTERNAL-ONLY']),
      columns: [...headers.map(title => ({ title, render: value => value instanceof Date ? value.toISOString() : value })), { title: 'Internal', visible: false }],
      select: true, colReorder: true, order: [[0, 'desc']], pageLength: 8 });
    table.search('WEB-').draw(); table.rows({ search: 'applied' }).select(); table.row(excludedRecord).deselect(); table.colReorder.order([1, 0, 2, 3, 4, 5, 6]);
  } else {
    document.body.innerHTML = '<h1>Selected service review</h1><div id="professional-grid" style="height:500px"></div>';
    const ids = ['service', 'site', 'status', 'latency', 'availability', 'checked'];
    const view = { protocol: 'canopyx/2', type: 'grid', id: 'professional-services', label: 'Service inventory', source: { kind: 'inline' },
      rowHeight: 30, pageSize: 7, presentation: 'paged', pageLength: 8, selection: 'multiple', search: true,
      reporting: { selectionControls: true }, columns: ids.map((id, index) => ({ id, title: headers[index], sortable: true, width: [160, 110, 120, 130, 130, 190][index],
        kind: index === 5 ? 'datetime' : [3, 4].includes(index) ? 'number' : 'text', wrap: true })),
      rows: records.map((row, index) => ({ id: 'record-' + index, cells: Object.fromEntries(ids.map((id, column) => [id,
        column === 2 ? { value: row[column], tone: row[column] === 'Offline' ? 'danger' : row[column] === 'Attention' ? 'warning' : 'success' }
          : column >= 3 ? { value: column === 5 ? row[column].toISOString() : row[column], text: format(row[column], column) }
          : row[column]])) })) };
    grid = CanopyX.createDataGrid(document.querySelector('#professional-grid'), { view }); await grid.ready;
    await grid.setQuery({ search: 'WEB-', sort: [{ column: 'service', direction: 'desc' }] });
    grid.setColumnOrder(order.map(index => ids[index]));
    grid.selectAllMatching(); grid.setRowSelected('record-' + excludedRecord, false);
  }
  const results = [];
  try {
    for (const delivery of ['blob', 'sink']) {
      const source = await capture(), progress = [], chunks = []; let enumerations = 0, returned = 0, produced = 0, writes = 0, accepted = 0;
      async function* rows() { enumerations++; try { for await (const row of source.rows) { produced++; yield row; } } finally { returned++; } }
      const report = await example.serviceReliabilityDocument(api, { input, rows: rows(), columns: source.columns, onProgress: event => progress.push(event.rows) });
      let blob, result;
      if (delivery === 'blob') blob = await api.writePdfDocument(report.blocks, report.options);
      else {
        const sink = new WritableStream({ async write(chunk) {
          const before = produced; await new Promise(resolve => setTimeout(resolve, 1));
          check(produced === before, 'Rows advanced while destination was awaiting acknowledgement.');
          chunks.push(chunk.slice()); accepted += chunk.length; writes++;
        } });
        result = await api.writePdfDocumentTo(report.blocks, sink, report.options); check(!sink.locked, 'Completed sink stayed locked.');
        check(result.rows === expected.length && result.columns === 6 && result.bytes === accepted, 'Composed sink result differs.');
        blob = new Blob(chunks, { type: 'application/pdf' });
      }
      check(enumerations === 1 && returned === 1 && produced === expected.length, 'Report source was replayed or not released.');
      check(progress.at(-1) === expected.length, 'Final progress lost the selected count.');
      const data = new Uint8Array(await blob.arrayBuffer()); let binary = '';
      for (let offset = 0; offset < data.length; offset += 8192) binary += String.fromCharCode(...data.subarray(offset, offset + 8192));
      await writeProfessionalPdf(profile + '-' + delivery + '.pdf', btoa(binary), JSON.stringify(contract));
      results.push({ delivery, rows: produced, enumerations, returned, writes, bytes: data.length });
    }
    const controller = new AbortController(), reason = Error('Selected report cancelled during inventory'), source = await capture(controller.signal);
    let returned = 0, produced = 0, release;
    const released = new Promise(resolve => { release = resolve; });
    async function* cancelledRows() { try { for await (const row of source.rows) { if (++produced === 20) controller.abort(reason); yield row; } } finally { returned++; release(); } }
    const report = await example.serviceReliabilityDocument(api, { input, rows: cancelledRows(), columns: source.columns, signal: controller.signal });
    const sink = new WritableStream({ write() {} }); let caught;
    try { await api.writePdfDocumentTo(report.blocks, sink, report.options); } catch (error) { caught = error; }
    check(caught === reason && produced === 20 && !sink.locked, 'Cancellation lost its reason/progress/stream lifetime: ' + JSON.stringify({ sameReason: caught === reason, error: String(caught), produced, locked: sink.locked }));
    // Failure preserves a nonblocking cancellation path; observe normal async iterator cleanup separately.
    await new Promise((resolve, reject) => {
      const watchdog = setTimeout(() => reject(Error('Selected capture did not release after cancellation.')), 10000);
      released.then(() => { clearTimeout(watchdog); resolve(); });
    });
    check(returned === 1, 'Cancelled selected source was not returned once.');
    return { profile, results, cancelledAfterRows: produced, captureScope: 'native filter + selection + current column order', noWholeTableExportMatrix: true };
  } finally { table?.destroy(); grid?.destroy(); }
}
