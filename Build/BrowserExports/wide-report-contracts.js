// Wide inventories exercise real grid captures; document layout remains in the PDF engine.
async function runWideReportContracts({ moduleBase, profile, stack, regular, bold }) {
  const api = profile === 'datatables' ? globalThis.OfficeIMO : await import(moduleBase + '/index.js');
  const execution = await import(moduleBase + '/execution/index.js');
  const check = (condition, message) => { if (!condition) throw Error(message); };
  const decode = value => Uint8Array.from(atob(value), character => character.charCodeAt(0));
  const names = ['Service ID', 'Site code', 'Team', 'OS', 'Health', 'IP address', 'Port', 'CPU load', 'Memory MB', 'Disk GB',
    'Latency ms', 'Uptime pct', 'TLS issuer', 'Expires UTC', 'Protocol', 'Build', 'Zone', 'Backup', 'Check UTC', 'Ticket'];
  const keys = names.map((_, index) => 'c' + index), count = 61, width = profile === 'hfx' ? 20 : 12;
  function record(index) {
    return ['S' + String(index).padStart(4, '0'), 'LOC' + String(index % 4).padStart(2, '0'), 'TEAM' + index % 7, 'Linux' + index % 3,
      index % 11 ? 'Healthy' : 'Offline', '10.1.' + Math.floor(index / 250) + '.' + (index % 250 + 1), 8000 + index,
      index + .25, 512 + index, 1024 + index, index + 2.25, 99 + index / 1000, 'CA' + index % 3,
      '2027-01-' + String(index % 28 + 1).padStart(2, '0'), 'HTTPS', 'v' + index, 'ZONE' + index % 2, 'Daily',
      '10:' + String(index % 60).padStart(2, '0'), 'INC' + index];
  }
  const records = Array.from({ length: count }, (_, index) => record(index).slice(0, width));
  const groupKeys = ['portfolio', 'environment', 'owner'];
  const groups = index => ['Portfolio ' + Math.floor(index / 32), 'Environment ' + Math.floor(index / 16), 'Owner ' + Math.floor(index / 8)];
  const columns = keys.slice(0, width).map((key, index) => ({ key, header: names[index], value: row => row[index], wrapText: true }));
  const display = value => String(value ?? '');
  let table, grid, queries = [], heldPage, sourceSignal, holdPage = false;
  let expected = records, selectedOrder = Array.from({ length: width }, (_, index) => index), structure;
  if (profile === 'datatables') {
    const sourceApi = await import(moduleBase + '/integrations/datatables/index.js');
    document.body.innerHTML = '<table id="wide-table"><thead><tr><th colspan="2">Identity</th><th colspan="4">Configuration</th><th colspan="6">Measurements</th><th>Internal</th></tr><tr>' +
      names.slice(0, width).map(name => '<th>' + name + '</th>').join('') + '<th>Internal</th></tr></thead></table>';
    table = new DataTable('#wide-table', { data: records.map((row, index) => [index % 4 === 3 ? 'UNSELECTED-' + index : row[0], ...row.slice(1), 'PRIVATE-ONLY']),
      columns: [...columns.map(column => ({ title: column.header })), { title: 'Internal', visible: false }], select: true, colReorder: true, pageLength: 7, order: [[0, 'desc']] });
    table.search('S0').draw(); table.rows({ search: 'applied' }).select(); table.row(60).deselect();
    selectedOrder = [0, 1, 3, 2, 4, 5, 8, 6, 7, 9, 10, 11]; table.colReorder.order([...selectedOrder, width]);
    expected = records.filter((_, index) => index % 4 !== 3 && index !== 60).reverse().map(row => selectedOrder.map(index => row[index]));
    const columnOptions = {};
    table.columns().indexes().toArray().forEach(index => { const original = table.column(index).dataSrc(); if (original < width) {
      const { value: _value, ...descriptor } = columns[original]; columnOptions[index] = descriptor;
    } });
    const captured = sourceApi.createDataTablesExport(DataTable, table, { headings: 'structured', batchRows: 7, maxBatchCells: 84,
      columnOptions, exportOptions: { columns: ':visible', orthogonal: 'export' } });
    check(captured.columns.map(column => column.key).join() === selectedOrder.map(index => keys[index]).join(), 'Selected current DataTables column order differs.');
    structure = captured;
  } else if (profile === 'canopyx') {
    const navigation = stack, sourceApi = await import(moduleBase + '/integrations/canopyx/index.js');
    document.body.innerHTML = '<div id="wide-grid" style="height:500px"></div>';
    const view = { protocol: 'canopyx/2', type: 'grid', id: 'wide-' + navigation, label: 'Wide service inventory', source: { kind: 'host', navigation },
      rowHeight: 28, pageSize: 7, presentation: 'paged', pageLength: 5, selection: 'none', reporting: {},
      columns: [...columns.map((column, index) => ({ id: column.key, title: column.header, kind: index >= 6 ? 'number' : 'text', width: 140 })),
        ...groupKeys.map(id => ({ id, title: id, hidden: true, kind: 'text' }))] };
    const source = { navigation, async query(query, { signal }) {
      sourceSignal = signal; const offset = navigation === 'cursor' ? Number(query.cursor ?? 0) : query.offset;
      queries.push(offset);
      if (holdPage && offset > 0) { heldPage?.(); await new Promise((_, reject) => signal.addEventListener('abort', () => reject(signal.reason), { once: true })); }
      const items = records.slice(offset, offset + query.limit).map((row, position) => ({ id: 'r' + (offset + position), cells: Object.fromEntries([
        ...row.map((value, index) => [keys[index], { value, text: display(value) }]), ...groups(offset + position).map((value, index) => [groupKeys[index], value])]) }));
      return { protocol: 'canopyx/2', type: navigation === 'cursor' ? 'gridCursorPage' : 'gridPage', view: query.view, revision: 'wide-v1', offset,
        total: navigation === 'cursor' ? null : count, matched: navigation === 'cursor' ? null : count, items,
        ...(navigation === 'cursor' && offset + items.length < count ? { nextCursor: String(offset + items.length) } : {}) };
    } };
    grid = CanopyX.createDataGrid(document.querySelector('#wide-grid'), { view, source }); await grid.ready;
    structure = { capture(signal) {
      const native = grid.prepareExport({ scope: 'all', values: 'raw', columns: [...keys.slice(0, width), ...groupKeys] });
      check(native.request.recordCount === (navigation === 'cursor' ? null : count), 'Native unknown/known record-count contract differs.');
      return sourceApi.createCanopyExport(native, 'pdf', { signal });
    } };
  }
  const grouped = profile === 'canopyx', title = 'Wide selected service inventory';
  const input = { grouped, title, width, columns: selectedOrder.map(index => ({ key: keys[index], header: names[index] })),
    headerRows: profile === 'datatables' ? structure.headerStructure : undefined,
    regular: decode(regular), bold: decode(bold), groupKeys };
  // This factory is also used unchanged inside the real export worker.
  async function writeReport(api, input, rows, sink, signal, onProgress) {
    const columns = input.columns.map((column, index) => ({ ...column, value: row => row[index], wrapText: true }));
    const fonts = { regular: await api.PdfFont.create(input.regular, { signal }), bold: await api.PdfFont.create(input.bold, { signal }) };
    const shared = { fonts, signal, onProgress, pageSize: 'A4', orientation: 'landscape', margins: 30, fontSize: 9, padding: 4,
      pageNumbers: true, limits: { maxRows: 1000, maxCells: 30000, maxPages: 100, maxTableBands: 8, maxTableWindowRows: 9 } };
    const tableOptions = { columns, columnWidths: Array(columns.length).fill(84), wideTable: { mode: 'bands', minimumColumnWidth: 84, repeatColumns: [0, 1] },
      ...(input.headerRows ? { headerRows: input.headerRows } : {}), headerPresentation: { background: '173C75', color: 'FFFFFF', bold: true } };
    async function* blocks() {
      yield { kind: 'section', title: input.title };
      if (!input.grouped) yield api.pdfTable(rows, tableOptions);
      else {
        const report = api.createGroupedReport(rows, { columns, groupBy: input.groupKeys.map((key, level) => ({ key, value: row => row[columns.length + level] })), signal });
        yield api.pdfGroupedTable(report, { labelColumn: { header: 'Group', width: 22 },
          subtotalLabel: group => 'Subtotal ' + group.label + ' (' + group.rows + ')', grandTotalLabel: summary => 'All records (' + summary.rows + ')',
          continuationLabel: group => group.label + ' continued', columnWidths: [120, ...Array(columns.length).fill(84)],
          wideTable: { mode: 'bands', minimumColumnWidth: 84, repeatColumns: [0, 1] },
          headerPresentation: tableOptions.headerPresentation });
      }
      yield { kind: 'section', title: 'Report complete' };
      yield { kind: 'paragraph', text: 'Selected scope is complete.' };
    }
    return api.writePdfDocumentTo(blocks(), sink, shared);
  }
  let produced = 0, returned = 0, enumerations = 0, reads = 0, writes = 0, sourceReturned;
  const progress = [], chunks = [];
  async function* rows(signal) {
    enumerations++;
    try {
      if (profile === 'hfx') for (const row of expected) {
        const captured = [];
        row.forEach((value, index) => Object.defineProperty(captured, index, { enumerable: true, get() { reads++; return value; } }));
        produced++; yield captured;
      } else if (profile === 'datatables') for await (const row of structure.rows) { produced++; yield row; }
      else {
        queries = []; const capture = structure.capture(signal);
        for await (const row of capture.rows()) { produced++; yield row.map(value => value instanceof api.ExportCell ? value.value : value); }
      }
    } finally { returned++; sourceReturned?.(); }
  }
  const sink = { async write(chunk) { const before = produced; await new Promise(resolve => setTimeout(resolve, 1));
    check(before === produced, 'Input advanced while byte output awaited acknowledgement.'); chunks.push(chunk.slice()); writes++; } };
  let workerState;
  try {
    let completion;
    if (profile !== 'hfx') completion = await writeReport(api, input, rows(), sink, undefined, event => progress.push(event.rows));
    else {
      const body = 'import * as api from ' + JSON.stringify(moduleBase + '/index.js') + '; import * as execution from ' + JSON.stringify(moduleBase + '/execution/index.js') + ';\n' +
        writeReport.toString() + '\nexecution.handleWorkerExports(self,({input,rows,sink,signal,onProgress})=>writeReport(api,input,rows,sink,signal,onProgress));';
      const url = URL.createObjectURL(new Blob([body], { type: 'text/javascript' })), worker = new Worker(url, { type: 'module' });
      let maxRows = 0, maxChunk = 0;
      const endpoint = { postMessage(message, transfer) { if (message.kind === 'rows') maxRows = Math.max(maxRows, message.rows.length); worker.postMessage(message, transfer); },
        addEventListener: (...args) => worker.addEventListener(...args), removeEventListener: (...args) => worker.removeEventListener(...args) };
      const output = event => { if (event.data.kind === 'output') maxChunk = Math.max(maxChunk, event.data.bytes.length); }; worker.addEventListener('message', output);
      try { completion = await execution.executeWorkerExport(endpoint, rows(), sink, { input, batchRows: 7, maxBatchCells: 140, maxChunkBytes: 4096, onProgress: event => progress.push(event.rows) }); }
      finally { worker.terminate(); URL.revokeObjectURL(url); }
      workerState = { maxRows, maxChunk }; check(maxRows <= 7 && maxChunk <= 4096, 'Wide worker exceeded row/chunk ownership bounds.');
    }
    check(enumerations === 1 && returned === 1 && produced === expected.length, 'Wide source was replayed, truncated or retained.');
    check(profile !== 'hfx' || reads === expected.length * width, 'Wide source cell getters were read more than once.');
    check(progress.at(-1) === expected.length && completion.rows === expected.length, 'Horizontal repetition changed logical record progress/count.');
    check(profile !== 'canopyx' || new Set(queries).size === queries.length, 'Wide bands replayed native source pages.');
    const blob = new Blob(chunks, { type: 'application/pdf' }), data = new Uint8Array(await blob.arrayBuffer()); let binary = '';
    for (let offset = 0; offset < data.length; offset += 8192) binary += String.fromCharCode(...data.subarray(offset, offset + 8192));
    const contract = { profile, stack, grouped, headers: input.columns.map(column => column.header), rows: expected.map(row => row.map(display)),
      groupPaths: grouped ? expected.map((_, index) => groups(index)) : [], widths: grouped ? [120, ...Array(width).fill(84)] : Array(width).fill(84),
      repeated: [0, 1], title, structured: profile === 'datatables', forbidden: ['UNSELECTED-', 'PRIVATE-ONLY'] };
    await writeWideArtifact(profile + '-' + stack + '.pdf', btoa(binary), JSON.stringify(contract));
    if (profile === 'canopyx' && stack === 'cursor') {
      const controller = new AbortController(), reason = Error('Wide native page cancelled'); let notified;
      const pendingPage = new Promise(resolve => { notified = resolve; }); heldPage = notified; holdPage = true;
      const released = new Promise(resolve => { sourceReturned = resolve; });
      const task = writeReport(api, input, rows(controller.signal), { write() {} }, controller.signal);
      let watchdog;
      try { await Promise.race([pendingPage, new Promise((_, reject) => { watchdog = setTimeout(() => reject(Error('Native page did not enter pending I/O.')), 10000); })]); }
      finally { clearTimeout(watchdog); }
      controller.abort(reason); const error = await task.then(() => null, error => error);
      check(error === reason && sourceSignal.aborted, 'Wide cancellation did not abort the pending native page with its reason.');
      try { await Promise.race([released, new Promise((_, reject) => { watchdog = setTimeout(() => reject(Error('Cancelled native capture was not released.')), 10000); })]); }
      finally { clearTimeout(watchdog); }
      check(returned === 2, 'Cancelled native capture was not returned once.');
    }
    return { profile, stack, rows: expected.length, columns: width, enumerations, returned, writes, reads, pages: queries.length, worker: workerState,
      selection: profile === 'datatables' ? 'native filter + selected exclusion + current reorder + structured headings' : profile === 'canopyx' ? 'native ' + stack + ' + hidden hierarchy' : 'HFX-shaped operational inventory' };
  } finally { table?.destroy(); grid?.destroy(); }
}
