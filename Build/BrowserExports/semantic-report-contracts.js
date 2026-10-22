// Real host captures are compared with source-authored records; no PDF reader is used here.
async function runSemanticReportContracts({ moduleBase, profile, stack, regular, bold, language = 'en' }) {
  const api = profile === 'datatables' ? globalThis.OfficeIMO : await import(moduleBase + '/index.js');
  const execution = await import(moduleBase + '/execution/index.js');
  const check = (condition, message) => { if (!condition) throw Error(message); };
  const texts = {
    en: { tag: 'en-US', names: ['Service ID', 'Site code', 'Service detail', 'Team', 'Health', 'OS', 'Latency ms', 'Checked UTC', 'Memory MB', 'Disk GB', 'Zone', 'Ticket'],
      groups: ['Portfolio', 'Environment', 'Owner', 'Location'], title: 'Semantic selected service inventory', scope: 'Scope: selected service records. ',
      documentation: 'Report documentation', pageHeader: 'Semantic operational report', caption: 'Table 1. Selected operational records', groupHeader: 'Group',
      subtotal: 'Subtotal', grandTotal: 'Grand total', continuation: 'continued', complete: 'Report complete', completeText: 'Selected scope is complete.' },
    pl: { tag: 'pl-PL', names: ['Identyfikator usługi', 'Kod lokalizacji', 'Opis usługi', 'Zespół', 'Stan', 'System', 'Opóźnienie ms', 'Sprawdzone UTC', 'Pamięć MB', 'Dysk GB', 'Strefa', 'Zgłoszenie'],
      groups: ['Portfel', 'Środowisko', 'Właściciel', 'Lokalizacja'], title: 'Semantyczny raport wybranych usług', scope: 'Zakres: wybrane rekordy usług. ',
      documentation: 'Dokumentacja raportu', pageHeader: 'Semantyczny raport operacyjny', caption: 'Tabela 1. Wybrane rekordy operacyjne', groupHeader: 'Grupa',
      subtotal: 'Suma częściowa', grandTotal: 'Suma końcowa', continuation: 'ciąg dalszy', complete: 'Raport zakończony', completeText: 'Wybrany zakres jest kompletny.' }
  };
  const text = texts[language]; check(text && (language === 'en' || profile === 'hfx'), 'Unsupported semantic fixture language/profile.');
  const names = text.names;
  const keys = names.map((_, index) => 'c' + index), groupKeys = ['portfolio', 'environment', 'owner', 'location'], count = 13;
  const groupPath = index => [text.groups[0] + ' ' + Math.floor(index / 12), text.groups[1] + ' ' + Math.floor(index / 6),
    text.groups[2] + ' ' + Math.floor(index / 6), text.groups[3] + ' ' + Math.floor(index / 3)];
  const detail = index => index === 4 ? Array.from({ length: 46 }, (_, line) => 'D' + String(index).padStart(4, '0') + 'L' + String(line).padStart(3, '0')).join('\n') : 'Detail-' + index;
  const records = Array.from({ length: count }, (_, index) => ['S' + String(index).padStart(4, '0'), index % 2 ? 'Łódź' : 'Warsaw', detail(index), 'Team-' + index,
    index % 5 ? 'Healthy' : 'Offline', 'Linux-' + index, index + .25, '2026-10-22T08:' + String(index).padStart(2, '0') + ':00.000Z',
    512 + index, 1024 + index, 'Zone-' + index, 'INC-' + index]);
  const descriptors = keys.map((key, index) => ({ key, header: names[index], type: [6, 8, 9].includes(index) ? 'number' : index === 7 ? 'date' : 'string', wrapText: true }));
  const scalar = value => value instanceof api.ExportCell ? value.value : value;
  const display = value => value == null ? '' : value instanceof Date ? value.toISOString() : String(value);
  let table, grid, source, projectionCalls = 0, dateValues = 0, queries = [], holdPage = false, pendingPage, sourceSignal;
  let expected = records, selectedOrder = keys.map((_, index) => index), headerRows, footerRows;
  if (profile === 'datatables') {
    const integration = await import(moduleBase + '/integrations/datatables/index.js');
    document.body.innerHTML = '<table id="semantic-table"><thead><tr><th rowspan="2">Service ID</th><th rowspan="2">Site code</th>' +
      '<th colspan="4">Service context</th><th colspan="6">Measurements</th><th rowspan="2">Internal</th></tr><tr>' +
      names.slice(2).map(name => '<th>' + name + '</th>').join('') + '</tr></thead><tfoot><tr><th rowspan="2">Selected totals</th>' +
      '<th rowspan="2">Export scope</th><th colspan="4">Operational scope</th><th colspan="6">Measurements checked</th><th rowspan="2">Internal</th></tr>' +
      '<tr><th colspan="4">Approved: Łódź</th><th colspan="6">Service review complete</th></tr></tfoot></table>';
    table = new DataTable('#semantic-table', { data: records.map((row, index) => [index % 4 === 3 ? 'UNSELECTED-' + index : row[0], ...row.slice(1), 'PRIVATE-ONLY']),
      columns: [...descriptors.map(column => ({ title: column.header, render: (value, type) => type === 'display' && typeof value === 'number' ? value.toFixed(2) + ' display' : value })),
        { title: 'Internal', visible: false }], select: true, colReorder: true, pageLength: 4, order: [[0, 'desc']] });
    table.search('S0').draw(); table.rows({ search: 'applied' }).select(); table.row(12).deselect();
    selectedOrder = [0, 1, 3, 2, 4, 5, 7, 6, 8, 9, 10, 11]; table.colReorder.order([...selectedOrder, keys.length]);
    expected = records.filter((_, index) => index % 4 !== 3 && index !== 12).reverse().map(row => selectedOrder.map(index => row[index]));
    const columnOptions = {};
    table.columns().indexes().toArray().forEach(index => {
      const original = table.column(index).dataSrc(); if (original < keys.length) columnOptions[index] = descriptors[original];
    });
    source = integration.createDataTablesExport(DataTable, table, { headings: 'structured', includeFooter: true, batchRows: 3, maxBatchCells: 36,
      columnOptions, exportOptions: { columns: ':visible', orthogonal: 'export' }, project(value, context) {
        projectionCalls++; const original = table.column(context.sourceColumnIndex).dataSrc();
        if (original === 7) { dateValues++; return new Date(value); }
        return value;
      } });
    check(source.columns.map(column => column.key).join() === selectedOrder.map(index => keys[index]).join(), 'Current selected DataTables column order differs.');
    check(source.rowCount === expected.length && source.headerStructure && source.footerStructure, 'Native structured selected capture is incomplete.');
    headerRows = source.headerStructure; footerRows = source.footerStructure;
  } else if (profile === 'canopyx') {
    const integration = await import(moduleBase + '/integrations/canopyx/index.js');
    document.body.innerHTML = '<div id="semantic-grid" style="height:400px"></div>';
    const view = { protocol: 'canopyx/2', type: 'grid', id: 'semantic-' + stack, label: 'Semantic operational inventory', source: { kind: 'host', navigation: stack },
      rowHeight: 28, pageSize: 3, presentation: 'paged', pageLength: 3, selection: 'none', reporting: {},
      columns: [...descriptors.map(column => ({ id: column.key, title: column.header, kind: column.type === 'date' ? 'datetime' : column.type === 'number' ? 'number' : 'text', width: 140 })),
        ...groupKeys.map(id => ({ id, title: id, hidden: true, kind: 'text' }))] };
    const host = { navigation: stack, async query(query, { signal }) {
      sourceSignal = signal; const offset = stack === 'cursor' ? Number(query.cursor ?? 0) : query.offset; queries.push(offset);
      if (holdPage && offset > 0) { pendingPage?.(); await new Promise((_, reject) => signal.addEventListener('abort', () => reject(signal.reason), { once: true })); }
      const items = records.slice(offset, offset + query.limit).map((row, ordinal) => ({ id: 'record-' + (offset + ordinal), cells: Object.fromEntries([
        ...row.map((value, index) => [keys[index], { value, text: display(value) }]),
        ...groupPath(offset + ordinal).map((value, level) => [groupKeys[level], value])]) }));
      return { protocol: 'canopyx/2', type: stack === 'cursor' ? 'gridCursorPage' : 'gridPage', view: query.view, revision: 'semantic-v1', offset,
        total: stack === 'cursor' ? null : count, matched: stack === 'cursor' ? null : count, items,
        ...(stack === 'cursor' && offset + items.length < count ? { nextCursor: String(offset + items.length) } : {}) };
    } };
    grid = CanopyX.createDataGrid(document.querySelector('#semantic-grid'), { view, source: host }); await grid.ready;
    source = { capture(signal) {
      const native = grid.prepareExport({ scope: 'all', values: 'raw', columns: [...keys, ...groupKeys], locale: 'en-US', timeZone: 'utc' });
      check(native.request.recordCount === (stack === 'cursor' ? null : count), 'Native known/unknown count changed.');
      check(native.request.columns.map(column => column.id).join() === [...keys, ...groupKeys].join(), 'Native hidden grouping intake changed.');
      return integration.createCanopyExport(native, 'pdf', { signal });
    } };
  }
  const grouped = profile !== 'datatables', title = text.title;
  const input = { title, text, grouped, standalone: profile === 'canopyx' && stack === 'offset', groupKeys, columns: selectedOrder.map(index => descriptors[index]),
    headerRows, footerRows, regular: Uint8Array.from(atob(regular), value => value.charCodeAt(0)), bold: Uint8Array.from(atob(bold), value => value.charCodeAt(0)) };

  // The exact same public document composition runs in-process and in the real export worker.
  async function writeSemanticReport(api, input, rows, sink, signal, onProgress) {
    const columns = input.columns.map((column, index) => ({ ...column, value: row => column.key === 'c0'
      ? new api.ExportCell(row[index] instanceof api.ExportCell ? row[index].value : row[index], { link: { target: 'https://example.org/service/' + (row[index] instanceof api.ExportCell ? row[index].value : row[index]) } })
      : row[index] }));
    const fonts = { regular: await api.PdfFont.create(input.regular, { signal }), bold: await api.PdfFont.create(input.bold, { signal }) };
    const configuration = { fonts, signal, onProgress, tagged: { language: input.text.tag }, pageSize: 'A4', orientation: 'landscape', margins: 30, fontSize: 9,
      padding: 4, pageNumbers: true, pageHeader: input.text.pageHeader, limits: { maxRows: 500, maxCells: 20000, maxPages: 150,
        maxTableBands: 8, maxTableWindowRows: 3, maxStructureElements: 20000, maxMarkedContentReferences: 50000 } };
    const options = { columns, columnWidths: Array(columns.length).fill(84), wideTable: { mode: 'bands', minimumColumnWidth: 84, repeatColumns: [0, 1] },
      ...(input.headerRows ? { headerRows: input.headerRows, footer: { rows: input.footerRows } } : {}) };
    const groupOptions = { labelColumn: { header: input.text.groupHeader, width: 100 }, columnWidths: [100, ...Array(columns.length).fill(84)],
      wideTable: { mode: 'bands', minimumColumnWidth: 84, repeatColumns: [0, 1] },
      subtotalLabel: group => input.text.subtotal + ' ' + group.label + ' (' + group.rows + ')', grandTotalLabel: summary => input.text.grandTotal + ' (' + summary.rows + ')',
      continuationLabel: group => group.label + ' ' + input.text.continuation, formatValue: value => value == null ? '' : value instanceof Date ? value.toISOString() : String(value) };
    const report = () => api.createGroupedReport(rows, { columns, groupBy: input.groupKeys.map((key, level) => ({ key, value: row => row[columns.length + level] })),
      totals: { c6: 'sum', c8: 'max' }, signal });
    if (input.standalone) return api.writeGroupedPdfTo(report(), sink, { ...configuration, ...groupOptions, title: input.title });
    async function* blocks() {
      yield { kind: 'section', title: input.title };
      yield { kind: 'paragraph', text: [{ text: input.text.scope }, { text: input.text.documentation, link: { target: 'https://officeimo.com/report-engine/' } }] };
      if (input.grouped) yield api.pdfGroupedTable(report(), groupOptions);
      else yield { ...api.pdfTable(rows, options), caption: { text: input.text.caption } };
      yield { kind: 'section', title: input.text.complete };
      yield { kind: 'paragraph', text: input.text.completeText };
    }
    return api.writePdfDocumentTo(blocks(), sink, configuration);
  }
  let produced = 0, returned = 0, enumerations = 0, getterReads = 0, writes = 0, released;
  const chunks = [], progress = [];
  async function* rows(signal) {
    enumerations++;
    try {
      if (profile === 'hfx') for (let ordinal = 0; ordinal < expected.length; ordinal++) {
        const values = [...expected[ordinal].map((value, index) => index === 7 ? new Date(value) : value), ...groupPath(ordinal)], row = [];
        values.forEach((value, index) => Object.defineProperty(row, index, { enumerable: true, get() { getterReads++; return value; } }));
        produced++; yield row;
      } else if (profile === 'datatables') for await (const row of source.rows) { produced++; yield row; }
      else {
        queries = []; const captured = source.capture(signal);
        for await (const row of captured.rows()) { produced++; yield row.map(scalar); }
      }
    } finally { returned++; released?.(); }
  }
  const sink = { async write(chunk) { const before = produced; await new Promise(resolve => setTimeout(resolve, 1));
    check(before === produced, 'Native source advanced during byte acknowledgement.'); chunks.push(chunk.slice()); writes++; } };
  let workerState;
  try {
    let completion;
    if (profile !== 'hfx') completion = await writeSemanticReport(api, input, rows(), sink, undefined, event => progress.push(event.rows));
    else {
      const body = 'import * as api from ' + JSON.stringify(moduleBase + '/index.js') + '; import * as execution from ' + JSON.stringify(moduleBase + '/execution/index.js') + ';\n' +
        writeSemanticReport.toString() + '\nexecution.handleWorkerExports(self,({input,rows,sink,signal,onProgress})=>writeSemanticReport(api,input,rows,sink,signal,onProgress));';
      const url = URL.createObjectURL(new Blob([body], { type: 'text/javascript' })), worker = new Worker(url, { type: 'module' });
      let maxRows = 0, maxChunk = 0;
      const endpoint = { postMessage(message, transfer) { if (message.kind === 'rows') maxRows = Math.max(maxRows, message.rows.length); worker.postMessage(message, transfer); },
        addEventListener: (...args) => worker.addEventListener(...args), removeEventListener: (...args) => worker.removeEventListener(...args) };
      worker.addEventListener('message', event => { if (event.data.kind === 'output') maxChunk = Math.max(maxChunk, event.data.bytes.length); });
      try { completion = await execution.executeWorkerExport(endpoint, rows(), sink, { input, batchRows: 3, maxBatchCells: 48, maxChunkBytes: 4096, onProgress: event => progress.push(event.rows) }); }
      finally { worker.terminate(); URL.revokeObjectURL(url); }
      workerState = { maxRows, maxChunk }; check(maxRows <= 3 && maxChunk <= 4096, 'Semantic worker exceeded row/chunk bounds.');
    }
    check(enumerations === 1 && returned === 1 && produced === expected.length, 'Semantic source was replayed, truncated or retained.');
    check(profile !== 'hfx' || getterReads === expected.length * (keys.length + groupKeys.length), 'Source getters were recaptured.');
    check(profile !== 'datatables' || projectionCalls === expected.length * keys.length && dateValues === expected.length, 'DataTables scalar projection was repeated or lost date values.');
    check(progress.at(-1) === expected.length && completion.rows === expected.length, 'Semantic repetition changed logical progress.');
    check(profile !== 'canopyx' || new Set(queries).size === queries.length, 'Tagged bands replayed a native page.');
    const data = new Uint8Array(await new Blob(chunks).arrayBuffer()); let binary = '';
    for (let offset = 0; offset < data.length; offset += 8192) binary += String.fromCharCode(...data.subarray(offset, offset + 8192));
    const contract = { profile, stack, grouped, standalone: input.standalone, language: text.tag, title,
      labels: { groupHeader: text.groupHeader, subtotal: text.subtotal, grandTotal: text.grandTotal },
      aggregates: grouped ? { c6: 'sum', c8: 'max' } : {},
      columns: input.columns.map(column => ({ key: column.key, header: column.header })),
      rows: expected.map(row => row.map(display)), rawRows: expected, groupPaths: grouped ? expected.map((_, index) => groupPath(index)) : [],
      headers: headerRows, footers: footerRows, repeatColumns: [0, 1],
      links: ['https://example.org/service/', ...(input.standalone ? [] : ['https://officeimo.com/report-engine/'])],
      forbidden: ['UNSELECTED-', 'PRIVATE-ONLY', ' display'], expectedFragmentToken: 'D0004L045' };
    await writeSemanticArtifact(profile + '-' + stack + (language === 'en' ? '' : '-' + language) + '.pdf', btoa(binary), JSON.stringify(contract));
    const acceptedQueries = queries.slice();
    if (profile === 'canopyx' && stack === 'cursor') {
      const controller = new AbortController(), reason = Error('Semantic native page cancelled'); let notified;
      const pending = new Promise(resolve => { notified = resolve; }); pendingPage = notified; holdPage = true;
      const release = new Promise(resolve => { released = resolve; });
      const task = writeSemanticReport(api, input, rows(controller.signal), { write() {} }, controller.signal); let watchdog;
      try { await Promise.race([pending, new Promise((_, reject) => { watchdog = setTimeout(() => reject(Error('Native semantic query did not suspend.')), 10000); })]); }
      finally { clearTimeout(watchdog); }
      controller.abort(reason); const error = await task.then(() => null, error => error);
      check(error === reason && sourceSignal.aborted, 'Native semantic cancellation lost the original reason.');
      try { await Promise.race([release, new Promise((_, reject) => { watchdog = setTimeout(() => reject(Error('Cancelled native iterator was not returned.')), 10000); })]); }
      finally { clearTimeout(watchdog); }
      check(returned === 2, 'Cancelled native iterator was not returned once.');
    }
    return { profile, stack, rows: expected.length, columns: keys.length, levels: grouped ? 4 : 0, enumerations, returned, writes,
      getterReads, projectionCalls, dateValues, nativeQueries: acceptedQueries, worker: workerState,
      selection: profile === 'datatables' ? 'native selected + filtered + current reordered + orthogonal typed projection + span header/footer' :
        profile === 'canopyx' ? 'actual native ' + stack + ' query callback + explicit hidden four-level metadata' : 'HFX-shaped worker operational assessment' };
  } finally { table?.destroy(); grid?.destroy(); }
}
