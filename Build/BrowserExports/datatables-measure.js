// One export operation per request. PowerForge owns repetition, warmups and ordering.
let comparisonTable, comparisonSpec, comparisonBlob;
const comparisonValue = (row, column, unique, format, textProfile) => column % 3 === 0 ? row * 10 + column
  : column % 3 === 1 ? (format === 'pdf' ? `Łódź${unique ? row : row % 100}-${column}` : `Łódź${textProfile === 'bmp' ? '' : ' 🧪'} ${unique ? 'row-' + row : 'group-' + row % 100}-column-${column}`)
    : (row % 10000) + column / 100;

globalThis.prepareDataTablesMeasurement = function (spec) {
  comparisonBlob = undefined;
  if (comparisonTable) comparisonTable.destroy();
  document.body.innerHTML = '<table id="measurement"></table>';
  comparisonSpec = spec;
  if (spec.format === 'pdf') preparePdfMeasurement(spec.pdfFonts);
  const data = Array.from({ length: spec.rows }, (_, row) => Array.from({ length: spec.columns }, (_, column) => comparisonValue(row, column, spec.unique, spec.format, spec.textProfile)));
  comparisonTable = new DataTable('#measurement', { data, order: [], deferRender: true, pageLength: 10,
    autoWidth: false, columns: Array.from({ length: spec.columns }, (_, column) => ({ title: (spec.format === 'pdf' ? 'Column' : 'Column ') + (column + 1), type: column % 3 === 1 ? 'string' : 'num' })),
    layout: { topStart: null, topEnd: null, bottomStart: null, bottomEnd: null } });
  DataTable.Buttons.jszip(JSZip);
  const nativeExporters = ['excelHtml5','csvHtml5'].every(name=>typeof DataTable.ext.buttons[name]?.action==='function');
  if (!nativeExporters) throw new Error('Pinned comparison stack did not register native HTML5 exporters.');
  return { rows: comparisonTable.rows().count(), dataTables: DataTable.version, buttons: DataTable.Buttons.version, nativeExporters };
};

globalThis.runDataTablesMeasurement = async function (lane, format, diagnosticYields = false) {
  comparisonBlob = undefined;
  let gatherMs = 0, compressionMs = null, maxTimerGapMs = 0, heap = performance.memory?.usedJSHeapSize ?? null;
  let tick = performance.now();
  const sample = () => { const now = performance.now(); maxTimerGapMs = Math.max(maxTimerGapMs, now - tick); tick = now;
    if (performance.memory) heap = Math.max(heap ?? 0, performance.memory.usedJSHeapSize); };
  const timer = setInterval(sample, 16), originalGather = comparisonTable.buttons.exportData;
  const originalUrl = URL.createObjectURL, originalZip = JSZip.prototype.generateAsync;
  const originalTimeout = globalThis.setTimeout;
  const originalStrip = DataTable.Buttons.stripData, originalCells = comparisonTable.cells;
  let stripMs = 0, stripCalls = 0, renderMs = 0, indexesMs = 0;
  const timeoutDelays = [];
  if (diagnosticYields) {
    DataTable.Buttons.stripData = function (...args) {
      const start = performance.now(); try { return originalStrip.apply(this, args); }
      finally { stripMs += performance.now() - start; stripCalls++; }
    };
    comparisonTable.cells = function (...args) {
      const cells = originalCells.apply(this, args), render = cells.render, indexes = cells.indexes;
      cells.render = function (...args) { const start = performance.now(); try { return render.apply(this, args); } finally { renderMs += performance.now() - start; } };
      cells.indexes = function (...args) { const start = performance.now(); try { return indexes.apply(this, args); } finally { indexesMs += performance.now() - start; } };
      return cells;
    };
  }
  if (diagnosticYields) globalThis.setTimeout = function (callback, delay, ...args) {
    if (typeof callback !== 'function' || delay !== 0 || callback.name !== 'finish') return originalTimeout.call(this, callback, delay, ...args);
    const started = performance.now();
    return originalTimeout.call(this, function (...values) { timeoutDelays.push(performance.now() - started); return callback.apply(this, values); }, delay, ...args);
  };
  comparisonTable.buttons.exportData = function (...args) {
    const start = performance.now(); try { return originalGather.apply(this, args); } finally { gatherMs += performance.now() - start; }
  };
  let completeBlob, failBlob;
  const completedBlob = new Promise((resolve, reject) => { completeBlob = resolve; failBlob = reject; });
  const failedGeneration = event => { event.preventDefault(); failBlob(event.reason); };
  // Native PDF actions call Buttons' completion callback before asynchronous layout/encoding ends.
  URL.createObjectURL = function (blob) { comparisonBlob = blob; completeBlob(blob); return originalUrl.call(this, blob); };
  JSZip.prototype.generateAsync = function (...args) {
    const start = performance.now(); return originalZip.apply(this, args).then(blob => { compressionMs = performance.now() - start; return blob; });
  };
  const started = performance.now();
  try {
    if (lane === 'native') {
      const definition = DataTable.ext.buttons[format === 'xlsx' ? 'excelHtml5' : format === 'pdf' ? 'pdfHtml5' : 'csvHtml5'];
      const config = { ...definition, title: null, messageTop: null, messageBottom: null, filename: 'Comparison', sheetName: 'Data',
        footer: false, header: true, bom: false, newline: '\r\n', exportOptions: { modifier: { order: 'index', search: 'none', selected: null }, escapeExcelFormula: true } };
      if (format === 'xlsx' && !comparisonSpec.fullWidthScan) config.customize = workbook => {
        for (const column of workbook.xl.worksheets['sheet1.xml'].getElementsByTagName('col')) column.setAttribute('width', '20');
      };
      if (format === 'pdf') {
        config.pageSize = 'A3'; config.orientation = 'landscape'; config.download = 'download';
        config.customize = doc => customizePdfMeasurement(doc, comparisonSpec);
        window.addEventListener('unhandledrejection', failedGeneration);
      }
      await new Promise((resolve, reject) => {
        try { definition.action(null, comparisonTable, null, config, resolve); } catch (error) { reject(error); }
      });
      await completedBlob;
      if (!(comparisonBlob instanceof Blob)) throw new Error('Native export did not produce a captured Blob.');
    } else {
      const fullSizing = format === 'xlsx' && comparisonSpec.fullWidthScan, cells = comparisonSpec.rows * comparisonSpec.columns;
      comparisonBlob = await OfficeIMO.exportDataTable(DataTable, comparisonTable, format, {
        mode: lane, headings: 'leaf', includeFooter: false,
        exportOptions: { modifier: { order: 'index', search: 'none', selected: null }, escapeExcelFormula: true },
        columnOptions: fullSizing ? undefined : Object.fromEntries(Array.from({ length: comparisonSpec.columns }, (_, column) => [column, { width: 20 }])),
        sheet: { autoFilter: false, autoSize: fullSizing ? { sampleRows: comparisonSpec.rows, minWidth: 6, maxWidth: 54 } : { sampleRows: 0 } }, csv: { quote: 'all' },
        ...(fullSizing ? { limits: { maxBufferedCells: cells, maxBufferedCharacters: cells * 192 + comparisonSpec.rows * 96 } } : {}),
        ...(format === 'pdf' ? { pdf: pdfMeasurementOptions(comparisonSpec) } : {})
      });
    }
    const exportMs = performance.now() - started;
    sample(); await new Promise(resolve => setTimeout(resolve, 0)); sample();
    return { exportMs, gatherMs, compressionMs, outputBytes: comparisonBlob.size, maxTimerGapMs, peakSampledJsHeapBytes: heap,
      rows: comparisonSpec.rows, columns: comparisonSpec.columns, lane, format,
      ...(diagnosticYields ? { diagnosticYields: {
        timeouts: timeoutDelays.length, totalTimeoutMs: timeoutDelays.reduce((a, b) => a + b, 0), maxTimeoutMs: Math.max(0, ...timeoutDelays),
        schedulerPostTask: typeof globalThis.scheduler?.postTask === 'function',
        stripMs, stripCalls, renderMs, indexesMs } } : {}) };
  } finally {
    clearInterval(timer); comparisonTable.buttons.exportData = originalGather;
    URL.createObjectURL = originalUrl; JSZip.prototype.generateAsync = originalZip;
    window.removeEventListener('unhandledrejection', failedGeneration);
    globalThis.setTimeout = originalTimeout;
    DataTable.Buttons.stripData = originalStrip; comparisonTable.cells = originalCells;
  }
};

globalThis.deliverDataTablesMeasurement = async function () {
  if (!(comparisonBlob instanceof Blob)) throw new Error('No completed comparison output.');
  for (let first = 0; first < comparisonBlob.size; first += 65536) {
    const bytes = new Uint8Array(await comparisonBlob.slice(first, first + 65536).arrayBuffer()); let binary = '';
    for (let offset = 0; offset < bytes.length; offset += 8192) binary += String.fromCharCode(...bytes.subarray(offset, offset + 8192));
    await acceptComparisonChunk(btoa(binary));
  }
  comparisonBlob = undefined;
};
