// One export operation per request. PowerForge owns repetition, warmups and ordering.
let comparisonTable, comparisonSpec, comparisonBlob;
const comparisonValue = (row, column, unique) => column % 3 === 0 ? row * 10 + column
  : column % 3 === 1 ? (unique ? `Łódź 🧪 row-${row}-column-${column}` : `Łódź 🧪 group-${row % 100}-column-${column}`)
    : (row % 10000) + column / 100;

globalThis.prepareDataTablesMeasurement = function (spec) {
  comparisonBlob = undefined;
  if (comparisonTable) comparisonTable.destroy();
  document.body.innerHTML = '<table id="measurement"></table>';
  comparisonSpec = spec;
  const data = Array.from({ length: spec.rows }, (_, row) => Array.from({ length: spec.columns }, (_, column) => comparisonValue(row, column, spec.unique)));
  comparisonTable = new DataTable('#measurement', { data, order: [], deferRender: true, pageLength: 10,
    autoWidth: false, columns: Array.from({ length: spec.columns }, (_, column) => ({ title: 'Column ' + (column + 1), type: column % 3 === 1 ? 'string' : 'num' })),
    layout: { topStart: null, topEnd: null, bottomStart: null, bottomEnd: null } });
  DataTable.Buttons.jszip(JSZip);
  return { rows: comparisonTable.rows().count(), dataTables: DataTable.version, buttons: DataTable.Buttons.version };
};

globalThis.runDataTablesMeasurement = async function (lane, format) {
  comparisonBlob = undefined;
  let gatherMs = 0, compressionMs = null, maxTimerGapMs = 0, heap = performance.memory?.usedJSHeapSize ?? null;
  let tick = performance.now();
  const sample = () => { const now = performance.now(); maxTimerGapMs = Math.max(maxTimerGapMs, now - tick); tick = now;
    if (performance.memory) heap = Math.max(heap ?? 0, performance.memory.usedJSHeapSize); };
  const timer = setInterval(sample, 16), originalGather = comparisonTable.buttons.exportData;
  const originalUrl = URL.createObjectURL, originalZip = JSZip.prototype.generateAsync;
  comparisonTable.buttons.exportData = function (...args) {
    const start = performance.now(); try { return originalGather.apply(this, args); } finally { gatherMs += performance.now() - start; }
  };
  URL.createObjectURL = function (blob) { comparisonBlob = blob; return originalUrl.call(this, blob); };
  JSZip.prototype.generateAsync = function (...args) {
    const start = performance.now(); return originalZip.apply(this, args).then(blob => { compressionMs = performance.now() - start; return blob; });
  };
  const started = performance.now();
  try {
    if (lane === 'native') {
      const definition = DataTable.ext.buttons[format === 'xlsx' ? 'excelHtml5' : 'csvHtml5'];
      const config = { ...definition, title: null, messageTop: null, messageBottom: null, filename: 'Comparison', sheetName: 'Data',
        footer: false, header: true, bom: false, newline: '\r\n', exportOptions: { modifier: { order: 'index', search: 'none', selected: null }, escapeExcelFormula: true } };
      if (format === 'xlsx') config.customize = workbook => {
        for (const column of workbook.xl.worksheets['sheet1.xml'].getElementsByTagName('col')) column.setAttribute('width', '20');
      };
      await new Promise((resolve, reject) => {
        try { definition.action(null, comparisonTable, null, config, resolve); } catch (error) { reject(error); }
      });
      if (!(comparisonBlob instanceof Blob)) throw new Error('Native export did not produce a captured Blob.');
    } else {
      comparisonBlob = await OfficeIMO.exportDataTable(DataTable, comparisonTable, format, {
        mode: lane, headings: 'leaf', includeFooter: false,
        exportOptions: { modifier: { order: 'index', search: 'none', selected: null }, escapeExcelFormula: true },
        columnOptions: Object.fromEntries(Array.from({ length: comparisonSpec.columns }, (_, column) => [column, { width: 20 }])),
        sheet: { autoFilter: false, autoSize: { sampleRows: 0 } }, csv: { quote: 'all' }
      });
    }
    const exportMs = performance.now() - started;
    sample(); await new Promise(resolve => setTimeout(resolve, 0)); sample();
    return { exportMs, gatherMs, compressionMs, outputBytes: comparisonBlob.size, maxTimerGapMs, peakSampledJsHeapBytes: heap,
      rows: comparisonSpec.rows, columns: comparisonSpec.columns, lane, format };
  } finally {
    clearInterval(timer); comparisonTable.buttons.exportData = originalGather;
    URL.createObjectURL = originalUrl; JSZip.prototype.generateAsync = originalZip;
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
