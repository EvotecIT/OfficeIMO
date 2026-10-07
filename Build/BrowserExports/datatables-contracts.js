// Real DataTables interoperability proof, independent of timing/benchmark policy.
globalThis.runDataTablesContracts = async function (workerScript) {
  const O = globalThis.OfficeIMO, D = globalThis.DataTable;
  const check = (ok, message) => { if (!ok) throw new Error(message); };
  const nativeExporters = ['excelHtml5','csvHtml5'].every(name=>typeof D.ext.buttons[name]?.action==='function');
  check(nativeExporters,'pinned stack registers native HTML5 exporters');
  document.body.innerHTML = '<table id="server"></table>';
  const server = new D('#server', { serverSide: true, columns: [{ title: 'Name' }],
    ajax: (request, callback) => callback({ draw: request.draw, recordsTotal: 200, recordsFiltered: 200, data: [['loaded']] }) });
  let rejectedServer = false;
  try { O.createDataTablesExport(D, server); } catch (error) { rejectedServer = /separate full-data source/.test(error.message); }
  check(rejectedServer && O.createDataTablesExport(D, server, { serverSide: 'loaded' }).rowCount === 1, 'explicit server-side loaded scope');
  server.destroy();
  document.body.innerHTML = '<h1>OfficeIMO table export</h1><table id="report"><thead><tr><th colspan="2">Identity and amount</th><th colspan="2">State</th></tr><tr><th>Name</th><th>Amount</th><th>Status</th><th>Hidden</th></tr></thead><tfoot><tr><th>Totals</th><th>20</th><th>Checked</th><th>Internal</th></tr></tfoot></table>';
  const data = [["visible-a", 12.5, "=status", "internal"], ["excluded", 100, "Other", "internal"],
    ["visible-c", 7.5, "Łódź 🧪", "internal"], ["visible-d", 3, "Other", "internal"]];
  const table = new D('#report', { data, columns: [{ type: "string" }, { type: "num" }, { type: "string" }, { visible: false }],
    select: true, colReorder: true, order: [[1, 'asc']], layout: { topStart: null } });
  table.search('visible').draw(); table.rows([0, 2]).select();
  const options = { exportOptions: { columns: ':visible' }, columnOptions: { 1: { type: 'number', format: '0.00' } },
    project: (v, c) => c.columnIndex === 1 ? new O.ExportCell(v, { presentation: { background: 'E2F0D9' } }) : v };
  const produced = [], reports = [];
  for (const mode of ['batched', 'compatibility']) {
    const source = O.createDataTablesExport(D, table, { ...options, mode });
    const rows = []; for await (const row of source.rows) rows.push(row.map(v => v instanceof O.ExportCell ? v.value : v));
    check(JSON.stringify(rows) === JSON.stringify([['visible-c', 7.5, 'Łódź 🧪'], ['visible-a', 12.5, '=status']]), mode + ' selection/order/values');
    check(source.columns.length === 3 && source.headers.length === 2 && source.footer[0] === 'Totals', mode + ' headings/footer');
    const csv = await O.exportDataTable(D, table, 'csv', { ...options, mode });
    const text = await csv.text();
    check(text === "Identity and amount,,State\r\nName,Amount,Status\r\nvisible-c,7.5,Łódź 🧪\r\nvisible-a,12.5,'=status\r\nTotals,20,Checked\r\n", mode + ' CSV value contract');
    await deliver(mode + '.csv', csv);
    const book = await O.exportDataTable(D, table, 'xlsx', { ...options, mode, sheet: { table: { style: 'TableStyleMedium2' }, print: { repeatHeaders: true } } });
    await deliver(mode + '.xlsx', book); reports.push({ mode, rows: source.rowCount, columns: source.columns.length });
  }
  // A real column reorder and explicit selected:null must also work without relying on cell-array ordering.
  table.colReorder.order([1, 0, 2, 3]);
  const reordered = O.createDataTablesExport(D, table, { exportOptions: { columns: ':visible', modifier: { selected: null } }, headings: 'leaf' });
  const reorderedRows = []; for await (const row of reordered.rows) reorderedRows.push(row);
  check(JSON.stringify(reorderedRows) === JSON.stringify([[3, 'visible-d', 'Other'], [7.5, 'visible-c', 'Łódź 🧪'], [12.5, 'visible-a', '=status']]), 'reordered columns and selection override');
  for (const mode of ['batched','compatibility']) {
    const controller = new AbortController(); let projected = 0;
    const cancelled = O.createDataTablesExport(D, table, { mode, signal: controller.signal, headings: 'leaf',
      project: v => { projected++; controller.abort('cancelled'); return v; } });
    let cancellation; try { for await (const row of cancelled.rows) void row; } catch (error) { cancellation = error; }
    check(cancellation === 'cancelled' && projected === 1, mode + ' cancellation bounds projection');
  }
  const customized = O.createDataTablesExport(D,table,{mode:'compatibility',headings:'leaf',
    exportOptions:{customizeData:data=>data.body.reverse()}});
  const customizedRows=[]; for await(const row of customized.rows) customizedRows.push(row);
  check(customizedRows[0][0]===12.5 && customizedRows[1][0]===7.5,'synchronous customization ignores return value');
  const downloads = [], failures = [];
  O.registerDataTablesButtons(D, { filename: 'Selected report', save: async (blob, filename) => { downloads.push(filename); await deliver('button-' + filename, blob); }, onError: error => failures.push(String(error)) });
  const buttons = new D.Buttons(table, { buttons: ['officeimoExcel', 'officeimoCsv'].map(extend => ({extend,
    filename:extend==='officeimoExcel'?'*':(config,api)=>{check(api.table().node()===table.table().node() && typeof config.filename==='function','filename callback arguments');return 'Report-*';},
    officeimo:{headings:'leaf',exportOptions:{columns:':visible'}}})) });
  document.body.prepend(buttons.container()[0] ?? buttons.container());
  for (let index = 0; index < 2; index++) {
    const node = table.button(index).node(); (node[0] ?? node).click();
    const started = performance.now();
    while (downloads.length <= index || table.button(index).processing()) {
      if (failures.length) throw new Error('Registered button failed: '+failures.join('; '));
      if (performance.now()-started > 10000) throw new Error('Registered button did not finish.');
      await new Promise(resolve => setTimeout(resolve,5));
    }
  }
  check(downloads.length === 2 && failures.length === 0, 'registered buttons save and clear processing after actual clicks');
  check(downloads[0]===document.title+'.xlsx' && downloads[1]==='Report-'+document.title+'.csv','Buttons document-title filenames');
  const workerElement = document.createElement('table'); document.body.append(workerElement); workerElement.hidden=true;
  const workerTable = new D(workerElement,{data:Array.from({length:130},(_,row)=>[row,'Łódź 🧪 '+row,row/10]),order:[],
    columns:[{title:'Index',type:'num'},{title:'Text',type:'string'},{title:'Amount',type:'num'}],layout:{topStart:null,topEnd:null,bottomStart:null,bottomEnd:null}});
  const originalCells = workerTable.cells; let predicateCalls=0;
  workerTable.cells=function(selector,...args){
    const bounded=typeof selector==='function'?(...values)=>{predicateCalls++;return selector(...values);}:selector;
    return originalCells.call(this,bounded,...args);
  };
  const workers = [];
  for (const compression of ['auto','store']) {
    const predicatesBefore=predicateCalls;
    const source = O.createDataTablesExport(D, workerTable, { headings:'leaf',includeFooter:false,batchRows:32,
      project:(v,c) => c.columnIndex === 0 ? new O.ExportCell(v,{presentation:{background:'E2F0D9'}}) : v });
    const result = await runDataTablesWorker(source,workerScript,compression);
    check(!result.failure && result.produced === 130 && result.maxBatch === 64 && result.maxChunk <= 65536,'bounded multi-batch worker rows/output');
    check(predicateCalls-predicatesBefore<=source.rowCount,'projection rescanned table: '+(predicateCalls-predicatesBefore)+' predicates for '+source.rowCount+' rows');
    await deliver('worker-' + compression + '.xlsx',result.blob); workers.push({compression,produced:result.produced,maxBatch:result.maxBatch,maxChunk:result.maxChunk,selectionPredicateCalls:predicateCalls-predicatesBefore});
  }
  const cancelledWorker = await runDataTablesWorker(O.createDataTablesExport(D,workerTable,{headings:'leaf',includeFooter:false}),workerScript,'store',65);
  check(cancelledWorker.failure?.includes('worker cancelled'),'worker cancellation');
  workerTable.destroy(); workerElement.remove();
  return { dataTables: D.version, buttons: D.Buttons.version, nativeExporters, reports, downloads, produced, workers, assertions: 21 };
  async function deliver(name, blob) {
    const bytes = new Uint8Array(await blob.arrayBuffer()); let binary = '';
    for (let first = 0; first < bytes.length; first += 8192) binary += String.fromCharCode(...bytes.subarray(first, first + 8192));
    await writeFixture(name, btoa(binary)); produced.push(name);
  }
};
