import test from "node:test";
import assert from "node:assert/strict";
import { ExportCell } from "../dist/core/index.js";
import { createDataTablesExport, exportDataTable, writeDataTableTo, registerDataTablesButtons } from "../dist/integrations/datatables/index.js";
import { readZip } from "./zip-reader.mjs";
import { inspectPdf } from "./pdf-reader.mjs";
async function collect(source) { const rows = []; for await (const row of source) rows.push(row); return rows; }

test("DataTables rejects workbook-local style IDs before reading the source or destination", async () => {
  const { host, table } = fixture();
  let sourceRead = false, written = false;
  table.page.info = () => { sourceRead = true; throw new Error("source read"); };
  assert.throws(() => createDataTablesExport(host, table, { columnOptions: { 0: { style: 1 } } }), /Workbook-local/);
  await assert.rejects(writeDataTableTo(host, table, "xlsx", { write() { written = true; } }, { sheet: { headerStyle: 1 } }), /Workbook-local/);
  assert.equal(sourceRead, false); assert.equal(written, false);
});

test("adapter portable presentation rejects component IDs and advanced Cells in sibling paths", async () => {
  const { host, table } = fixture();
  for (const sheet of [{ alternatingRowStyle: { font: 0 } }, { title: { text: "Title", style: { fill: 0 } } },
    { footer: { style: { border: 0 } } }, { rowStyle: () => ({ numberFormat: 0 }) }, { cellStyle: () => ({ font: 0 }) }])
    await assert.rejects(exportDataTable(host, table, "xlsx", { sheet }), /Workbook-local/);
  const { Cell } = await import("../dist/xlsx/index.js");
  await assert.rejects(exportDataTable(host, table, "xlsx", { sheet: { footer: { values: [new Cell(1, 0)] } } }), TypeError);
  await assert.rejects(exportDataTable(host, table, "xlsx", { columnOptions: { 1: { type: "custom" } }, workbook: { cellValueWriters: { custom: v => new Cell(v, 0) } } }), TypeError);
});

// Contract-shaped external API: deliberately returns batch cells in a different order.
function fixture({ data = [["second", 12.5], ["first", 7.5]], selected = [1, 0], serverSide = false, grouped = true, nodeRows } = {}) {
  const calls = [], apiArray = values => ({ toArray: () => values });
  const host = { Buttons: { stripData: v => v }, ext: { buttons: {} } };
  const table = {
    page: { info: () => ({ serverSide }) },
    table: () => ({}),
    rows: selector => ({ indexes: () => apiArray(Array.isArray(selector) ? selector : [...selected]), count: () => selected.length }),
    columns: () => ({ indexes: () => apiArray([0, 1]) }),
    cells(rows, columns) {
      assert.deepEqual(rows, []); assert.deepEqual(columns, []);
      let positions = [];
      return { iterator(type, callback) { assert.equal(type, 'table'); callback(); },
        pop() { positions = []; },
        push(requested) { calls.push([...new Set(requested.map(p => p.row))]); positions = [...requested].reverse(); },
        render: () => apiArray(positions.map(p => data[p.row][p.column])),
        indexes: () => apiArray(positions), nodes: () => apiArray(nodeRows ? positions.filter(p => nodeRows.includes(p.row)) : positions.map(() => null)) };
    },
    buttons: { exportInfo: options => ({filename:(options.filename ?? "Export").replaceAll('*','Fixture title'),
      title: (options.title ?? "").replaceAll('*','Fixture title'), messageTop: options.messageTop ?? "", messageBottom: options.messageBottom ?? ""}), exportData(options) {
      const result = { header: ["Name", "Amount"], body: options.rows?.length === 0 ? [] : selected.map(row => [...data[row]]),
        footer: ["Totals", 20], footerStructure: [[{ title: "Totals", colspan: 1, rowspan: 1 }, { title: "20", colspan: 1, rowspan: 1 }]],
        headerStructure: grouped ? [[{ title: "Metrics", colspan: 2, rowspan: 1 }, null],
          [{ title: "Name", colspan: 1, rowspan: 1 }, { title: "Amount", colspan: 1, rowspan: 1 }]] : [] };
      options.customizeData?.(result); return result;
    } }
  };
  return { host, table, calls, data };
}

test("batched export preserves scope, reordered cell indexes, typed presentation, grouped headings and footer", async () => {
  const { host, table, calls } = fixture();
  const contexts = [], options = { batchRows: 1, maxBatchCells: 2, columnOptions: { 1: { type: "number", format: "0.00" } },
    project: (v, c) => { contexts.push(c); return c.sourceColumnIndex === 1 ? new ExportCell(v, { text: "=display", presentation: { background: "C6EFCE" } }) : v; } };
  const source = createDataTablesExport(host, table, options), rows = await collect(source.rows);
  assert.equal(source.rowCount, 2); assert.deepEqual(source.headers, [["Metrics", ""], ["Name", "Amount"]]);
  assert.equal(rows[0][0], "first"); assert.equal(rows[0][1].value, 7.5); assert.deepEqual(calls, [[1], [0]]);
  assert.deepEqual(contexts.map(c => [c.sourceRowIndex, c.sourceColumnIndex, c.rowIndex, c.columnIndex]), [[1, 1, 0, 1], [1, 0, 0, 0], [0, 1, 1, 1], [0, 0, 1, 0]]);
  assert.throws(() => source.rows[Symbol.asyncIterator](), /only once/);
  assert.throws(() => createDataTablesExport(host, table, { columnOptions: { 1: { value: () => "unexpected" } } }), /resolve values with project/);
  const zip = await readZip(await exportDataTable(host, table, "xlsx", options));
  assert.match(zip.get("xl/worksheets/sheet1.xml").content, /<mergeCell ref="A1:B1"/);
  assert.match(zip.get("xl/worksheets/sheet1.xml").content, /<c r="B3" s="\d+"><v>7.5<\/v>/);
  assert.match(zip.get("xl/styles.xml").content, /C6EFCE/);
  const csv = await exportDataTable(host, table, "csv", { ...options, csv: { valueMode: "display" } });
  assert.equal(await csv.text(), "Metrics,\r\nName,Amount\r\nfirst,'=display\r\nsecond,'=display\r\nTotals,20\r\n");
});

test("compatibility delegates whole-matrix customization and rejects ambiguous source-index projection", async () => {
  const { host, table, calls } = fixture();
  const exportOptions = { customizeData: data => { data.body.reverse(); data.body.push(["extra", 3]); } };
  const source = createDataTablesExport(host, table, { mode: "compatibility", exportOptions });
  assert.deepEqual(await collect(source.rows), [["second", 12.5], ["first", 7.5], ["extra", 3]]);
  assert.equal(source.rowCount, 3); assert.equal(calls.length, 0);
  assert.throws(() => createDataTablesExport(host, table, { exportOptions }), /compatibility mode/);
  assert.throws(() => createDataTablesExport(host, table, { mode: "compatibility", exportOptions, project: v => v }), /source indexes/);
  const reversed = createDataTablesExport(host, table, { mode: "compatibility", exportOptions: { customizeData: data => data.body.reverse() } });
  assert.deepEqual(await collect(reversed.rows), [["second", 12.5], ["first", 7.5]]);
  assert.throws(() => createDataTablesExport(host, table, { mode: "compatibility", exportOptions: { customizeData: async () => { throw new Error("async customization"); } } }), /synchronous/);
  await new Promise(resolve => setTimeout(resolve, 0));
});

test("projection budgets, preflight row limits and explicit server-side scope fail before output", async () => {
  const { host, table } = fixture({ serverSide: true });
  assert.throws(() => createDataTablesExport(host, table), /separate full-data source/);
  assert.equal(createDataTablesExport(host, table, { serverSide: "loaded" }).rowCount, 2);
  assert.throws(() => createDataTablesExport(host, table, { serverSide: "loaded", maxBatchCells: 1 }), /selected row/);
  assert.throws(() => createDataTablesExport(host, table, { serverSide: "loaded", limits: { maxRows: 1 } }), /maxRows/);
  assert.throws(() => createDataTablesExport(host, table, { serverSide: "loaded", batchRows: 0 }), /batchRows/);
  const cells = table.cells;
  table.cells = (...args) => ({ ...cells(...args), iterator(_type, callback) { callback(); callback(); } });
  assert.throws(() => createDataTablesExport(host, table, { serverSide: "loaded" }), /exactly one/);
});

test("CSV row limits count data rows and progress excludes grouped headings and footer", async () => {
  const { host, table } = fixture(); const progress = [], chunks = [];
  const result = await writeDataTableTo(host, table, "csv", { write: bytes => { chunks.push(bytes.slice()); } },
    { limits: { maxRows: 2 }, onProgress: p => progress.push(p) });
  assert.equal(result.rows, 2); assert.equal(result.bytes, Buffer.concat(chunks).length);
  assert.equal(progress.at(-1).phase, "complete"); assert.equal(progress.at(-1).rows, 2);
  await assert.rejects(writeDataTableTo(host, table, "csv", { write() {} }, { limits: { maxOutputBytes: 1 } }), /maxOutputBytes/);
});

test("cancellation stops the next projection batch and preserves the caller's reason", async () => {
  const { host, table, calls } = fixture(); const controller = new AbortController(), reason = new Error("cancel export");
  const source = createDataTablesExport(host, table, { batchRows: 1, signal: controller.signal });
  const iterator = source.rows[Symbol.asyncIterator](); assert.equal((await iterator.next()).value[0], "first");
  controller.abort(reason); await assert.rejects(iterator.next(), error => error === reason); assert.equal(calls.length, 1);
  await assert.rejects(exportDataTable(host, table, "xlsx", { signal: controller.signal }), error => error === reason);
  for (const mode of ["batched", "compatibility"]) {
    const cancelled = new AbortController(); let projected = 0;
    const source = createDataTablesExport(host, table, { mode, signal: cancelled.signal,
      project: v => { projected++; cancelled.abort(reason); return v; } });
    await assert.rejects(collect(source.rows), error => error === reason);
    assert.equal(projected, 1, mode + " must stop before the next projection");
  }
});

test("body formatting maps sparse DOM nodes and rejects async callback results", async () => {
  const deferred = fixture({nodeRows:[0]}), observed=[];
  const source = createDataTablesExport(deferred.host,deferred.table,{exportOptions:{format:{body:(v,row,column,node)=>{
    assert.deepEqual(node,row===0?{row,column}:undefined);observed.push({row,column,node});return v;
  }}}});
  assert.deepEqual(await collect(source.rows),[['first',7.5],['second',12.5]]);
  assert.equal(observed.length,4);
  const { host, table } = fixture();
  for (const options of [{ project: async () => { throw new Error("async project"); } },
    { exportOptions: { format: { body: async () => { throw new Error("async format"); } } } }])
    await assert.rejects(exportDataTable(host, table, "xlsx", options), /synchronous/);
  await new Promise(resolve => setTimeout(resolve, 0));
});

test("unsupported heading/footer shapes are rejected and explicit leaf/omission choices work", () => {
  const { host, table } = fixture(); const native = table.buttons.exportData;
  table.buttons.exportData = options => { const data = native(options); data.headerStructure[0][0].rowspan = 2; return data; };
  assert.throws(() => createDataTablesExport(host, table), /vertical spans/);
  assert.equal(createDataTablesExport(host, table, { headings: "leaf" }).headers.length, 1);
  table.buttons.exportData = options => { const data = native(options); data.footerStructure.push([{}, {}]); return data; };
  assert.throws(() => createDataTablesExport(host, table), /single footer row/);
  assert.equal(createDataTablesExport(host, table, { includeFooter: false }).footer, undefined);
  table.buttons.exportData = options => { const data = native(options); data.headerStructure[0] = [
    { title: "Same", colspan: 1, rowspan: 1 }, { title: "Same", colspan: 1, rowspan: 1 }]; return data; };
  assert.throws(() => createDataTablesExport(host, table), /equal group paths/);
  table.buttons.exportData = options => { const data = native(options); data.headerStructure = [[{ title: "Both", colspan: 2, rowspan: 1 }, null]]; return data; };
  assert.throws(() => createDataTablesExport(host, table), /one leaf header/);
  table.buttons.exportData = options => { const data = native(options); data.footerStructure = [[{ title: "Total", colspan: 2, rowspan: 1 }, null]]; return data; };
  assert.throws(() => createDataTablesExport(host, table), /one leaf header/);
});

test("buttons complete once after saving and before asynchronous error reporting", async () => {
  const { host, table } = fixture(); const events = [];
  registerDataTablesButtons(host, { filename: "Report", save: async (_blob, name) => { events.push(name); }, onError: async error => { events.push(error.message); } });
  async function action(config) {
    await new Promise(resolve => { host.ext.buttons.officeimoCsv.action(null, table, null, config, () => { events.push("done"); resolve(); }); });
    await new Promise(resolve => setTimeout(resolve, 0));
  }
  await action({}); assert.deepEqual(events, ["Report.csv", "done"]); events.length = 0;
  await action({filename:'*'}); assert.deepEqual(events,['Fixture title.csv','done']); events.length=0;
  await action({filename:'Report.csv'}); assert.deepEqual(events,['Report.csv','done']); events.length=0;
  await action({filename:(config,api)=> { assert.equal(api,table);assert.equal(typeof config.filename,'function');return 'Report-*'; }});
  assert.deepEqual(events,['Report-Fixture title.csv','done']); events.length=0;
  await action({ customize() {} }); assert.equal(events[0], "done"); assert.match(events[1], /customize/);
  events.length = 0;
  await action({ get officeimo() { throw new Error('configuration getter'); } });
  assert.deepEqual(events, ['done','configuration getter']);
  events.length = 0;
  await action({officeimo:{filename:async () => { throw new Error('async filename'); }}});
  assert.equal(events[0],'done'); assert.match(events[1],/synchronous/);
  assert.throws(() => registerDataTablesButtons(host), /already registered/);
});

test("DataTables PDF uses the shared projection, grouped headings, footer and button delivery", async () => {
  const { host, table } = fixture();
  const blob = await exportDataTable(host, table, "pdf", { pdf: { title: "Report", compression: false } });
  const pdf = await inspectPdf(blob);
  for (const value of ["Report", "Metrics", "Name", "Amount", "first", "second", "7.5", "12.5", "Totals", "20"]) assert.ok(pdf.text.includes(value), value);
  const events = []; let delivered;
  registerDataTablesButtons(host, { filename: "Report", save: async (file, name) => { delivered = file; events.push(name); }, onError: e => { throw e; } });
  await new Promise(resolve => host.ext.buttons.officeimoPdf.action(null, table, null,
    { title: "*", messageTop: "Above", messageBottom: "Below", orientation: "landscape", pageSize: "LETTER", header: false, footer: false }, () => { events.push("done"); resolve(); }));
  assert.deepEqual(events, ["Report.pdf", "done"]);
  const buttonPdf = await inspectPdf(delivered);
  for (const value of ["Fixture title", "Above", "Below", "first", "second"]) assert.ok(buttonPdf.text.includes(value), value);
  assert.ok(!buttonPdf.text.includes("Metrics")); assert.ok(!buttonPdf.text.includes("Totals"));
  assert.match(buttonPdf.pages[0].body, /MediaBox \[0 0 792 612\]/);
});

test("PDF button metadata inherits omitted values and clears explicit native null overrides", async () => {
  const { host, table } = fixture(); let delivered, failure;
  const defaults = { title: "Default title", messageTop: "Default above", messageBottom: "Default below", compression: false };
  registerDataTablesButtons(host, { pdf: defaults, save: file => { delivered = file; }, onError: error => { failure = error; } });
  async function action(configuration) {
    delivered = undefined; failure = undefined;
    await new Promise(resolve => host.ext.buttons.officeimoPdf.action(null, table, null, configuration, resolve));
    assert.equal(failure, undefined); assert.ok(delivered);
    return (await inspectPdf(delivered)).text;
  }
  const keys = ["title", "messageTop", "messageBottom"];
  const inherited = await action({});
  for (const key of keys) assert.ok(inherited.includes(defaults[key]), key);
  for (const cleared of keys) {
    const text = await action({ [cleared]: null });
    for (const key of keys) assert.equal(text.includes(defaults[key]), key !== cleared, cleared + " / " + key);
    assert.ok(text.includes("first")); assert.ok(text.includes("second"));
  }
  const omitted = await action({ title: null, messageTop: null, messageBottom: null });
  for (const key of keys) assert.ok(!omitted.includes(defaults[key]), key);
  const replaced = await action({ title: "*", messageTop: "", messageBottom: "New below" });
  assert.ok(replaced.includes("Fixture title")); assert.ok(replaced.includes("New below"));
  for (const key of keys) assert.ok(!replaced.includes(defaults[key]), key);
  assert.deepEqual(defaults, { title: "Default title", messageTop: "Default above", messageBottom: "Default below", compression: false });
});

test("PDF preserves DataTables vertical headers, blank spans and multiple footer rows", async () => {
  const { host, table } = fixture(); const native = table.buttons.exportData;
  table.buttons.exportData = options => {
    const data = native(options);
    data.headerStructure = [[{ title: "Name", colspan: 1, rowspan: 2 }, { title: "", colspan: 1, rowspan: 1 }], [null, { title: "Amount", colspan: 1, rowspan: 1 }]];
    data.footerStructure = [[{ title: "Totals", colspan: 1, rowspan: 2 }, { title: "20", colspan: 1, rowspan: 1 }], [null, { title: "Approved", colspan: 1, rowspan: 1 }]];
    return data;
  };
  const pdf = await inspectPdf(await exportDataTable(host, table, "pdf", { columnOptions: { 0: { header: "Customer" } }, pdf: { compression: false } }));
  for (const text of ["Customer", "Amount", "first", "second", "Totals", "20", "Approved"]) assert.ok(pdf.text.includes(text), text);
  await assert.rejects(exportDataTable(host, table, "csv", { headings: "structured" }), /require PDF/);
  table.buttons.exportData = options => { const data = native(options); data.headerStructure = [[{ title: "Both", colspan: 2, rowspan: 1 }, null]]; return data; };
  await assert.rejects(exportDataTable(host, table, "pdf", { columnOptions: { 1: { header: "Override" } } }), /spanning leaf/);
});
