import test from "node:test";
import assert from "node:assert/strict";
import { writeXlsx, writeXlsxTo, writeCsv, writeCsvTo, writePdfTo, Workbook, ExportCell, Cell } from "../dist/index.js";
import { BlobByteSink } from "../dist/core/index.js";
import { readZip } from "./zip-reader.mjs";

const rows = [{ person: { name: "Łódź 🧪" }, amount: 12.5, seen: new Date("2026-10-07T00:00:00Z") }];
const columns = [{ header: "Name", key: "name", value: row => row.person.name },
  { header: "Amount", key: "amount", type: "number", format: "0.00" },
  { header: "Seen", key: "seen", type: "date", format: "yyyy-mm-dd" }];

for (const [format, writeTo] of [["csv", writeCsvTo], ["xlsx", writeXlsxTo], ["pdf", writePdfTo]])
  test(format + " preserves consumer failures and releases the destination while producer cleanup fails or waits", { timeout: 2000 }, async () => {
    for (const cleanup of ["throw", "reject", "wait"]) {
      const failure = new Error("consumer failed"), cleanupFailure = new Error("cleanup failed"); let returned = 0, produced = 0;
      const source = { [Symbol.iterator]() { return {
        next() { produced++; return { done: false, value: ["Value"] }; },
        return() { returned++; if (cleanup === "throw") throw cleanupFailure;
          return cleanup === "reject" ? Promise.reject(cleanupFailure) : new Promise(() => {}); }
      }; } };
      const stream = new WritableStream({ write() {} });
      await assert.rejects(writeTo(source, stream, { columns: [{ header: "V", value() { throw failure; } }] }), error => error === failure);
      assert.equal(returned, 1); assert.equal(produced, 1); assert.equal(stream.locked, false);
    }
    await new Promise(resolve => setImmediate(resolve));
  });

test("one-table XLSX preflights required portable columns without reading rows or borrowing a destination", async () => {
  let read = false, acquired = false, written = false;
  const source = { [Symbol.iterator]() { read = true; return [][Symbol.iterator](); } };
  const destination = { getWriter() { acquired = true; throw new Error("borrowed"); }, write() { written = true; } };
  for (const options of [{}, { columns: null }, { columns: {} }, { columns: [{ header: "V", style: 0 }] },
    { columns: [{ header: "V" }], sheet: { headerStyle: 0 } }]) {
    await assert.rejects(writeXlsx([], options), TypeError);
    await assert.rejects(writeXlsxTo(source, destination, options), TypeError);
  }
  assert.equal(read, false); assert.equal(acquired, false); assert.equal(written, false);
  const empty = await writeXlsx([], { columns: [] }); assert.ok(empty.size > 0);
});

test("table helper presentation stays portable across static and callback style paths", async () => {
  for (const sheet of [{ alternatingRowStyle: { font: 0 } }, { title: { text: "Title", style: { fill: 0 } } },
    { footer: { style: { border: 0 } } }, { rowStyle: () => ({ numberFormat: 0 }) }, { cellStyle: () => ({ font: 0 }) }])
    await assert.rejects(writeXlsx(rows, { columns, sheet }), /Workbook-local/);
  const { Cell } = await import("../dist/xlsx/index.js");
  await assert.rejects(writeXlsx(rows, { columns, sheet: { footer: { values: [new Cell(1, 0)] } } }), TypeError);
  await assert.rejects(writeXlsx(rows, { columns: [{ header: "Amount", key: "amount", type: "custom" }], cellValueWriters: { custom: v => new Cell(v, 0) } }), TypeError);
});

for (const [format, write, writeTo] of [["csv", writeCsv, writeCsvTo], ["xlsx", writeXlsx, writeXlsxTo]])
  test(format + " rejects advanced Cells in every selected row path and releases the source/destination", async () => {
    const cell = new Cell(123, 1);
    for (const [row, selected] of [[[cell], [{ header: "Amount" }]],
      [{ amount: cell }, [{ header: "Amount", key: "amount" }]],
      [{ amount: cell }, [{ header: "Amount", value: row => row.amount }]]]) {
      let returned = 0, produced = 0, closed = 0, aborted = 0;
      function* source() { try { produced++; yield row; produced++; yield row; } finally { returned++; } }
      await assert.rejects(write(source(), { columns: selected }), TypeError);
      const destination = new WritableStream({ write() {}, close() { closed++; }, abort() { aborted++; } });
      await assert.rejects(writeTo(source(), destination, { columns: selected }), TypeError);
      assert.equal(returned, 2); assert.equal(produced, 2); assert.equal(destination.locked, false);
      assert.equal(closed, 0); assert.equal(aborted, 0);
    }
    const selected = [{ header: "Amount", value: row => row[0].amount }];
    const blob = await write([[{ amount: 123 }]], { columns: selected });
    if (format === "csv") assert.equal(await blob.text(), "Amount\r\n123\r\n");
    else assert.match((await readZip(blob)).get("xl/worksheets/sheet1.xml").content, /<v>123<\/v>/);
  });

test("advanced Cell keys and getters preserve registered styles through the shared projector", async () => {
  const book = new Workbook(), style = book.styles.add({ numberFormat: "0.000", fill: { color: "C6EFCE" } });
  let calls = 0;
  const sheet = book.addWorksheet("Styled", { columns: [{ header: "Key", key: "amount" },
    { header: "Getter", value: row => { calls++; return new Cell(row.amount.value + 1, style); } }],
    autoSize: {}, table: {}, footer: { totals: { amount: "sum" } },
    conditionalFormats: [{ type: "cellIs", range: { column: "amount" }, operator: "greaterThan", value: 0, style: { font: { bold: true } } }] });
  await sheet.addRows([{ amount: new Cell(123, style) }]);
  const archive = await readZip(await book.toBlob()), xml = archive.get("xl/worksheets/sheet1.xml").content;
  assert.equal(calls, 1);
  assert.match(xml, new RegExp('r="A2" s="' + style + '"><v>123</v>'));
  assert.match(xml, new RegExp('r="B2" s="' + style + '"><v>124</v>'));
  assert.match(xml, /<f>SUBTOTAL\(109,A2:A2\)<\/f><v>123<\/v>/);
  assert.match(xml, /sqref="A2:A2"/);
});

for (const compression of ["auto", "store"]) test("one-table XLSX shares the advanced report engine: " + compression, async () => {
  const options = { columns, compression, dateMode: "utc", sheet: { name: "Report", title: { text: "Results" },
    table: { name: "Results" }, footer: { values: ["Total"], totals: { amount: "sum" } },
    print: { repeatHeaders: true }, conditionalFormats: [{ type: "cellIs", range: { column: "amount" },
      operator: "greaterThan", value: 10, style: { fill: { color: "C6EFCE" } } }] } };
  const buffered = await readZip(await writeXlsx(rows, options));
  const sink = new BlobByteSink(), result = await writeXlsxTo(rows, sink, options), blob = sink.toBlob();
  const streamed = await readZip(blob);
  assert.deepEqual(result, { rows: 1, columns: 3, bytes: blob.size });
  for (const part of ["xl/worksheets/sheet1.xml", "xl/styles.xml", "xl/tables/table1.xml", "xl/workbook.xml"])
    assert.equal(buffered.get(part).content, streamed.get(part).content);
  const xml = streamed.get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /Łódź 🧪/); assert.match(xml, /r="B3" s="\d+"><v>12.5<\/v>/);
  assert.match(xml, /sqref="B3:B3"/); assert.match(xml, /<f>SUBTOTAL\(109,B3:B3\)<\/f><v>12.5<\/v>/);
  assert.match(streamed.get("xl/styles.xml").content, /formatCode="yyyy-mm-dd"/);
});

test("shared getters preserve portable typed values and CSV raw/display modes", async () => {
  const projected = [{ header: "Name", value: row => row.person.name },
    { header: "Amount", value: row => new ExportCell(row.amount, { text: "=12.50 USD", presentation: { background: "C6EFCE" } }), format: "0.00" }];
  assert.equal(await (await writeCsv(rows, { columns: projected })).text(), "Name,Amount\r\nŁódź 🧪,12.5\r\n");
  assert.equal(await (await writeCsv(rows, { columns: projected, valueMode: "display" })).text(), "Name,Amount\r\nŁódź 🧪,'=12.50 USD\r\n");
  const xml = (await readZip(await writeXlsx(rows, { columns: projected }))).get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /r="B2" s="\d+"><v>12.5<\/v>/);
});

test("data indexes stay zero-based across append calls and report headings", async () => {
  const getters = [], styles = [], writers = [], formatters = [];
  const c = [{ header: "Name", key: "name", groups: ["Group"], type: "recordName",
    value: (row, context) => { getters.push(context); return row.person.name; } }];
  const book = new Workbook({ cellValueWriters: { recordName: (value, context) => { writers.push(context); return value; } } });
  const sheet = book.addWorksheet("Report", { columns: c, title: { text: "Report" },
    cellStyle: context => { styles.push(context); } });
  await sheet.addRows(rows); await sheet.addRows(rows); await book.toBlob();
  for (const contexts of [getters, styles, writers])
    assert.deepEqual(contexts.map(c => [c.rowIndex, c.columnIndex, c.worksheetRow, c.sheetName]), [[0, 0, 4, "Report"], [1, 0, 5, "Report"]]);
  getters.length = 0;
  await writeCsv([...rows, ...rows], { columns: [{ ...c[0], valueFormatter: (value, context) => { formatters.push(context); return value; } }] });
  for (const contexts of [getters, formatters]) assert.deepEqual(contexts.map(c => [c.rowIndex, c.columnIndex, c.worksheetRow]), [[0, 0, undefined], [1, 0, undefined]]);
});

test("workbook progress reports global rows and explicit sheet-local counts", async () => {
  const events = [], book = new Workbook({ onProgress: event => events.push(event) });
  await book.addWorksheet("First", { columns }).addRows(rows);
  await book.addWorksheet("Second", { columns }).addRows([...rows, ...rows]);
  await book.toBlob();
  const second = events.filter(p => p.sheetName === "Second");
  assert.ok(second.length); assert.equal(second.at(-1).rows, 3); assert.equal(second.at(-1).sheetRows, 2);
  assert.equal(events.at(-1).rows, 3);
});

for (const [format, write, writeTo] of [["csv", writeCsv, writeCsvTo], ["xlsx", writeXlsx, writeXlsxTo]]) {
  test(format + " snapshots projection, handles fresh sources and counts accepted bytes", async () => {
    let returned = 0;
    const projection = [{ header: "Name", value: row => row.person.name }];
    async function* fresh() { try { projection[0].value = () => "changed"; yield rows[0]; } finally { returned++; } }
    const expected = await write(fresh(), { columns: projection });
    assert.equal(returned, 1);
    if (format === "csv") assert.equal(await expected.text(), "Name\r\nŁódź 🧪\r\n");
    else assert.match((await readZip(expected)).get("xl/worksheets/sheet1.xml").content, /Łódź 🧪/);
    const sink = new BlobByteSink(), result = await writeTo([], sink, { columns });
    assert.deepEqual(result, { rows: 0, columns: 3, bytes: sink.toBlob().size });
  });

  test(format + " borrows native streams without closing or aborting them", async () => {
    const chunks = []; let closes = 0, aborts = 0;
    const stream = new WritableStream({ write: chunk => { chunks.push(new Uint8Array(chunk)); }, close: () => { closes++; }, abort: () => { aborts++; } });
    const result = await writeTo(rows, stream, { columns });
    assert.equal(stream.locked, false); assert.equal(result.bytes, new Blob(chunks).size);
    assert.equal(closes, 0); assert.equal(aborts, 0);
    const writer = stream.getWriter(); await writer.close(); writer.releaseLock(); assert.equal(closes, 1);
    const failure = new Error("destination failed"), failing = new WritableStream({ write() { throw failure; }, abort() { aborts++; } });
    await assert.rejects(writeTo(rows, failing, { columns }), error => error === failure);
    assert.equal(failing.locked, false); assert.equal(aborts, 0);
  });

  test(format + " cancellation releases a blocked stream and closes its source", async () => {
    const controller = new AbortController(), failure = new Error("cancelled");
    let started, returned = false, produced = 0;
    const waiting = new Promise(resolve => { started = resolve; });
    const stream = new WritableStream({ write() { if (produced) { started(); return new Promise(() => {}); } } });
    function* source() { try { for (let i = 0; i < 2000; i++) { produced++; yield ["🧪".repeat(200)]; } } finally { returned = true; } }
    const pending = writeTo(source(), stream, { columns: [{ header: "Value" }], signal: controller.signal, compression: "store", sheet: { autoSize: { sampleRows: 0 } } });
    await waiting; controller.abort(failure); await assert.rejects(pending, error => error === failure);
    assert.equal(stream.locked, false); assert.equal(returned, true);
  });

  test(format + " rejects asynchronous getters and stops callback work on cancellation", async () => {
    await assert.rejects(write(rows, { columns: [{ header: "Value", value: async () => { throw new Error("async getter"); } }] }), TypeError);
    const controller = new AbortController(), failure = new Error("cancelled in getter"); let calls = 0, returned = false;
    function* source() { try { yield rows[0]; yield rows[0]; } finally { returned = true; } }
    await assert.rejects(write(source(), { signal: controller.signal, columns: [
      { header: "First", value: () => { controller.abort(failure); return 1; } },
      { header: "Second", value: () => { calls++; return 2; } }] }), error => error === failure);
    assert.equal(calls, 0); assert.equal(returned, true);
    await new Promise(resolve => setImmediate(resolve));
  });
}

test("single-table header suppression and explicit wider-array selection", async () => {
  const xml = (await readZip(await writeXlsx([[3]], { columns: [{ header: "N" }], sheet: { includeHeader: false } }))).get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /r="A1" s="\d+"><v>3<\/v>/); assert.doesNotMatch(xml, /autoFilter/);
  const selected = [{ header: "N", value: row => row[2] }];
  assert.equal(await (await writeCsv([[1, 2, 3]], { columns: selected })).text(), "N\r\n3\r\n");
  await assert.rejects(writeCsv([[1, 2, 3]], { columns: [{ header: "N" }] }), RangeError);
});
