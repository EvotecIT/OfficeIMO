import test from "node:test";
import assert from "node:assert/strict";
import { ExportCell, pause } from "../dist/core/index.js";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { Workbook } from "../dist/xlsx/index.js";
import { readZip } from "./zip-reader.mjs";

for (const scenario of ["small appends", "full width sample"]) {
  test("Excel " + scenario + " lets task cancellation stop the source before all rows are accepted", async () => {
    const original = Object.getOwnPropertyDescriptor(performance, "now");
    let clock = performance.now() + 1000, produced = 0, returned = false;
    Object.defineProperty(performance, "now", { configurable: true, value: () => clock });
    const controller = new AbortController(), reason = new Error("task cancellation"), total = 200;
    let timer, book;
    try {
      await pause();
      book = new Workbook({ compression: "store", signal: controller.signal, sink: { write() {} },
        limits: { maxBufferedCells: total, maxBufferedCharacters: total * 256 } });
      const sheet = book.addWorksheet("Data", { columns: [{ header: "ID", value: row => { clock += 5; return row[0]; } }],
        autoSize: { sampleRows: scenario === "full width sample" ? total : 0 } });
      timer = setTimeout(() => controller.abort(reason), 0);
      await assert.rejects(async () => {
        if (scenario === "small appends") {
          for (let i = 0; i < total; i++) { produced++; await sheet.addRows([[i]]); }
        } else {
          function* rows() { try { for (let i = 0; i < total; i++) { produced++; yield [i]; } } finally { returned = true; } }
          await sheet.addRows(rows());
        }
        await book.finish();
      }, error => error === reason);
      assert.ok(produced < total, "timer cancellation must interrupt the append/sample work");
      if (scenario === "full width sample") assert.equal(returned, true);
    } finally {
      clearTimeout(timer); await book?.discard(reason);
      if (original) Object.defineProperty(performance, "now", original); else delete performance.now;
      await pause();
    }
  });
}

test("typed workbook stores literal text, date serials and styles with valid ZIP payloads", async () => {
  const events = [], book = new Workbook({ creator: "A<&\"", title: "T<>&", dateMode: "utc", onProgress: p => events.push(p) });
  const sheet = book.addWorksheet("'Bad[]:*?/\\'", { columns: [
    { header: "Text", width: 28, wrapText: true, alignment: "left" },
    { header: "Number", type: "number", format: "0.00" },
    { header: "Date", type: "date", format: "yyyy-mm-dd hh:mm" },
    { header: "Bool", type: "boolean" }, { header: "Empty" }
  ], freezeHeader: true, autoFilter: true, headerFill: "#D9E1F2" });
  assert.equal(sheet.name, "Bad_______");
  await sheet.addRows([["A<&\"'\r\n\t🧪שלום\u0001\ud800_x0041_", -1.25, new Date("1900-01-01T00:00:00Z"), true, null]]);
  await sheet.addRows([["=literal", NaN, new Date("1900-02-28T12:00:00Z"), false, undefined],
    [" ", Infinity, new Date("1900-03-01T00:00:00Z"), true, null]]);
  const blob = await book.toBlob(), zip = await readZip(blob);
  assert.equal(await book.toBlob(), blob);
  const xml = zip.get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /&#13;&#10;&#9;🧪שלום_<\/t><\/r><r><t xml:space="preserve">x0041_/);
  assert.match(xml, /ySplit="1"/); assert.match(xml, /autoFilter ref="A1:E4"/);
  assert.match(xml, /width="28"/); assert.match(xml, /<v>1<\/v>/);
  assert.match(xml, /<v>59.5<\/v>/); assert.match(xml, /<v>61<\/v>/);
  assert.doesNotMatch(xml, /NaN|Infinity|\u0001|\ud800|<f>/u);
  assert.match(zip.get("xl/styles.xml").content, /formatCode="yyyy-mm-dd hh:mm"/);
  assert.match(zip.get("xl/styles.xml").content, /rgb="FFD9E1F2"/);
  assert.equal(events.at(-1).rows, 3);
  assert.throws(() => book.addWorksheet("late"), /finalized/);
});

test("empty workbook, one cell, unique names and stored fallback", async () => {
  const original = globalThis.CompressionStream;
  try {
    globalThis.CompressionStream = undefined;
    const book = new Workbook();
    const names = [book.addWorksheet("  ").name, book.addWorksheet("sheet").name, book.addWorksheet("'  '").name,
      book.addWorksheet("History").name, book.addWorksheet("a".repeat(30) + "🧪").name];
    assert.deepEqual(names, ["Sheet", "sheet (2)", "Sheet (3)", "History_", "a".repeat(30)]);
    const one = book.addWorksheet("One", { columns: [{ header: "V" }], includeHeader: false });
    await one.addRows([["🧪"]]);
    const zip = await readZip(await book.toBlob());
    assert.ok([...zip.values()].every(e => e.method === 0));
    assert.match(zip.get("xl/worksheets/sheet6.xml").content, /r="A1"/);
    const empty = await readZip(await new Workbook().toBlob());
    assert.match(empty.get("xl/worksheets/sheet1.xml").content, /<sheetData><\/sheetData>/);
  } finally { globalThis.CompressionStream = original; }
});

test("OOXML attributes preserve literal escapes and validate the decoded worksheet names", async () => {
  const book = new Workbook({ compression: "store" });
  const requests = ["_x0041_", "A", "_x003A_", "_x003a_", "Bad:Name", "Bad_Name", "_x005F_x0041_"];
  const expected = ["_x0041_", "A", "_x003A_", "_x003a_ (2)", "Bad_Name", "Bad_Name (2)", "_x005F_x0041_"];
  for (const name of requests) book.addWorksheet(name);
  book.styles.add({ font: { name: "Font_x0041_" }, numberFormat: '"_x003A_"0' });
  const zip = await readZip(await book.toBlob());
  const decode = text => text.replace(/_x([0-9a-f]{4})_/gi, (_, hex) => String.fromCharCode(parseInt(hex, 16)));
  const names = [...zip.get("xl/workbook.xml").content.matchAll(/<sheet name="([^"]+)"/g)].map(match => decode(match[1]));
  assert.deepEqual(names, expected);
  assert.deepEqual(book.worksheets.map(sheet => sheet.name), expected);
  assert.equal(new Set(names.map(name => name.toLowerCase())).size, names.length);
  assert.ok(names.every(name => name.length <= 31 && !/[\[\]:*?/\\]/.test(name)));
  const styles = zip.get("xl/styles.xml").content;
  assert.match(styles, /name val="Font_x005F_x0041_"/);
  assert.match(styles, /formatCode="&quot;_x005F_x003A_&quot;0"/);
});

test("XLSX enforces string, width, column, date and declared-type boundaries", async () => {
  const book = new Workbook();
  assert.throws(() => book.addWorksheet("Too wide", { columns: Array.from({ length: 16385 }, () => ({ header: "V" })) }), RangeError);
  assert.throws(() => book.addWorksheet("Width", { columns: [{ header: "V", width: 256 }] }), RangeError);
  assert.throws(() => book.addWorksheet("Bad header", { columns: [{ header: "x".repeat(32768) }] }), RangeError);
  assert.throws(() => book.addWorksheet("Bad filter", { includeHeader: false, autoFilter: true }), TypeError);
  const long = book.addWorksheet("Long", { columns: [{ header: "V" }] });
  await long.addRows([["x".repeat(32767)]]);
  await assert.rejects(long.addRows([["x".repeat(32768)]]), RangeError);
  await assert.rejects(book.toBlob(), RangeError);
  const typed = new Workbook().addWorksheet("Typed", { columns: [{ header: "V", type: "number" }] });
  await assert.rejects(typed.addRows([["2"]]), TypeError);
  const dates = new Workbook().addWorksheet("Dates", { columns: [{ header: "V" }] });
  await assert.rejects(dates.addRows([[new Date("1800-01-01T00:00:00Z")]]), RangeError);
});

test("maximum column projects the last cell to XFD", async () => {
  const book = new Workbook({ compression: "store" });
  const cols = Array.from({ length: 16384 }, () => ({ header: "V" }));
  const sheet = book.addWorksheet("Wide", { columns: cols, includeHeader: false });
  const row = Array(16384).fill(null); row[16383] = "last";
  await sheet.addRows([row]);
  assert.match((await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content, /r="XFD1"[^>]*t="inlineStr"/);
});

test("XLSX rejects overlapping appends/finalization and snapshots object projection", async () => {
  const columns = [{ header: "V", key: "value" }], book = new Workbook();
  const sheet = book.addWorksheet("Data", { columns }); columns[0].key = "other";
  let release;
  async function* rows() { await new Promise(r => { release = r; }); yield { value: "kept" }; }
  const writing = sheet.addRows(rows());
  await assert.rejects(sheet.addRows([]), /Await/);
  assert.throws(() => book.toBlob(), /Await/);
  while (!release) await new Promise(r => setTimeout(r, 0));
  release(); await writing;
  assert.match((await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content, /kept/);
});

test("XLSX cancels mid-stream and never finalizes a partial workbook", async () => {
  const controller = new AbortController(), book = new Workbook({ signal: controller.signal });
  const sheet = book.addWorksheet("Rows", { columns: [{ header: "V" }] });
  let closed = false;
  function* rows() { try { for (let i = 0; i < 100000; i++) yield ["text" + i]; } finally { closed = true; } }
  const writing = sheet.addRows(rows());
  setTimeout(() => controller.abort(), 10);
  await assert.rejects(writing, { name: "AbortError" });
  assert.equal(closed, true);
  await assert.rejects(book.toBlob(), { name: "AbortError" });
});

test("classic scripts compose in either order without a module loader", async () => {
  for (const order of [["datatables", "xlsx", "csv", "pdf"], ["pdf", "csv", "xlsx", "datatables"]]) {
    const context = vm.createContext({ TextEncoder, Blob, Date, performance, setTimeout, DOMException });
    let core;
    for (const kind of order) {
      vm.runInContext(await readFile(new URL("../bundles/officeimo-" + kind + ".js", import.meta.url), "utf8"), context);
      if (core) assert.equal(context.OfficeIMO.core, core);
      core = context.OfficeIMO.core;
    }
    assert.equal(typeof context.OfficeIMO.Workbook, "function");
    assert.equal(typeof context.OfficeIMO.writeCsv, "function");
    assert.equal(typeof context.OfficeIMO.writePdf, "function");
    assert.equal(typeof context.OfficeIMO.saveBlob, "function");
    assert.equal(typeof context.OfficeIMO.registerDataTablesButtons, "function");
    assert.throws(() => new context.OfficeIMO.Workbook().addWorksheet("Pending", { dataValidation: [] }), error =>
      error instanceof context.OfficeIMO.core.NotSupportedError && error instanceof context.OfficeIMO.core.OfficeIMOError);
    vm.runInContext(await readFile(new URL("../bundles/officeimo.js", import.meta.url), "utf8"), context);
    assert.equal(context.OfficeIMO.core, core);
    assert.equal(context.OfficeIMO.Cell, context.OfficeIMO.xlsx.Cell);
  }
});

test("standalone ES module assets execute the same writer contracts", async () => {
  for (const name of ["officeimo", "officeimo-xlsx", "officeimo-csv", "officeimo-pdf"]) {
    const module = await import("../bundles/" + name + ".mjs");
    if (module.Workbook) {
      const book = new module.Workbook({ compression: "store" });
      await book.addWorksheet("Module", { columns: [{ header: "V" }] }).addRows([["Łódź 🧪"]]);
      assert.match((await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content, /Łódź 🧪/);
    }
    if (module.writeCsv) assert.equal(await (await module.writeCsv([["=literal"]], { columns: [{ header: "V" }] })).text(), "V\r\n'=literal\r\n");
  }
});

test("automatic width sampling fits wide tables within cell and encoded text budgets", async () => {
  const columns = Array.from({ length: 1001 }, (_, i) => ({ header: "C" + i }));
  const book = new Workbook({ compression: "store" });
  await book.addWorksheet("Wide", { columns, autoSize: {} }).addRows(Array.from({ length: 100 }, () => columns.map(() => 1)));
  const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
  assert.equal([...xml.matchAll(/<row /g)].length, 101);
  assert.match(xml, /r="ALM101"[^>]*><v>1<\/v>/);
});

test("automatic sampling can flush mid-row without repeating callbacks, links or totals", async () => {
  for (const limits of [{ maxBufferedCells: 1 }, { maxBufferedCharacters: 50 }]) {
    let calls = 0;
    const book = new Workbook({ compression: "store", limits, cellValueWriters: { amount: v => { calls++; return new ExportCell(v, { presentation: { background: "C6EFCE" } }); } } });
    const sheet = book.addWorksheet("Budget", { columns: [{ header: "Name" }, { header: "Amount", key: "amount", type: "amount", format: "0.00" }],
      autoSize: {}, footer: { totals: { amount: "sum" } }, hyperlinks: [{ cell: "A2", target: "https://example.com" }] });
    await sheet.addRows([["Łódź 🧪", 12.5]]); await sheet.addRows([["second", 7.5]]);
    const parts = await readZip(await book.toBlob()), xml = parts.get("xl/worksheets/sheet1.xml").content;
    assert.equal(calls, 2); assert.match(xml, /Łódź 🧪/); assert.match(xml, /<f>SUBTOTAL\(109,B2:B3\)<\/f><v>20<\/v>/);
    assert.equal([...xml.matchAll(/<hyperlink /g)].length, 1); assert.equal([...xml.matchAll(/<row /g)].length, 4);
    assert.match(parts.get("xl/styles.xml").content, /formatCode="0.00"/);
  }
  for (const limits of [{ maxBufferedCells: 1 }, { maxBufferedCharacters: 50 }]) {
    const book = new Workbook({ limits });
    await assert.rejects(book.addWorksheet("Explicit", { columns: [{ header: "A" }, { header: "B" }], autoSize: { sampleRows: 2 } }).addRows([["value", 1]]), { code: "RESOURCE_LIMIT" });
  }
});

test("an installed CompressionStream without deflate-raw support falls back to stored ZIP", async () => {
  const original = globalThis.CompressionStream;
  try {
    globalThis.CompressionStream = class { constructor() { throw new TypeError("No raw deflate support"); } };
    const book = new Workbook();
    await book.addWorksheet("Stored", { columns: [{ header: "V" }] }).addRows([["fallback"]]);
    assert.ok([...(await readZip(await book.toBlob())).values()].every(e => e.method === 0));
  } finally { globalThis.CompressionStream = original; }
});
