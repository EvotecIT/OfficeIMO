import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { createWorkbook } from "../dist/xlsx/index.js";
import { readZip } from "./zip-reader.mjs";

test("typed workbook stores literal text, date serials and styles with valid ZIP payloads", async () => {
  const events = [], book = createWorkbook({ creator: "A<&\"", title: "T<>&", dateMode: "utc", onProgress: p => events.push(p) });
  const sheet = book.addSheet("'Bad[]:*?/\\'", { columns: [
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
  assert.throws(() => book.addSheet("late"), /finalized/);
});

test("empty workbook, one cell, unique names and stored fallback", async () => {
  const original = globalThis.CompressionStream;
  try {
    globalThis.CompressionStream = undefined;
    const book = createWorkbook();
    const names = [book.addSheet("  ").name, book.addSheet("sheet").name, book.addSheet("'  '").name,
      book.addSheet("History").name, book.addSheet("a".repeat(30) + "🧪").name];
    assert.deepEqual(names, ["Sheet", "sheet (2)", "Sheet (3)", "History_", "a".repeat(30)]);
    const one = book.addSheet("One", { columns: [{ header: "V" }], includeHeader: false });
    await one.addRows([["🧪"]]);
    const zip = await readZip(await book.toBlob());
    assert.ok([...zip.values()].every(e => e.method === 0));
    assert.match(zip.get("xl/worksheets/sheet6.xml").content, /r="A1"/);
    const empty = await readZip(await createWorkbook().toBlob());
    assert.match(empty.get("xl/worksheets/sheet1.xml").content, /<sheetData><\/sheetData>/);
  } finally { globalThis.CompressionStream = original; }
});

test("XLSX enforces string, width, column, date and declared-type boundaries", async () => {
  const book = createWorkbook();
  assert.throws(() => book.addSheet("Too wide", { columns: Array.from({ length: 16385 }, () => ({ header: "V" })) }), RangeError);
  assert.throws(() => book.addSheet("Width", { columns: [{ header: "V", width: 256 }] }), RangeError);
  assert.throws(() => book.addSheet("Bad header", { columns: [{ header: "x".repeat(32768) }] }), RangeError);
  assert.throws(() => book.addSheet("Bad filter", { includeHeader: false, autoFilter: true }), TypeError);
  const long = book.addSheet("Long", { columns: [{ header: "V" }] });
  await long.addRows([["x".repeat(32767)]]);
  await assert.rejects(long.addRows([["x".repeat(32768)]]), RangeError);
  await assert.rejects(book.toBlob(), RangeError);
  const typed = createWorkbook().addSheet("Typed", { columns: [{ header: "V", type: "number" }] });
  await assert.rejects(typed.addRows([["2"]]), TypeError);
  const dates = createWorkbook().addSheet("Dates", { columns: [{ header: "V" }] });
  await assert.rejects(dates.addRows([[new Date("1800-01-01T00:00:00Z")]]), RangeError);
});

test("maximum column projects the last cell to XFD", async () => {
  const book = createWorkbook({ compression: "store" });
  const cols = Array.from({ length: 16384 }, () => ({ header: "V" }));
  const sheet = book.addSheet("Wide", { columns: cols, includeHeader: false });
  const row = Array(16384).fill(null); row[16383] = "last";
  await sheet.addRows([row]);
  assert.match((await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content, /r="XFD1"[^>]*t="inlineStr"/);
});

test("XLSX rejects overlapping appends/finalization and snapshots object projection", async () => {
  const columns = [{ header: "V", key: "value" }], book = createWorkbook();
  const sheet = book.addSheet("Data", { columns }); columns[0].key = "other";
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
  const controller = new AbortController(), book = createWorkbook({ signal: controller.signal });
  const sheet = book.addSheet("Rows", { columns: [{ header: "V" }] });
  let closed = false;
  function* rows() { try { for (let i = 0; i < 100000; i++) yield ["text" + i]; } finally { closed = true; } }
  const writing = sheet.addRows(rows());
  setTimeout(() => controller.abort(), 10);
  await assert.rejects(writing, { name: "AbortError" });
  assert.equal(closed, true);
  assert.throws(() => book.toBlob(), { name: "AbortError" });
});

test("classic scripts compose in either order without a module loader", async () => {
  for (const order of [["xlsx", "csv"], ["csv", "xlsx"]]) {
    const context = vm.createContext({ TextEncoder, Blob, Date, performance, setTimeout, DOMException });
    let core;
    for (const kind of order) {
      vm.runInContext(await readFile(new URL("../bundles/officeimo-" + kind + ".js", import.meta.url), "utf8"), context);
      if (core) assert.equal(context.OfficeIMO.core, core);
      core = context.OfficeIMO.core;
    }
    assert.equal(typeof context.OfficeIMO.createWorkbook, "function");
    assert.equal(typeof context.OfficeIMO.writeCsv, "function");
    assert.equal(typeof context.OfficeIMO.saveBlob, "function");
    assert.throws(() => context.OfficeIMO.createWorkbook().addWorksheet("Pending", { mergedCells: [] }), error =>
      error instanceof context.OfficeIMO.core.NotSupportedError && error instanceof context.OfficeIMO.core.OfficeIMOError);
    vm.runInContext(await readFile(new URL("../bundles/officeimo.js", import.meta.url), "utf8"), context);
    assert.equal(context.OfficeIMO.core, core);
    assert.equal(context.OfficeIMO.Cell, context.OfficeIMO.xlsx.Cell);
  }
});

test("standalone ES module assets execute the same writer contracts", async () => {
  for (const name of ["officeimo", "officeimo-xlsx", "officeimo-csv"]) {
    const module = await import("../bundles/" + name + ".mjs");
    if (module.createWorkbook) {
      const book = module.createWorkbook({ compression: "store" });
      await book.addWorksheet("Module", { columns: [{ header: "V" }] }).addRows([["Łódź 🧪"]]);
      assert.match((await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content, /Łódź 🧪/);
    }
    if (module.writeCsv) assert.equal(await (await module.writeCsv([["=literal"]], { columns: [{ header: "V" }] })).text(), "V\r\n'=literal\r\n");
  }
});

test("an installed CompressionStream without deflate-raw support falls back to stored ZIP", async () => {
  const original = globalThis.CompressionStream;
  try {
    globalThis.CompressionStream = class { constructor() { throw new TypeError("No raw deflate support"); } };
    const book = createWorkbook();
    await book.addSheet("Stored", { columns: [{ header: "V" }] }).addRows([["fallback"]]);
    assert.ok([...(await readZip(await book.toBlob())).values()].every(e => e.method === 0));
  } finally { globalThis.CompressionStream = original; }
});
