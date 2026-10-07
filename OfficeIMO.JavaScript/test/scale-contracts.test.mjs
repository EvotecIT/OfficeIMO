import test from "node:test";
import assert from "node:assert/strict";
import { Workbook, StyleRegistry } from "../dist/xlsx/index.js";
import { writeCsvTo } from "../dist/csv/index.js";
import { readZip } from "./zip-reader.mjs";
import { ZipWriter } from "../dist/zip/index.js";
import { OpcPackage, relationshipTypes } from "../dist/opc/index.js";
import { BlobByteSink } from "../dist/core/index.js";

test("appendable ZIP and OPC parts retain descriptors, catalogs and sequential ownership", async () => {
  const sink = new BlobByteSink(), zip = new ZipWriter(sink, { compression: "store" });
  const first = await zip.openEntry("first.txt"); await first.write(new TextEncoder().encode("first"));
  await assert.rejects(zip.openEntry("overlap.txt"), { code: "INVALID_STATE" });
  await assert.rejects(zip.finish(), { code: "INVALID_STATE" });
  await first.close(); await first.close(); await zip.add("second.txt", new TextEncoder().encode("second")); await zip.finish();
  const entries = await readZip(sink.toBlob()); assert.equal(entries.get("first.txt").content, "first"); assert.equal(entries.get("second.txt").content, "second");
  const output = new BlobByteSink(), opc = new OpcPackage({ sink: output });
  const part = await opc.openPart("/data.txt", "text/plain"); await part.write(new TextEncoder().encode("data")); await part.close();
  opc.addRelationship("/", { id: "document", type: relationshipTypes.officeDocument, target: "/data.txt" });
  const length = await opc.finish(), blob = output.toBlob(); assert.equal(blob.size, length);
  const packaged = await readZip(blob); assert.equal(packaged.get("data.txt").content, "data");
  assert.match(packaged.get("[Content_Types].xml").content, /PartName="\/data.txt" ContentType="text\/plain"/);
  assert.match(packaged.get("_rels/.rels").content, /Target="data.txt"/);
});

test("XLSX emits rows before input completes, awaits its sink and finishes sequential worksheets", async () => {
  const chunks = [];
  let produced = 0, release, blocked;
  const waiting = new Promise(resolve => { blocked = resolve; });
  const gate = new Promise(resolve => { release = resolve; });
  const decoder = new TextDecoder();
  let held = false;
  const sink = { async write(bytes) {
    // Hold actual row output, rather than a ZIP header whose flush timing can vary.
    if (!held && decoder.decode(bytes).includes('r="A2"')) { held = true; blocked(); await gate; }
    chunks.push(new Uint8Array(bytes));
  } };
  const book = new Workbook({ sink, compression: "store" });
  const first = book.addWorksheet("First", { columns: [{ header: "Value" }] });
  function* rows() { for (let i = 0; i < 1000; i++) { produced++; yield ["Łódź 🧪 ".repeat(80) + i]; } }
  const append = first.addRows(rows());
  await waiting;
  assert.ok(produced > 0 && produced < 1000, "backpressure must reach the source before all rows are retained");
  release(); await append;
  await book.addWorksheet("Second", { columns: [{ header: "Number", type: "number" }] }).addRows([[125.75]]);
  await assert.rejects(first.addRows([["late"]]), /closed/);
  const result = await book.finish();
  assert.equal(result.rows, 1001); assert.equal(result.sheets, 2);
  assert.equal(result.bytes, chunks.reduce((sum, chunk) => sum + chunk.length, 0));
  assert.deepEqual(await book.finish(), result);
  assert.throws(() => book.toBlob(), /caller-owned/);
  const zip = await readZip(new Blob(chunks));
  assert.match(zip.get("xl/worksheets/sheet1.xml").content, /r="A1001"/);
  assert.match(zip.get("xl/worksheets/sheet2.xml").content, /<v>125.75<\/v>/);
  assert.match(zip.get("xl/workbook.xml").content, /name="Second"/);
});

test("style component caching keeps equivalent definitions stable and snapshots mutable inputs", () => {
  const registry = new StyleRegistry();
  const patch = { fill: { color: "#ff0000" }, font: { bold: true } };
  const first = registry.add(patch);
  for (let i = 0; i < 1000; i++) assert.equal(registry.add({ font: { bold: true }, fill: { color: "FFFF0000" } }), first);
  patch.fill.color = "00FF00";
  const second = registry.add(patch);
  assert.notEqual(first, second);
  const xml = registry.toXml();
  assert.match(xml, /<fonts count="2"/); assert.match(xml, /<fills count="4"/);
  assert.match(xml, /FFFF0000/); assert.match(xml, /FF00FF00/);
});

test("XLSX and CSV resource ceilings stop their sources and reject oversized output chunks", async () => {
  for (const format of ["xlsx", "csv"]) {
    let returned = false;
    function* rows() { try { yield [1]; yield [2]; yield [3]; } finally { returned = true; } }
    const columns = [{ header: "Value", type: "number" }];
    if (format === "xlsx") {
      const book = new Workbook({ limits: { maxRows: 2 } });
      await assert.rejects(book.addWorksheet("Data", { columns }).addRows(rows()), { code: "RESOURCE_LIMIT" });
      await assert.rejects(book.toBlob(), { code: "RESOURCE_LIMIT" });
    } else await assert.rejects(writeCsvTo(rows(), { write() {} }, { columns, limits: { maxRows: 2 } }), { code: "RESOURCE_LIMIT" });
    assert.ok(returned);
  }
  let accepted = 0;
  const sink = { write(bytes) { accepted += bytes.length; } };
  const book = new Workbook({ sink, compression: "store", limits: { maxOutputBytes: 100 } });
  await assert.rejects(book.addWorksheet("Data", { columns: [{ header: "Value" }] }).addRows([["x".repeat(1000)]]), { code: "RESOURCE_LIMIT" });
  assert.ok(accepted <= 100);
  await assert.rejects(writeCsvTo([["long text"]], sink, { columns: [{ header: "Value" }], limits: { maxTextCharacters: 8 } }), { code: "RESOURCE_LIMIT" });
  const small = new Workbook({ limits: { maxStyles: 1, maxSheets: 1 } });
  assert.throws(() => small.styles.add({ fill: { color: "FF0000" } }), { code: "RESOURCE_LIMIT" });
  small.addWorksheet("Only");
  assert.throws(() => small.addWorksheet("Extra"), { code: "RESOURCE_LIMIT" });
});

for (const compression of ["auto", "store"]) test("streamed XLSX preserves sink failure and cancellation: " + compression, async () => {
  const failure = new Error("sink failed");
  const failed = new Workbook({ compression, sink: { write() { throw failure; } } });
  await assert.rejects(failed.addWorksheet("Data", { columns: [{ header: "Value" }] }).addRows([[1]]), error => error === failure);
  await assert.rejects(failed.finish(), error => error === failure);
  const controller = new AbortController();
  let started;
  const blocked = new Promise(resolve => { started = resolve; });
  const cancelled = new Workbook({ compression, signal: controller.signal, sink: { write() { started(); return new Promise(() => {}); } } });
  const append = cancelled.addWorksheet("Data", { columns: [{ header: "Value" }] }).addRows([[1]]);
  await blocked; controller.abort(failure);
  await assert.rejects(append, error => error === failure);
});
