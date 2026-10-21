import test from "node:test";
import assert from "node:assert/strict";
import { getEventListeners } from "node:events";
import { MessageChannel } from "node:worker_threads";
import { BlobByteSink, ExportCell, createGroupedReport } from "../dist/core/index.js";
import { writeCsv, writeCsvTo } from "../dist/csv/index.js";
import { writeXlsx } from "../dist/xlsx/index.js";
import { writePdf } from "../dist/pdf/index.js";
import { ZipWriter } from "../dist/zip/index.js";
import { OpcPackage } from "../dist/opc/index.js";
import { executeWorkerExport, handleWorkerExports } from "../dist/execution/index.js";
import { readZip } from "./zip-reader.mjs";
import { inspectPdf } from "./pdf-reader.mjs";

const writers = [
  ["CSV", writeCsv, blob => blob.text()],
  ["XLSX", writeXlsx, async blob => (await readZip(blob)).get("xl/worksheets/sheet1.xml").content],
  ["PDF", writePdf, async blob => (await inspectPdf(blob)).text]
];

test("table writers capture selected column fields once, including structural class getters", async () => {
  for (const [format, write, text] of writers) {
    let keys = 0, headers = 0;
    const column = { get header() { headers++; return "Value"; }, get key() { return ++keys === 1 ? "selected" : "other"; } };
    const output = await text(await write([{ selected: "expected-value", other: "wrong-value" }], { columns: [column] }));
    assert.match(output, /expected-value/, format); assert.doesNotMatch(output, /wrong-value/, format);
    assert.equal(keys, 1, format); assert.equal(headers, 1, format);
    class StructuralColumn { get header() { return "Value"; } get key() { return "selected"; } }
    assert.match(await text(await write([{ selected: "class-value" }], { columns: [new StructuralColumn()] })), /class-value/, format);
  }
});

test("hard row ceilings capture accessors once and retain structural limit profiles", async () => {
  let reads = 0;
  await assert.rejects(writeCsv([[1], [2]], { columns: [{ header: "Value" }],
    limits: { get maxRows() { return ++reads === 1 ? 1 : 100; } } }), { code: "RESOURCE_LIMIT" });
  assert.equal(reads, 1);
  class LimitProfile { get maxRows() { return 1; } }
  await assert.rejects(writeCsv([[1], [2]], { columns: [{ header: "Value" }], limits: new LimitProfile() }), { code: "RESOURCE_LIMIT" });
  const grouped = createGroupedReport([["A", 1], ["A", 2]], {
    columns: [{ key: "group", header: "Group" }, { key: "value", header: "Value" }], groupBy: "group", limits: new LimitProfile()
  });
  await assert.rejects(async () => { for await (const _ of grouped.events) { /* Consume through the public source. */ } }, { code: "RESOURCE_LIMIT" });
});

test("portable cells retain the same display and permitted link values that were validated", async () => {
  let textReads = 0, targetReads = 0;
  const cell = new ExportCell("raw", {
    get text() { return ++textReads === 1 ? "captured display" : 123; },
    link: { get target() { return ++targetReads < 4 ? "https://example.com/record" : "javascript:alert(1)"; } }
  });
  assert.equal(textReads, 1); assert.equal(targetReads, 1);
  assert.equal(cell.link.target, "https://example.com/record");
  assert.equal(await (await writeCsv([[cell]], { columns: [{ header: "Value" }], valueMode: "display" })).text(), "Value\r\ncaptured display\r\n");
});

test("CSV Blob and destination output share captured display, null and notification policy", async () => {
  for (const streamed of [false, true]) {
    const progress = [], options = { columns: [{ header: "Value" }], valueMode: "raw", nullValue: "original", onProgress: event => progress.push(event) };
    async function* rows() {
      options.valueMode = "display"; options.nullValue = "changed";
      options.onProgress = () => { throw Error("A replacement callback belongs to another operation."); };
      yield [new ExportCell(12, { text: "display-only" })]; yield [null];
    }
    let blob;
    if (streamed) { const sink = new BlobByteSink(); await writeCsvTo(rows(), sink, options); blob = sink.toBlob(); }
    else blob = await writeCsv(rows(), options);
    assert.equal(await blob.text(), "Value\r\n12\r\noriginal\r\n");
    assert.equal(progress.at(-1).phase, "complete"); assert.equal(progress.at(-1).rows, 2);
  }
});

test("ZIP and OPC retain constructor ceilings and XML policy until finalization", async () => {
  for (const packageMode of [false, true]) {
    const options = { compression: "store", maxOutputBytes: 0 };
    const output = packageMode ? new OpcPackage(options) : new ZipWriter(undefined, options);
    options.maxOutputBytes = 10000;
    if (packageMode) {
      output.addPart({ uri: "/data.xml", contentType: "application/xml", data: "<data/>" });
      await assert.rejects(output.toBlob(), { code: "RESOURCE_LIMIT" });
    } else await assert.rejects(output.add("data", Uint8Array.of(1)), { code: "RESOURCE_LIMIT" });
  }
  const options = { invalidCharacterPolicy: "reject" }, output = new OpcPackage(options);
  options.invalidCharacterPolicy = "strip";
  assert.throws(() => output.setProperties({ title: "invalid\u0001" }), { code: "INVALID_XML" });
});

test("worker cancellation retains its original signal and releases that signal's listener", { timeout: 5000 }, async () => {
  const { port1, port2 } = new MessageChannel();
  const handler = handleWorkerExports(port2, ({ rows, sink, signal }) => writeCsvTo(rows, sink, { columns: [{ header: "Value" }], signal }));
  const controller = new AbortController(), replacement = new AbortController(), reason = new Error("original signal cancelled");
  const options = { signal: controller.signal };
  let started, returned = 0, pending;
  const ready = new Promise(resolve => { started = resolve; });
  const rows = { [Symbol.asyncIterator]: () => ({ next() { started(); return new Promise(() => {}); }, return() { returned++; return { done: true }; } }) };
  try {
    pending = executeWorkerExport(port1, rows, { write() {} }, options);
    await ready; options.signal = replacement.signal; controller.abort(reason);
    let timeout;
    try {
      await assert.rejects(Promise.race([pending, new Promise((_, reject) => { timeout = setTimeout(() => reject(Error("The original signal was ignored.")), 1000); })]), error => error === reason);
    } finally { clearTimeout(timeout); }
    await new Promise(setImmediate);
    assert.equal(returned, 1); assert.equal(getEventListeners(controller.signal, "abort").length, 0);
  } finally { handler.dispose(); await pending?.catch(() => {}); port1.close(); port2.close(); }
});
