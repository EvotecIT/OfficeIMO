import test from "node:test";
import assert from "node:assert/strict";
import { MessageChannel, Worker } from "node:worker_threads";
import { readFile } from "node:fs/promises";
import { executeWorkerExport, handleWorkerExports } from "../dist/execution/index.js";
import { ExportCell, PdfFont, writeCsvTo, writePdfDocument } from "../dist/index.js";
import { BlobByteSink, NotSupportedError, OfficeIMOError } from "../dist/core/index.js";
import { TrueTypeFont } from "../dist/pdf/truetype.js";
import { inspectPdf } from "./pdf-reader.mjs";
import { readZip } from "./zip-reader.mjs";

const fontBytes = new Uint8Array(await readFile(new URL("../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf", import.meta.url)));

function observed(port) {
  const listeners = new Set(), outgoing = [], incoming = [];
  return { outgoing, incoming, get listenerCount() { return listeners.size; },
    postMessage(data) { outgoing.push(data); port.postMessage(data); },
    addEventListener(type, listener) { const wrapped = event => { if (type === "message") incoming.push(event.data); listener(event); }; listeners.add({ type, listener, wrapped }); port.addEventListener(type, wrapped); },
    removeEventListener(type, listener) { const entry = [...listeners].find(entry => entry.type === type && entry.listener === listener); if (entry) { port.removeEventListener(type, entry.wrapped); listeners.delete(entry); } },
    start() { port.start(); }
  };
}
async function local(operation, run) {
  const { port1, port2 } = new MessageChannel(), endpoint = observed(port1), handler = handleWorkerExports(port2, operation);
  try { await run(endpoint, handler); }
  finally { handler.dispose(); port1.close(); port2.close(); }
}

test("worker execution preserves portable typed cells, display, presentation and links through real workers", { timeout: 15000 }, async () => {
  for (const format of ["csv", "xlsx", "pdf"]) {
    const { port1, port2 } = new MessageChannel(), endpoint = observed(port1);
    const worker = new Worker(new URL("./execution-worker.mjs", import.meta.url), { workerData: { port: port2 }, transferList: [port2] });
    const sink = new BlobByteSink(), sourceBytes = fontBytes.slice(), values = [new ExportCell(12.5, { text: "12.50 USD", presentation: { bold: true, background: "E2F0D9" }, link: { target: "https://example.com/12" } }), new Date("2026-10-19T12:00:00Z"), "Łódź"];
    try {
      const completed = await executeWorkerExport(endpoint, [values], sink, { input: { format, columns: [{ header: "Amount" }, { header: "Date" }, { header: "City" }], fontBytes: sourceBytes }, maxChunkBytes: 257, batchRows: 2 });
      assert.equal(completed.rows, 1); assert.equal(completed.columns, 3); assert.ok(completed.bytes > 0);
      assert.equal(sourceBytes.byteLength, fontBytes.byteLength, "Font input remains caller-owned and attached");
      assert.deepEqual(sourceBytes, fontBytes); assert.equal(endpoint.listenerCount, 0);
      assert.ok(endpoint.incoming.filter(data => data.kind === "output").every(data => data.bytes.length <= 257 && data.bytes.buffer.byteLength <= 257));
      const blob = sink.toBlob(); assert.equal(blob.size, completed.bytes);
      if (format === "csv") assert.match(await blob.text(), /12\.5,2026-10-19T12:00:00\.000Z,Łódź/);
      if (format === "xlsx") {
        const parts = await readZip(blob), sheet = parts.get("xl/worksheets/sheet1.xml").content;
        assert.match(sheet, /<v>12\.5<\/v>/); assert.match(sheet, /Łódź/);
        assert.match(parts.get("xl/worksheets/_rels/sheet1.xml.rels").content, /https:\/\/example\.com\/12/);
      }
      if (format === "pdf") assert.match((await inspectPdf(blob)).text, /12\.50 USD/);
    } finally { port1.close(); await worker.terminate(); }
  }
});

test("worker rows and destination acknowledgements bound production and preserve source ownership", { timeout: 5000 }, async () => {
  let produced = 0, release, began;
  const blocked = new Promise(resolve => { began = resolve; }), gate = new Promise(resolve => { release = resolve; });
  await local(({ rows, sink, signal }) => writeCsvTo(rows, sink, { columns: [{ header: "Value" }], signal }), async endpoint => {
    function* source() { for (let i = 0; i < 100; i++) { produced++; yield ["x".repeat(1000) + i]; } }
    let writes = 0;
    const destination = { async write() { if (++writes === 1) { began(); await gate; } } };
    const pending = executeWorkerExport(endpoint, source(), destination, { batchRows: 3, maxBatchCells: 3, maxBatchCharacters: 4000, maxChunkBytes: 8192 });
    await blocked; const stopped = produced;
    assert.ok(stopped < 100);
    await new Promise(resolve => setImmediate(resolve)); assert.equal(produced, stopped);
    release(); assert.equal((await pending).rows, 100);
    assert.ok(endpoint.outgoing.filter(data => data.kind === "rows").every(data => data.rows.length <= 3));
    assert.equal(endpoint.listenerCount, 0);
  });
});

test("worker cancellation releases pending source and sink operations and allows endpoint reuse", { timeout: 5000 }, async () => {
  await local(({ rows, sink, signal }) => writeCsvTo(rows, sink, { columns: [{ header: "Value" }], signal }), async endpoint => {
    for (const phase of ["source", "sink"]) {
      const controller = new AbortController(), original = new Error(phase + " cancelled");
      let ready, returned = 0;
      const started = new Promise(resolve => { ready = resolve; });
      const source = phase === "source" ? { [Symbol.asyncIterator]: () => ({ next() { ready(); return new Promise(() => {}); }, return() { returned++; return { done: true }; } }) } : [["x".repeat(40000)]];
      const destination = { write() { if (phase === "sink") { ready(); return new Promise(() => {}); } } };
      const pending = executeWorkerExport(endpoint, source, destination, { signal: controller.signal });
      await started; controller.abort(original);
      await assert.rejects(pending, error => error === original);
      await new Promise(resolve => setImmediate(resolve));
      if (phase === "source") assert.equal(returned, 1);
      assert.equal(endpoint.listenerCount, 0);
      assert.equal((await executeWorkerExport(endpoint, [["reused"]], { write() {} })).rows, 1);
    }
  });
});

test("worker output clones only each bounded chunk of a large backing buffer", { timeout: 5000 }, async () => {
  const source = new Uint8Array(1024 * 1024); source.fill(42);
  await local(async ({ sink }) => {
    await sink.write(source.subarray(17, source.length - 17));
    return { rows: 0, columns: 0, bytes: source.length - 34 };
  }, async endpoint => {
    let accepted = 0;
    await executeWorkerExport(endpoint, [], { write(bytes) { accepted += bytes.length; assert.ok(bytes.every(value => value === 42)); } }, { maxChunkBytes: 65536 });
    assert.equal(accepted, source.length - 34);
    const chunks = endpoint.incoming.filter(data => data.kind === "output");
    assert.ok(chunks.length > 1);
    assert.ok(chunks.every(data => data.bytes.byteOffset === 0 && data.bytes.buffer.byteLength <= 65536));
    assert.equal(source.byteLength, 1024 * 1024, "The caller's backing buffer remains attached");
  });
});

test("worker disposal closes a pending producer without requiring a caller signal", { timeout: 5000 }, async () => {
  await local(({ rows, sink, signal }) => writeCsvTo(rows, sink, { columns: [{ header: "Value" }], signal }), async (endpoint, handler) => {
    let ready, returned = 0;
    const started = new Promise(resolve => { ready = resolve; });
    const rows = { [Symbol.asyncIterator]: () => ({ next() { ready(); return new Promise(() => {}); }, return() { returned++; return { done: true }; } }) };
    const pending = executeWorkerExport(endpoint, rows, { write() {} });
    await started;
    handler.dispose(new Error("worker disposed while source pending"));
    await assert.rejects(pending, { message: "worker disposed while source pending" });
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(returned, 1);
    assert.equal(endpoint.listenerCount, 0);
  });
});

test("worker errors, progress failures and handler disposal keep their terminal failure contract", { timeout: 5000 }, async () => {
  await local(async ({ input, onProgress, signal }) => {
    if (input === "remote") throw new OfficeIMOError("RESOURCE_LIMIT", "worker budget");
    if (input === "unsupported") throw new NotSupportedError("complex shaping");
    onProgress({ phase: "rows", rows: 1 });
    if (input === "dispose") await new Promise((_, reject) => signal.addEventListener("abort", () => reject(signal.reason), { once: true }));
    return { rows: 1, columns: 1, bytes: 0 };
  }, async (endpoint, handler) => {
    await assert.rejects(executeWorkerExport(endpoint, [], { write() {} }, { input: "remote" }), { code: "RESOURCE_LIMIT", message: "worker budget" });
    await assert.rejects(executeWorkerExport(endpoint, [], { write() {} }, { input: "unsupported" }), { code: "NOT_SUPPORTED", feature: "complex shaping" });
    const original = new Error("host progress");
    await assert.rejects(executeWorkerExport(endpoint, [], { write() {} }, { onProgress() { throw original; } }), error => error === original);
    let ready; const started = new Promise(resolve => { ready = resolve; });
    const pending = executeWorkerExport(endpoint, [], { write() {} }, { input: "dispose", onProgress() { ready(); } });
    await started; handler.dispose(); await assert.rejects(pending, { name: "AbortError" });
    assert.equal(endpoint.listenerCount, 0);
  });
});

test("worker setup and row budgets reject oversized or live configuration before unbounded cloning", async () => {
  await local(async ({ rows }) => { let count = 0; for await (const _ of rows) count++; return { rows: count, columns: 1, bytes: 0 }; }, async endpoint => {
    await assert.rejects(executeWorkerExport(endpoint, [], { write() {} }, { input: { huge: "x".repeat(100) }, maxSetupBytes: 10 }), { code: "RESOURCE_LIMIT" });
    await assert.rejects(executeWorkerExport(endpoint, [], { write() {} }, { input: { font: new PdfFont(fontBytes) } }), /reconstruct font instances/);
    await assert.rejects(executeWorkerExport(endpoint, [], { write() {} }, { input: new Array(1000000), maxSetupNodes: 10 }), { code: "RESOURCE_LIMIT" });
    const accessor = []; Object.defineProperty(accessor, "0", { get() { throw Error("Setup accessor must not run"); }, enumerable: true });
    await assert.rejects(executeWorkerExport(endpoint, [], { write() {} }, { input: accessor }), /accessors/);
    await assert.rejects(executeWorkerExport(endpoint, [["oversized"]], { write() {} }, { maxBatchCharacters: 2 }), { code: "RESOURCE_LIMIT" });
    assert.equal(endpoint.listenerCount, 0);
  });
});

test("worker progress coalesces to a bounded acknowledged queue and preserves the final update", async () => {
  await local(async ({ onProgress }) => {
    for (let i = 0; i < 10000; i++) onProgress({ phase: "rows", rows: i });
    return { rows: 9999, columns: 0, bytes: 0 };
  }, async endpoint => {
    const progress = [];
    const completed = await executeWorkerExport(endpoint, [], { write() {} }, { onProgress(event) { progress.push(event.rows); } });
    assert.equal(completed.rows, 9999); assert.equal(progress.at(-1), 9999);
    assert.ok(progress.length <= 2); assert.equal(endpoint.listenerCount, 0);
  });
});

test("async font preparation owns its snapshot and matches synchronous subset bytes", async () => {
  const source = fontBytes.slice(), pending = PdfFont.create(source), prepared = await pending;
  source.fill(0); assert.deepEqual(prepared.toBytes(), fontBytes);
  const program = new TrueTypeFont(fontBytes), characters = new Map([..."Łódź report"].map(char => [char.codePointAt(0), program.glyph(char.codePointAt(0))]));
  assert.deepEqual(await program.subsetCooperatively(new Set(characters.values()), characters), program.subset(new Set(characters.values()), characters));
  await assert.rejects(PdfFont.create(fontBytes, { maxBytes: 1 }), /byte limit/);
  const controller = new AbortController(), original = new Error("font preparation cancelled"); controller.abort(original);
  await assert.rejects(PdfFont.create(fontBytes, { signal: controller.signal }), error => error === original);
});

test("font parsing and subset finalization cooperate with cancellation", async () => {
  const now = Object.getOwnPropertyDescriptor(performance, "now"), scheduler = globalThis.scheduler;
  try {
    for (const phase of ["prepare", "subset"]) {
      const controller = new AbortController(), original = new Error(phase + " font cancelled"); let clock = 0;
      Object.defineProperty(performance, "now", { configurable: true, value: () => clock += 40 });
      globalThis.scheduler = { postTask: async callback => { controller.abort(original); callback(); } };
      if (phase === "prepare") await assert.rejects(PdfFont.create(fontBytes, { signal: controller.signal }), error => error === original);
      else {
        const program = new TrueTypeFont(fontBytes), characters = new Map([[65, program.glyph(65)]]);
        await assert.rejects(program.subsetCooperatively(new Set(characters.values()), characters, controller.signal), error => error === original);
      }
    }
  } finally { if (now) Object.defineProperty(performance, "now", now); else delete performance.now; globalThis.scheduler = scheduler; }
});

test("raw RGBA compaction preserves caller bytes and enforces the owned working-byte ceiling", async () => {
  const data = Uint8Array.of(1,2,3,4, 5,6,7,8, 9,10,11,12), original = data.slice();
  const block = { kind: "image", image: { format: "rgba", data, width: 3, height: 1 } };
  const pdf = await inspectPdf(await writePdfDocument([block], { compression: false, limits: { maxImageWorkingBytes: 15 } }));
  assert.deepEqual(data, original);
  const images = [...pdf.objects.values()].filter(object => /\/Subtype \/Image/.test(object.body));
  assert.deepEqual([...images.find(object => /DeviceRGB/.test(object.body)).stream], [1,2,3,5,6,7,9,10,11]);
  assert.deepEqual([...images.find(object => /DeviceGray/.test(object.body)).stream], [4,8,12]);
  await assert.rejects(writePdfDocument([block], { limits: { maxImageWorkingBytes: 14 } }), { code: "RESOURCE_LIMIT" });
});

test("PNG owned-buffer budgets include source snapshot, output planes and two RGBA scanlines", async () => {
  for (const format of ["rgb", "rgba"]) {
    const data = new Uint8Array(await readFile(new URL("./fixtures/pdf/" + format + "-pillow.png", import.meta.url)));
    const maximum = data.length + 12 * 8 * (format === "rgba" ? 4 : 3) + (format === "rgba" ? 12 * 8 : 0);
    const block = { kind: "image", image: { format: "png", data } };
    const pdf = await inspectPdf(await writePdfDocument([block], { compression: false, limits: { maxImageWorkingBytes: maximum } }));
    assert.ok(pdf.objects.size > 0);
    await assert.rejects(writePdfDocument([block], { limits: { maxImageWorkingBytes: maximum - 1 } }), { code: "RESOURCE_LIMIT" });
  }
});
