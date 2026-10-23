import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { writePdfDocumentTo, writePdfDocument, PdfFont } from "../dist/pdf/index.js";
import { PdfObjects } from "../dist/pdf/objects.js";
import { inspectPdf } from "./pdf-reader.mjs";

const raw = () => {
  const data = new Uint8Array(512 * 256 * 3); let state = 123456789;
  for (let i = 0; i < data.length; i++) { state ^= state << 13; state ^= state >>> 17; state ^= state << 5; data[i] = state & 255; }
  return data;
};

test("compressed PDF streams reach the sink before completion and retain caller image/font snapshots", async () => {
  const native = globalThis.CompressionStream;
  let compressorFinished = false, reachedBeforeCompletion = false;
  globalThis.CompressionStream = class {
    constructor(format) {
      const stream = new native(format);
      this.writable = stream.writable;
      this.readable = stream.readable.pipeThrough(new TransformStream({ transform(chunk, controller) { controller.enqueue(chunk); }, flush() { compressorFinished = true; } }));
    }
  };
  try {
    const data = raw(), original = data.slice();
    const fontBytes = new Uint8Array(await readFile(new URL("../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf", import.meta.url)));
    const originalFont = fontBytes.slice(), font = new PdfFont(fontBytes), chunks = [];
    const result = await writePdfDocumentTo([
      { kind: "image", image: { format: "rgb", width: 512, height: 256, data } },
      { kind: "paragraph", text: "Captured report font" }
    ], { async write(chunk) {
      reachedBeforeCompletion ||= !compressorFinished;
      chunks.push(chunk.slice()); data.fill(0); fontBytes.fill(0); await Promise.resolve();
    } }, { fonts: { regular: font }, pageNumbers: false });
    assert.ok(reachedBeforeCompletion, "A destination must observe bytes while the compressor is still producing output");
    const pdf = await inspectPdf(new Blob(chunks));
    assert.equal(result.bytes, pdf.bytes.length);
    const image = [...pdf.objects.values()].find(object => /\/Subtype \/Image/.test(object.body));
    assert.deepEqual(new Uint8Array(image.stream), original);
    assert.deepEqual(font.toBytes(), originalFont);
    assert.match(pdf.text, /Captured report font/);
    assert.ok([...pdf.objects.values()].some(object => /\/Length \d+ 0 R \/Filter \/FlateDecode/.test(object.body)));
  } finally { globalThis.CompressionStream = native; }
});

test("native PDF WritableStreams retain accepted chunks without byte corruption", async () => {
  const chunks = [], accepted = [], data = raw(); let closed = false;
  const destination = new WritableStream({
    write(chunk) { chunks.push(chunk); accepted.push(chunk.slice()); },
    close() { closed = true; }
  });
  const result = await writePdfDocumentTo([
    { kind: "image", image: { format: "rgb", width: 512, height: 256, data } },
    { kind: "paragraph", text: "Retained native stream report" }
  ], destination, { compression: false, pageNumbers: false });
  assert.ok(chunks.length > 1, "The artifact must span multiple output writes");
  assert.deepEqual(Buffer.concat(chunks), Buffer.concat(accepted), "Accepted chunks must remain stable after later PDF writes");
  const pdf = await inspectPdf(new Blob(chunks));
  assert.equal(result.bytes, pdf.bytes.length);
  assert.match(pdf.text, /Retained native stream report/);
  const image = [...pdf.objects.values()].find(object => /\/Subtype \/Image/.test(object.body));
  assert.deepEqual(new Uint8Array(image.stream), data);
  assert.equal(destination.locked, false); assert.equal(closed, false);
  const writer = destination.getWriter();
  try { await writer.close(); } finally { writer.releaseLock(); }
  assert.equal(closed, true);
});

test("compressed output preserves sink/codec failures, cancellation and output ceilings", async () => {
  const source = raw();
  for (const mode of ["sink", "abort", "limit"]) {
    const controller = new AbortController(), reason = new Error(mode + " stopped"), objects = new PdfObjects({ write() {
      if (mode === "sink") throw reason;
      if (mode === "abort") { controller.abort(reason); return new Promise(() => {}); }
    } }, controller.signal, mode === "limit" ? 1000 : undefined);
    const pending = objects.stream(objects.reserve(), source);
    await assert.rejects(pending, mode === "limit" ? { code: "RESOURCE_LIMIT" } : error => error === reason);
  }
  const native = globalThis.CompressionStream, reason = new Error("codec stopped");
  globalThis.CompressionStream = class {
    constructor() { const stream = new TransformStream({ transform() { throw reason; } }); this.readable = stream.readable; this.writable = stream.writable; }
  };
  try {
    const objects = new PdfObjects({ write() {} });
    await assert.rejects(objects.stream(objects.reserve(), source), error => error === reason);
  } finally { globalThis.CompressionStream = native; }
});

test("PDF image resource ceilings are captured before asynchronous document sources execute", async () => {
  const block = { kind: "image", image: { format: "rgba", width: 3, height: 1, data: Uint8Array.of(1,2,3,4,5,6,7,8,9,10,11,12) } };
  for (const [initial, changed] of [[14, undefined], [15, 0]]) {
    const limits = { maxImageWorkingBytes: initial };
    async function* blocks() { limits.maxImageWorkingBytes = changed; yield block; }
    if (initial === 14) await assert.rejects(writePdfDocument(blocks(), { limits }), { code: "RESOURCE_LIMIT" });
    else await inspectPdf(await writePdfDocument(blocks(), { limits }));
  }
});

test("large logical object token streams preserve UTF-8 across bounded chunks and await destination acceptance", async () => {
  const parts = ["A".repeat(32759) + "\ud83d", "\ude80" + "B".repeat(100000) + "\ud83d\ude80"];
  const expected = Buffer.from("1 0 obj\n" + parts.join("") + "\nendobj\n"), chunks = [];
  let entered, accept, completed = false;
  const observed = new Promise(resolve => { entered = resolve; }), held = new Promise(resolve => { accept = resolve; });
  const objects = new PdfObjects({ async write(bytes) { chunks.push(bytes.slice()); entered(); await held; } });
  const pending = objects.object(objects.reserve(), parts).then(() => { completed = true; });
  await observed; assert.equal(completed, false, "A full output chunk must wait for the destination");
  accept(); await pending; await objects.flush();
  assert.deepEqual(Buffer.concat(chunks), expected);
  assert.ok(chunks.every(chunk => chunk.length <= 32768));
});
