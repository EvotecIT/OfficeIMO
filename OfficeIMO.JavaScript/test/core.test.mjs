import test from "node:test";
import assert from "node:assert/strict";
import { BlobByteSink, ChunkedTextSink, writeBytes, detectFeatures, saveBlob, NotSupportedError, pause } from "../dist/core/index.js";

test("cooperative pauses yield tasks and release concurrent message channels", async () => {
  const NativeChannel = globalThis.MessageChannel;
  try {
    let closed = 0;
    globalThis.MessageChannel = class extends NativeChannel {
      constructor() {
        super();
        for (const port of [this.port1, this.port2]) {
          const close = port.close.bind(port);
          port.close = () => { closed++; close(); };
        }
      }
    };
    let completed = 0;
    const pending = Promise.all(Array.from({ length: 3 }, () => pause().then(() => { completed++; })));
    await Promise.resolve(); assert.equal(completed, 0, "a microtask alone must not resume exports");
    await pending; assert.equal(completed, 3); assert.equal(closed, 6);
    globalThis.MessageChannel = undefined;
    let resumed = false;
    const timer = pause().then(() => { resumed = true; });
    await Promise.resolve(); assert.equal(resumed, false);
    await timer; assert.equal(resumed, true);
  } finally {
    globalThis.MessageChannel = NativeChannel;
  }
});

test("a delayed message channel cannot hold up the timer yield or leak ports", async () => {
  const NativeChannel = globalThis.MessageChannel;
  let callback, closed = 0;
  try {
    globalThis.MessageChannel = class {
      port1 = { set onmessage(value) { callback = value; }, close() { closed++; } };
      port2 = { postMessage() {}, close() { closed++; } };
    };
    await pause(); assert.equal(closed, 2);
    callback(); assert.equal(closed, 2, "a late message must not settle or clean up twice");
  } finally { globalThis.MessageChannel = NativeChannel; }
});

test("text pipeline stages share a completed yield and yield again after further work", async () => {
  const original = Object.getOwnPropertyDescriptor(performance, "now");
  let clock = performance.now() + 1000;
  Object.defineProperty(performance, "now", { configurable: true, value: () => clock });
  try {
    const first = new ChunkedTextSink({ write() {} }), second = new ChunkedTextSink({ write() {} });
    clock += 1000;
    assert.equal(first.append("first stage"), true);
    await first.flush();
    assert.equal(second.append("next stage"), false, "a stage must not immediately repeat the completed task yield");
    clock += 1000;
    assert.equal(second.append("more work"), true);
    await second.close();
  } finally {
    if (original) Object.defineProperty(performance, "now", original);
    else delete performance.now;
    await pause();
  }
});

test("text sinks preserve Unicode at chunk and append boundaries and await the destination", async () => {
  const chunks = [], sink = { async write(bytes) { await Promise.resolve(); chunks.push(bytes.slice()); } };
  const writer = new ChunkedTextSink(sink, undefined, 4);
  writer.append("abc\ud83e"); await writer.flush();
  writer.append("\uddeaשלום"); await writer.close();
  assert.equal(Buffer.concat(chunks).toString("utf8"), "abc🧪שלום");
  assert.ok(chunks.every(c => c.length <= 12));
  for (const bytes of [new Uint8Array([1, 2, 3]), Buffer.from([1, 2, 3])]) {
    const blob = new BlobByteSink();
    await writeBytes(bytes, blob); bytes.fill(0);
    assert.deepEqual(new Uint8Array(await blob.toBlob().arrayBuffer()), new Uint8Array([1, 2, 3]));
    assert.throws(() => blob.write(bytes), { code: "INVALID_STATE" });
  }
});

test("pending sink cancellation propagates the original reason", async () => {
  const controller = new AbortController(), error = new Error("stop");
  const writing = writeBytes(new Uint8Array([1]), { write: () => new Promise(() => {}) }, controller.signal);
  controller.abort(error); await assert.rejects(writing, e => e === error);
});

test("feature detection and typed errors work without a DOM", () => {
  assert.equal(detectFeatures().download, false);
  assert.equal(detectFeatures().blob, true);
  assert.throws(() => saveBlob(new Blob(), "file.csv"), { code: "PLATFORM_UNAVAILABLE" });
  assert.equal(new NotSupportedError("mergedCells").code, "NOT_SUPPORTED");
});
