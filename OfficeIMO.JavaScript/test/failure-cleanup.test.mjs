import test from "node:test";
import assert from "node:assert/strict";
import { randomFillSync } from "node:crypto";
import { ZipWriter } from "../dist/zip/index.js";
import { inputRows, detectFeatures } from "../dist/core/index.js";
import { readZip } from "./zip-reader.mjs";

test("ZIP rejects malformed Unicode without aliasing valid emitted names", async () => {
  const zip = new ZipWriter();
  for (const name of ["file\ud800.txt", "file\udc00.txt", "file\ud800a.txt"])
    await assert.rejects(zip.add(name, new Uint8Array()), /well-formed Unicode/);
  for (const name of ["file\ufffd.txt", "file🧪.txt"]) await zip.add(name, new TextEncoder().encode(name));
  const entries = await readZip(await zip.toBlob());
  assert.deepEqual([...entries.keys()], ["file\ufffd.txt", "file🧪.txt"]);
  for (const [name, entry] of entries) assert.equal(entry.content, name);
});

test("compressed producer failure releases stalled output without caller cancellation", {
  timeout: 3000, skip: !detectFeatures().deflateRaw
}, async () => {
  let writes = 0, began;
  const started = new Promise(resolve => { began = resolve; });
  const original = new Error("producer failed"), bytes = randomFillSync(new Uint8Array(20000));
  const zip = new ZipWriter({ write() {
    if (++writes === 1) return;
    began(); return new Promise(() => {});
  } });
  await assert.rejects(zip.add("data.bin", async sink => {
    await sink.write(bytes); await started; throw original;
  }), error => error === original);
  await assert.rejects(zip.finish(), { code: "INVALID_STATE" });
});

test("iterator cleanup cannot replace producer failure or cancellation reason", async () => {
  for (const cleanup of [() => { throw new Error("cleanup"); }, () => Promise.reject(new Error("cleanup"))]) {
    const original = new Error("producer"), controller = new AbortController();
    const rows = { [Symbol.asyncIterator]: () => ({ next: () => new Promise(() => {}), return: cleanup }) };
    const pending = inputRows(rows, controller.signal).next();
    controller.abort(original); await assert.rejects(pending, error => error === original);
    const failing = { [Symbol.asyncIterator]: () => ({ next: () => Promise.reject(original), return: cleanup }) };
    await assert.rejects(inputRows(failing).next(), error => error === original);
  }
});
