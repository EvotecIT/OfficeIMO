import test from "node:test";
import assert from "node:assert/strict";
import { Crc32, ZipWriter } from "../dist/zip/index.js";
import { readZip } from "./zip-reader.mjs";

for (const compression of ["auto", "store"]) test("ZIP streams with backpressure and valid descriptors: " + compression, async () => {
  const chunks = [], zip = new ZipWriter({ async write(bytes) { await Promise.resolve(); chunks.push(bytes.slice()); } }, { compression });
  await zip.add("test/Łódź.txt", async sink => {
    await sink.write(new TextEncoder().encode("abc")); await sink.write(new TextEncoder().encode("🧪"));
  });
  assert.ok(chunks.length > 2, "data is emitted before finalization");
  await zip.add("empty", new Uint8Array()); await zip.finish();
  const entries = await readZip(new Blob(chunks));
  assert.equal(entries.get("test/Łódź.txt").content, "abc🧪");
  assert.equal(entries.get("empty").content, "");
  if (compression === "store") assert.ok([...entries.values()].every(e => e.method === 0));
  await assert.rejects(zip.add("late", new Uint8Array()), { code: "INVALID_STATE" });
});

test("ZIP rejects unsafe paths, duplicate entries and overlapping writes", async () => {
  const zip = new ZipWriter();
  for (const name of ["../escape", "a/../b", "/root", "a\\b", "C:/root", "a//b", "nul\u0000"])
    await assert.rejects(zip.add(name, new Uint8Array()), TypeError);
  await zip.add("valid", new Uint8Array([1]));
  await assert.rejects(zip.add("valid", new Uint8Array()), /Duplicate/);
  let release;
  const adding = zip.add("pending", () => new Promise(r => { release = r; }));
  while (!release) await new Promise(r => setTimeout(r, 0));
  await assert.rejects(zip.finish(), { code: "INVALID_STATE" });
  release(); await adding;
  await readZip(await zip.toBlob());
  const crc = new Crc32(); crc.update(new TextEncoder().encode("123")); crc.update(new TextEncoder().encode("456789"));
  assert.equal(crc.value, 0xcbf43926);
});

test("ZIP64 entry-count policy fails before writing a sentinel-sized directory", async () => {
  const zip = new ZipWriter({ write() {} }, { compression: "store" });
  for (let i = 0; i < 65534; i++) await zip.add("e" + i, new Uint8Array());
  await assert.rejects(zip.add("overflow", new Uint8Array()), { code: "ZIP64_REQUIRED" });
  await zip.finish();
});

test("a sink failure or cancellation prevents ZIP finalization", async () => {
  const bad = new ZipWriter({ write() { throw new Error("disk full"); } });
  await assert.rejects(bad.add("data", new Uint8Array([1])), /disk full/);
  await assert.rejects(bad.finish(), { code: "INVALID_STATE" });
  const controller = new AbortController(), cancelled = new ZipWriter(undefined, { signal: controller.signal });
  const adding = cancelled.add("pending", () => new Promise(() => {}));
  controller.abort(); await assert.rejects(adding, { name: "AbortError" });
  await assert.rejects(cancelled.toBlob(), { name: "AbortError" });
});
