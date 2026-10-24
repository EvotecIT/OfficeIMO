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
});

test("CRC-32 matches standard and independent vectors across update splits and byte-view offsets", () => {
  const expected = bytes => {
    let crc = 0xffffffff;
    for (const byte of bytes) {
      crc ^= byte;
      for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
    return (crc ^ 0xffffffff) >>> 0;
  };
  const standard = new TextEncoder().encode("123456789"), crc = new Crc32();
  assert.equal(crc.value, 0);
  crc.update(standard.subarray(0, 3)); crc.update(standard.subarray(3));
  assert.equal(crc.value, 0xcbf43926);
  for (const length of [0, 1, 3, 4, 5, 7, 8, 9, 15, 16, 17, 255, 256, 257]) {
    for (const storage of [new Uint8Array(length + 7), Buffer.alloc(length + 7)]) {
      const bytes = storage.subarray(3, 3 + length);
      for (let index = 0; index < length; index++) bytes[index] = (index * 73 + 19) & 255;
      for (const split of new Set([0, Math.min(1, length), Math.min(3, length), Math.min(4, length), Math.floor(length / 2), length])) {
        const checksum = new Crc32();
        checksum.update(bytes.subarray(0, split));
        assert.equal(checksum.value, expected(bytes.subarray(0, split)));
        checksum.update(bytes.subarray(split));
        assert.equal(checksum.value, expected(bytes));
      }
    }
  }
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

test("invalid ZIP producer bytes fail without waiting for cleanup or finalizing output", { timeout: 2000 }, async () => {
  let returned = 0, produced = 0;
  const source = { [Symbol.iterator]() { return {
    next() { produced++; return { done: false, value: "not bytes" }; },
    return() { returned++; return new Promise(() => {}); }
  }; } };
  const zip = new ZipWriter();
  await assert.rejects(zip.add("data", source), TypeError);
  assert.equal(returned, 1); assert.equal(produced, 1);
  await assert.rejects(zip.finish(), { code: "INVALID_STATE" });
});
