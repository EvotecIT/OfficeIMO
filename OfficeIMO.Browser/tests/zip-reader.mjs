// Node's independent ZIP consumer checks headers and inflates platform-produced payloads.
import assert from "node:assert/strict";
import { inflateRawSync } from "node:zlib";

export async function readZip(blob) {
  const bytes = Buffer.from(await blob.arrayBuffer());
  const end = bytes.length - 22;
  assert.equal(bytes.readUInt32LE(end), 0x06054b50);
  const count = bytes.readUInt16LE(end + 10), entries = new Map();
  let offset = bytes.readUInt32LE(end + 16);
  for (let i = 0; i < count; i++) {
    assert.equal(bytes.readUInt32LE(offset), 0x02014b50);
    const length = bytes.readUInt16LE(offset + 28);
    const name = bytes.toString("utf8", offset + 46, offset + 46 + length);
    const local = bytes.readUInt32LE(offset + 42), compressed = bytes.readUInt32LE(offset + 20);
    assert.equal(bytes.readUInt32LE(local), 0x04034b50);
    const start = local + 30 + bytes.readUInt16LE(local + 26) + bytes.readUInt16LE(local + 28);
    const method = bytes.readUInt16LE(offset + 10);
    const payload = bytes.subarray(start, start + compressed);
    const content = method === 8 ? inflateRawSync(payload) : payload;
    assert.equal(content.length, bytes.readUInt32LE(offset + 24));
    let crc = 0xffffffff;
    for (const byte of content) {
      crc ^= byte;
      for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
    }
    assert.equal((crc ^ 0xffffffff) >>> 0, bytes.readUInt32LE(offset + 16), "ZIP CRC-32 differs");
    assert.equal(bytes.readUInt32LE(local + 14), bytes.readUInt32LE(offset + 16));
    entries.set(name, { method, content: content.toString("utf8"), bytes: content });
    offset += 46 + length + bytes.readUInt16LE(offset + 30) + bytes.readUInt16LE(offset + 32);
  }
  return entries;
}
