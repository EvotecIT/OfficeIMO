import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { PdfFont } from "../dist/pdf/index.js";

const source = new Uint8Array(await readFile(new URL("../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf", import.meta.url)));

function directoryEntry(bytes, tag) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let i = 0; i < view.getUint16(4); i++) {
    const p = 12 + i * 16;
    if (String.fromCharCode(...bytes.subarray(p, p + 4)) === tag) return p;
  }
  throw Error("Missing fixture font table " + tag);
}

function aliasedFont(format, records) {
  const original = new DataView(source.buffer), entry = directoryEntry(source, "cmap");
  let subtable;
  if (format === 4) {
    const cmap = original.getUint32(entry + 8);
    for (let i = 0; i < original.getUint16(cmap + 2); i++) {
      const p = cmap + original.getUint32(cmap + 8 + i * 8);
      if (original.getUint16(p) === 4) { subtable = source.slice(p, p + original.getUint16(p + 2)); break; }
    }
    assert.ok(subtable);
  } else {
    const groups = 128;
    subtable = new Uint8Array(16 + groups * 12);
    const table = new DataView(subtable.buffer);
    table.setUint16(0, 12); table.setUint32(4, subtable.length); table.setUint32(12, groups);
    for (let i = 0; i < groups; i++) {
      const p = 16 + i * 12;
      table.setUint32(p, 65 + i); table.setUint32(p + 4, 65 + i); table.setUint32(p + 8, i + 1);
    }
  }
  const tableOffset = source.length, subOffset = tableOffset + 4 + records * 8;
  const bytes = new Uint8Array(subOffset + subtable.length); bytes.set(source); bytes.set(subtable, subOffset);
  const view = new DataView(bytes.buffer);
  view.setUint32(entry + 8, tableOffset); view.setUint32(entry + 12, bytes.length - tableOffset);
  view.setUint16(tableOffset + 2, records);
  for (let i = 0; i < records; i++) {
    const p = tableOffset + 4 + i * 8;
    view.setUint16(p, 3); view.setUint16(p + 2, format === 12 ? 10 : 1);
    view.setUint32(p + 4, subOffset - tableOffset);
  }
  return { bytes, subOffset };
}

function parseReads(bytes) {
  // Count bounded binary reads instead of using a host-dependent timing assertion.
  // The public constructor copies the input, so identify its reader by byte length.
  const u16 = DataView.prototype.getUint16, u32 = DataView.prototype.getUint32;
  let reads = 0;
  try {
    DataView.prototype.getUint16 = function(...args) { if (this.byteLength === bytes.length) reads++; return u16.apply(this, args); };
    DataView.prototype.getUint32 = function(...args) { if (this.byteLength === bytes.length) reads++; return u32.apply(this, args); };
    const font = new PdfFont(bytes);
    assert.deepEqual(font.toBytes(), bytes);
    return reads;
  } finally { DataView.prototype.getUint16 = u16; DataView.prototype.getUint32 = u32; }
}

for (const format of [4, 12]) test("PDF fonts bound validation of aliased Unicode cmap format " + format, t => {
  const single = aliasedFont(format, 1), repeated = aliasedFont(format, 128);
  const baseline = parseReads(single.bytes), aliased = parseReads(repeated.bytes);
  t.diagnostic(JSON.stringify({ format, baseline, aliased }));
  assert.ok(aliased <= baseline + 128 * 8, `Repeated encoding records must add directory work, not revalidate the shared map: ${baseline} versus ${aliased} reads`);
  const view = new DataView(repeated.bytes.buffer);
  if (format === 12) view.setUint32(repeated.subOffset + 16, 0x10ffff);
  else { const segments = view.getUint16(repeated.subOffset + 6) / 2; view.setUint16(repeated.subOffset + 16 + segments * 2, 0xffff); }
  assert.throws(() => new PdfFont(repeated.bytes), /Invalid cmap/, "Aliasing must not hide an invalid first subtable");
});
