import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { PdfFont, writePdf, writePdfDocument } from "../dist/pdf/index.js";
import { TrueTypeFont } from "../dist/pdf/truetype.js";
import { inspectPdf } from "./pdf-reader.mjs";

const source = new Uint8Array(await readFile(new URL("../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf", import.meta.url)));

// A sparse synthetic Unicode cmap over the existing licensed fixture's original glyphs.
function withCharacters(characters) {
  const bytes = new Uint8Array(source.length + 28 + characters.length * 12); bytes.set(source);
  const view = new DataView(bytes.buffer), entry = directoryEntry(bytes, "cmap"), offset = source.length;
  view.setUint32(entry + 8, offset); view.setUint32(entry + 12, bytes.length - offset);
  view.setUint16(offset + 2, 1); view.setUint16(offset + 4, 3); view.setUint16(offset + 6, 10); view.setUint32(offset + 8, 12);
  view.setUint16(offset + 12, 12); view.setUint32(offset + 16, 16 + characters.length * 12); view.setUint32(offset + 24, characters.length);
  characters.forEach(([scalar, glyph], index) => {
    const p = offset + 28 + index * 12;
    view.setUint32(p, scalar); view.setUint32(p + 4, scalar); view.setUint32(p + 8, glyph);
  });
  return new PdfFont(bytes);
}

test("PDF preserves missing no-break thin spaces through the font's blank thin-space glyph", async () => {
  const program = new TrueTypeFont(source), thin = program.glyph(0x2009), narrow = program.glyph(0x202f), text = "\u2009\u202f";
  const fonts = { regular: withCharacters([[0x2009, thin]]) };
  const options = { fonts, pageNumbers: false, compression: false };
  for (const blob of [await writePdf([[text]], { ...options, columns: [{ header: "" }], includeHeader: false, columnWidths: [100] }),
    await writePdfDocument([{ kind: "paragraph", text }], options)]) {
    const pdf = await inspectPdf(blob);
    assert.equal(pdf.text, text, "Distinct Unicode scalars must survive extraction even when they share an inkless glyph");
    const cid = [...pdf.objects.values()].find(object => /\/Subtype \/CIDFontType2/.test(object.body));
    const map = pdf.objects.get(Number(/\/CIDToGIDMap (\d+) 0 R/.exec(cid.body)[1])).stream;
    assert.equal(map.readUInt16BE(2), thin); assert.equal(map.readUInt16BE(4), thin);
    const widths = /\/W \[1 \[([^\]]+)/.exec(cid.body)[1].trim().split(/\s+/).map(Number);
    assert.deepEqual(widths, [program.widths[thin], program.widths[thin]].map(width => Number(width.toFixed(3))));
  }
  // A declared narrow-space glyph takes precedence; missing or inked substitutes still fail.
  const native = await inspectPdf(await writePdfDocument([{ kind: "paragraph", text: "\u202f" }], { ...options, fonts: { regular: withCharacters([[0x2009, thin], [0x202f, narrow]]) } }));
  const nativeCid = [...native.objects.values()].find(object => /\/Subtype \/CIDFontType2/.test(object.body));
  assert.deepEqual(/\/W \[1 \[([^\]]+)/.exec(nativeCid.body)[1].trim().split(/\s+/).map(Number), [Number(program.widths[narrow].toFixed(3))]);
  for (const characters of [[[0x20, program.glyph(0x20)]], [[0x2009, program.glyph(65)]]])
    await assert.rejects(writePdfDocument([{ kind: "paragraph", text: "\u202f" }], { ...options, fonts: { regular: withCharacters(characters) } }), /no glyph for U\+202F/);
});

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
