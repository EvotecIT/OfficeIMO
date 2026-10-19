import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { PdfFont, preflightPdfText, writePdf, writePdfDocument } from "../dist/pdf/index.js";
import { TrueTypeFont } from "../dist/pdf/truetype.js";
import { NotSupportedError } from "../dist/core/index.js";
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
  assert.deepEqual(preflightPdfText(text, { fonts }), { supported: true });
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
  const nativeFonts = { regular: withCharacters([[0x2009, thin], [0x202f, narrow]]) };
  assert.deepEqual(preflightPdfText("\u202f", { fonts: nativeFonts }), { supported: true });
  const native = await inspectPdf(await writePdfDocument([{ kind: "paragraph", text: "\u202f" }], { ...options, fonts: nativeFonts }));
  const nativeCid = [...native.objects.values()].find(object => /\/Subtype \/CIDFontType2/.test(object.body));
  assert.deepEqual(/\/W \[1 \[([^\]]+)/.exec(nativeCid.body)[1].trim().split(/\s+/).map(Number), [Number(program.widths[narrow].toFixed(3))]);
  for (const characters of [[[0x20, program.glyph(0x20)]], [[0x2009, program.glyph(65)]]]) {
    const missingFonts = { regular: withCharacters(characters) }, issue = preflightPdfText("\u202f", { fonts: missingFonts });
    assert.equal(issue.supported, false); assert.equal(issue.code, "missing-glyph"); assert.equal(issue.index, 0); assert.equal(issue.scalar, 0x202f);
    await assert.rejects(writePdfDocument([{ kind: "paragraph", text: "\u202f" }], { ...options, fonts: missingFonts }), { message: new NotSupportedError(issue.message).message });
  }
});

test("PDF preflight reports the first writer text issue at the original UTF-16 index", async () => {
  const program = new TrueTypeFont(source), glyph = program.glyph(65);
  const fonts = { regular: withCharacters([[0x20, program.glyph(0x20)], [65, glyph], [0x1f9ea, glyph]]) };
  for (const [text, selection, code, index, scalar] of [
    ["City Ł", {}, "embedded-font-required", 5, 0x141],
    ["A\u0001Z", { fonts }, "unsupported-control", 1, 1],
    ["A\u0301Z", { fonts }, "unsupported-layout", 1, 0x301],
    ["Aمرحبا", { fonts }, "unsupported-layout", 1, 0x645],
    ["A\u1100", { fonts }, "unsupported-layout", 1, 0x1100],
    ["A🧪Z", { fonts }, "missing-glyph", 3, 90],
    ["A\r\n\t🧪\ud800", { fonts }, "unpaired-surrogate", 6, undefined],
    ["A\r🧪\udc00", { fonts }, "unpaired-surrogate", 4, undefined]
  ]) {
    const result = preflightPdfText(text, selection);
    assert.equal(result.supported, false, text); assert.equal(result.code, code); assert.equal(result.index, index); assert.equal(result.scalar, scalar);
    await assert.rejects(writePdfDocument([{ kind: "paragraph", text }], { ...selection, pageNumbers: false }), error => {
      assert.equal(error.message, code === "unpaired-surrogate" ? result.message : new NotSupportedError(result.message).message);
      assert.equal(error.name, code === "unpaired-surrogate" ? "TypeError" : "NotSupportedError"); return true;
    });
  }
  const text = "A\r\n\t🧪\rA\nA";
  assert.deepEqual(preflightPdfText(text, { fonts }), { supported: true });
  const pdf = await inspectPdf(await writePdfDocument([{ kind: "paragraph", text }], { fonts, pageNumbers: false }));
  assert.deepEqual(pdf.pages[0].lines, ["A", "    🧪", "A", "A"]);
  assert.deepEqual(preflightPdfText("\r\n\r\n"), { supported: true });
  const missingSpace = { regular: withCharacters([[65, glyph]]) }, tab = preflightPdfText("A\t", { fonts: missingSpace });
  assert.equal(tab.supported, false); assert.equal(tab.code, "missing-glyph"); assert.equal(tab.index, 1); assert.equal(tab.scalar, 9);
  await assert.rejects(writePdfDocument([{ kind: "paragraph", text: "A\t" }], { fonts: missingSpace, pageNumbers: false }), { message: new NotSupportedError(tab.message).message });
});

test("PDF preflight and table headings select the same exact and synthetic emphasis faces", async () => {
  const program = new TrueTypeFont(source), face = scalar => withCharacters([[scalar, program.glyph(scalar)]]);
  const fonts = { regular: face(65), bold: face(66), italic: face(67), boldItalic: face(68) };
  for (const [text, bold, italic] of [["A", false, false], ["B", true, false], ["C", false, true], ["D", true, true]]) {
    assert.deepEqual(preflightPdfText(text, { fonts, bold, italic }), { supported: true });
    const pdf = await inspectPdf(await writePdfDocument([{ kind: "paragraph", text, presentation: { bold, italic } }], { fonts, pageNumbers: false }));
    assert.equal(pdf.text, text);
  }
  const heading = await inspectPdf(await writePdf([], { columns: [{ header: "B" }], fonts, pageNumbers: false }));
  assert.equal(heading.text, "B");
  const exact = preflightPdfText("A", { fonts, bold: true });
  assert.equal(exact.supported, false); assert.equal(exact.code, "missing-glyph");
  await assert.rejects(writePdfDocument([{ kind: "paragraph", text: "A", presentation: { bold: true } }], { fonts, pageNumbers: false }), { message: new NotSupportedError(exact.message).message });
  const fallbackFonts = { regular: fonts.regular };
  assert.deepEqual(preflightPdfText("A", { fonts: fallbackFonts, bold: true, italic: true }), { supported: true });
  const synthetic = await inspectPdf(await writePdfDocument([{ kind: "paragraph", text: "A", presentation: { bold: true, italic: true } }], { fonts: fallbackFonts, pageNumbers: false }));
  assert.equal(synthetic.text, "A"); assert.match(synthetic.pages[0].content, /2 Tr/); assert.match(synthetic.pages[0].content, /1 0 0\.213 1/);
});

test("PDF text preflight does not claim layout or document resource success", async () => {
  assert.deepEqual(preflightPdfText("Report"), { supported: true });
  await assert.rejects(writePdfDocument([{ kind: "paragraph", text: "Report" }], { limits: { maxCellCharacters: 1 } }), /maxCellCharacters/);
  const fonts = { regular: new PdfFont(source) };
  assert.deepEqual(preflightPdfText("Report", { fonts }), { supported: true });
  await assert.rejects(writePdfDocument([{ kind: "paragraph", text: "Report" }], { fonts, limits: { maxFontBytes: 1 } }), /maxFontBytes/);
  assert.throws(() => preflightPdfText(null), TypeError);
  assert.throws(() => preflightPdfText("Report", { bold: "yes" }), TypeError);
  assert.throws(() => preflightPdfText("Report", { fonts: {} }), /fonts.regular/);
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
