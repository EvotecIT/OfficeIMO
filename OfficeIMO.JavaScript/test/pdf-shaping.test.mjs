import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { PdfFont, writePdf, writePdfDocument } from "../dist/pdf/index.js";
import { TrueTypeFont } from "../dist/pdf/truetype.js";
import { inspectPdf } from "./pdf-reader.mjs";
import { pdfStructureElements } from "./pdf-structure-reader.mjs";
import { createHarfBuzzTextShaper } from "../dist/pdf/providers/harfbuzz.js";
import { isUnicode13ProfileScalar } from "../dist/pdf/providers/unicode13-profile.js";

const bytes = new Uint8Array(await readFile(new URL("../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf", import.meta.url)));
const font = new PdfFont(bytes), otherFont = new PdfFont(bytes), program = new TrueTypeFont(bytes), id = program.glyph(65);
const options = { fonts: { regular: font }, pageNumbers: false, compression: false };
const paragraph = text => ({ kind: "paragraph", text });
const glyph = (start, end, extra = {}) => ({ id, clusterStart: start, clusterEnd: end, xAdvance: 500, yAdvance: 0, xOffset: 0, yOffset: 0, ...extra });
const provider = callback => ({ shapeLine: request => callback(request) });
const nominal = provider(request => [{ font, glyphs: [...request.text.slice(request.start, request.end)].map((char, index) => glyph(request.start + index, request.start + index + 1)) }]);

test("shaped glyphs retain positioned outlines, logical ActualText and nonempty ToUnicode cluster mappings", async () => {
  const textShaper = provider(() => [{ font, glyphs: [glyph(0, 2, { xOffset: 20, yOffset: 30, xAdvance: 0 }), glyph(0, 2)] }]);
  const pdf = await inspectPdf(await writePdfDocument([paragraph("Aa")], { ...options, textShaper, fontSize: 10 }));
  assert.match(pdf.pages[0].content, /\/ActualText <FEFF00410061>/);
  const cmap = [...pdf.objects.values()].find(object => /beginbfchar/.test(object.body)).body;
  assert.match(cmap, /1 beginbfchar\n<0001> <00410061>/);
  assert.doesNotMatch(cmap, /<\d+> <>/);
  assert.match(pdf.pages[0].content, /36\.2 [\d.]+ Tm <0001> Tj/);
  assert.match(pdf.pages[0].content, /36 [\d.]+ Tm <0002> Tj/);
});

test("core validates provider fonts, glyphs, horizontal metrics and complete scalar-safe cluster partitions", async () => {
  for (const output of [
    [{ font: otherFont, glyphs: [glyph(0, 2)] }],
    [{ font, glyphs: [glyph(0, 2, { id: 0 })] }],
    [{ font, glyphs: [glyph(0, 2, { id: program.glyphCount })] }],
    [{ font, glyphs: [glyph(0, 2, { xAdvance: NaN })] }],
    [{ font, glyphs: [glyph(0, 2, { yAdvance: 1 })] }],
    [{ font, glyphs: [glyph(0, 1)] }],
    [{ font, glyphs: [glyph(0, 2), glyph(1, 2)] }],
    [{ font, glyphs: [glyph(0, 3)] }]
  ]) await assert.rejects(writePdfDocument([paragraph("AB")], { ...options, textShaper: provider(() => output) }), TypeError);
  await assert.rejects(writePdfDocument([paragraph("🧪")], { ...options, textShaper: provider(() => [{ font, glyphs: [glyph(0, 1), glyph(1, 2)] }]) }), /UTF-16/);
});

test("indivisible shaped clusters reject conflicting presentation and link boundaries", async () => {
  const textShaper = provider(() => [{ font, glyphs: [glyph(0, 2)] }]);
  for (const tail of [{ text: "a", color: "ff0000" }, { text: "a", bold: true }, { text: "a", link: { target: "https://example.com" } }])
    await assert.rejects(writePdfDocument([paragraph([{ text: "A" }, tail])], { ...options, textShaper }), /indivisible shaped cluster/);
  await assert.rejects(writePdfDocument([paragraph([{ text: "\ud83e" }, { text: "\uddea" }])], { ...options, textShaper }), /UTF-16 scalar/);
});

test("bounded shaped wrapping preserves logical whitespace, blank lines, page links and table text", async () => {
  const pdf = await inspectPdf(await writePdfDocument([paragraph([{ text: "A B C D E F G H I J ".repeat(8) + "\n\nK\tL", link: { target: "https://example.com/report" } }])], {
    ...options, textShaper: nominal, tagged: { language: "en" }, fontSize: 10, pageSize: { width: 130, height: 100 }, margins: 10
  }));
  assert.ok(pdf.pages.length > 1);
  assert.ok(pdf.pages.every(page => /\/Annots/.test(page.body)));
  assert.ok(pdfStructureElements(pdf).some(([, object]) => /\/S \/Link/.test(object.body)));
  const table = await inspectPdf(await writePdf([{ value: "AB" }], { ...options, textShaper: nominal, columns: [{ header: "H", key: "value" }] }));
  assert.equal((table.pages[0].content.match(/\/ActualText/g) ?? []).length, 2);
});

test("shaping limits and cancellation fail before fallback or unbounded provider work", async () => {
  let calls = 0;
  const textShaper = provider(request => { calls++; return nominal.shapeLine(request); });
  await assert.rejects(writePdfDocument([paragraph("A".repeat(4097))], { ...options, textShaper }), /4,096/);
  assert.equal(calls, 0);
  await assert.rejects(writePdfDocument([paragraph("A")], { textShaper }), /embedded TrueType/);
  await assert.rejects(writePdfDocument([paragraph("A\u0001")], { ...options, textShaper }), /control/);
  const controller = new AbortController(), reason = Error("stop shaped layout");
  const pending = writePdfDocument([paragraph("A")], { ...options, signal: controller.signal, textShaper: provider(() => { controller.abort(reason); return new Promise(() => {}); }) });
  await assert.rejects(pending, error => error === reason);
  const failure = Error("provider failed");
  await assert.rejects(writePdfDocument([paragraph("A")], { ...options, textShaper: provider(() => { throw failure; }) }), error => error === failure);
});

test("provider method is captured before asynchronous input and legacy decorations stay scalar", async () => {
  const textShaper = { shapeLine: nominal.shapeLine };
  async function* blocks() { textShaper.shapeLine = () => { throw Error("replaced"); }; yield paragraph("A"); }
  await writePdfDocument(blocks(), { ...options, textShaper });
  await assert.rejects(writePdfDocument([paragraph("A")], { ...options, textShaper: nominal, title: "مرحبا" }), /scalar layout/);
});

test("explicit run faces share the operation font budget, snapshot identity and apply synthetic emphasis", async () => {
  const run = { text: "A", font, bold: true, italic: true };
  async function* blocks() { yield { kind: "paragraph", text: [run], keepWithNext: true }; run.font = otherFont; yield paragraph("B"); }
  const pdf = await inspectPdf(await writePdfDocument(blocks(), { ...options, limits: { maxFontBytes: bytes.length } }));
  assert.match(pdf.pages[0].content, /2 Tr/); assert.match(pdf.pages[0].content, /1 0 0\.213 1/);
  await assert.rejects(writePdfDocument([paragraph([{ text: "A", font: otherFont }])], { ...options, limits: { maxFontBytes: bytes.length } }), /maxFontBytes/);
  await assert.rejects(writePdfDocument([paragraph([{ text: "A", font: {} }])], options), /must be a PdfFont/);
  await assert.rejects(writePdfDocument([], { ...options, fonts: { regular: font, bold: otherFont }, limits: { maxFontBytes: bytes.length } }), /maxFontBytes/);
});

test("optional adapter is dependency-free at import and disposed providers reject further work", () => {
  class Handle {}
  const shaper = createHarfBuzzTextShaper({ harfbuzz: { Blob: Handle, Face: Handle, Font: Handle, Buffer: Handle, shape() {} }, bidi: { getEmbeddingLevels() {}, getReorderedIndices() {}, getBidiCharTypeName() {} } });
  shaper.dispose(); shaper.dispose();
  assert.throws(() => shaper.shapeLine({ text: "A", start: 0, end: 1, runs: [{ start: 0, end: 1, font }], direction: "ltr" }), /disposed/);
});

test("optional provider's normative repertoire rejects later Arabic assignments independently of host Unicode data", () => {
  assert.equal(isUnicode13ProfileScalar(0x627), true);
  assert.equal(isUnicode13ProfileScalar(0x301), true);
  assert.equal(isUnicode13ProfileScalar(0x870), false, "Arabic Extended-B was assigned after Unicode13");
  class Handle {}
  const shaper = createHarfBuzzTextShaper({ harfbuzz: { Blob: Handle, Face: Handle, Font: Handle, Buffer: Handle, shape() {} },
    bidi: { getEmbeddingLevels: text => ({ levels: new Uint8Array(text.length), paragraphs: [{ start: 0, end: text.length - 1, level: 0 }] }),
      getReorderedIndices: text => [...text].map((_, i) => i), getBidiCharTypeName: () => "L" } });
  assert.throws(() => shaper.shapeLine({ text: "\u0870", start: 0, end: 1, runs: [{ start: 0, end: 1, font }], direction: "auto" }), /Unicode 13\.0 profile/);
  shaper.dispose();
});

test("tagged shaped table headings retain one semantic header and mark every repeated heading as an artifact", async () => {
  const pdf = await inspectPdf(await writePdf(Array.from({ length: 30 }, () => ({ value: "Body row" })), {
    ...options, textShaper: nominal, tagged: { language: "en" }, columns: [{ header: "Heading", key: "value" }],
    pageSize: { width: 180, height: 150 }, orientation: "landscape", margins: 10
  }));
  assert.ok(pdf.pages.length >= 3, "The fixture must exercise several repeated headings");
  assert.equal(pdfStructureElements(pdf).filter(([, object]) => /\/S \/TH\b/.test(object.body)).length, 1,
    "The column heading has one logical TH even when painted on several pages");
  for (const [index, page] of pdf.pages.entries()) {
    const stack = []; let textObjects = 0, firstTextOwner;
    for (const token of page.content.matchAll(/\/(\w+)\s*<<([\s\S]*?)>>\s*BDC|\bEMC\b|\bBT\b/g)) {
      if (token[1]) stack.push({ role: token[1], properties: token[2] });
      else if (token[0] === "EMC") assert.ok(stack.pop(), "Marked content must be balanced");
      else {
        const owner = stack.find(mark => mark.role === "Artifact" || /\/MCID\b/.test(mark.properties));
        assert.ok(owner, "Every tagged text object must have a semantic or artifact owner");
        firstTextOwner ??= owner.role; textObjects++;
      }
    }
    assert.equal(stack.length, 0); assert.ok(textObjects >= 2);
    assert.equal(firstTextOwner, index ? "Artifact" : "TH", "Repeated heading text must use the pagination artifact contract");
  }
});

test("synthetic italic shears positioned mark origins together with their base glyph", async () => {
  const textShaper = provider(() => [{ font, glyphs: [glyph(0, 2, { xOffset: 20, yOffset: 30, xAdvance: 0 }),
    glyph(0, 2, { xOffset: -10, yOffset: -50, xAdvance: 0 }), glyph(0, 2)] }]);
  const matrices = async italic => {
    const pdf = await inspectPdf(await writePdfDocument([paragraph([{ text: "Aa", font, italic }])], { ...options, textShaper, fontSize: 10 }));
    return [...pdf.pages[0].content.matchAll(/1 0 ([\d.-]+) 1 ([\d.-]+) ([\d.-]+) Tm/g)].map(match => ({ shear: Number(match[1]), x: Number(match[2]), y: Number(match[3]) }));
  };
  const upright = await matrices(false), italic = await matrices(true);
  assert.equal(upright.length, 3); assert.equal(italic.length, 3);
  const baseline = upright.at(-1).y;
  for (let i = 0; i < upright.length; i++) {
    assert.equal(italic[i].y, upright[i].y);
    assert.ok(Math.abs(italic[i].x - (upright[i].x + italic[i].shear * (upright[i].y - baseline))) <= .001,
      "The same affine shear must preserve positive and negative mark attachment offsets");
  }
});
