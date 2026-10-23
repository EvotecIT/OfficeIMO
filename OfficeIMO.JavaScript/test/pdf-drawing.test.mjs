import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { PdfFont, writePdfDocument, writePdfDocumentTo } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";
import { pdfStructureElements } from "./pdf-structure-reader.mjs";

const fonts = Object.fromEntries(await Promise.all([["regular", "Regular"], ["bold", "Bold"], ["italic", "Italic"], ["boldItalic", "BoldItalic"]].map(async ([slot, file]) =>
  [slot, new PdfFont(new Uint8Array(await readFile(new URL("../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-" + file + ".ttf", import.meta.url))))])));
const path = { kind: "path", stroke: "008000", strokeWidth: 2, commands: [{ op: "move", x: -20, y: 0 }, { op: "line", x: 200, y: 100 }] };
const text = { kind: "text", text: "Łódź", x: 10, y: 25, font: "regular", fontSize: 14, color: "000000" };
const drawing = elements => ({ version: 1, width: 200, height: 100, elements });
const block = elements => ({ kind: "drawing", drawing: drawing(elements) });
const options = { fonts, pageNumbers: false, compression: false };

test("portable drawings preserve painter order, upright embedded text, clipping and stroke semantics", async () => {
  const filled = { kind: "path", fill: "FF0000", fillRule: "evenodd", commands: [{ op: "move", x: 0, y: 0 }, { op: "line", x: 200, y: 0 }, { op: "line", x: 200, y: 100 }, { op: "close" }] };
  const pdf = await inspectPdf(await writePdfDocument([block([filled, text, { kind: "group", clip: { x: 5, y: 8, width: 120, height: 50 }, elements: [
    { ...path, lineCap: "round", lineJoin: "bevel" }, { ...text, text: "Bold", font: "bold", y: 40 },
    { ...text, text: "Italic", font: "italic", y: 60 }, { ...text, text: "Both", font: "boldItalic", y: 80 }
  ] }, { ...filled, fill: "0000FF" }])], options));
  assert.equal(pdf.pages.length, 1); assert.equal(pdf.text, "Łódź\nBold\nItalic\nBoth");
  const content = pdf.pages[0].content;
  assert.match(content, /1 0 0 -1 36 805\.89 cm 0 0 200 100 re W n/);
  assert.match(content, /BT \/F1 14 Tf 0 Tr 1 0 0 -1 10 25 Tm/);
  assert.match(content, /5 8 120 50 re W n/);
  assert.match(content, /1 J\n2 j\n10 M\n-20 0 m\n200 100 l\nS/);
  assert.match(content, /h\nf\*/);
  assert.ok(content.indexOf("1 0 0 rg") < content.indexOf("BT /F1"));
  assert.ok(content.indexOf("BT /F1") < content.indexOf("5 8 120 50 re W n"));
  assert.ok(content.indexOf("5 8 120 50 re W n") < content.indexOf("0 0 1 rg"), "Later paint overlays earlier text and paths");
  assert.equal([...pdf.objects.values()].filter(object => /\/Subtype \/Type0/.test(object.body)).length, 4, "All four exact faces use the existing embedded-font owner");
});

test("drawings scale and paginate as indivisible graphics and use the figure alternative contract", async () => {
  const blocks = [{ kind: "paragraph", text: "Before\nBefore\nBefore\nBefore\nBefore" },
    { ...block([path, text]), alternativeText: "Sales diagram" }, { ...block([path]), decorative: true }];
  const pdf = await inspectPdf(await writePdfDocument(blocks, { ...options, tagged: { language: "en-US" }, pageSize: { width: 220, height: 180 }, orientation: "landscape", margins: 24 }));
  assert.equal(pdf.pages.length, 3);
  assert.match(pdf.pages[1].content, /0\.86 0 0 -0\.86 24 156 cm/);
  assert.match(pdf.pages[1].content, /\/Figure << \/MCID 0 >> BDC/);
  assert.ok(!pdf.pages[2].content.includes("/MCID"));
  assert.match(pdf.pages[2].content, /\/Artifact << \/Type \/Layout >> BDC/);
  assert.equal(pdfStructureElements(pdf).filter(([, object]) => /\/S \/Figure /.test(object.body)).length, 1);
  await assert.rejects(writePdfDocument([block([path])], { ...options, tagged: { language: "en-US" } }), /alternativeText/);
});

test("drawing text requires exact embedded faces and preserves the scalar writer's rejection rules", async () => {
  await assert.rejects(writePdfDocument([block([text])], { pageNumbers: false }), /exact embedded font/);
  for (const slot of ["bold", "italic", "boldItalic"])
    await assert.rejects(writePdfDocument([block([{ ...text, font: slot }])], { ...options, fonts: { regular: fonts.regular } }), /exact embedded font/);
  for (const [value, error] of [["Line\nBreak", /single baseline/], ["Line\u2028Break", /single baseline/], ["Tab\t", /control/], ["a\u0301", /shaping/], ["مرحبا", /shaping/], ["\ud800", /surrogate/], ["😀", /glyph/]])
    await assert.rejects(writePdfDocument([block([{ ...text, text: value }])], options), error);
});

test("portable drawing schemas reject unsupported fields, geometry, versions and cyclic/deep groups", async () => {
  for (const [value, error] of [
    [{ ...block([path]), opacity: .5 }, /drawing block properties/],
    [{ ...block([path]), drawing: { ...drawing([path]), transform: [1, 0, 0, 1, 0, 0] } }, /drawing properties/],
    [{ ...block([path]), drawing: { ...drawing([path]), version: 2 } }, /version 1/],
    [block([{ ...path, dash: [2, 2] }]), /path properties/],
    [block([{ ...text, rotation: 15 }]), /text properties/],
    [block([{ kind: "group", clip: { x: 0, y: 0, width: 10, height: 10, radius: 2 }, elements: [path] }]), /clip properties/],
    [block([{ ...path, commands: [{ op: "move", x: 0, y: 0, z: 1 }] }]), /command properties/],
    [block([{ ...path, commands: [{ op: "quadratic", x: 0, y: 0 }] }]), /Unsupported drawing path command/],
    [block([{ ...path, commands: [{ op: "line", x: 0, y: 0 }] }]), /begin with move/],
    [block([{ kind: "image", data: "external" }]), /element kind/],
    [block([{ ...text, x: Infinity }]), /coordinates/],
    [block([{ ...text, x: -1000001 }]), /coordinates/],
    [{ ...block([path]), drawing: { ...drawing([path]), width: 14401 } }, /Drawing width/],
    [block([{ ...path, lineCap: "triangle" }]), /lineCap/]
  ]) await assert.rejects(writePdfDocument([value], options), error);
  const cyclic = { kind: "group", elements: [] }; cyclic.elements.push(cyclic);
  await assert.rejects(writePdfDocument([block([cyclic])], options), /cycles/);
  let nested = path; for (let i = 0; i < 32; i++) nested = { kind: "group", elements: [nested] };
  await assert.rejects(writePdfDocument([block([nested])], options), /32 element levels/);
});

test("drawing resource budgets span blocks and share text and path accounting with existing blocks", async () => {
  const vector = { kind: "vector", width: 10, height: 10, stroke: "000000", commands: [{ op: "move", x: 0, y: 0 }, { op: "line", x: 10, y: 10 }] };
  await writePdfDocument([vector, block([path])], { ...options, limits: { maxVectorCommands: 4, maxDrawingElements: 1 } });
  for (const [blocks, limits, error] of [
    [[vector, block([path])], { maxVectorCommands: 3 }, /maxVectorCommands/],
    [[block([path]), vector], { maxVectorCommands: 3 }, /maxVectorCommands/],
    [[block([path]), block([path])], { maxDrawingElements: 1 }, /maxDrawingElements/],
    [[block([{ kind: "group", elements: [path] }])], { maxDrawingElements: 1 }, /maxDrawingElements/],
    [[block([text])], { maxCellCharacters: 3 }, /maxCellCharacters/],
    [[{ kind: "paragraph", text: "Before" }, block([text])], { maxTextCharacters: 9 }, /maxTextCharacters/],
    [[block(Array.from({ length: 50 }, () => path))], { maxPageBytes: 512 }, /maxPageBytes/]
  ]) await assert.rejects(writePdfDocument(blocks, { ...options, limits }), error);
});

test("drawing command and text getters retain the values admitted by their resource limits", async () => {
  const commands = [{ op: "move", x: 0, y: 0 }], extra = { op: "line", x: 9, y: 9 };
  let commandReads = 0, textReads = 0;
  const capturedPath = { ...path, get commands() { return ++commandReads === 1 ? commands : [...commands, extra]; } };
  const capturedText = { ...text, get text() { return ++textReads === 1 ? "A" : "BBBB"; } };
  const pdf = await inspectPdf(await writePdfDocument([block([capturedPath, capturedText]), block([path])], {
    ...options, limits: { maxVectorCommands: 3, maxCellCharacters: 1 }
  }));
  assert.equal(pdf.text, "A");
  assert.doesNotMatch(pdf.pages[0].content, /9 9 l/);
  assert.equal((pdf.pages[0].content.match(/(?:^|\n)0 0 m/g) ?? []).length, 1);
  await assert.rejects(writePdfDocument([block([{ ...path, get commands() { return [...commands, extra]; } }])], {
    ...options, limits: { maxVectorCommands: 1 }
  }), /maxVectorCommands/);
  await assert.rejects(writePdfDocument([block([{ ...text, get text() { return "BBBB"; } }])], {
    ...options, limits: { maxCellCharacters: 1 }
  }), /maxCellCharacters/);
});

test("drawing command operations, paint, font slots and group clips use their captured fields", async () => {
  let opReads = 0, fillReads = 0, strokeReads = 0, fontReads = 0, clipReads = 0, kindReads = 0;
  const capturedPath = { ...path, get kind() { return ++kindReads === 1 ? "path" : "text"; },
    get fill() { return ++fillReads === 1 ? "FF0000" : "0000FF"; },
    get stroke() { return ++strokeReads === 1 ? "008000" : "0000FF"; },
    commands: [{ get op() { return ++opReads === 1 ? "move" : "line"; }, x: 0, y: 0 }] };
  const capturedText = { ...text, text: "Regular", get font() { return ++fontReads === 1 ? "regular" : "bold"; } };
  const group = { kind: "group", get clip() { return ++clipReads === 1 ? { x: 3, y: 4, width: 12, height: 13 } : { x: Infinity }; }, elements: [capturedPath, capturedText] };
  const pdf = await inspectPdf(await writePdfDocument([block([group])], { ...options, fonts: { regular: fonts.regular } }));
  assert.equal(pdf.text, "Regular");
  assert.match(pdf.pages[0].content, /3 4 12 13 re W n/);
  assert.match(pdf.pages[0].content, /1 0 0 rg\n0 0\.502 0 RG/);
  assert.match(pdf.pages[0].content, /0 0 m\nB/);
  assert.doesNotMatch(pdf.pages[0].content, /0 0 1 [rR][gG]|0 0 l/);
});

test("drawing preflight and serialization yield to cancellation and release active source/destination ownership", async () => {
  const now = Object.getOwnPropertyDescriptor(performance, "now"), scheduler = globalThis.scheduler;
  try {
    for (const [drawingBlock, abortYield] of [[block([{ ...text, text: "A".repeat(10000) }]), 1],
      [block([{ ...path, commands: [{ op: "move", x: 0, y: 0 }, ...Array.from({ length: 3500 }, () => ({ op: "line", x: 1, y: 1 }))] }]), 4]]) {
      let clock = 0, yields = 0, returned = 0, writes = 0; const abort = new AbortController();
      Object.defineProperty(performance, "now", { configurable: true, value: () => clock += 40 });
      globalThis.scheduler = { postTask: async callback => { if (++yields === abortYield) abort.abort(Error("drawing cancelled")); callback(); } };
      const source = { *[Symbol.iterator]() { try { yield drawingBlock; } finally { returned++; } } };
      const destination = new WritableStream({ write() { writes++; } });
      await assert.rejects(writePdfDocumentTo(source, destination, { ...options, signal: abort.signal }), /drawing cancelled/);
      assert.equal(yields, abortYield); assert.equal(writes, 0); assert.equal(returned, 1); assert.equal(destination.locked, false);
    }
  } finally { if (now) Object.defineProperty(performance, "now", now); else delete performance.now; globalThis.scheduler = scheduler; }
});

test("drawing preflight snapshots validated path style before a cooperative yield", async () => {
  const now = Object.getOwnPropertyDescriptor(performance, "now"), scheduler = globalThis.scheduler;
  const mutable = { ...path, fill: "FFFFFF", fillRule: "evenodd", lineCap: "round", lineJoin: "bevel",
    commands: [{ op: "move", x: 0, y: 0 }, ...Array.from({ length: 1100 }, () => ({ op: "line", x: 1, y: 1 }))] };
  try {
    let clock = 0; Object.defineProperty(performance, "now", { configurable: true, value: () => clock += 40 });
    globalThis.scheduler = { postTask: async callback => { mutable.lineCap = "unsupported"; mutable.lineJoin = "unsupported"; mutable.fillRule = "nonzero"; callback(); } };
    const pdf = await inspectPdf(await writePdfDocument([block([{ kind: "group", clip: { x: -10000, y: -10000, width: 50000, height: 50000 }, elements: [mutable] }])], options));
    assert.match(pdf.pages[0].content, /1 J\n2 j\n10 M/);
    assert.match(pdf.pages[0].content, /B\*/);
    assert.match(pdf.pages[0].content, /-10000 -10000 50000 50000 re W n/);
  } finally { if (now) Object.defineProperty(performance, "now", now); else delete performance.now; globalThis.scheduler = scheduler; }
});
