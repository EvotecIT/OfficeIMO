import test from "node:test";
import assert from "node:assert/strict";
import { writePdfDocument, writePdfDocumentTo, pdfTable } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const options = { pageSize: { width: 240, height: 180 }, margins: 20, pageNumbers: false, compression: false, tagged: { language: "en" } };
const columns = [{ header: "Asset", key: "name" }];
function elements(pdf, role) {
  return [...pdf.objects.entries()].filter(([, object]) => new RegExp("/Type /StructElem /S /" + role + " ").test(object.body));
}

test("streamed numbered findings keep hanging bodies, one marker and one semantic item across page breaks", async () => {
  let calls = 0, returned = 0;
  async function* items() {
    calls++;
    try {
      for (let i = 0; i < 4; i++) yield { text: [
        { text: "Finding " + i + ": ", bold: true },
        { text: "Long remediation body that wraps onto additional lines. ".repeat(3) },
        { text: "Open action " + i, link: { target: "https://example.com/action/" + i } }
      ] };
    } finally { returned++; }
  }
  const pdf = await inspectPdf(await writePdfDocument([
    { kind: "section", title: "Findings" },
    { kind: "list", ordered: true, start: 7, items: items(), itemStyle: "finding" }
  ], { ...options, paragraphStyles: { finding: { fontSize: 10, spaceAfter: 8 } } }));
  assert.equal(calls, 1); assert.equal(returned, 1); assert.ok(pdf.pages.length >= 3);
  for (let i = 7; i <= 10; i++) assert.equal(pdf.pages.flatMap(page => page.lines).filter(line => line === i + ".").length, 1);
  assert.equal(elements(pdf, "L").length, 1); assert.equal(elements(pdf, "LI").length, 4);
  assert.equal(elements(pdf, "Lbl").length, 4); assert.equal(elements(pdf, "LBody").length, 4);
  for (const [id, label] of elements(pdf, "Lbl")) {
    const parent = Number(/\/P (\d+) 0 R/.exec(label.body)[1]);
    assert.ok(elements(pdf, "LI").some(([item]) => item === parent), "Every label belongs to its list item");
    assert.ok([...pdf.objects.values()].some(object => object.body.includes("/K [") && object.body.includes(id + " 0 R")));
  }
  const itemText = pdf.pages.flatMap(page => [...page.content.matchAll(/1 0 0 1 ([\d.]+) [\d.]+ Tm <(4c6f6e672072656d6564696174696f6e|626f64792074686174|6164646974696f6e616c206c696e6573)/g)]);
  assert.ok(itemText.length > 0); assert.ok(itemText.every(match => Number(match[1]) >= 40), "Continuation text retains its hanging body offset");
  assert.equal([...pdf.objects.values()].filter(object => /\/Subtype \/Link/.test(object.body)).length, 4);
});

test("list limits, invalid keeps and cancellation close item and block sources without retaining the destination", async () => {
  let itemsReturned = 0, blocksReturned = 0;
  async function* items() { try { yield { text: "One" }; yield { text: "Two" }; } finally { itemsReturned++; } }
  async function* blocks() { try { yield { kind: "list", items: items() }; yield { kind: "paragraph", text: "Unreached" }; } finally { blocksReturned++; } }
  const sink = new WritableStream({ write() {} });
  await assert.rejects(writePdfDocumentTo(blocks(), sink, { ...options, limits: { maxListItems: 1 } }), /maxListItems/);
  assert.equal(itemsReturned, 1); assert.equal(blocksReturned, 1); assert.equal(sink.locked, false);
  await assert.rejects(writePdfDocument([{ kind: "list", items: [{ text: "One" }] }, { kind: "list", items: [{ text: "Two" }] }],
    { ...options, limits: { maxListItems: 1 } }), /maxListItems/);
  await assert.rejects(writePdfDocument([{ kind: "list", items: [{ text: "One" }], itemStyle: "kept" }],
    { ...options, paragraphStyles: { kept: { keepWithNext: true } } }), /keepWithNext/);
  await assert.rejects(writePdfDocument([{ kind: "list", ordered: true, start: 100, indent: 1, items: [{ text: "One" }] }], options), /marker and gap/);
  await assert.rejects(writePdfDocument([{ kind: "list", items: [{ text: "Line\n".repeat(30), keepTogether: true }] }], options), /exceeds a fresh page/);
  const controller = new AbortController(), reason = Error("Finding export cancelled"); let closed = 0;
  async function* many() { try { for (let i = 0; i < 1000; i++) yield { text: "Finding " + i }; } finally { closed++; } }
  const destination = new WritableStream({ write() { controller.abort(reason); } });
  await assert.rejects(writePdfDocumentTo([{ kind: "list", items: many() }], destination, { ...options, signal: controller.signal }), error => error === reason);
  assert.equal(closed, 1); assert.equal(destination.locked, false);
});

test("table captions snapshot before enumeration and travel once with their first heading and data fragment", async () => {
  const caption = { text: "Table 1. Selected assets", style: "caption" };
  function* rows() { caption.text = "Mutated caption"; for (let i = 0; i < 18; i++) yield { name: "NODE-" + i }; }
  const pdf = await inspectPdf(await writePdfDocument([
    { kind: "paragraph", text: "Opening summary", spaceAfter: 90 },
    { ...pdfTable(rows(), { columns }), caption }
  ], { ...options, paragraphStyles: { caption: { fontSize: 9, spaceAfter: 4 } } }));
  const page = pdf.pages.find(page => page.lines.includes("Table 1. Selected assets"));
  assert.ok(page); assert.ok(page.lines.includes("Asset")); assert.ok(page.lines.includes("NODE-0"));
  assert.equal(pdf.pages.flatMap(page => page.lines).filter(line => line === "Table 1. Selected assets").length, 1);
  assert.ok(!pdf.text.includes("Mutated caption")); assert.ok(pdf.pages.filter(page => page.lines.includes("Asset")).length > 1);
  const [tableId] = elements(pdf, "Table")[0];
  assert.equal(elements(pdf, "Caption").length, 1);
  assert.match(elements(pdf, "Caption")[0][1].body, new RegExp("/P " + tableId + " 0 R"));
});

test("captioned image, vector and portable drawing retain actual Figure ownership and caption capture", async () => {
  const caption = { text: "Figure 1. Service status" };
  const commands = [{ op: "move", x: 0, y: 0 }, { op: "line", x: 20, y: 10 }];
  const pdf = await inspectPdf(await writePdfDocument([
    { kind: "image", alternativeText: "One green service", width: 20, height: 20, caption,
      image: { format: "rgb", width: 1, height: 1, get data() { caption.text = "Changed after capture"; return Uint8Array.of(0, 128, 0); } } },
    { kind: "vector", width: 20, height: 10, commands, stroke: "008000", alternativeText: "Status trend", caption: { text: "Figure 2. Trend" } },
    { kind: "drawing", drawing: { version: 1, width: 20, height: 10, elements: [{ kind: "path", commands, stroke: "008000" }] },
      alternativeText: "Owner-produced status chart", caption: { text: "Figure 3. Chart" } }
  ], options));
  assert.ok(pdf.text.includes("Figure 1. Service status")); assert.ok(!pdf.text.includes("Changed after capture"));
  const figures = new Set(elements(pdf, "Figure").map(([id]) => id)), captions = elements(pdf, "Caption");
  assert.equal(figures.size, 3); assert.equal(captions.length, 3);
  for (const [, caption] of captions) assert.ok(figures.has(Number(/\/P (\d+) 0 R/.exec(caption.body)[1])));
  for (const graphic of [{ kind: "image", image: { format: "rgb", width: 1, height: 1, data: Uint8Array.of(0, 0, 0) } },
    { kind: "vector", width: 20, height: 10, commands, stroke: "008000" }])
    await assert.rejects(writePdfDocument([{ ...graphic, decorative: true, caption: { text: "Cannot describe decoration" } }], options), /cannot be decorative/);
});


test("a caption retains the first fragment of a continued data row or ordinary footer", async () => {
  const text = "FIRST\nSecond\nThird\nFourth\nFifth";
  for (const footer of [false, true]) {
    const table = pdfTable(footer ? [] : [{ name: text }], { columns, ...(footer ? { footer: { values: [text] } } : {}) });
    const pdf = await inspectPdf(await writePdfDocument([
      { kind: "paragraph", text: "Before\nBefore" }, { ...table, caption: { text: "Table caption" } }
    ], { ...options, pageSize: { width: 240, height: 160 }, orientation: "landscape" }));
    const page = pdf.pages.find(page => page.lines.includes("Table caption"));
    assert.ok(page.lines.includes("Asset")); assert.ok(page.lines.includes("FIRST"), "Caption and heading cannot be stranded before the first visible value");
    assert.ok(pdf.text.includes("Fifth"));
  }
});
