import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { writePdf, writePdfTo, PdfFont, ExportCell } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const columns = [{ header: "Name", key: "name" }, { header: "Amount", key: "amount" }];
const fontBytes = new Uint8Array(await readFile(new URL("../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf", import.meta.url)));
const regular = new PdfFont(fontBytes);
test("PDF table API preserves projected display values, styling, grouped headings and numeric totals", async () => {
  const report = await writePdf([{ name: { label: "Łódź — Zażółć gęślą jaźń" }, amount: 12.5 }, { name: { label: "Δοκιμή Москва" }, amount: 7.5 }], {
    fonts: { regular }, title: "Raport Łódź", columns: [{ header: "Name", value: r => r.name.label, groups: ["Report"] },
      { key: "amount", header: "Amount", groups: ["Report"], value: r => new ExportCell(r.amount, { text: r.amount.toFixed(2), presentation: { color: "008000", background: "fff000", bold: true, italic: true } }) }],
    footer: { values: ["Totals"], totals: { amount: "sum" } }, compression: false
  });
  assert.equal(report.type, "application/pdf");
  const pdf = await inspectPdf(report);
  for (const text of ["Raport Łódź", "Łódź — Zażółć gęślą jaźń", "Δοκιμή Москва", "12.50", "7.50", "Totals", "20"]) assert.ok(pdf.text.includes(text), text);
  assert.match(pdf.pages[0].content, /2 Tr/); assert.match(pdf.pages[0].content, /1 0 0\.\d+ 1 [\d.]+ [\d.]+ Tm/);
  const embedded = [...pdf.objects.values()].filter(o => /\/Length1 /.test(o.body));
  assert.ok(embedded.length >= 1);
  for (const object of embedded) {
    assert.ok(object.stream.length < fontBytes.length / 2, "The export must contain a real font subset");
    let sum = 0;
    for (let i = 0; i < object.stream.length; i += 4) { const b = object.stream; sum = (sum + ((b[i] ?? 0) * 0x1000000 + ((b[i+1] ?? 0) << 16) + ((b[i+2] ?? 0) << 8) + (b[i+3] ?? 0))) >>> 0; }
    assert.equal(sum, 0xb1b0afba, "TrueType checksum adjustment");
  }
});

test("PDF paginates oversized rows without losing long words and repeats every heading", async () => {
  const long = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".repeat(400);
  const pdf = await inspectPdf(await writePdf([[long]], { columns: [{ header: "Heading", groups: ["Group"] }],
    pageSize: "A5", columnWidths: [160], fonts: { regular }, pageHeader: c => "Header " + c.pageNumber, pageFooter: "Footer", compression: false }));
  assert.ok(pdf.pages.length > 5);
  let body = "";
  for (let i = 0; i < pdf.pages.length; i++) {
    const lines = pdf.pages[i].lines;
    assert.deepEqual(lines.slice(0, 5), ["Header " + (i + 1), "Footer", "Page " + (i + 1) + " of ", "Group", "Heading"]);
    body += lines.slice(5).join("");
  }
  assert.equal(body, long);
});

test("PDF streams complete pages before producer exhaustion and releases borrowed destinations", async () => {
  let writes = 0, completed = false, returned = false;
  const chunks = [], destination = new WritableStream({ async write(bytes) { assert.equal(completed, false); writes++; chunks.push(bytes.slice()); } });
  async function* source() { try { for (let i = 0; i < 180; i++) { if (i === 150) assert.ok(writes > 0); yield { name: "Row " + i, amount: i }; } } finally { returned = true; } }
  const result = await writePdfTo(source(), destination, { columns, footer: { totals: { amount: "sum" } } }); completed = true;
  assert.equal(destination.locked, false); assert.equal(returned, true);
  assert.deepEqual(result, { rows: 180, columns: 2, bytes: chunks.reduce((n, b) => n + b.length, 0) });
  const pdf = await inspectPdf(new Blob(chunks)); assert.ok(pdf.text.includes("Row 179")); assert.ok(pdf.text.includes("16110"));
});

test("PDF native compression and missing-compressor fallback preserve the same page text", async () => {
  const compressed = await inspectPdf(await writePdf([["Hello €", 12]], { columns, compression: true }));
  const native = globalThis.CompressionStream;
  try {
    globalThis.CompressionStream = undefined;
    const fallback = await inspectPdf(await writePdf([["Hello €", 12]], { columns }));
    assert.equal(compressed.text, fallback.text);
    assert.ok(compressed.bytes.length < fallback.bytes.length);
    assert.ok(!fallback.bytes.includes(Buffer.from("/FlateDecode")));
  } finally { globalThis.CompressionStream = native; }
});

test("PDF resource policies fail explicitly without truncation and async callbacks are rejected", async () => {
  for (const limits of [{ maxRows: 0 }, { maxCells: 1 }, { maxTextCharacters: 1 }, { maxOutputBytes: 10 }, { maxPages: 0 }, { maxCellCharacters: 2 }, { maxRowLines: 0 }, { maxPageBytes: 10 }])
    await assert.rejects(writePdf([["Hello", 1]], { columns, limits }), /exceeded/);
  await assert.rejects(writePdf([["Łódź", 1]], { columns }), /TrueType/);
  await assert.rejects(writePdf([["مرحبا", 1]], { columns, fonts: { regular } }), /shaping/);
  await assert.rejects(writePdf([["e\u0301", 1]], { columns, fonts: { regular } }), /shaping/);
  await assert.rejects(writePdf([["\ud800", 1]], { columns, fonts: { regular } }), /surrogate/);
  await assert.rejects(writePdf([["x", 1]], { columns, formatValue: async () => { throw Error("observed"); } }), /synchronous/);
  await assert.rejects(writePdf([["x", 1]], { columns, pageHeader: async () => { throw Error("observed"); } }), /synchronous/);
  await assert.rejects(writePdf([["Hello", 1]], { columns, fonts: { regular }, limits: { maxFontBytes: 10 } }), /maxFontBytes/);
  await assert.rejects(writePdf([["Hello", 1]], { columns, columnWidths: [400, 400], wideTable: "reject" }), /printable/);
  await assert.rejects(writePdf([["Hello", 1]], { columns, pageSize: {width:80,height:200} }), /narrow/);
});

test("PDF cancellation returns a pending source and releases a hung sink's stream lock", async () => {
  const sourceAbort = new AbortController(); let returned = 0;
  const source = { [Symbol.asyncIterator]() { return { next() { setTimeout(() => sourceAbort.abort(Error("source abort")), 0); return new Promise(() => {}); }, return() { returned++; return Promise.resolve({done:true}); } }; } };
  await assert.rejects(writePdf(source, { columns, signal: sourceAbort.signal }), /source abort/); assert.equal(returned, 1);
  const sinkAbort = new AbortController(), destination = new WritableStream({ write() { setTimeout(() => sinkAbort.abort(Error("sink abort")), 0); return new Promise(() => {}); } });
  await assert.rejects(writePdfTo(Array.from({length:100}, () => ["Row", 1]), destination, { columns, signal: sinkAbort.signal }), /sink abort/);
  assert.equal(destination.locked, false);
});

test("PDF font programs are immutable and malformed/restricted fonts never reach a document", () => {
  assert.throws(() => new PdfFont(new Uint8Array([0,1,2])), /TrueType/);
  const bytes = fontBytes.slice(), view = new DataView(bytes.buffer);
  let os2;
  for (let i = 0; i < view.getUint16(4); i++) { const p = 12 + i * 16; if (String.fromCharCode(...bytes.subarray(p, p+4)) === "OS/2") os2 = view.getUint32(p+8); }
  view.setUint16(os2 + 8, 2);
  assert.throws(() => new PdfFont(bytes), /permissions/);
  const other = fontBytes.slice(); const safe = new PdfFont(other); other.fill(0);
  assert.ok(safe instanceof PdfFont);
});

test("PDF rectangular headings repeat and spanning multirow footers preserve every anchor", async () => {
  const headerRows = [[{ value: "Region", rowSpan: 2 }, { value: "Metrics" }], [null, { value: "Amount" }]];
  const footer = { rows: [[{ value: "Totals", rowSpan: 2 }, { value: "20" }], [null, { value: "Approved" }]] };
  const pdf = await inspectPdf(await writePdf(Array.from({ length: 90 }, (_, i) => ["Row " + i, i]), {
    columns, headerRows, footer, compression: false
  }));
  assert.ok(pdf.pages.length > 1);
  for (const page of pdf.pages) for (const text of ["Region", "Metrics", "Amount"]) assert.equal(page.lines.filter(t => t === text).length, 1);
  for (const text of ["Totals", "20", "Approved"]) assert.ok(pdf.pages.at(-1).lines.includes(text));
  assert.ok(pdf.text.includes("Row 89"));
  for (const bad of [
    [[null, { value: "x" }]],
    [[{ value: "x", columnSpan: 2 }, { value: "overlap" }]],
    [[{ value: "x", rowSpan: 2 }, { value: "y" }]],
    [[{ value: "x", columnSpan: 0 }, { value: "y" }]]
  ]) await assert.rejects(writePdf([], { columns, headerRows: bad }), /span|covered|column|overlap/i);
});

test("PDF full-row header spans and footer page breaks use the declared row heights", async () => {
  const pdf = await inspectPdf(await writePdf(Array.from({length:6}, (_, i) => ["Data" + i, i]), { columns, title: "Title", compression: false,
    pageSize: { width: 220, height: 240 }, margins: 25, fontSize: 8,
    headerRows: [[{ value: "Long wrapped heading across all cells and two rows", columnSpan: 2, rowSpan: 2 }, null], [null, null]],
    footer: { rows: [[{ value: "First footer" }, { value: "Value" }], [{ value: "Second footer" }, { value: "End" }]] }
  }));
  assert.equal(pdf.pages.length, 2); assert.ok(pdf.pages[0].lines.includes("Data5")); assert.ok(pdf.pages[1].lines.includes("End"));
  const content = pdf.pages[0].content;
  const headerBottom = Number(/25 ([\d.]+) 170 [\d.]+ re S/.exec(content)[1]);
  const dataBaseline = Number(/1 0 0 1 29 ([\d.]+) Tm <4461746130>/.exec(content)[1]);
  assert.ok(dataBaseline + 8 <= headerBottom, "Data text must start below both rows of the spanning heading");
});

test("PDF fonts and portable cells compose across standalone ESM assemblies without exposing mutable bytes", async () => {
  const standalone = await import('../bundles/officeimo-pdf.mjs');
  const integration = await import('../bundles/officeimo-datatables.mjs');
  const font = new standalone.PdfFont(fontBytes);
  const copy = font.toBytes(); copy.fill(0);
  assert.notEqual(font.toBytes()[1], 0);
  const report = await inspectPdf(await writePdf([[new standalone.ExportCell('Łódź')]], {columns:[{header:'City'}],fonts:{regular:font}}));
  assert.ok(report.text.includes('Łódź'));
  const other = new integration.ExportCell('Gdańsk');
  const second = await inspectPdf(await standalone.writePdf([[other]], {columns:[{header:'City'}],fonts:{regular}}));
  assert.ok(second.text.includes('Gdańsk'));
});
