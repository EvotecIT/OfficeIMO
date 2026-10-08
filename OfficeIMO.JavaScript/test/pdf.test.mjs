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

test("empty report metadata preserves direct PDF layout and the data cell budget", async () => {
  const options = { columns, fonts: { regular }, compression: false, pageNumbers: false, limits: { maxCells: 4 } };
  const rows = [["Hello", 1]], baseline = await inspectPdf(await writePdf(rows, options));
  for (const metadata of [{ title: "" }, { messageTop: "" }, { messageBottom: "" }, { pageHeader: "" }, { pageFooter: "" },
    { pageHeader: () => "" }, { pageFooter: () => "" }, { title: "", messageTop: "", messageBottom: "", pageHeader: "", pageFooter: "" }]) {
    const report = await inspectPdf(await writePdf(rows, { ...options, ...metadata }));
    assert.deepEqual(report.pages.map(page => page.content), baseline.pages.map(page => page.content));
  }
  for (const key of ["title", "messageTop", "messageBottom", "pageHeader", "pageFooter"])
    await assert.rejects(writePdf(rows, { ...options, [key]: "Visible metadata" }), /maxCells/);
});

test("empty page decorations permit zero margins while visible decorations retain layout checks", async () => {
  const options = { columns, compression: false, pageNumbers: false, margins: 0, limits: { maxCells: 4 } };
  const baseline = await inspectPdf(await writePdf([["Hello", 1]], options));
  for (const key of ["pageHeader", "pageFooter"]) {
    for (const value of ["", () => ""]) {
      const actual = await inspectPdf(await writePdf([["Hello", 1]], { ...options, [key]: value }));
      assert.deepEqual(actual.pages.map(page => page.content), baseline.pages.map(page => page.content));
    }
    await assert.rejects(writePdf([["Hello", 1]], { ...options, [key]: "Visible" }), /margin/);
    await assert.rejects(writePdf([["Hello", 1]], { ...options, [key]: () => "Visible" }), /margin/);
    await assert.rejects(writePdf([["Hello", 1]], { ...options, [key]: async () => "" }), /synchronous/);
    await assert.rejects(writePdf([["Hello", 1]], { ...options, [key]: () => null }), /synchronous string/);
  }
  await assert.rejects(writePdf([["Hello", 1]], { ...options, pageNumbers: true }), /margin/);
  const visited = [];
  await assert.rejects(writePdf(Array.from({ length: 120 }, () => ["Hello", 1]), {
    ...options, limits: undefined, pageSize: "A5", pageHeader: ({ pageNumber }) => { visited.push(pageNumber); return pageNumber === 1 ? "" : "Visible"; }
  }), /margin/);
  assert.deepEqual(visited, [1, 2], "Callbacks are resolved once on each page, including later failures");
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

function fontTable(bytes, tag) {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  for (let i = 0; i < view.getUint16(4); i++) {
    const p = 12 + i * 16;
    if (String.fromCharCode(...bytes.subarray(p, p + 4)) === tag) return view.getUint32(p + 8);
  }
  throw Error("Missing fixture font table " + tag);
}

test("PDF subsets retain Unicode glyph IDs and composite metrics, including trailing short records", async () => {
  for (const short of [false, true]) {
    const source = fontBytes.slice(), view = new DataView(source.buffer), hhea = fontTable(source, "hhea"), hmtx = fontTable(source, "hmtx");
    const glyphCount = view.getUint16(fontTable(source, "maxp") + 4);
    if (short) {
      view.setUint16(hhea + 34, 2);
      source.fill(0, hmtx, hmtx + 8 + (glyphCount - 2) * 2);
      view.setUint16(hmtx, 600); view.setUint16(hmtx + 4, 800);
      for (let glyph = 2; glyph < glyphCount; glyph++) view.setInt16(hmtx + 8 + (glyph - 2) * 2, 12);
    }
    const text = "Łódź Δ Москва", pdf = await inspectPdf(await writePdf([[text]], { columns: [{ header: "" }],
      includeHeader: false, pageNumbers: false, compression: false, fonts: { regular: new PdfFont(source) } }));
    assert.ok(pdf.text.includes(text));
    const embedded = [...pdf.objects.values()].find(object => /\/Length1 /.test(object.body)).stream;
    const subset = new DataView(embedded.buffer, embedded.byteOffset, embedded.byteLength);
    const metricCount = view.getUint16(hhea + 34), subsetMetrics = fontTable(embedded, "hmtx"), subsetLoca = fontTable(embedded, "loca");
    const cmap = fontTable(embedded, "cmap"), mapping = cmap + subset.getUint32(cmap + 8);
    assert.equal(subset.getUint16(mapping), 12);
    const mapped = new Map();
    for (let i = 0; i < subset.getUint32(mapping + 12); i++) {
      const p = mapping + 16 + i * 12, first = subset.getUint32(p), last = subset.getUint32(p + 4), glyph = subset.getUint32(p + 8);
      for (let scalar = first; scalar <= last; scalar++) mapped.set(scalar, glyph + scalar - first);
    }
    assert.deepEqual([...mapped.keys()].sort((a,b) => a-b), [...new Set([...text].map(char => char.codePointAt(0)))].sort((a,b) => a-b));
    let outlines = 0;
    for (let glyph = 0; glyph < glyphCount; glyph++) {
      if (glyph && subset.getUint32(subsetLoca + glyph * 4) === subset.getUint32(subsetLoca + (glyph + 1) * 4)) continue;
      outlines++;
      const advance = Math.min(glyph, metricCount - 1) * 4;
      const bearing = glyph < metricCount ? glyph * 4 + 2 : metricCount * 4 + (glyph - metricCount) * 2;
      assert.equal(subset.getUint16(subsetMetrics + advance), view.getUint16(hmtx + advance));
      assert.equal(subset.getInt16(subsetMetrics + bearing), view.getInt16(hmtx + bearing));
    }
    assert.ok(outlines > mapped.size, "Composite dependencies retain outlines and their original metrics");
    for (const scalar of [...text]) assert.ok(mapped.get(scalar.codePointAt(0)) > 0);
  }
});

test("PDF honors the caller font's no-subsetting embedding permission", async () => {
  const source = fontBytes.slice(), view = new DataView(source.buffer);
  view.setUint16(fontTable(source, "OS/2") + 8, 0x100);
  const pdf = await inspectPdf(await writePdf([["Łódź"]], { columns: [{ header: "" }], includeHeader: false,
    pageNumbers: false, compression: false, fonts: { regular: new PdfFont(source) } }));
  const embedded = [...pdf.objects.values()].find(object => /\/Length1 /.test(object.body)).stream;
  assert.deepEqual(new Uint8Array(embedded), source);
});

test("PDF point and equal widths do not require a zero glyph from a sparse Unicode font", async () => {
  const bytes = fontBytes.slice(), view = new DataView(bytes.buffer), cmap = fontTable(bytes, "cmap");
  for (let i = 0; i < view.getUint16(cmap + 2); i++) {
    const offset = cmap + view.getUint32(cmap + 8 + i * 8);
    if (view.getUint16(offset) !== 4) continue;
    const segments = view.getUint16(offset + 6) / 2;
    for (let j = 0; j < segments; j++) {
      const end = offset + 14 + j * 2, start = offset + 16 + segments * 2 + j * 2;
      if (view.getUint16(start) <= 48 && view.getUint16(end) >= 48) {
        // Keep the Greek fixture glyphs; replace this ASCII segment with an absent zero.
        view.setUint16(start, 48); view.setUint16(end, 48);
        view.setInt16(offset + 16 + segments * 4 + j * 2, -48);
        view.setUint16(offset + 16 + segments * 6 + j * 2, 0);
      }
    }
  }
  const sparse = new PdfFont(bytes), options = { columns: [{ header: "" }], fonts: { regular: sparse }, includeHeader: false, pageNumbers: false };
  for (const widths of [undefined, [120]]) {
    const pdf = await inspectPdf(await writePdf([["Δ"]], { ...options, columnWidths: widths }));
    assert.equal(pdf.text, "Δ");
  }
  await assert.rejects(writePdf([["Δ"]], { ...options, columns: [{ header: "", width: 12 }] }), /no glyph for U\+30/);
});

test("PDF mixed font rows and paragraphs respect each selected face's vertical metrics", async () => {
  const bytes = fontBytes.slice(), view = new DataView(bytes.buffer), hhea = fontTable(bytes, "hhea"), head = fontTable(bytes, "head");
  const units = view.getUint16(head + 18);
  view.setInt16(hhea + 4, units * 2); view.setInt16(hhea + 6, -units / 2);
  const tall = new PdfFont(bytes);
  const pdf = await inspectPdf(await writePdf([[new ExportCell("Tall", { presentation: { bold: true } })], ["Next"]], {
    columns: [{ header: "Heading" }], fonts: { regular, bold: tall }, fontSize: 10, pageNumbers: false, compression: false,
    messageBottom: "End", headerPresentation: { bold: false }
  }));
  const content = pdf.pages[0].content;
  const rectangles = [...content.matchAll(/36 ([\d.]+) [\d.]+ ([\d.]+) re S/g)].map(m => ({ bottom: Number(m[1]), height: Number(m[2]) }));
  assert.ok(rectangles[1].height >= 33.99, "Tall face needs 26 points of line space plus padding");
  const baselines = [...content.matchAll(/1 0 0 1 40 ([\d.]+) Tm/g)].map(m => Number(m[1]));
  assert.ok(baselines[1] - 5 >= rectangles[1].bottom, "Tall face descent stays inside its row");
  assert.ok(baselines[2] < rectangles[1].bottom, "Following row starts below the tall face");
  const spanned = await inspectPdf(await writePdf([["Data", "Other"]], { columns, fonts: { regular, bold: tall }, fontSize: 10,
    title: "Title", pageNumbers: false, compression: false,
    headerRows: [[{ value: "Spanned", rowSpan: 2 }, { value: "Top" }], [null, { value: "Leaf" }]] }));
  const drawn = [...spanned.pages[0].content.matchAll(/1 0 0 1 ([\d.]+) ([\d.]+) Tm/g)].map(m => ({ x: Number(m[1]), y: Number(m[2]) }));
  const tableTop = [...spanned.pages[0].content.matchAll(/36 ([\d.]+) [\d.]+ ([\d.]+) re S/g)].map(m => ({ bottom: Number(m[1]), height: Number(m[2]) }));
  assert.ok(tableTop[0].bottom + tableTop[0].height <= drawn[0].y - 7.5, "Tall title descent ends before the spanning header");
  assert.ok(drawn.at(-2).y < tableTop[0].bottom, "Data starts below the complete spanning header");
  await assert.rejects(writePdf([], { columns, fonts: { regular, bold: tall }, title: "Too tall", includeHeader: false,
    pageNumbers: false, pageSize: { width: 200, height: 100 }, orientation: "landscape", margins: 20, fontSize: 30 }), /paragraph line/);
});

test("PDF page numbers fit the declared page bounds at large font sizes", async () => {
  const pdf = await inspectPdf(await writePdf([["Value"]], { columns: [{ header: "Heading" }], fontSize: 36,
    pageSize: { width: 700, height: 900 }, margins: { left: 36, right: 36, top: 36, bottom: 80 }, pageFooter: "End", compression: false }));
  const placement = /q 1 0 0 1 ([\d.]+) [\d.]+ cm \/TotalPages Do/.exec(pdf.pages[0].content);
  const form = [...pdf.objects.values()].find(o => /\/Subtype \/Form/.test(o.body));
  const width = Number(/\/BBox \[-1 [-\d.]+ ([\d.]+)/.exec(form.body)[1]);
  assert.ok(Number(placement[1]) + width <= 665, "Total page count stays within the right margin");
  assert.ok(pdf.text.includes("Page 1 of "));
  const many = await inspectPdf(await writePdf(Array.from({ length: 1000 }, () => ["Row"]), { columns: [{ header: "" }],
    includeHeader: false, fontSize: 16, pageSize: { width: 260, height: 100 }, orientation: "landscape",
    margins: { top: 20, bottom: 30, left: 16, right: 16 }, limits: { maxPages: 1024 }, compression: false }));
  assert.equal(many.pages.length, 1000);
  assert.ok(many.pages.at(-1).lines.includes("Page 1000 of "));
});

test("PDF message-only continuation pages omit table headings and retain decorations", async () => {
  const message = "Afterword ".repeat(1000);
  const pdf = await inspectPdf(await writePdf([["Data"]], { columns: [{ header: "Table heading" }], pageSize: "A5", compression: false,
    pageHeader: "Report", pageFooter: "Footer", messageBottom: message }));
  assert.ok(pdf.pages.length > 2);
  assert.equal(pdf.pages[0].lines.filter(t => t === "Table heading").length, 1);
  for (const page of pdf.pages.slice(1)) {
    assert.ok(!page.lines.includes("Table heading"));
    assert.ok(page.lines.includes("Report") && page.lines.includes("Footer"));
  }
  assert.equal(pdf.pages.flatMap(p => p.lines).filter(t => t.includes("Afterword")).join("").replaceAll(" ", ""), message.replaceAll(" ", ""));
});
