import test from "node:test";
import assert from "node:assert/strict";
import { createGroupedReport, ExportCell } from "../dist/core/index.js";
import { writePdf, writePdfTo, writePdfDocument, pdfTable, pdfGroupedTable } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const columns = Array.from({ length: 7 }, (_, index) => ({ header: `Header-${index}`, key: `c${index}` }));
const base = { columns, columnWidths: columns.map(() => 80), wideTable: { mode: "bands", repeatColumns: [0] },
  pageSize: { width: 300, height: 360 }, margins: 20, fontSize: 9, padding: 4, pageNumbers: false, compression: false };
const data = count => Array.from({ length: count }, (_, row) => Object.fromEntries(columns.map((_, index) => [`c${index}`, index === 0 ? `ID-${row}` : `R${row}C${index}`])));

test("horizontal bands preserve readable widths and synchronized bounded windows with one source/value/format pass", async () => {
  let reads = 0, formats = 0, opened = 0, closed = 0;
  const rows = data(5);
  function* source() { opened++; try { for (const row of rows) yield Object.fromEntries(columns.map((_, index) => [`c${index}`, { value: row[`c${index}`] }])); } finally { closed++; } }
  const access = columns.map((column, index) => ({ ...column, value(row) { reads++; return row[`c${index}`].value; } }));
  const pdf = await inspectPdf(await writePdf(source(), { ...base, columns: access, limits: { maxTableWindowRows: 2 }, formatValue(value) { formats++; return String(value); } }));
  assert.equal(opened, 1); assert.equal(closed, 1); assert.equal(reads, 35); assert.equal(formats, 35);
  assert.equal(pdf.pages.length, 9, "Three column bands for each of the three logical windows");
  for (let window = 0; window < 3; window++) {
    const expected = rows.slice(window * 2, window * 2 + 2).map(row => row.c0);
    for (let band = 0; band < 3; band++) {
      const page = pdf.pages[window * 3 + band];
      assert.deepEqual(page.lines.filter(line => /^ID-/.test(line)), expected);
      assert.ok(page.lines.includes("Header-0"));
      for (const index of [band * 2 + 1, band * 2 + 2]) {
        assert.ok(page.lines.includes(`Header-${index}`));
        for (const row of rows.slice(window * 2, window * 2 + 2)) assert.ok(page.lines.includes(row[`c${index}`]));
      }
      assert.match(page.content, /80 [\d.]+ re S/, "Declared point widths are never fit-scaled");
    }
  }
  const nonleading = await inspectPdf(await writePdf(data(1), { ...base, wideTable: { mode: "bands", repeatColumns: [2] } }));
  assert.deepEqual(nonleading.pages[0].lines.slice(0, 3), ["Header-2", "Header-0", "Header-1"], "An explicitly repeated nonleading identifier becomes the physical prefix");
});

test("structured band headings and footers repeat captured intersections and retain vertical span geometry", async () => {
  const headers = [
    [{ value: "Identity", rowSpan: 2 }, { value: "Metrics", columnSpan: 6 }, null, null, null, null, null],
    [null, ...columns.slice(1).map(column => ({ value: column.header }))]
  ];
  const footer = [[{ value: "Summary", columnSpan: 7 }, null, null, null, null, null, null]];
  const pdf = await inspectPdf(await writePdf(data(1), { ...base, headerRows: headers, footer: { rows: footer } }));
  assert.equal(pdf.pages.length, 3);
  for (const [band, page] of pdf.pages.entries()) {
    assert.ok(page.lines.includes("Identity")); assert.ok(page.lines.includes("Metrics"));
    assert.equal(page.lines.filter(line => line === "Summary").length, band === 0 ? 1 : 2, "The span repeats at identifier and separate metric intersections");
    assert.ok(page.lines.includes("ID-0"));
  }
});

test("hierarchical composed bands repeat the full captured path while counts and numeric aggregates remain logical", async () => {
  const sourceColumns = [{ header: "ID", key: "id" }, ...Array.from({ length: 5 }, (_, index) => ({ header: `Value-${index}`, key: `v${index}`, type: "number" }))];
  const rows = Array.from({ length: 3 }, (_, row) => ({ id: `Service-${row}`, ...Object.fromEntries(Array.from({ length: 5 }, (_, index) => [`v${index}`, row + 1])) }));
  let read = 0;
  function* source() { for (const row of rows) { read++; yield row; } }
  const grouped = createGroupedReport(source(), { columns: sourceColumns, groupBy: [{ key: "region", value: () => "Region" }, { key: "site", value: () => "Site" }], totals: { v4: "sum" } });
  const table = pdfGroupedTable(grouped, { labelColumn: { header: "Summary" }, subtotalLabel: group => `Total-${group.level}-${group.rows}`, grandTotalLabel: summary => `Grand-${summary.rows}`,
    columnWidths: [80, 80, 80, 80, 80, 80, 80], wideTable: { mode: "bands", repeatColumns: [0, 1] } });
  const pdf = await inspectPdf(await writePdfDocument([{ kind: "paragraph", text: "Report" }, table], { ...base, columns: undefined, limits: { maxRows: 6, maxTableWindowRows: 2 } }));
  assert.equal(read, 3);
  const detail = pdf.pages.filter(page => page.lines.some(line => /^Service-/.test(line)));
  for (const page of detail.filter(page => !page.lines.includes("Grand-3"))) { assert.ok(page.lines.includes("Region")); assert.ok(page.lines.includes("Site")); }
  assert.ok(pdf.text.includes("Grand-3")); assert.ok(pdf.text.includes("6"));
  const grandPages = pdf.pages.filter(page => page.lines.includes("Grand-3"));
  assert.equal(grandPages.length, 5, "Every final metric band retains the logical grand-total label");
});

test("band limits count physical repeated cells, annotations and pages, and reject unsupported geometry before source consumption", async () => {
  for (const limits of [{ maxTableBands: 2 }, { maxTableWindowRows: 0 }, { maxPages: 2 }, { maxCells: 17 }, { maxHyperlinks: 2 }]) {
    let consumed = 0;
    function* source() { consumed++; yield { ...data(1)[0], c0: new ExportCell("ID-0", { link: { target: "https://example.com/0" } }) }; }
    await assert.rejects(writePdf(source(), { ...base, limits }), { code: "RESOURCE_LIMIT" });
    if (limits.maxTableBands !== undefined || limits.maxTableWindowRows !== undefined) assert.equal(consumed, 0);
  }
  for (const wideTable of [{ mode: "bands", repeatColumns: [1, 1] }, { mode: "bands", repeatColumns: [7] }, { mode: "bands", repeatColumns: [0, 1, 2, 3] }, { mode: "bands", minimumColumnWidth: NaN }])
    await assert.rejects(writePdf(data(1), { ...base, wideTable }), /repeatColumns|Repeated|width/);
  await assert.rejects(writePdf(data(1), { ...base, tagged: { language: "en" } }), /Tagged PDF horizontal bands/);
  await assert.rejects(writePdf(data(1), { ...base, wideTable: "reject" }), /widths exceed/);
  assert.equal((await inspectPdf(await writePdf(data(1), { ...base, wideTable: "fit" }))).pages.length, 1);
});

test("band policy accessors are captured once and sink backpressure/cancellation closes the source", async () => {
  const reads = { policy: 0, mode: 0, width: 0, repeat: 0 };
  const policy = { get mode() { reads.mode++; return "bands"; }, get minimumColumnWidth() { reads.width++; return 80; }, get repeatColumns() { reads.repeat++; return [0]; } };
  let produced = 0, closed = 0, release, entered;
  const blocked = new Promise(resolve => { entered = resolve; }), hold = new Promise(resolve => { release = resolve; });
  const controller = new AbortController(), reason = new Error("cancel band window");
  async function* source() { try { for (const row of data(100)) { produced++; yield row; } } finally { closed++; } }
  const pending = writePdfTo(source(), { async write() { if (produced) { entered(); await hold; } } }, { ...base,
    get wideTable() { reads.policy++; return policy; }, signal: controller.signal, limits: { maxTableWindowRows: 2 } });
  await blocked;
  const before = produced; await Promise.resolve(); assert.equal(produced, before); assert.ok(produced <= 3, "One bounded window reaches output before the source completes");
  controller.abort(reason); release(); await assert.rejects(pending, error => error === reason); assert.equal(closed, 1);
  assert.deepEqual(reads, { policy: 1, mode: 1, width: 1, repeat: 1 });
});


test("standalone band windows finish before bottom messages, preserving title and footer order", async () => {
  const pdf = await inspectPdf(await writePdf(data(1), { ...base, title: "Before-title", messageTop: "Before-table",
    footer: { values: ["Footer", ...columns.slice(1).map(() => "")] }, messageBottom: "After-table" }));
  assert.ok(pdf.text.indexOf("Before-title") < pdf.text.indexOf("Before-table"));
  assert.ok(pdf.text.indexOf("Before-table") < pdf.text.indexOf("ID-0"));
  assert.ok(pdf.text.lastIndexOf("Footer") < pdf.text.indexOf("After-table"));
  assert.ok(pdf.text.lastIndexOf("R0C6") < pdf.text.indexOf("After-table"));
});

test("long logical rows split into identical band windows and repeat short identifiers without losing detail lines", async () => {
  const row = data(1)[0], lines = Array.from({ length: 50 }, (_, index) => `Line-${index}`); row.c1 = lines.join("\n");
  const pdf = await inspectPdf(await writePdf([row], base));
  assert.ok(pdf.pages.length >= 6); assert.equal(pdf.pages.length % 3, 0, "Every vertical window emits all three horizontal bands");
  for (const page of pdf.pages) assert.ok(page.lines.includes("ID-0"), "The same identifier remains visible on every horizontal/vertical fragment");
  assert.deepEqual(pdf.pages.flatMap(page => page.lines).filter(line => /^Line-/.test(line)), lines);
});
