import test from "node:test";
import assert from "node:assert/strict";
import { writeFile } from "node:fs/promises";
import { createWorkbook, Cell, ExportCell } from "../dist/xlsx/index.js";
import { readZip } from "./zip-reader.mjs";

test("titles shift grouped tables, totals, frozen panes and print repetition together", async () => {
  for (const streamed of [false, true]) {
    const chunks = [], book = createWorkbook({ ...(streamed ? { sink: { write(bytes) { chunks.push(bytes.slice()); } } } : {}), limits: { maxMergedRanges: 2, maxCells: 18, maxTextCharacters: 55 } });
    const title = { text: "Łódź 🧪 report", style: { fill: { color: "D9E1F2" }, font: { size: 20 } }, height: 32 };
    const sheet = book.addSheet("Quarter's report", { title, columns: [{ header: "Name", groups: ["Metrics"] }, { header: "Amount", key: "amount", groups: ["Metrics"], type: "number", format: "0.00" }, { header: "Count", type: "number", groups: ["Status"] }], table: { name: "Report" }, freezeHeader: true, freezeColumns: 1, autoSize: { sampleRows: 1 }, footer: { values: ["Total"], totals: { amount: "sum" } }, print: { repeatHeaders: true } });
    title.text = "changed"; title.style.font.size = 1;
    await sheet.addRows([["one", 12, 1], ["two", 8, 2]]);
    sheet.addHyperlink({ cell: "A1", target: "https://example.com/" });
    let blob;
    if (streamed) { await sheet.close(); await book.finish(); blob = new Blob(chunks); } else blob = await book.toBlob();
    const zip = await readZip(blob), xml = zip.get("xl/worksheets/sheet1.xml").content;
    assert.match(xml, /Łódź 🧪 report/); assert.doesNotMatch(xml, /changed/);
    assert.match(xml, /<row r="1" ht="32" customHeight="1">/);
    assert.match(xml, /ySplit="3" topLeftCell="B4"/);
    assert.match(xml, /mergeCell ref="A1:C1"/); assert.match(xml, /mergeCell ref="A2:B2"/);
    assert.match(xml, /<f>SUBTOTAL\(109,B4:B5\)<\/f><v>20<\/v>/);
    assert.match(zip.get("xl/tables/table1.xml").content, /ref="A3:C6"/);
    assert.match(zip.get("xl/workbook.xml").content, /\$2:\$3/);
    assert.match(zip.get("xl/styles.xml").content, /<sz val="20"/);
    if (!streamed && process.env.OFFICEIMO_REPORT_FIXTURES) await writeFile(process.env.OFFICEIMO_REPORT_FIXTURES + "/title-table.xlsx", new Uint8Array(await blob.arrayBuffer()));
  }
});

test("explicit horizontal and vertical merges preserve their anchors in sampled rows", async () => {
  const ranges = ["A1:B2", "A3:B3"], book = createWorkbook();
  const sheet = book.addSheet("Regions", { columns: [{ header: "A" }, { header: "B" }, { header: "C", type: "number" }], includeHeader: false, mergedCells: ranges, autoSize: { sampleRows: 3 } });
  ranges[0] = "A1:C3";
  await sheet.addRows([[new ExportCell("top", { presentation: { bold: true } }), null, 10], [null, new Cell(""), 20], ["bottom", undefined, 30]]);
  const blob = await book.toBlob(), xml = (await readZip(blob)).get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /mergeCell ref="A1:B2"/); assert.match(xml, /mergeCell ref="A3:B3"/);
  assert.match(xml, /r="C3"[^>]*><v>30<\/v>/);
  if (process.env.OFFICEIMO_REPORT_FIXTURES) await writeFile(process.env.OFFICEIMO_REPORT_FIXTURES + "/merged-regions.xlsx", new Uint8Array(await blob.arrayBuffer()));
});

test("merge validation rejects overlaps, table intersections, bounds and hidden data", async () => {
  const options = { columns: [{ header: "A" }, { header: "B" }, { header: "C" }], includeHeader: false };
  for (const ranges of [["A1:B2", "B2:C3"], ["A1:A1"], ["B2:A1"], ["a1:B2"], ["A1:D2"], ["A1:B1048577"]]) assert.throws(() => createWorkbook().addSheet("Bad", { ...options, mergedCells: ranges }));
  assert.throws(() => createWorkbook().addSheet("Bad", { columns: options.columns, table: {}, mergedCells: ["A2:B2"] }), /native table/);
  assert.throws(() => createWorkbook().addSheet("Bad", { columns: options.columns, title: { text: "title" }, mergedCells: ["B1:C1"] }), /overlap/);
  for (const value of ["hidden", 0, false, new Date("2026-10-06"), new ExportCell("hidden"), new Cell("hidden")]) {
    const book = createWorkbook(), sheet = book.addSheet("Hidden", { ...options, autoSize: {}, mergedCells: ["A1:B1"] });
    await assert.rejects(sheet.addRows([["anchor", value, null]]), /would hide/);
    await assert.rejects(book.finish(), /would hide/);
  }
  const book = createWorkbook(), sheet = book.addSheet("Bounds", { ...options, mergedCells: ["A1:B2"] });
  await sheet.addRows([["anchor", null, 1]]); await assert.rejects(book.toBlob(), /exported rows/);
  assert.throws(() => createWorkbook().addSheet("Link", { ...options, mergedCells: ["A1:B1"], hyperlinks: [{ cell: "B1", target: "https://example.com/" }] }), /covered/);
  const links = createWorkbook().addSheet("Link", { ...options, mergedCells: ["A1:B1"] });
  assert.throws(() => links.addHyperlink({ cell: "B1", target: "https://example.com/" }), /covered/);
});

test("titles without leaf headers and empty reports retain exact layout budgets", async () => {
  const book = createWorkbook({ limits: { maxCells: 4, maxTextCharacters: 6, maxMergedRanges: 1 } });
  await book.addSheet("Title", { columns: [{ header: "unused" }, { header: "unused" }], title: { text: "Title" }, includeHeader: false }).addRows([[1, 2]]);
  const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /r="A2"[^>]*><v>1<\/v>/); assert.doesNotMatch(xml, /unused/);
  const limited = createWorkbook({ limits: { maxMergedRanges: 1 } });
  limited.addSheet("One", { columns: [{ header: "A" }, { header: "B" }], title: { text: "first" } });
  assert.throws(() => limited.addSheet("Two", { columns: [{ header: "A" }, { header: "B" }], title: { text: "second" } }), { code: "RESOURCE_LIMIT" });
  limited.addSheet("Two", { columns: [{ header: "A" }] }); // rejected registration leaves the name available
  await limited.finish();
  const text = createWorkbook({ limits: { maxTextCharacters: 4 } });
  await assert.rejects(text.addSheet("Title", { columns: [{ header: "A" }], title: { text: "Title" }, includeHeader: false }).addRows([]), { code: "RESOURCE_LIMIT" });
});

test("covered empty strings are blank cells in declared typed columns", async () => {
  for (const type of ["number", "date", "boolean"]) for (const empty of ["", new Cell(""), new ExportCell("")]) {
    const book = createWorkbook(), sheet = book.addSheet("Blank", { columns: [{ header: "A" }, { header: "B", type }], includeHeader: false, mergedCells: ["A1:B1"] });
    await sheet.addRows([["anchor", empty]]);
    const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
    assert.match(xml, /<c r="B1" s="\d+"\/>/);
  }
  let calls = 0;
  const book = createWorkbook({ cellValueWriters: { milliseconds: value => { calls++; return Number(value) / 1000; } } });
  const sheet = book.addSheet("Custom", { columns: [{ header: "A" }, { header: "B", type: "milliseconds" }, { header: "C", type: "milliseconds" }], includeHeader: false, mergedCells: ["A1:B8"], autoSize: {} });
  const empties = ["", null, undefined, new Cell(""), new Cell(null), new ExportCell(""), new ExportCell(null), new ExportCell(undefined)];
  await sheet.addRows(empties.map((empty, i) => [i ? null : "anchor", empty, 1000]));
  const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
  assert.equal(calls, 8); assert.equal([...xml.matchAll(/r="C\d+"[^>]*><v>1<\/v>/g)].length, 8);
  const reject = createWorkbook({ cellValueWriters: { domain: () => 0 } });
  await assert.rejects(reject.addSheet("Hidden", { columns: [{ header: "A" }, { header: "B", type: "domain" }], includeHeader: false, mergedCells: ["A1:B1"] }).addRows([["anchor", "nonempty"]]), /would hide/);
});
