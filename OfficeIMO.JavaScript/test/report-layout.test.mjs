import test from "node:test";
import assert from "node:assert/strict";
import { writeFile, mkdir } from "node:fs/promises";
import { createHash } from "node:crypto";
import { createWorkbook, ExportCell, Cell } from "../dist/xlsx/index.js";
import { writeCsv } from "../dist/csv/index.js";
import { readZip } from "./zip-reader.mjs";

const columns = [
  { header: "Name", key: "name", groups: ["Identity"], width: 24 },
  { header: "Amount", key: "amount", groups: ["Metrics", "Money"], type: "number", format: "0.00" },
  { header: "Count", key: "count", groups: ["Metrics", "Money"], type: "number" },
  { header: "Date", key: "date", groups: ["Metrics", "Time"], type: "date", format: "yyyy-mm-dd" }
];
test("resolved presentation preserves typed values and supports explicit CSV raw/display modes", async () => {
  const value = new ExportCell(12.5, { text: "=display", presentation: { background: "FF0000", bold: true } });
  const book = createWorkbook(), sheet = book.addSheet("Values", { columns: [{ header: "Amount", type: "number", format: "0.00" }] });
  await sheet.addRows([[value]]);
  const zip = await readZip(await book.toBlob());
  assert.match(zip.get("xl/worksheets/sheet1.xml").content, /<c r="A2" s="\d+"><v>12.5<\/v>/);
  assert.match(zip.get("xl/styles.xml").content, /formatCode="0.00"/);
  assert.equal(await (await writeCsv([[value]], { columns: [{ header: "Amount" }] })).text(), "Amount\r\n12.5\r\n");
  assert.equal(await (await writeCsv([[value]], { columns: [{ header: "Amount" }], valueMode: "display" })).text(), "Amount\r\n'=display\r\n");
});
test("grouped report headings, cached totals, bounded sizes and print settings form a coherent table", async () => {
  const book = createWorkbook({ dateMode: "utc" });
  const sheet = book.addSheet("Report's _x0041_", { columns, table: { name: "ReportData" }, freezeHeader: true, freezeColumns: 1,
    autoSize: { sampleRows: 2, minWidth: 8, maxWidth: 30 }, footer: { values: ["Totals"], totals: { amount: "sum", count: "average" }, style: { font: { bold: true }, fill: { color: "D9E1F2" } } },
    print: { repeatHeaders: true, header: "Report & data", footer: "Measured totals", paper: "A4", orientation: "landscape" } });
  await sheet.addRows([{ name: "first", amount: new ExportCell(12.5, { text: "12.50 USD", presentation: { background: "E2F0D9" } }), count: 2, date: new Date("2026-10-06T00:00:00Z") }]);
  await sheet.addRows([{ name: "second", amount: 7.5, count: 4, date: new Date("2026-10-07T00:00:00Z") }, { name: "last", amount: null, count: null, date: null }]);
  const blob = await book.toBlob(), zip = await readZip(blob), xml = zip.get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /ySplit="3" topLeftCell="B4"/);
  assert.match(xml, /<mergeCell ref="B1:D1"\/>/); assert.match(xml, /<mergeCell ref="B2:C2"\/>/);
  assert.match(xml, /width="24"/); assert.match(xml, /<c r="B7"[^>]*><f>SUBTOTAL\(109,B4:B6\)<\/f><v>20<\/v>/);
  assert.match(xml, /<c r="C7"[^>]*><f>IF\(COUNT\(C4:C6\)=0,&quot;&quot;,SUBTOTAL\(101,C4:C6\)\)<\/f><v>3<\/v>/);
  assert.match(xml, /<pageSetUpPr fitToPage="1"/); assert.match(xml, /paperSize="9" orientation="landscape" fitToWidth="1" fitToHeight="0"/);
  assert.match(xml, /Report &amp;&amp; data/);
  assert.match(zip.get("xl/tables/table1.xml").content, /ref="A3:D7" totalsRowCount="1"><autoFilter ref="A3:D6"/);
  assert.match(zip.get("xl/tables/table1.xml").content, /name="Count" totalsRowFunction="custom"><totalsRowFormula>IF\(COUNT\(C4:C6\)=0,&quot;&quot;,SUBTOTAL\(101,C4:C6\)\)/);
  assert.match(zip.get("xl/workbook.xml").content, /_xlnm.Print_Titles[^>]*>&apos;Report&apos;&apos;s _x005F_x0041_&apos;!\$1:\$3/);
  if (process.env.OFFICEIMO_REPORT_FIXTURES) { await mkdir(process.env.OFFICEIMO_REPORT_FIXTURES, { recursive: true }); await writeFile(process.env.OFFICEIMO_REPORT_FIXTURES + "/layout.xlsx", new Uint8Array(await blob.arrayBuffer())); }
});
test("oversized Unicode is preserved in ordered safe chunks and linked without external relationships", async () => {
  const text = "a".repeat(32766) + "🧪" + "Łódź\r\nשלום_x0041_".repeat(2500);
  const book = createWorkbook({ oversizedText: "preserve" }), sheet = book.addSheet("Text overflow", { columns: [{ header: "Text" }] });
  await sheet.addRows([[text]]); const blob = await book.toBlob(), zip = await readZip(blob);
  const main = zip.get("xl/worksheets/sheet1.xml").content, overflow = zip.get("xl/worksheets/sheet2.xml").content;
  assert.match(main, /location="&apos;Text overflow \(2\)&apos;!D2"/); assert.match(main, /Full text:/);
  const chunks = [...overflow.matchAll(/<c r="D(?:[2-9]|\d{2,})"[^>]*>[\s\S]*?<\/c>/g)].map(m => [...m[0].matchAll(/<t[^>]*>([\s\S]*?)<\/t>/g)].map(t => t[1].replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code))).replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">")).join(""));
  const hash = value => createHash("sha256").update(value).digest("hex");
  assert.equal(hash(chunks.join("")), hash(text)); assert.ok(chunks.every(chunk => chunk.length <= 32767 && !/[\ud800-\udbff]$/.test(chunk)));
  assert.ok(!zip.has("xl/worksheets/_rels/sheet1.xml.rels"));
  if (process.env.OFFICEIMO_REPORT_FIXTURES) await writeFile(process.env.OFFICEIMO_REPORT_FIXTURES + "/preserved.xlsx", new Uint8Array(await blob.arrayBuffer()));
});
test("layout bounds and preservation ceilings fail visibly without a partial Blob", async () => {
  assert.throws(() => createWorkbook().addSheet("Bad", { columns, autoSize: { sampleRows: 10001 } }), RangeError);
  assert.throws(() => createWorkbook().addSheet("Bad", { columns, footer: { totals: { unknown: "sum" } } }), TypeError);
  for (const options of [{ oversizedText: "preserve", limits: { maxOverflowCharacters: 32767 } }, { oversizedText: "preserve", limits: { maxHyperlinks: 0 } }, { oversizedText: "preserve", limits: { maxSheets: 1 } }]) {
    const book = createWorkbook(options);
    await assert.rejects(book.addSheet("Long", { columns: [{ header: "Text" }] }).addRows([["x".repeat(32768)]]), { code: "RESOURCE_LIMIT" });
    await assert.rejects(book.toBlob(), { code: "RESOURCE_LIMIT" });
  }
  const book = createWorkbook({ limits: { maxBufferedCells: 1 } });
  await assert.rejects(book.addSheet("Sample", { columns: [{ header: "A" }, { header: "B" }], autoSize: {} }).addRows([[1, 2]]), { code: "RESOURCE_LIMIT" });
});
test("empty averages stay blank and footer keys do not traverse object prototypes", async () => {
  const book = createWorkbook(), sheet = book.addSheet("Empty", { columns: [{ header: "Average", key: "average" }, { header: "toString" }], footer: { totals: { average: "average" } } });
  const zip = await readZip(await book.toBlob()), xml = zip.get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /<c r="A2"[^>]*t="str"><f>&quot;&quot;<\/f><v><\/v>/); assert.match(xml, /<c r="B2"[^>]*\/>/);
  assert.equal(sheet.rowCount, 0);
});
