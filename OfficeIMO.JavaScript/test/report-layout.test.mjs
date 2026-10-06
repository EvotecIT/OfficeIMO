import test from "node:test";
import assert from "node:assert/strict";
import { writeFile, mkdir } from "node:fs/promises";
import { createHash } from "node:crypto";
import { createWorkbook, ExportCell, Cell } from "../dist/xlsx/index.js";
import { writeCsv } from "../dist/csv/index.js";
import { readZip } from "./zip-reader.mjs";
import * as standaloneXlsx from "../bundles/officeimo-xlsx.mjs";
import * as standaloneCsv from "../bundles/officeimo-csv.mjs";

const columns = [
  { header: "Name", key: "name", groups: ["Identity"], width: 24 },
  { header: "Amount", key: "amount", groups: ["Metrics", "Money"], type: "number", format: "0.00" },
  { header: "Count", key: "count", groups: ["Metrics", "Money"], type: "number" },
  { header: "Date", key: "date", groups: ["Metrics", "Time"], type: "date", format: "yyyy-mm-dd" }
];
test("standalone module assemblies share resolved cells across independently loaded exports", async () => {
  const value = new standaloneXlsx.ExportCell(12.5, { text: "12.50 USD" });
  assert.equal(await (await standaloneCsv.writeCsv([[value]], { columns: [{ header: "Amount" }], valueMode: "display" })).text(), "Amount\r\n12.50 USD\r\n");
});
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

test("sampling validates and budgets each row at ingress and calls converters once", async () => {
  for (const limits of [{ maxCells: 2 }, { maxTextCharacters: 2 }]) {
    let produced = 0, returned = false;
    function* source() { try { for (let i = 0; i < 100; i++) { produced++; yield ["a"]; } } finally { returned = true; } }
    const book = createWorkbook({ limits });
    await assert.rejects(book.addSheet("Sample", { columns: [{ header: "A" }], autoSize: { sampleRows: 100 } }).addRows(source()), { code: "RESOURCE_LIMIT" });
    assert.equal(produced, 2); assert.equal(returned, true); await assert.rejects(book.toBlob(), { code: "RESOURCE_LIMIT" });
  }
  const invalid = createWorkbook();
  await assert.rejects(invalid.addSheet("Invalid", { columns: [{ header: "A" }], autoSize: {} }).addRows([[{}]]), TypeError);
  let calls = 0;
  const book = createWorkbook({ cellValueWriters: { custom(value) { calls++; return value + 1; } } });
  const sheet = book.addSheet("Once", { columns: [{ header: "A", type: "custom" }], autoSize: { sampleRows: 100 }, footer: { totals: { A: "sum" } } });
  await sheet.addRows([[1], [2]]); await sheet.close(); await sheet.close();
  const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
  assert.equal(calls, 2); assert.match(xml, /<c r="A4"[^>]*><f>SUBTOTAL\(109,A2:A3\)<\/f><v>5<\/v>/);
});

test("preservation, footer and Blob output limits reject while accepting source rows", async () => {
  for (const limits of [{ maxTextCharacters: 40000 }, { maxCells: 10 }, { maxRows: 1 }]) {
    const book = createWorkbook({ oversizedText: "preserve", limits });
    await assert.rejects(book.addSheet("Long", { columns: [{ header: "A" }], autoSize: {} }).addRows([["x".repeat(100000)]]), { code: "RESOURCE_LIMIT" });
    await assert.rejects(book.toBlob(), { code: "RESOURCE_LIMIT" });
  }
  const footer = createWorkbook({ limits: { maxCells: 2 } });
  await assert.rejects(footer.addSheet("Footer", { columns: [{ header: "A" }], autoSize: {}, footer: { values: ["End"] } }).addRows([[1]]), { code: "RESOURCE_LIMIT" });
  let produced = 0, returned = false;
  function* rows() { try { for (let i = 0; i < 1000; i++) { produced++; yield ["x".repeat(1000)]; } } finally { returned = true; } }
  const bounded = createWorkbook({ compression: "store", limits: { maxOutputBytes: 10000 } });
  await assert.rejects(bounded.addSheet("Bounded", { columns: [{ header: "A" }] }).addRows(rows()), { code: "RESOURCE_LIMIT" });
  assert.ok(produced < 1000); assert.equal(returned, true);
  await assert.rejects(writeCsv([[null]], { columns: [{ header: "A" }], nullValue: "Unavailable", limits: { maxTextCharacters: 3 } }), { code: "RESOURCE_LIMIT" });
  const preservedFooter = createWorkbook({ oversizedText: "preserve", limits: { maxTextCharacters: 140000 } });
  await preservedFooter.addSheet("Footer", { columns: [{ header: "A" }], footer: { values: ["x".repeat(100000)] } }).addRows([[1]]);
  const footerZip = await readZip(await preservedFooter.toBlob());
  assert.match(footerZip.get("xl/worksheets/sheet1.xml").content, /<hyperlink ref="A3" location=/);
  assert.match(footerZip.get("xl/worksheets/sheet2.xml").content, /<c r="E5"[^>]*><v>4<\/v>/);
});

test("footer caches aggregate emitted date serials and numeric counts use General format", async () => {
  const operations = ["sum", "count", "average", "min", "max"];
  const book = createWorkbook({ dateMode: "utc" });
  const sheet = book.addSheet("Dates", { columns: operations.map(header => ({ header, type: "date", format: "yyyy-mm-dd" })).concat({ header: "Inferred" }),
    autoSize: {}, footer: { totals: Object.fromEntries(operations.map(op => [op, op]).concat([["Inferred", "count"]])) } });
  await sheet.addRows([operations.map(() => new Date("2026-10-06T00:00:00Z")).concat(new ExportCell(new Date("2026-10-06T00:00:00Z"))),
    operations.map(() => new Date("2026-10-08T00:00:00Z")).concat(new Date("2026-10-08T00:00:00Z")),
    operations.map(() => new Date(NaN)).concat(null)]);
  const blob = await book.toBlob(), xml = (await readZip(blob)).get("xl/worksheets/sheet1.xml").content;
  const serial = Number(xml.match(/<c r="A2"[^>]*><v>([^<]+)<\/v>/)[1]);
  const expected = [serial * 2 + 2, 2, serial + 1, serial, serial + 2, 2];
  for (let i = 0; i < expected.length; i++) assert.equal(Number(xml.match(new RegExp('<c r="' + String.fromCharCode(65 + i) + '5"[^>]*><f>.*?<\\/f><v>([^<]+)<\\/v>'))[1]), expected[i]);
  assert.match(xml, /<c r="B5" s="0">/); assert.match(xml, /<c r="F5" s="0">/);
  if (process.env.OFFICEIMO_REPORT_FIXTURES) await writeFile(process.env.OFFICEIMO_REPORT_FIXTURES + "/dates.xlsx", new Uint8Array(await blob.arrayBuffer()));
});

test("count, min, max and average do not require a finite sum", async () => {
  const operations = ["count", "min", "max", "average"];
  const book = createWorkbook(), sheet = book.addSheet("Large", { columns: operations.map(header => ({ header })), footer: { totals: Object.fromEntries(operations.map(op => [op, op])) } });
  await sheet.addRows([operations.map(() => 1e308), operations.map(() => 1e308)]);
  const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
  for (let i = 0; i < operations.length; i++) assert.match(xml, new RegExp('<c r="' + String.fromCharCode(65 + i) + '4"[^>]*><f>.*?<\\/f><v>' + (i ? '1e\\+308' : '2') + '<\\/v>'));
  const mixed = createWorkbook(); await mixed.addSheet("Mixed", { columns: [{ header: "A" }], footer: { totals: { A: "average" } } }).addRows([[1e308], [-1e308]]);
  assert.match((await readZip(await mixed.toBlob())).get("xl/worksheets/sheet1.xml").content, /<v>0<\/v>/);
  const sum = createWorkbook(); await assert.rejects(sum.addSheet("Sum", { columns: [{ header: "A" }], footer: { totals: { A: "sum" } } }).addRows([[1e308], [1e308]]), RangeError);
});

test("hyperlinks may address explicit footer values but not rows after the footer", async () => {
  const book = createWorkbook(), sheet = book.addSheet("Footer", { columns: [{ header: "A" }], footer: { values: ["Details"] } });
  await sheet.addRows([[1]]); sheet.addHyperlink({ cell: "A3", target: "https://evotec.xyz" });
  const blob = await book.toBlob(); assert.match((await readZip(blob)).get("xl/worksheets/sheet1.xml").content, /<hyperlink ref="A3" r:id="link1"/);
  if (process.env.OFFICEIMO_REPORT_FIXTURES) await writeFile(process.env.OFFICEIMO_REPORT_FIXTURES + "/footer-link.xlsx", new Uint8Array(await blob.arrayBuffer()));
  const invalid = createWorkbook(), other = invalid.addSheet("Footer", { columns: [{ header: "A" }], footer: { values: ["End"] } });
  await other.addRows([[1]]); other.addHyperlink({ cell: "A4", target: "https://evotec.xyz" }); await assert.rejects(invalid.toBlob(), RangeError);
});

test("cancellation between appends and finalization preserves the original reason", async () => {
  for (const closed of [false, true]) for (const operation of ["finish", "toBlob"]) {
    const controller = new AbortController(), reason = new Error("between operations");
    const book = createWorkbook({ signal: controller.signal }), sheet = book.addSheet("Cancel", { columns: [{ header: "A" }] });
    await sheet.addRows([["accepted"]]); if (closed) await sheet.close(); controller.abort(reason);
    await assert.rejects(book[operation](), error => error === reason);
    await assert.rejects(sheet.addRows([["later"]]), error => error === reason);
  }
});
