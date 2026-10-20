import test from "node:test";
import assert from "node:assert/strict";
import { Workbook, ExportCell } from "../dist/xlsx/index.js";
import { writeCsv, writeCsvTo } from "../dist/csv/index.js";
import { readZip } from "./zip-reader.mjs";

const vertical = () => [
  [{ value: new ExportCell("Name", { presentation: { background: "14532D", color: "FFFFFF" } }), rowSpan: 2 }, { value: "Metrics", columnSpan: 2 }, null],
  [null, { value: "Amount" }, { value: "Count" }]
];
const columns = () => [{ header: "Name" }, { header: "Amount", type: "number", format: "0.00" }, { header: "Count", type: "number" }];

test("rectangular XLSX headings and appended footers preserve spans, raw cells, styles and data coordinates", async () => {
  for (const streamed of [false, true]) {
    const chunks = [], book = new Workbook({ compression: "store", ...(streamed ? { sink: { write(bytes) { chunks.push(bytes.slice()); } } } : {}) });
    const headerRows = vertical(), footerRows = [
      [{ value: "Totals", columnSpan: 2 }, null, { value: new ExportCell(3, { text: "three", presentation: { bold: true } }), rowSpan: 2 }],
      [{ value: "Checked" }, { value: 25 }, null]
    ], calls = [];
    const sheet = book.addWorksheet("Report", { columns: columns(), title: { text: "Report" }, headerRows, footer: { rows: footerRows }, freezeHeader: true,
      autoSize: { sampleRows: 1 }, print: { repeatHeaders: true }, cellStyle(context) { calls.push([context.rowIndex, context.columnIndex, context.worksheetRow]); } });
    headerRows[0][1].value = "mutated"; footerRows[0][0].value = "mutated";
    await sheet.addRows([["one", new ExportCell(12.5, { text: "twelve", link: { target: "https://example.com/one" } }), 1], ["two", 12.5, 2]]);
    const result = await book.finish(), blob = streamed ? new Blob(chunks) : await book.toBlob(), zip = await readZip(blob), xml = zip.get("xl/worksheets/sheet1.xml").content;
    assert.equal(result.rows, 2); assert.equal(sheet.rowCount, 2);
    assert.deepEqual(calls, [[0, 0, 4], [0, 1, 4], [0, 2, 4], [1, 0, 5], [1, 1, 5], [1, 2, 5]]);
    assert.match(xml, /ySplit="3" topLeftCell="A4"/);
    for (const ref of ["A1:C1", "A2:A3", "B2:C2", "A6:B6", "C6:C7"]) assert.ok(xml.includes('mergeCell ref="' + ref + '"'), ref);
    assert.match(xml, /r="B4"[^>]*><v>12.5<\/v>/); assert.match(xml, /hyperlink ref="B4"/);
    assert.match(xml, /r="C6"[^>]*><v>3<\/v>/); assert.match(xml, /r="B7"[^>]*><v>25<\/v>/);
    assert.doesNotMatch(xml, /mutated|twelve|three/);
    assert.match(zip.get("xl/styles.xml").content, /fgColor rgb="FF14532D"/);
    assert.match(zip.get("xl/workbook.xml").content, /\$2:\$3/); assert.match(zip.get("xl/workbook.xml").content, /\$A\$1:\$C\$7/);
  }
});

test("separate equal heading groups remain separate and structured footers stay outside a native table", async () => {
  const book = new Workbook(), names = ["A", "B", "C", "D"], sheet = book.addWorksheet("Native", {
    columns: names.map(header => ({ header })), title: { text: "Title" }, headerRows: [
      [{ value: "Same", columnSpan: 2 }, null, { value: "Same", columnSpan: 2 }, null], names.map(value => ({ value }))
    ], table: {}, autoFilter: true, footer: { rows: [[{ value: "First", columnSpan: 4 }, null, null, null], names.map(value => ({ value }))] }
  });
  await sheet.addRows([[1, 2, 3, 4], [5, 6, 7, 8]]);
  const zip = await readZip(await book.toBlob()), xml = zip.get("xl/worksheets/sheet1.xml").content, table = zip.get("xl/tables/table1.xml").content;
  for (const ref of ["A2:B2", "C2:D2", "A6:D6"]) assert.ok(xml.includes('mergeCell ref="' + ref + '"'));
  assert.match(table, /ref="A3:D5" totalsRowShown="0"/); assert.match(table, /autoFilter ref="A3:D5"/);
  assert.equal((table.match(/tableColumn id=/g) ?? []).length, 4);
  assert.doesNotMatch(xml, /<autoFilter/);
});

test("structured XLSX validation prevents hidden metadata, invalid native headings and budget bypasses", async () => {
  for (const option of [{ table: {} }, { autoFilter: true }]) assert.throws(() => new Workbook().addWorksheet("Bad", { columns: columns(), headerRows: vertical(), ...option }), /unmerged final heading/);
  assert.throws(() => new Workbook().addWorksheet("Bad", { columns: columns(), footer: { rows: vertical(), values: [] } }), /cannot accompany/);
  assert.throws(() => new Workbook().addWorksheet("Bad", { columns: columns(), headerRows: [[{ value: "hidden", columnSpan: 3 }, { value: "lost" }, null]] }), /Covered/);
  const book = new Workbook({ limits: { maxMergedRanges: 1 } });
  book.addWorksheet("One", { columns: [{ header: "A" }, { header: "B" }], footer: { rows: [[{ value: "one", columnSpan: 2 }, null]] } });
  assert.throws(() => book.addWorksheet("Two", { columns: [{ header: "A" }, { header: "B" }], footer: { rows: [[{ value: "two", columnSpan: 2 }, null]] } }), { code: "RESOURCE_LIMIT" });
  for (const limits of [{ maxCells: 14 }, { maxTextCharacters: 1 }]) {
    const bounded = new Workbook({ limits }), sheet = bounded.addWorksheet("Bounded", { columns: columns(), headerRows: vertical(), footer: { rows: [[{ value: "End" }, { value: 1 }, { value: 2 }]] } });
    await assert.rejects(sheet.addRows([["one", 1, 2], ["two", 3, 4]]), { code: "RESOURCE_LIMIT" });
  }
  const links = new Workbook(), sheet = links.addWorksheet("Links", { columns: [{ header: "A" }, { header: "B" }], footer: { rows: [[{ value: "end", columnSpan: 2 }, null]] } });
  await sheet.addRows([[1, 2]]); sheet.addHyperlink({ cell: "B3", target: "https://example.com" });
  await assert.rejects(links.toBlob(), /covered merged cell/);
  const overlap = new Workbook(), other = overlap.addWorksheet("Overlap", { columns: columns(), mergedCells: ["A3:B3"], footer: { rows: [[{ value: "End", columnSpan: 3 }, null, null]] } });
  await other.addRows([[1, 2, 3]]); await assert.rejects(overlap.toBlob(), /overlap/);
  const suppressed = new Workbook(), hidden = suppressed.addWorksheet("Suppressed", { columns: columns().map(column => ({ ...column, groups: ["unused"] })), headerRows: vertical(), includeHeader: false });
  await hidden.addRows([["row", 1, 2]]);
  const xml = (await readZip(await suppressed.toBlob())).get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /r="A1"[^>]*><is><t xml:space="preserve">row/); assert.doesNotMatch(xml, /Metrics|unused|mergeCells/);
});

test("CSV explicitly flattens every physical heading/footer row and retains body context and counts", async () => {
  for (const spanMode of ["repeat", "anchors"]) for (const valueMode of ["raw", "display"]) {
    const calls = [], footer = { rows: [[{ value: new ExportCell(10, { text: "ten" }), columnSpan: 2 }, null, { value: "end", rowSpan: 2 }], [{ value: "note" }, { value: "last" }, null]] };
    const options = { columns: columns().map(column => ({ ...column, valueFormatter(value, context) { calls.push([context.rowIndex, context.columnIndex]); return value; } })), headerRows: vertical(), footer, spanMode, valueMode };
    const chunks = [], result = await writeCsvTo([["row", 12.5, 1]], { write(chunk) { chunks.push(chunk.slice()); } }, options);
    const text = new TextDecoder().decode(Buffer.concat(chunks));
    const metric = spanMode === "repeat" ? "Name,Metrics,Metrics\r\nName,Amount,Count\r\n" : "Name,Metrics,\r\n,Amount,Count\r\n";
    const amount = valueMode === "raw" ? "10" : "ten";
    assert.equal(text, metric + "row,12.5,1\r\n" + amount + "," + (spanMode === "repeat" ? amount : "") + ",end\r\nnote,last," + (spanMode === "repeat" ? "end" : "") + "\r\n");
    assert.equal(result.rows, 1); assert.equal(result.bytes, Buffer.byteLength(text)); assert.deepEqual(calls, [[0, 0], [0, 1], [0, 2]]);
  }
});

test("CSV rejects implicit span loss before reading data and applies protection and limits to metadata", async () => {
  let produced = 0, writes = 0;
  function* rows() { produced++; yield [1, 2, 3]; }
  await assert.rejects(writeCsvTo(rows(), { write() { writes++; } }, { columns: columns(), headerRows: vertical() }), /explicit spanMode/);
  assert.equal(produced, 0); assert.equal(writes, 0);
  const blob = await writeCsv([[1]], { columns: [{ header: "unused" }], headerRows: [[{ value: "=Heading" }]], footer: { rows: [[{ value: "@Footer" }]] } });
  assert.equal(await blob.text(), "'=Heading\r\n1\r\n'@Footer\r\n");
  await assert.rejects(writeCsv([], { columns: columns(), headerRows: vertical(), spanMode: "repeat", limits: { maxCells: 5 } }), { code: "RESOURCE_LIMIT" });
  const suppressed = await writeCsv([[1, 2, 3]], { columns: columns(), headerRows: vertical(), includeHeader: false });
  assert.equal(await suppressed.text(), "1,2,3\r\n");
});
