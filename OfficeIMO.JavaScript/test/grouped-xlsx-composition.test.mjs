import test from "node:test";
import assert from "node:assert/strict";
import { createGroupedReport } from "../dist/core/grouped-report.js";
import { addGroupedXlsx, writeGroupedXlsx } from "../dist/xlsx/grouped.js";
import { Workbook } from "../dist/xlsx/workbook.js";
import { readZip } from "./zip-reader.mjs";

const columns = [
  { key: "amount", header: "Kwota", type: "number" },
  { key: "name", header: "選択", type: "string" },
  { key: "mean", header: "Średnia", type: "number" },
  { key: "minimum", header: "Minimum", type: "number" },
  { key: "maximum", header: "Maximum", type: "number" },
  { key: "count", header: "Count", type: "number" },
  { key: "active", header: "Active", type: "boolean" },
  { key: "when", header: "Date", type: "date" }
];
const rows = [
  { region: "North", department: "Ops", team: "East", lane: "A", amount: 10, name: "Alpha", mean: 2, minimum: 2, maximum: 2, count: 1, active: true, when: new Date("2026-01-02T00:00:00Z") },
  { region: "North", department: "Ops", team: "East", lane: "A", amount: 20, name: "Beta", mean: 6, minimum: 6, maximum: 6, count: null, active: false, when: null },
  { region: "North", department: "Ops", team: "East", lane: "B", amount: 30, name: "Alpha", mean: 10, minimum: 10, maximum: 10, count: 1, active: true, when: null },
  { region: "North", department: "Ops", team: "West", lane: "A", amount: 40, name: "Beta", mean: null, minimum: null, maximum: null, count: null, active: false, when: null }
];
const totals = { amount: "sum", mean: "average", minimum: "min", maximum: "max", count: "count" };
const labels = { labelColumn: { header: "分析" }, subtotalLabel: group => "Σ " + group.label, grandTotalLabel: () => "Grand" };
function source(input = rows, options = {}) {
  return createGroupedReport(input, { columns, totals,
    groupBy: ["region", "department", "team", "lane"].map(key => ({ key, value: row => row[key] })), ...options });
}
function cell(xml, column, row) { return xml.match(new RegExp('<c r="' + column + row + '"[^>]*(?:\\/>|>[\\s\\S]*?<\\/c>)'))?.[0]; }

test("composed four-level analysis shares raw aggregates, finite coordinates and detail-only editable native ranges", async () => {
  let reads = 0, returned = 0;
  async function* input() { try { for (const row of rows) { reads++; yield row; } } finally { returned++; } }
  const chunks = [], progress = [], book = new Workbook({ compression: "store", dateMode: "utc", sink: { write: bytes => chunks.push(bytes.slice()) }, onProgress: event => progress.push(event) });
  await book.addWorksheet("Introduction", { columns: [{ header: "Report" }] }).addRows([["Before analysis"], ["Context"], ["Scope"]]);
  const result = await addGroupedXlsx(book, source(input()), { ...labels, collapseGroups: true,
    sheet: { name: "Analysis", title: { text: "Editable report" }, freezeHeader: true },
    conditionalFormats: [{ type: "cellIs", range: { column: "mean" }, operator: "greaterThan", value: 5, style: { font: { bold: true } } }],
    dataValidation: [{ type: "list", column: 2, items: ["Alpha", "Beta"] }] });
  assert.deepEqual([result.rows, result.columns, reads, returned, result.groups.length], [4, 9, 4, 1, 7]);
  assert.deepEqual(result.grandSummary, { rows: 4, groups: 7, totals: [100, null, 6, 2, 10, 2, null, null] });
  const east = result.groups.find(group => group.label === "East");
  assert.equal(east.rows, 3); assert.equal(east.totals[2], 6, "Three raw values determine the parent mean, rather than its two unequal leaf means");
  assert.deepEqual(east.path.map(group => group.label), ["North", "Ops", "East"]);
  assert.equal(result.groups.find(group => group.label === "West").totals[2], null);
  assert.ok(Object.isFrozen(result) && Object.isFrozen(result.groups) && Object.isFrozen(east) && Object.isFrozen(east.totals));
  assert.ok(Object.isFrozen(result.detailRanges) && result.detailRanges.every(Object.isFrozen));
  await assert.rejects(result.worksheet.addRows([["Accidental append"]]), /closed/);
  const overview = book.addWorksheet("Overview", { columns: [{ header: "Group" }, { header: "Mean", type: "number" }] });
  await overview.addRows(result.groups.filter(group => group.level === 2).map(group => [group.label, group.totals[2]]));
  assert.equal(book.rowCount, rows.length + 5);
  const completion = await book.finish();
  assert.equal(completion.sheets, 3); assert.equal(completion.rows, rows.length + 5);
  assert.equal(progress.at(-1).rows, rows.length + 5);
  assert.ok(progress.filter(event => event.sheetName === "Analysis").every(event => event.sheetRows <= rows.length));
  const parts = await readZip(new Blob(chunks)), xml = parts.get("xl/worksheets/sheet2.xml").content;
  assert.match(parts.get("xl/worksheets/sheet1.xml").content, /Before analysis/);
  assert.match(parts.get("xl/worksheets/sheet3.xml").content, /East/);
  assert.match(xml, /outlineLevelRow="4"/);
  const first = result.detailRanges[0].firstRow;
  assert.match(xml, new RegExp('<row r="' + first + '" outlineLevel="4" hidden="1"'));
  assert.match(cell(xml, "B", first), /<v>10<\/v>/);
  assert.match(cell(xml, "H", first), /t="b"><v>1<\/v>/);
  assert.match(cell(xml, "I", first), /<v>46024<\/v>/);
  assert.match(cell(xml, "D", east.summaryRow), new RegExp('IF\\(SUBTOTAL\\(2,D' + east.firstDetailRow + ':D' + east.lastDetailRow + '\\)=0,&quot;&quot;,SUBTOTAL\\(1,D' + east.firstDetailRow + ':D' + east.lastDetailRow + '\\)\\)<\\/f><v>6<\\/v>'));
  assert.match(cell(xml, "D", result.grandTotalRow), /<v>6<\/v>/);
  const ranges = column => result.detailRanges.map(range => column + range.firstRow + ":" + column + range.lastRow).join(" ");
  assert.ok(xml.includes('<conditionalFormatting sqref="' + ranges("D") + '">'));
  assert.ok(xml.includes(' sqref="' + ranges("C") + '"><formula1>'));
  assert.doesNotMatch(xml, /<autoFilter|<tableParts|>region<|>department<|>team<|>lane</);
  assert.equal(result.detailRanges.reduce((count, range) => count + range.lastRow - range.firstRow + 1, 0), 4);
  for (const group of result.groups) {
    assert.ok(group.startRow < group.firstDetailRow && group.firstDetailRow <= group.lastDetailRow && group.lastDetailRow < group.summaryRow);
    assert.ok(!result.detailRanges.some(range => group.summaryRow >= range.firstRow && group.summaryRow <= range.lastRow));
  }
});

test("completed grouped data ranges exclude generated rows and reject discontiguous details", async () => {
  const book = new Workbook();
  const single = await addGroupedXlsx(book, source(rows.slice(0, 2)), { ...labels, sheet: { name: "Team's analysis", title: { text: "Context" } } });
  const { firstRow, lastRow } = single.detailRanges[0];
  const expected = "'Team''s analysis'!$B$" + firstRow + ":$B$" + lastRow;
  assert.equal(single.worksheet.dataRange("amount"), expected);
  assert.equal(single.worksheet.dataRange(2), expected);
  const ordinary = book.addWorksheet("Plain", { columns: [{ key: "amount", header: "Amount", type: "number" }], title: { text: "Context" } });
  await ordinary.addRows([[2], [4]]); await ordinary.close();
  assert.equal(ordinary.dataRange("amount"), "'Plain'!$A$3:$A$4");
  const multiple = await addGroupedXlsx(book, source(), { ...labels, sheet: { name: "Multiple" } });
  assert.throws(() => multiple.worksheet.dataRange("amount"), error => error.code === "INVALID_STATE" && /discontiguous/.test(error.message));
  for (const includeGrandTotal of [true, false]) {
    const empty = await addGroupedXlsx(book, source([]), { ...labels, includeGrandTotal, sheet: { name: "Empty " + includeGrandTotal } });
    assert.equal(empty.worksheet.dataRange("amount"), undefined);
  }
  const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
  const values = Array.from({ length: lastRow - firstRow + 1 }, (_, index) => Number(cell(xml, "B", firstRow + index).match(/<v>([^<]+)<\/v>/)[1]));
  assert.deepEqual(values, [10, 20]);
  assert.equal(values.reduce((sum, value) => sum + value, 0), 30, "The returned reference covers raw detail amounts once");
});

test("grouped detail selectors capture inherited metadata once and cannot address generated or hidden columns", async () => {
  let columnReads = 0, throughReads = 0, rangeReads = 0;
  const range = Object.create({ get column() { columnReads++; return "amount"; }, get through() { throughReads++; return "mean"; } });
  const rule = Object.assign(Object.create({ get range() { rangeReads++; return range; } }), { type: "dataBar", color: "#88aa44" });
  const book = new Workbook();
  const result = await addGroupedXlsx(book, source(), { ...labels, conditionalFormats: [rule], includeGrandTotal: false });
  assert.deepEqual([columnReads, throughReads, rangeReads], [1, 1, 1]);
  assert.equal(result.grandTotalRow, undefined); assert.equal(result.grandSummary.totals[0], 100);
  const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
  assert.ok(xml.includes('sqref="' + result.detailRanges.map(range => "B" + range.firstRow + ":D" + range.lastRow).join(" ") + '"'));
  for (const invalid of ["B2:B3", { column: 0 }, { column: "region" }, { column: 9 }]) {
    let pulls = 0;
    function* input() { pulls++; yield rows[0]; }
    const invalidBook = new Workbook();
    await assert.rejects(addGroupedXlsx(invalidBook, source(input()), { ...labels, conditionalFormats: [{ type: "dataBar", range: invalid, color: "88aa44" }] }), /detail|column|one-based/);
    assert.equal(pulls, 0); assert.equal(invalidBook.worksheets.length, 0);
  }
});

test("composed empty analysis omits detail editing rules and still exposes shared grand aggregates", async () => {
  const book = new Workbook(), result = await addGroupedXlsx(book, source([]), { ...labels, includeGrandTotal: false,
    conditionalFormats: [{ type: "dataBar", range: { column: 1 }, color: "88aa44" }],
    dataValidation: [{ type: "list", column: "name", items: ["Alpha"] }] });
  assert.deepEqual(result.detailRanges, []); assert.deepEqual(result.groups, []);
  assert.deepEqual(result.grandSummary, { rows: 0, groups: 0, totals: [0, null, null, null, null, 0, null, null] });
  const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
  assert.doesNotMatch(xml, /<conditionalFormatting|<dataValidations|<f>/);
});

test("grouped editing rule limits reject before observing oversized rule getters", async () => {
  let reads = 0;
  const oversized = Array.from({ length: 2 }, () => ({ get type() { reads++; return "dataBar"; } }));
  for (const property of ["conditionalFormats", "dataValidation"]) {
    const book = new Workbook({ limits: { maxConditionalFormats: 1, maxDataValidations: 1 } });
    await assert.rejects(addGroupedXlsx(book, source(), { ...labels, [property]: oversized }), { code: "RESOURCE_LIMIT" });
    assert.equal(reads, 0); assert.equal(book.worksheets.length, 0);
  }
});

test("source and workbook cancellation both return a held producer and poison only the failed workbook", { timeout: 5000 }, async () => {
  for (const owner of ["source", "book"]) {
    const controller = new AbortController(), reason = new Error(owner + " cancelled"), chunks = [];
    let returns = 0, pending;
    const entered = new Promise(resolve => { pending = resolve; });
    const input = { [Symbol.asyncIterator]() { return {
      next() { pending(); return new Promise(() => {}); },
      return() { returns++; return { done: true }; }
    }; } };
    const book = new Workbook({ compression: "store", ...(owner === "book" ? { signal: controller.signal } : {}), sink: { write: bytes => chunks.push(bytes.slice()) } });
    const writing = addGroupedXlsx(book, source(input, owner === "source" ? { signal: controller.signal } : {}), labels);
    await entered; controller.abort(reason);
    await assert.rejects(writing, error => error === reason);
    await new Promise(resolve => setImmediate(resolve)); assert.equal(returns, 1);
    await assert.rejects(book.finish(), error => error === reason);
    assert.throws(() => book.addWorksheet("After failure"), error => error === reason);
  }
});

test("an already cancelled grouped source registers and writes no worksheet", async () => {
  const controller = new AbortController(), reason = new Error("already cancelled"); controller.abort(reason);
  let writes = 0, reads = 0;
  function* input() { reads++; yield rows[0]; }
  const book = new Workbook({ sink: { write() { writes++; } } });
  await assert.rejects(addGroupedXlsx(book, source(input(), { signal: controller.signal }), { ...labels, sheet: { autoSize: { sampleRows: 0 } } }), error => error === reason);
  assert.deepEqual([reads, writes, book.worksheets.length], [0, 0, 0]);
});

test("source cancellation interrupts a held composed destination write and returns the active body iterator", { timeout: 5000 }, async () => {
  for (const owner of ["source", "book"]) {
    const controller = new AbortController(), reason = new Error(owner + " output cancelled");
    let returned = 0, writes = 0;
    async function* input() { try { yield* rows; } finally { returned++; } }
    const book = new Workbook({ compression: "store", ...(owner === "book" ? { signal: controller.signal } : {}), sink: {
      write() { writes++; queueMicrotask(() => controller.abort(reason)); return new Promise(() => {}); }
    } });
    await assert.rejects(addGroupedXlsx(book, source(input(), owner === "source" ? { signal: controller.signal } : {}), { ...labels, sheet: { autoSize: { sampleRows: 1 } } }), error => error === reason);
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(writes, 1); assert.equal(returned, 1);
    await assert.rejects(book.finish(), error => error === reason);
  }
});

test("discontiguous detail references enforce their native text bound and standalone shares the same composition", async () => {
  const { dataRowReference } = await import("../dist/xlsx/data-row-ranges.js");
  const ranges = Array.from({ length: 4000 }, (_, index) => ({ firstRow: index * 2 + 2, lastRow: index * 2 + 2 }));
  assert.throws(() => dataRowReference(1, 1, 1, 8000, ranges), { code: "RESOURCE_LIMIT" });
  const output = await writeGroupedXlsx(source(), { ...labels, conditionalFormats: [{ type: "dataBar", range: { column: "amount" }, color: "88aa44" }] });
  const xml = (await readZip(output)).get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /<conditionalFormatting sqref="B\d+:B\d+ B\d+:B\d+ B\d+:B\d+">/);
});
