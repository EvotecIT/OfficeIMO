import test from "node:test";
import assert from "node:assert/strict";
import { createGroupedReport } from "../dist/core/grouped-report.js";
import { writeGroupedXlsx, writeGroupedXlsxTo } from "../dist/xlsx/grouped.js";
import { Workbook } from "../dist/xlsx/workbook.js";
import { FormulaCell } from "../dist/xlsx/formula.js";
import { ExportCell } from "../dist/core/presentation.js";
import { readZip } from "./zip-reader.mjs";

const columns = [
  { key: "value", header: "Wartość", type: "number", format: "0.00" },
  { key: "site", header: "場所", type: "string" },
  { key: "mean", header: "Średnia", type: "number" },
  { key: "minimum", header: "أدنى", type: "number" },
  { key: "maximum", header: "Höchste", type: "number" },
  { key: "count", header: "数", type: "number", format: "0.00" }
];
const totals = { value: "sum", mean: "average", minimum: "min", maximum: "max", count: "count" };
const rows = [
  { value: new ExportCell(10, { text: "dziesięć" }), site: "東京", mean: 4, minimum: 6, maximum: 6, count: 10 },
  { value: 20, site: "東京", mean: 8, minimum: 2, maximum: 8, count: null },
  { value: 5, site: "Łódź", mean: null, minimum: null, maximum: null, count: null }
];
const labels = { labelColumn: { header: "集計", width: 24 }, subtotalLabel: group => "Σ " + group.label, grandTotalLabel: () => "ΣΣ" };
function source(input = rows) { return createGroupedReport(input, { columns, groupBy: "site", totals }); }
async function xml(blob) { return (await readZip(blob)).get("xl/worksheets/sheet1.xml").content; }

test("grouped Excel keeps every typed source column, finite subtotal references and native outline rows", async () => {
  const output = await writeGroupedXlsx(source(), { ...labels, collapseGroups: true, sheet: { name: "分析", title: { text: "可編集" }, freezeHeader: true } });
  const parts = await readZip(output), sheet = parts.get("xl/worksheets/sheet1.xml").content;
  assert.match(sheet, /<outlinePr summaryBelow="1" showOutlineSymbols="1"\/>/);
  assert.match(sheet, /<sheetFormatPr defaultRowHeight="15" outlineLevelRow="1"\/>/);
  assert.match(sheet, /<row r="4"[^>]*outlineLevel="1"[^>]*hidden="1"/);
  assert.match(sheet, /<row r="5"[^>]*outlineLevel="1"[^>]*hidden="1"/);
  assert.match(sheet, /<row r="6"[^>]*collapsed="1"/);
  assert.match(sheet, /<row r="8"[^>]*outlineLevel="1"[^>]*hidden="1"/);
  assert.match(sheet, /<c r="B4"[^>]*><v>10<\/v>/);
  assert.match(sheet, /<c r="B6"[^>]*><f>SUBTOTAL\(9,B4:B5\)<\/f><v>30<\/v>/);
  assert.match(sheet, /<c r="D6"[^>]*><f>IF\(SUBTOTAL\(2,D4:D5\)=0,&quot;&quot;,SUBTOTAL\(1,D4:D5\)\)<\/f><v>6<\/v>/);
  assert.match(sheet, /<c r="E6"[^>]*><f>IF\(SUBTOTAL\(2,E4:E5\)=0,&quot;&quot;,SUBTOTAL\(5,E4:E5\)\)<\/f><v>2<\/v>/);
  assert.match(sheet, /<c r="F6"[^>]*><f>IF\(SUBTOTAL\(2,F4:F5\)=0,&quot;&quot;,SUBTOTAL\(4,F4:F5\)\)<\/f><v>8<\/v>/);
  assert.match(sheet, /<c r="G6" s="0"><f>SUBTOTAL\(2,G4:G5\)<\/f><v>1<\/v>/);
  assert.match(sheet, /<c r="D9"[^>]*t="str"><f>IF\(SUBTOTAL\(2,D8:D8\)=0,&quot;&quot;,SUBTOTAL\(1,D8:D8\)\)<\/f><v><\/v>/);
  assert.match(sheet, /<c r="B10"[^>]*><f>SUBTOTAL\(9,B4:B8\)<\/f><v>35<\/v>/);
  assert.match(sheet, /<c r="D10"[^>]*><f>IF\(SUBTOTAL\(2,D4:D8\)=0,&quot;&quot;,SUBTOTAL\(1,D4:D8\)\)<\/f><v>6<\/v>/);
  assert.doesNotMatch(sheet, /<autoFilter|<tableParts|dziesięć|SUBTOTAL\(10[129]/);
  assert.match(parts.get("xl/workbook.xml").content, /<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/);
});

test("grouped sink results and progress count accepted detail rows while budgets include summaries", async () => {
  const chunks = [], notifications = [];
  let pulls = 0, returns = 0, closeCalls = 0, abortCalls = 0;
  async function* input() { try { for (const row of rows) { pulls++; yield row; } } finally { returns++; } }
  const destination = new WritableStream({ write(bytes) { chunks.push(bytes.slice()); }, close() { closeCalls++; }, abort() { abortCalls++; } });
  const result = await writeGroupedXlsxTo(source(input()), destination, { ...labels, onProgress: event => notifications.push(event), compression: "store" });
  assert.deepEqual([result.rows, result.columns, pulls, returns, closeCalls, abortCalls, destination.locked], [3, 7, 3, 1, 0, 0, false]);
  assert.equal(result.bytes, chunks.reduce((sum, bytes) => sum + bytes.length, 0));
  assert.equal(notifications.at(-1).phase, "complete"); assert.equal(notifications.at(-1).rows, 3);
  for (const event of notifications) assert.ok(event.rows <= 3 && (event.sheetRows === undefined || event.sheetRows <= 3));
  await readZip(new Blob(chunks));
  await assert.rejects(writeGroupedXlsx(source(), { ...labels, limits: { maxRows: 7 } }), { code: "RESOURCE_LIMIT" });
});

test("empty and no-grand-total grouped workbooks preserve headers and meaningful empty aggregates", async () => {
  const empty = await xml(await writeGroupedXlsx(source([]), labels));
  assert.match(empty, /<c r="A2"[\s\S]*?ΣΣ/);
  assert.match(empty, /<c r="B2"[^>]*><v>0<\/v>/);
  assert.match(empty, /<c r="D2"[^>]*\/>/);
  assert.doesNotMatch(empty, /<f>|outlineLevel="1"/);
  const noGrand = await xml(await writeGroupedXlsx(source(), { ...labels, includeGrandTotal: false, sheet: { includeHeader: false } }));
  assert.doesNotMatch(noGrand, /ΣΣ|<row r="8"/);
  assert.match(noGrand, /<c r="B4"[^>]*><f>SUBTOTAL\(9,B2:B3\)<\/f><v>30<\/v>/);
});

test("grouped XLSX rejects incompatible options before source activity and preserves caller sink failures", async () => {
  let pulls = 0, returns = 0;
  async function* input() { try { for (const row of rows) { pulls++; yield row; } } finally { returns++; } }
  for (const sheet of [{ table: {} }, { autoFilter: true }, { footer: {} }, { dataValidation: [] }, { rowOutline() {} }])
    await assert.rejects(writeGroupedXlsx(source(input()), { ...labels, sheet }), /owns|autofilter/);
  assert.equal(pulls, 0);
  await assert.rejects(writeGroupedXlsx(source(input()), { ...labels, subtotalLabel: null }), /labels must be functions/);
  assert.throws(() => createGroupedReport([], { columns: [{ key: "site", header: "Site" }, { key: "date", header: "Date", type: "date" }], groupBy: "site", totals: { date: "count" } }), /numeric|number/);
  const sentinel = new Error("destination rejected"); let writes = 0;
  await assert.rejects(writeGroupedXlsxTo(source(input()), { write() { writes++; throw sentinel; } }, { ...labels, sheet: { autoSize: { sampleRows: 1 } } }), error => error === sentinel);
  assert.ok(writes > 0); assert.equal(returns, 1);
});

test("native outline metadata captures callbacks and getters once and fails invalid levels", async () => {
  const book = new Workbook(), calls = [], reads = [];
  const sheet = book.addWorksheet("Outline", { columns: [{ header: "Value" }], autoSize: { sampleRows: 2 }, maxOutlineLevel: 2,
    rowOutline(context) { calls.push([context.rowIndex, context.worksheetRow, context.values[0]]); return {
      get level() { reads.push("level"); return context.rowIndex + 1; }, get hidden() { reads.push("hidden"); return false; }, get collapsed() { reads.push("collapsed"); return false; }
    }; } });
  await sheet.addRows([[1], [2]]);
  const output = await xml(await book.toBlob());
  assert.deepEqual(calls, [[0, 2, 1], [1, 3, 2]]); assert.equal(reads.length, 6);
  assert.match(output, /r="A2"/); assert.match(output, /<row r="3" outlineLevel="2"/);
  for (const options of [{ rowOutline() {} }, { maxOutlineLevel: 1 }, { rowOutline() {}, maxOutlineLevel: 8 }])
    assert.throws(() => new Workbook().addWorksheet("Invalid", { columns: [{ header: "Value" }], ...options }));
  for (const metadata of [{ level: -1 }, { level: 1.5 }, { level: 2 }, { hidden: 1 }, null]) {
    const invalid = new Workbook();
    await assert.rejects(invalid.addWorksheet("Invalid", { columns: [{ header: "Value" }], maxOutlineLevel: 1, rowOutline: () => metadata }).addRows([[1]]));
  }
});

test("blank formula caches survive authored cells without becoming a numeric zero", async () => {
  const book = new Workbook(), observed = [];
  await book.addWorksheet("Null cache", { columns: [{ header: "Numeric", type: "number" }],
    rowStyle(context) { observed.push(context.values[0]); } }).addRows([[new FormulaCell('IF(1=1,"",0)', null)]]);
  assert.deepEqual(observed, [null]);
  assert.match(await xml(await book.toBlob()), /<c r="A2"[^>]*t="str"><f>IF\(1=1,&quot;&quot;,0\)<\/f><v><\/v>/);
});

test("configured source cancellation interrupts a stalled grouped output write and returns its iterator", async () => {
  const sourceController = new AbortController(), writerController = new AbortController(), reason = new Error("source cancelled while output pending");
  let released = 0, writes = 0;
  async function* input() { try { yield rows[0]; yield rows[1]; } finally { released++; } }
  const report = createGroupedReport(input(), { columns, groupBy: "site", totals, signal: sourceController.signal });
  const destination = new WritableStream({ write() { writes++; queueMicrotask(() => sourceController.abort(reason)); return new Promise(() => {}); } });
  await assert.rejects(writeGroupedXlsxTo(report, destination, { ...labels, signal: writerController.signal, compression: "store", sheet: { autoSize: { sampleRows: 1 } } }), error => error === reason);
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(released, 1); assert.equal(writes, 1); assert.equal(destination.locked, false); assert.equal(writerController.signal.aborted, false);
});
