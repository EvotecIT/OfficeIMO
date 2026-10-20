import test from "node:test";
import assert from "node:assert/strict";
import { FormulaCell, Workbook } from "../dist/xlsx/index.js";
import { readZip } from "./zip-reader.mjs";

const columns = [{ key: "status", header: "Status" }, { key: "count", header: "Services", type: "number" }];
const bound = (sheet, overrides = {}) => ({ type: "columnClustered", row: 11, column: 1, width: 600, height: 320,
  data: { kind: "range", sheet, rowCount: 3, categoryColumn: "status", series: [{ column: "count" }], ...overrides } });
const xml = (zip, id = 1) => zip.get(`xl/charts/chart${id}.xml`).content.replaceAll("&apos;", "'");

test("visible table columns supply native formulas and caches in one writer pass, excluding report decorations", async () => {
  const book = new Workbook({ compression: "store", limits: { maxSheets: 1, maxChartPoints: 3 } });
  const sheet = book.addWorksheet("O'Brien overview", { columns: columns.map(column => ({ ...column, groups: ["Service review"] })),
    title: { text: "Selected services" }, footer: { values: ["Footer", 12345] }, table: { name: "VisibleServices" } });
  sheet.addChart(bound(sheet, { firstRow: 1 }));
  let enumerations = 0, released = 0;
  async function* rows() { enumerations++; try { yield* [["Excluded", 999], ["Healthy", 84], ["Attention", 8], ["Offline", 3], ["Outside", 4321]]; } finally { released++; } }
  await sheet.addRows(rows());
  const zip = await readZip(await book.toBlob()), chart = xml(zip);
  assert.equal(enumerations, 1); assert.equal(released, 1);
  assert.equal(book.rowCount, 5); assert.equal((await book.finish()).sheets, 1);
  assert.match(chart, /'O''Brien overview'!\$A\$5:\$A\$7/);
  assert.match(chart, /'O''Brien overview'!\$B\$5:\$B\$7/);
  assert.match(chart, /'O''Brien overview'!\$B\$3/);
  assert.match(chart, /<c:plotVisOnly val="1"/);
  assert.match(chart, /<c:ptCount val="3"/);
  for (const value of [84, 8, 3]) assert.match(chart, new RegExp(`<c:v[^>]*>${value}</c:v>`));
  assert.doesNotMatch(chart, /Excluded|Outside|Footer|>999<|>12345<|>4321</);
  assert.doesNotMatch(zip.get("xl/workbook.xml").content, /state="hidden"|Chart data/);
  assert.equal([...zip.keys()].filter(path => /^xl\/worksheets\/sheet\d+\.xml$/.test(path)).length, 1);
  assert.match(zip.get("xl/tables/table1.xml").content, /name="VisibleServices"/);
});

test("all four chart families bind another visible sheet and preserve explicit series names and native header references", async () => {
  const book = new Workbook({ compression: "store" }), source = book.addWorksheet("Visible source", { columns });
  const report = book.addWorksheet("Report", { columns: [{ header: "Report" }] });
  for (const type of ["columnClustered", "barClustered", "line", "pie"]) {
    const input = bound(source); input.type = type;
    if (type === "pie") input.data.series = [{ column: 2, name: "Selected count", color: "#203864" }];
    report.addChart(input);
  }
  await source.addRows([["Healthy", 84], ["Attention", 8], ["Offline", 3]]);
  await report.addRows([["Selected service review"]]);
  const zip = await readZip(await book.toBlob());
  for (let id = 1; id <= 4; id++) {
    const chart = xml(zip, id);
    assert.match(chart, /'Visible source'!\$A\$2:\$A\$4/);
    assert.match(chart, /'Visible source'!\$B\$2:\$B\$4/);
    if (id === 4) {
      assert.match(chart, /<c:tx><c:v>Selected count<\/c:v><\/c:tx>/);
      assert.match(chart, /<a:srgbClr val="203864"/);
    } else assert.match(chart, /'Visible source'!\$B\$1/);
  }
  assert.equal((await book.finish()).sheets, 2);
  assert.doesNotMatch(zip.get("xl/workbook.xml").content, /state="hidden"/);
});

test("bound registration captures caller metadata once and leaves source cells editable", async () => {
  const book = new Workbook({ compression: "store" }), sheet = book.addWorksheet("Overview", { columns });
  const calls = new Map(), once = (key, value) => { calls.set(key, (calls.get(key) ?? 0) + 1); return value; };
  const series = { get column() { return once("column", "count"); }, get name() { return once("name", "Selected"); } };
  const data = { kind: "range", get sheet() { return once("sheet", sheet); }, get rowCount() { return once("rowCount", 3); },
    get categoryColumn() { return once("categoryColumn", "status"); }, get series() { return once("series", [series]); } };
  const input = { ...bound(sheet), get data() { return once("data", data); } }; sheet.addChart(input);
  await sheet.addRows([["Healthy", 84], ["Attention", 8], ["Offline", 3]]);
  const zip = await readZip(await book.toBlob());
  for (const key of ["data", "sheet", "rowCount", "categoryColumn", "series", "column", "name"]) assert.equal(calls.get(key), 1, key);
  assert.match(xml(zip), /<c:tx><c:v>Selected<\/c:v><\/c:tx>/);
  assert.match(zip.get("xl/worksheets/sheet1.xml").content, /r="B2"[^>]*><v>84<\/v>/);
});

test("numeric formula caches populate a bound chart without treating a title as a series header", async () => {
  const book = new Workbook({ compression: "store" }), facts = book.addWorksheet("Facts", { columns: [{ header: "Status" }] });
  const sheet = book.addWorksheet("Overview", { columns, includeHeader: false, title: { text: "Selected services" } });
  sheet.addChart(bound(sheet));
  await facts.addRows([["Healthy"], ["Healthy"], ["Offline"]]);
  await sheet.addRows([["Healthy", new FormulaCell('COUNTIF(\'Facts\'!$A$2:$A$4,"Healthy")', 2)],
    ["Attention", new FormulaCell('COUNTIF(\'Facts\'!$A$2:$A$4,"Attention")', 0)],
    ["Offline", new FormulaCell('COUNTIF(\'Facts\'!$A$2:$A$4,"Offline")', 1)]]);
  const zip = await readZip(await book.toBlob()), chart = xml(zip);
  assert.match(chart, /'Overview'!\$A\$2:\$A\$4/); assert.match(chart, /'Overview'!\$B\$2:\$B\$4/);
  assert.match(chart, /<c:tx><c:v>Services<\/c:v><\/c:tx>/);
  assert.doesNotMatch(chart, /'Overview'!\$B\$1/);
  for (const value of [2, 0, 1]) assert.match(chart, new RegExp(`<c:v[^>]*>${value}</c:v>`));
  const cells = zip.get("xl/worksheets/sheet2.xml").content;
  assert.match(cells, /r="B2"[^>]*><f>COUNTIF/); assert.match(cells, /<\/f><v>2<\/v>/);
});

test("foreign, late, invalid and over-budget bindings fail without allocating hidden chart sheets", async () => {
  const other = new Workbook(), foreign = other.addWorksheet("Foreign", { columns });
  const book = new Workbook({ compression: "store", limits: { maxCharts: 1, maxChartPoints: 3, maxSheets: 1 } });
  const sheet = book.addWorksheet("Overview", { columns });
  for (const data of [{ sheet: foreign }, { rowCount: 0 }, { firstRow: -1 }, { firstRow: null }, { rowCount: Infinity }, { categoryColumn: 0 },
    { categoryColumn: "missing" }, { series: [{ column: 3 }] }, { series: [{ column: "status" }] }, { series: [{ column: "count", color: "transparent" }] }])
    assert.throws(() => sheet.addChart(bound(sheet, data)));
  sheet.addChart(bound(sheet));
  assert.throws(() => sheet.addChart(bound(sheet)), error => error.code === "RESOURCE_LIMIT");
  await sheet.addRows([["Healthy", 84], ["Attention", 8], ["Offline", 3]]);
  const zip = await readZip(await book.toBlob()); assert.ok(zip.has("xl/charts/chart1.xml")); assert.ok(!zip.has("xl/charts/chart2.xml"));
  const lateBook = new Workbook(), late = lateBook.addWorksheet("Late", { columns });
  await late.addRows([["Already written", 1]]);
  assert.throws(() => late.addChart(bound(late)), /before|already|written|range/i);
  await lateBook.discard(); await other.discard();
});

test("truncated, blank, nonnumeric and invalid pie source values reject the operation", async () => {
  for (const rows of [[["Only one", 1]], [["A", null], ["B", 2], ["C", 3]], [[7, 1], ["B", 2], ["C", 3]],
    [["A", NaN], ["B", 2], ["C", 3]]]) {
    const book = new Workbook(), sheet = book.addWorksheet("Overview", { columns }); sheet.addChart(bound(sheet));
    await assert.rejects(async () => { await sheet.addRows(rows); await book.toBlob(); }, /chart|categor|numeric|number|range/i);
  }
  for (const values of [[0, 0, 0], [1, -1, 2]]) {
    const book = new Workbook(), sheet = book.addWorksheet("Overview", { columns }); const input = bound(sheet); input.type = "pie"; sheet.addChart(input);
    await assert.rejects(async () => { await sheet.addRows(values.map((value, i) => [String(i), value])); await book.toBlob(); }, /positive|nonnegative/i);
  }
});

test("category cache retention consumes the shared chart text budget before copying", async () => {
  const book = new Workbook({ limits: { maxChartTextCharacters: 10 } }), sheet = book.addWorksheet("Overview", { columns });
  sheet.addChart(bound(sheet, { rowCount: 2 })); // Series heading costs eight characters; the categories need three more.
  await assert.rejects(async () => { await sheet.addRows([["A", 1], ["BB", 2]]); await book.toBlob(); }, error => error.code === "RESOURCE_LIMIT" && /maxChartTextCharacters/.test(error.message));
});

test("cancelling a bound source preserves its reason instead of finalizing an incomplete cache", { timeout: 5000 }, async () => {
  const controller = new AbortController(), reason = new Error("Stop selected analysis"), chunks = [];
  const book = new Workbook({ compression: "store", signal: controller.signal, sink: { write(bytes) { chunks.push(bytes.slice()); } } });
  const sheet = book.addWorksheet("Overview", { columns }); sheet.addChart(bound(sheet));
  let returned = 0, release; const released = new Promise(resolve => { release = resolve; });
  async function* rows() {
    try { yield ["Healthy", 84]; yield ["Attention", 8]; controller.abort(reason); yield ["Offline", 3]; }
    finally { returned++; release(); }
  }
  await assert.rejects(sheet.addRows(rows()), error => error === reason);
  await released; assert.equal(returned, 1);
  await assert.rejects(book.finish(), error => error === reason);
  assert.equal(Buffer.concat(chunks).indexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06])), -1);
});
