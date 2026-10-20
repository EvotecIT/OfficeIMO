import test from "node:test";
import assert from "node:assert/strict";
import { Workbook } from "../dist/xlsx/index.js";
import { readZip } from "./zip-reader.mjs";

const chart = () => ({ type: "line", data: { categories: ["A"], series: [{ name: "S", values: [1] }] }, row: 1, column: 1, width: 200, height: 100 });
const reentered = error => error.code === "INVALID_STATE" && /reentered/.test(error.message);

test("chart input getters cannot bypass workbook registration limits and a rejected capture leaves normal registration usable", async () => {
  for (const limits of [{ maxCharts: 1 }, { maxSheets: 2 }, { maxChartPoints: 1 }, { maxChartTextCharacters: 2 }]) {
    const book = new Workbook({ limits }), sheet = book.addWorksheet("Report"), source = chart();
    assert.throws(() => sheet.addChart({ ...source, get data() { sheet.addChart(chart()); return source.data; } }), reentered);
    sheet.addChart(chart());
    const zip = await readZip(await book.toBlob());
    assert.equal((await book.finish()).sheets, 2);
    assert.ok(zip.has("xl/charts/chart1.xml")); assert.ok(!zip.has("xl/charts/chart2.xml"));
    assert.equal((zip.get("xl/workbook.xml").content.match(/state="hidden"/g) ?? []).length, 1);
  }
  const book = new Workbook({ limits: { maxSheets: 3 } }), sheet = book.addWorksheet("Report"), source = chart();
  assert.throws(() => sheet.addChart({ ...source, get data() { book.addWorksheet("Nested"); return source.data; } }), reentered);
  book.addWorksheet("Recovery"); sheet.addChart(chart());
  const zip = await readZip(await book.toBlob());
  assert.equal((await book.finish()).sheets, 3); assert.doesNotMatch(zip.get("xl/workbook.xml").content, /Nested/);
});

test("worksheet option getters cannot reenter sheet or chart registration and failed attempts release the guard", async () => {
  for (const nested of ["sheet", "chart"]) {
    const book = new Workbook({ limits: { maxSheets: 3, maxCharts: 1 } }), report = book.addWorksheet("Report");
    const options = nested === "sheet" ? { columns: [{ get header() { book.addWorksheet("Nested"); return "Value"; } }] }
      : { columns: [{ header: "Value" }], alternatingRowStyle: { get font() { report.addChart(chart()); return { bold: true }; } } };
    assert.throws(() => book.addWorksheet("Outer", options), reentered);
    const recovered = book.addWorksheet("Outer", { columns: [{ header: "Value" }] });
    await recovered.addRows([[42]]); report.addChart(chart());
    const zip = await readZip(await book.toBlob());
    assert.equal((await book.finish()).sheets, 3); assert.equal((await book.finish()).rows, 1);
    assert.match(zip.get("xl/worksheets/sheet2.xml").content, /<v>42<\/v>/);
    assert.doesNotMatch(zip.get("xl/workbook.xml").content, /Nested|Outer \(2\)/);
    assert.ok(zip.has("xl/charts/chart1.xml")); assert.ok(!zip.has("xl/charts/chart2.xml"));
  }
});

test("chart and worksheet capture recheck preservation sheets allocated by a getter before committing", async () => {
  for (const registering of ["chart", "worksheet"]) {
    const book = new Workbook({ oversizedText: "preserve", limits: { maxSheets: 2 } });
    const report = book.addWorksheet("Report", { columns: [{ header: "Value" }], autoSize: { sampleRows: 1 } });
    let pending;
    const capture = () => { pending ??= report.addRows([["x".repeat(32768)]]); };
    const attempt = registering === "chart" ? () => {
      const source = chart();
      report.addChart({ ...source, get data() { capture(); return source.data; } });
    } : () => book.addWorksheet("Rejected", { columns: [{ get header() { capture(); return "Value"; } }] });
    assert.throws(attempt, error => error.code === "RESOURCE_LIMIT" && /maxSheets/.test(error.message));
    await pending;
    await report.addRows([[42]]);
    assert.throws(() => book.addWorksheet("No capacity"), error => error.code === "RESOURCE_LIMIT" && /maxSheets/.test(error.message));
    assert.throws(() => report.addChart(chart()), error => error.code === "RESOURCE_LIMIT" && /maxSheets/.test(error.message));
    const zip = await readZip(await book.toBlob()), result = await book.finish();
    assert.equal(result.sheets, 2); assert.equal(result.rows, 2);
    assert.match(zip.get("xl/workbook.xml").content, /name="Text overflow"/);
    assert.doesNotMatch(zip.get("xl/workbook.xml").content, /Rejected|No capacity|Chart data/);
    assert.ok(!zip.has("xl/charts/chart1.xml"));
    assert.match(zip.get("xl/worksheets/sheet1.xml").content, /<v>42<\/v>/);
  }
});


test("worksheet naming commits after option capture can allocate a preservation sheet", async () => {
  const book = new Workbook({ oversizedText: "preserve", limits: { maxSheets: 3 } });
  const report = book.addWorksheet("Report", { columns: [{ header: "Value" }], autoSize: { sampleRows: 1 } });
  let reads = 0, pending;
  const outer = book.addWorksheet("Text overflow", { get columns() {
    reads++; pending = report.addRows([["x".repeat(32768)]]);
    return [{ header: "Sheet", value: (_row, context) => context.sheetName }, { header: "Value", key: "value" }];
  } });
  await pending; await outer.addRows([{ value: 42 }]);
  assert.equal(reads, 1, "worksheet option properties are captured once before registration");
  assert.equal(outer.name, "Text overflow (2)");
  const zip = await readZip(await book.toBlob()), xml = zip.get("xl/workbook.xml").content;
  const names = [...xml.matchAll(/<sheet name="([^"]+)"/g)].map(match => match[1]);
  assert.deepEqual(names, ["Report", "Text overflow (2)", "Text overflow"]);
  assert.equal(new Set(names.map(name => name.toLowerCase())).size, names.length);
  assert.match(zip.get("xl/worksheets/sheet2.xml").content, /Text overflow \(2\)/);
  assert.match(zip.get("xl/worksheets/sheet2.xml").content, /<v>42<\/v>/);
});

test("chart capture rechecks category text charged by input getters without retaining a rejected chart", async () => {
  for (const kind of ["snapshot", "range"]) {
    const book = new Workbook({ limits: { maxChartTextCharacters: 5 } });
    const source = book.addWorksheet("Bound", { columns: [{ header: "Category", key: "category" }, { header: "S", key: "value" }], autoSize: { sampleRows: 1 } });
    const future = book.addWorksheet("Future", { columns: [{ header: "Category", key: "category" }, { header: "LONG", key: "value" }] });
    const target = book.addWorksheet("Report"), placement = { type: "line", row: 1, column: 1, width: 200, height: 100 };
    source.addChart({ ...placement, data: { kind: "range", sheet: source, rowCount: 1, categoryColumn: "category", series: [{ column: "value" }] } });
    let pending, reads = 0;
    assert.throws(() => target.addChart({ ...placement, get data() {
      reads++; pending = source.addRows([{ category: "AB", value: 1 }]);
      return kind === "snapshot" ? { categories: ["ABC"], series: [{ name: "S", values: [1] }] }
        : { kind: "range", sheet: future, rowCount: 1, categoryColumn: "category", series: [{ column: "value" }] };
    } }), error => error.code === "RESOURCE_LIMIT" && /maxChartTextCharacters/.test(error.message));
    await pending; assert.equal(reads, 1);
    target.addChart({ ...placement, data: { categories: ["Z"], series: [{ name: "S", values: [2] }] } });
    const zip = await readZip(await book.toBlob());
    assert.ok(zip.has("xl/charts/chart1.xml")); assert.ok(zip.has("xl/charts/chart2.xml")); assert.ok(!zip.has("xl/charts/chart3.xml"));
    assert.match(zip.get("xl/charts/chart1.xml").content, /<c:v(?: [^>]*)?>AB<\/c:v>/);
    assert.match(zip.get("xl/charts/chart2.xml").content, /<c:v(?: [^>]*)?>Z<\/c:v>/);
    assert.equal((zip.get("xl/workbook.xml").content.match(/state="hidden"/g) ?? []).length, 1);
  }
});
