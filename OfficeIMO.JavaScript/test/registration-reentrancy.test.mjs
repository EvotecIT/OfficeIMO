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
