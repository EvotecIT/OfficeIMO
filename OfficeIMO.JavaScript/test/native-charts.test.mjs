import test from "node:test";
import assert from "node:assert/strict";
import { Workbook } from "../dist/xlsx/index.js";
import { readZip } from "./zip-reader.mjs";

const chartData = () => ({ categories: ["Łódź <🧪>", "Warsaw", "Gdańsk"], series: [
  { name: "Actual & verified", values: [12.5, -2, 0], color: "#203864" },
  { name: "Target", values: [10, 10, 10] }
] });
const chart = (type = "columnClustered", data = chartData()) => ({ type, data, row: 5, column: 2, width: 640, height: 360,
  title: "Measured <report>", description: "Native chart 🧪" });
const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jfKsAAAAASUVORK5CYII=", "base64");
const chartXml = (zip, id) => zip.get(`xl/charts/chart${id}.xml`).content.replaceAll("&apos;", "'");

test("native charts share image drawings and retain editable hidden source cells with matching caches", async () => {
  const book = new Workbook({ compression: "store" });
  // Names are reserved in both directions: existing and future user worksheets cannot collide.
  const existing = book.addWorksheet("Chart data 1", { columns: [{ header: "Existing report" }] });
  await existing.addRows([["Preserved"]]);
  const report = book.addWorksheet("O'Brien report", { columns: [{ header: "Report" }] });
  await report.addRows([["Summary"]]);
  report.addImage({ data: png, row: 1, column: 1, width: 100, height: 50 });
  const source = chartData(), input = chart("columnClustered", source);
  input.dataLabels = { showValue: true, position: "outsideEnd", numberFormat: "0.00" };
  input.categoryAxis = { title: "Location" };
  input.valueAxis = { title: "Milliseconds", minimum: -10, maximum: 50, numberFormat: "0.0" };
  report.addChart(input);
  source.categories[0] = "Mutated category"; source.series[0].values[0] = 999;
  source.series[0].name = "Mutated series"; input.dataLabels.numberFormat = "0%"; input.valueAxis.maximum = 999;
  const future = book.addWorksheet("Chart data 1 (2)", { columns: [{ header: "Future report" }] });
  assert.equal(future.name, "Chart data 1 (2) (2)");
  await future.addRows([["Future"]]);
  const zip = await readZip(await book.toBlob());
  const workbook = zip.get("xl/workbook.xml").content;
  assert.match(workbook, /name="Chart data 1 \(2\)" sheetId="4" state="hidden" r:id="rId4"/);
  assert.equal(book.rowCount, 3, "Hidden source rows are not report progress rows.");
  const drawing = zip.get("xl/drawings/drawing2.xml").content;
  assert.equal((drawing.match(/<xdr:oneCellAnchor>/g) ?? []).length, 2);
  assert.match(drawing, /<xdr:cNvPr id="1" name="Image 1"/);
  assert.match(drawing, /<xdr:cNvPr id="2" name="Chart 1" descr="Native chart 🧪"/);
  assert.match(drawing, /<c:chart[^>]*r:id="chart1"/);
  assert.match(drawing, /<xdr:col>1<\/xdr:col>.*<xdr:row>4<\/xdr:row>/);
  assert.match(drawing, /cx="6096000" cy="3429000"/);
  const drawingLinks = zip.get("xl/drawings/_rels/drawing2.xml.rels").content;
  assert.match(drawingLinks, /Id="image1"[^>]*Target="..\/media\/sheet2-image1.png"/);
  assert.match(drawingLinks, /Id="chart1"[^>]*Type="[^"]+\/chart"[^>]*Target="..\/charts\/chart1.xml"/);
  const worksheetLinks = zip.get("xl/worksheets/_rels/sheet2.xml.rels").content;
  assert.match(worksheetLinks, /Id="drawing"[^>]*Target="..\/drawings\/drawing2.xml"/);
  const xml = chartXml(zip, 1);
  assert.match(xml, /<c:barDir val="col"\/><c:grouping val="clustered"\/>/);
  assert.match(xml, /'Chart data 1 \(2\)'!\$A\$2:\$A\$4/);
  assert.match(xml, /'Chart data 1 \(2\)'!\$B\$1/);
  assert.match(xml, /'Chart data 1 \(2\)'!\$B\$2:\$B\$4/);
  assert.match(xml, /<c:ptCount val="3"\/>/);
  assert.match(xml, /Łódź &lt;🧪&gt;/); assert.match(xml, /Actual &amp; verified/);
  assert.match(xml, /<c:v[^>]*>12.5<\/c:v>/); assert.match(xml, /<c:v[^>]*>-2<\/c:v>/);
  assert.match(xml, /<a:srgbClr val="203864"/);
  assert.match(xml, /<c:numFmt formatCode="0.00" sourceLinked="0"\/>/);
  assert.match(xml, /<c:scaling><c:orientation val="minMax"\/><c:max val="50"\/><c:min val="-10"\/><\/c:scaling>/);
  assert.match(xml, /<c:plotVisOnly val="0"/);
  assert.doesNotMatch(xml, /Mutated|>999<|val="999"/);
  const dataSheet = zip.get("xl/worksheets/sheet4.xml").content;
  assert.match(dataSheet, /r="A2"[^>]*><is><t xml:space="preserve">Łódź &lt;🧪&gt;/);
  assert.match(dataSheet, /r="B1"[^>]*><is><t xml:space="preserve">Actual &amp; verified/);
  assert.match(dataSheet, /r="B2"[^>]*><v>12.5<\/v>/);
  assert.match(dataSheet, /r="B3"[^>]*><v>-2<\/v>/);
  assert.match(dataSheet, /r="C4"[^>]*><v>10<\/v>/);
  const types = zip.get("[Content_Types].xml").content;
  assert.match(types, /PartName="\/xl\/charts\/chart1.xml" ContentType="application\/vnd.openxmlformats-officedocument.drawingml.chart\+xml"/);
  assert.match(types, /PartName="\/xl\/drawings\/drawing2.xml"/);
  assert.equal(await book.toBlob(), await book.toBlob(), "Repeated completion reuses the same archive.");
});

test("column, horizontal bar, line and pie charts have separate native parts and valid chart-local axes", async () => {
  const book = new Workbook({ compression: "store" });
  for (const type of ["columnClustered", "barClustered", "line", "pie"]) {
    const sheet = book.addWorksheet(type, { columns: [{ header: "Summary" }] });
    await sheet.addRows([[type]]);
    const data = type === "pie" ? { categories: ["Healthy", "Warning", "Offline"], series: [{ name: "Controllers", values: [8, 2, 0] }] } : chartData();
    sheet.addChart({ ...chart(type, data), legend: "bottom", dataLabels: type === "pie" ? { showValue: false, showPercent: true, showCategoryName: true, position: "bestFit" } : { position: type === "line" ? "above" : "insideEnd" } });
  }
  const blob = await book.toBlob(), zip = await readZip(blob);
  const types = ["columnClustered", "barClustered", "line", "pie"];
  for (let i = 0; i < types.length; i++) {
    const id = i + 1, xml = chartXml(zip, id);
    assert.match(zip.get(`xl/drawings/_rels/drawing${id}.xml.rels`).content, new RegExp(`Id="chart${id}"[^>]*Target="../charts/chart${id}.xml"`));
    assert.match(xml, new RegExp(`'Chart data ${id}'!\\$A\\$2:\\$A\\$4`));
    assert.match(xml, /<c:legendPos val="b"/);
    if (types[i] === "pie") {
      assert.match(xml, /<c:pieChart>/); assert.match(xml, /<c:showPercent val="1"/);
      assert.equal((xml.match(/<c:ser>/g) ?? []).length, 1);
      assert.doesNotMatch(xml, /<c:catAx>|<c:valAx>|<c:axId/);
    } else {
      assert.match(xml, /<c:catAx><c:axId val="1"/); assert.match(xml, /<c:valAx><c:axId val="2"/);
      assert.match(xml, /<c:crossAx val="2"/); assert.match(xml, /<c:crossAx val="1"/);
      assert.equal((xml.match(/<c:ser>/g) ?? []).length, 2);
      if (types[i] === "barClustered") {
        assert.match(xml, /<c:barDir val="bar"/); assert.match(xml, /<c:catAx>.*<c:axPos val="l"/); assert.match(xml, /<c:valAx>.*<c:axPos val="b"/);
      } else if (types[i] === "line") {
        assert.match(xml, /<c:lineChart><c:grouping val="standard"/);
        assert.match(xml, /<a:ln w="25400"><a:solidFill>/);
        assert.match(xml, /<c:dLblPos val="t"/);
      }
    }
  }
  assert.equal((zip.get("xl/workbook.xml").content.match(/state="hidden"/g) ?? []).length, 4);
  assert.equal((await book.finish()).rows, 4);
  assert.equal((await book.finish()).sheets, 8);
});

test("chart input getters and array values are captured once before asynchronous output", async () => {
  const calls = new Map(), once = (key, value) => { calls.set(key, (calls.get(key) ?? 0) + 1); return value; };
  const categories = ["First"], values = [2];
  Object.defineProperty(categories, 0, { get: () => once("category", "First"), enumerable: true });
  Object.defineProperty(values, 0, { get: () => once("value", 2), enumerable: true });
  const series = { get name() { return once("name", "Series"); }, get values() { return once("values", values); } };
  const data = { get categories() { return once("categories", categories); }, get series() { return once("series", [series]); } };
  const requested = { ...chart(), get data() { return once("data", data); }, get title() { return once("title", "Title"); } };
  const book = new Workbook({ compression: "store" }), sheet = book.addWorksheet("Report");
  sheet.addChart(requested);
  await readZip(await book.toBlob());
  for (const key of ["data", "title", "categories", "series", "name", "values", "category", "value"]) assert.equal(calls.get(key), 1, key + " is read once.");
});

test("snapshot chart array growth cannot exceed the preflighted numeric point budget", async () => {
  const categories = ["A"], series = [{ name: "S", values: [1, 2] }];
  Object.defineProperty(categories, 0, { get() { if (categories.length === 1) categories.push("B"); return "A"; }, enumerable: true });
  const book = new Workbook({ compression: "store", limits: { maxChartPoints: 1 } }), sheet = book.addWorksheet("Report");
  assert.throws(() => sheet.addChart(chart("line", { categories, series })), /match the category count/);
  let extraReads = 0;
  const selected = [{ name: "S", values: [1] }];
  Object.defineProperty(selected, 0, { get() {
    if (selected.length === 1) Object.defineProperty(selected, 1, { enumerable: true, get() { extraReads++; return { name: "Unexpected", values: [2] }; } });
    return { name: "S", values: [1] };
  }, enumerable: true });
  sheet.addChart(chart("line", { categories: ["A"], series: selected }));
  const zip = await readZip(await book.toBlob()), xml = chartXml(zip, 1);
  assert.equal(extraReads, 0); assert.equal((xml.match(/<c:ser>/g) ?? []).length, 1);
  assert.match(xml, /<c:numCache><c:formatCode>General<\/c:formatCode><c:ptCount val="1"/);
  assert.ok(!zip.has("xl/charts/chart2.xml"));
});

test("native chart language is explicit canonical metadata with no forced report locale", async () => {
  const book = new Workbook({ compression: "store" }), sheet = book.addWorksheet("Report");
  sheet.addChart(chart("line", { categories: ["Łódź"], series: [{ name: "Wynik", values: [1] }] }));
  sheet.addChart({ ...chart("line", { categories: ["Łódź"], series: [{ name: "Wynik", values: [1] }] }), language: "pl-pl" });
  for (const language of ["", "Polish language", "en_US", 42]) assert.throws(() => sheet.addChart({ ...chart(), language }), /language/);
  const zip = await readZip(await book.toBlob());
  assert.doesNotMatch(chartXml(zip, 1), /<c:lang/);
  assert.match(chartXml(zip, 2), /<c:lang val="pl-PL"\/>/);
});

test("streamed worksheets retain their charts after closing and use distinct source-sheet and relationship identities", async () => {
  const chunks = []; let writing = false;
  const book = new Workbook({ compression: "store", sink: { async write(bytes) {
    assert.equal(writing, false, "Destination writes are awaited."); writing = true;
    await Promise.resolve(); chunks.push(bytes.slice()); writing = false;
  } } });
  for (const name of ["First", "Second"]) {
    const sheet = book.addWorksheet(name, { columns: [{ header: "Status" }] });
    await sheet.addRows([[name]]); sheet.addChart(chart("line")); await sheet.close();
    assert.throws(() => sheet.addChart(chart()), error => error.code === "INVALID_STATE");
  }
  const result = await book.finish();
  assert.deepEqual(result, { rows: 2, sheets: 4, bytes: chunks.reduce((total, chunk) => total + chunk.length, 0) });
  assert.equal(await book.finish(), result);
  const zip = await readZip(new Blob(chunks, { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  for (const id of [1, 2]) {
    assert.match(zip.get(`xl/worksheets/sheet${id}.xml`).content, /<drawing r:id="drawing"/);
    assert.match(zip.get(`xl/drawings/_rels/drawing${id}.xml.rels`).content, new RegExp(`Id="chart${id}"`));
    assert.match(chartXml(zip, id), new RegExp(`'Chart data ${id}'!`));
    assert.match(zip.get(`xl/worksheets/sheet${id + 2}.xml`).content, /<v>12.5<\/v>/);
  }
});

test("chart rejection is explicit and does not consume source-sheet names, IDs or budgets", async () => {
  const book = new Workbook({ compression: "store", limits: { maxCharts: 1, maxChartPoints: 6, maxChartTextCharacters: 300, maxSheets: 2 } });
  const sheet = book.addWorksheet("Report");
  for (const invalid of [
    { ...chart(), type: "surface" }, { ...chart(), smooth: true },
    chart("pie"), chart("pie", { categories: ["A"], series: [{ name: "S", values: [-1] }] }),
    chart("pie", { categories: ["A"], series: [{ name: "S", values: [0] }] }),
    chart("line", { categories: ["A"], series: [{ name: "S", values: [Infinity] }] }),
    chart("line", { categories: ["A"], series: [{ name: "S", values: [] }] }),
    { ...chart(), row: 0 }, { ...chart(), width: -1 },
    { ...chart(), dataLabels: { showPercent: true } }, { ...chart(), dataLabels: { position: "above" } },
    { ...chart(), valueAxis: { minimum: 10, maximum: 0 } },
    { ...chart("pie", { categories: ["A"], series: [{ name: "S", values: [1] }] }), categoryAxis: {} },
    chart("line", { categories: ["A"], series: [{ name: "S", values: [1], xValues: [0] }] })
  ]) assert.throws(() => sheet.addChart(invalid));
  sheet.addChart(chart());
  assert.throws(() => sheet.addChart(chart()), error => error.code === "RESOURCE_LIMIT" && /maxCharts/.test(error.message));
  assert.throws(() => book.addWorksheet("Another"), error => error.code === "RESOURCE_LIMIT" && /maxSheets/.test(error.message));
  const zip = await readZip(await book.toBlob());
  assert.ok(zip.has("xl/charts/chart1.xml")); assert.ok(!zip.has("xl/charts/chart2.xml"));
  assert.match(zip.get("xl/workbook.xml").content, /name="Chart data 1"/);
  assert.throws(() => sheet.addChart(chart()), error => error.code === "INVALID_STATE");
});

test("chart source rows and caches honor workbook-wide point, text, cell, sheet and output budgets", async () => {
  for (const limits of [{ maxCharts: 0 }, { maxChartPoints: 5 }, { maxChartTextCharacters: 10 }, { maxSheets: 1 }]) {
    const book = new Workbook({ limits }), sheet = book.addWorksheet("Report");
    assert.throws(() => sheet.addChart(chart()), error => error.code === "RESOURCE_LIMIT");
  }
  const book = new Workbook({ limits: { maxChartPoints: 10 } }), sheet = book.addWorksheet("Report");
  sheet.addChart(chart());
  assert.throws(() => sheet.addChart(chart()), error => error.code === "RESOURCE_LIMIT" && /maxChartPoints/.test(error.message));
  for (const limits of [{ maxCells: 4 }, { maxRows: 2 }, { maxOutputBytes: 100 }]) {
    const capped = new Workbook({ compression: "store", limits }); capped.addWorksheet("Report").addChart(chart());
    await assert.rejects(capped.toBlob(), error => error.code === "RESOURCE_LIMIT");
  }
});

test("chart generation cancels and propagates destination failure without finalizing partial output", async () => {
  for (const cancel of [true, false]) {
    const controller = new AbortController(), failure = new Error(cancel ? "Cancel chart part" : "Chart destination failed");
    const chunks = [];
    const book = new Workbook({ compression: "store", signal: controller.signal, sink: { write(bytes) {
      chunks.push(bytes.slice());
      if (new TextDecoder().decode(bytes).includes("xl/charts/chart1.xml")) {
        if (cancel) controller.abort(failure); else throw failure;
      }
    } } });
    book.addWorksheet("Report").addChart(chart());
    await assert.rejects(book.finish(), error => error === failure);
    await assert.rejects(book.finish(), error => error === failure);
    const output = Buffer.concat(chunks);
    assert.ok(output.length > 0, "The caller owns partial bytes.");
    assert.equal(output.indexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06])), -1, "No completed ZIP directory is returned.");
  }
});
