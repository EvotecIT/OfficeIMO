import test from "node:test";
import assert from "node:assert/strict";
import { Workbook, Cell } from "../dist/xlsx/index.js";
import { writeCsv, writeCsvTo } from "../dist/csv/index.js";
import { BlobByteSink } from "../dist/core/index.js";
import { readZip } from "./zip-reader.mjs";

test("table presentation retains typed values and number/date formats across highlighting", async () => {
  const book = new Workbook({ compression: "store", dateMode: "utc" });
  const header = book.styles.add({ font: { bold: true, color: "FFFFFF" }, fill: { color: "203864" }, verticalAlignment: "center" });
  const columns = [{ header: "Name", key: "name", width: 24 }, { header: "Latency", key: "latency", type: "number", format: "0.000" },
    { header: "Seen", key: "seen", type: "date", format: "yyyy-mm-dd" }, { header: "Healthy", key: "healthy", type: "boolean" }];
  const seen = new Date("2026-10-06T00:00:00Z");
  const sheet = book.addWorksheet("Report", { columns, table: { name: "ControllerReport", style: "TableStyleMedium9" },
    headerStyle: header, freezeHeader: true, freezeColumns: 1, headerHeight: 30, rowHeight: 22,
    alternatingRowStyle: { fill: { color: "EAF1F8" } },
    rowStyle: ({ values }) => values[3] === false ? { fill: { color: "FCE4D6" }, font: { bold: true } } : undefined,
    cellStyle: ({ value, columnIndex }) => columnIndex === 2 && value > 100 ? { font: { color: "C00000", underline: true } } : undefined
  });
  await sheet.addRows([{ name: "Łódź 🧪", latency: 12.5, seen, healthy: true }]);
  await sheet.addRows([{ name: "Warsaw", latency: 125.75, seen, healthy: false }]);
  const zip = await readZip(await book.toBlob()), xml = zip.get("xl/worksheets/sheet1.xml").content, styles = zip.get("xl/styles.xml").content;
  assert.equal(sheet.rowCount, 2);
  assert.match(xml, /xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight"/);
  assert.match(xml, /r="1" ht="30" customHeight="1"/); assert.match(xml, /r="3" ht="22" customHeight="1"/);
  assert.match(xml, /min="2" max="2" width="20"/);
  assert.match(xml, /r="B3" s="\d+"><v>125.75<\/v>/); assert.match(xml, /r="D3" s="\d+" t="b"><v>0<\/v>/);
  const xfs = [...styles.matchAll(/<xf numFmtId="(\d+)"[^>]*>/g)].slice(1);
  const style = cell => Number(new RegExp('r="' + cell + '" s="(\\d+)"').exec(xml)[1]);
  assert.equal(xfs[style("B2")][1], xfs[style("B3")][1]);
  assert.equal(xfs[style("C2")][1], xfs[style("C3")][1]);
  assert.notEqual(xfs[style("C3")][1], "0");
  assert.match(styles, /rgb="FFC00000"/); assert.match(styles, /<u\/>/); assert.match(styles, /vertical="center"/);
  assert.match(xml, /<tablePart r:id="table"/);
  assert.match(zip.get("xl/tables/table1.xml").content, /name="ControllerReport" displayName="ControllerReport" ref="A1:D3"/);
  assert.match(zip.get("xl/tables/table1.xml").content, /name="TableStyleMedium9".*showRowStripes="1"/);
  assert.match(zip.get("xl/worksheets/_rels/sheet1.xml.rels").content, /Target="..\/tables\/table1.xml"/);
});

test("table identifiers and headers reject ambiguous input; empty exports remain header-only", async () => {
  const book = new Workbook(), columns = [{ header: "Value" }];
  for (const name of ["A1", "R2C3", "with space", "1Table"]) assert.throws(() => book.addWorksheet("Bad", { columns, table: { name } }), TypeError);
  for (const columns of [[{ header: "" }], [{ header: "V" }, { header: "v" }], [{ header: "A\u0001" }, { header: "A" }]])
    assert.throws(() => book.addWorksheet("Bad", { columns, table: {} }), TypeError);
  assert.throws(() => book.addWorksheet("Bad", { columns, table: { style: "TableStyleMedium29" } }), TypeError);
  book.addWorksheet("Empty", { columns, table: { name: "EmptyTable" } });
  assert.throws(() => book.addWorksheet("Duplicate", { columns, table: { name: "emptytable" } }), /Duplicate/);
  const zip = await readZip(await book.toBlob());
  assert.ok(!zip.has("xl/tables/table1.xml"));
  assert.doesNotMatch(zip.get("xl/worksheets/sheet1.xml").content, /tableParts|r="2"/);
});

test("style overlays preserve components, deduplicate and respect explicit Cell presentation", async () => {
  const book = new Workbook({ compression: "store" }), base = book.styles.add({ font: { name: "Arial", italic: true, color: "123456" },
    border: { bottom: { style: "thin" } }, numberFormat: "0.00%", wrapText: true });
  const overlay = { font: { bold: true }, border: { top: { style: "double" } }, verticalAlignment: "top" };
  const composed = book.styles.compose(base, overlay);
  assert.equal(composed, book.styles.compose(base, overlay));
  const sheet = book.addWorksheet("Cells", { columns: [{ header: "Rate", style: base }],
    rowStyle: () => ({ fill: { color: "FF0000" } }) });
  await sheet.addRows([[new Cell(0.25, composed)]]);
  const zip = await readZip(await book.toBlob()), styles = zip.get("xl/styles.xml").content;
  assert.match(zip.get("xl/worksheets/sheet1.xml").content, new RegExp('r="A2" s="' + composed + '"'));
  assert.match(styles, /<b\/><i\/>.*rgb="FF123456".*name val="Arial"/);
  assert.match(styles, /<top style="double">.*<bottom style="thin">/);
  assert.match(styles, /formatCode="0.00%"/);
});

test("CSV formatting protects formatted strings, quotes once and retains long Unicode data", async () => {
  const columns = [{ header: "Name", key: "name", valueFormatter: value => "=" + value },
    { header: "Healthy", key: "healthy", valueFormatter: value => value ? "Yes" : "No" },
    { header: "Value", key: "value" }];
  const rows = [{ name: 'Łódź, "🧪"', healthy: false, value: null }];
  const options = { columns, quote: "strings", nullValue: "missing" };
  const expected = '"Name","Healthy","Value"\r\n"\'=Łódź, ""🧪""","No","missing"\r\n';
  assert.equal(await (await writeCsv(rows, options)).text(), expected);
  const sink = new BlobByteSink(); await writeCsvTo(rows, sink, options);
  assert.equal(await sink.toBlob().text(), expected);
  assert.equal(await (await writeCsv([[-12.5, true, null]], { columns: [{ header: "N" }, { header: "B" }, { header: "E" }], quote: "all", includeHeader: false })).text(), '"-12.5","True",""\r\n');
  const long = "Łódź 🧪".repeat(10000);
  assert.equal(await (await writeCsv([[long]], { columns: [{ header: "V" }], includeHeader: false })).text(), long + "\r\n");
});

test("formatter and presentation failures return producers and prevent partial output", async () => {
  const error = new Error("format failed"); let returned = false;
  function* rows() { try { yield ["first"]; yield ["second"]; } finally { returned = true; } }
  await assert.rejects(writeCsv(rows(), { columns: [{ header: "V", valueFormatter: () => { throw error; } }] }), e => e === error);
  assert.equal(returned, true);
  const book = new Workbook(), sheet = book.addWorksheet("Failure", { columns: [{ header: "V" }], cellStyle: () => { throw error; } });
  await assert.rejects(sheet.addRows([["value"]]), e => e === error);
  await assert.rejects(book.toBlob(), e => e === error);
  const asynchronous = new Workbook().addWorksheet("Async", { columns: [{ header: "V" }], rowStyle: async () => { throw error; } });
  await assert.rejects(asynchronous.addRows([["value"]]), /synchronous CellStyle/);
  const styledBook = new Workbook({ compression: "store" }), styled = styledBook.addWorksheet("Styled", {
    columns: [{ header: "V" }], rowStyle: async () => { throw error; }
  });
  await assert.rejects(styled.addRows([[new Cell("value", 0)]]), /synchronous CellStyle/);
  await assert.rejects(styledBook.toBlob(), /synchronous CellStyle/);
  for (const callback of ["rowStyle", "cellStyle"]) {
    const invalid = new Workbook().addWorksheet("Invalid", { columns: [{ header: "V" }], [callback]: () => null });
    await assert.rejects(invalid.addRows([[new Cell("value", 0)]]), /CellStyle object/);
  }
  await assert.rejects(writeCsv([["value"]], { columns: [{ header: "V", valueFormatter: async () => { throw error; } }] }), /synchronous/);
});

test("report hyperlinks and PNG anchors use valid independent package relationships", async () => {
  const png = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jfKsAAAAASUVORK5CYII=", "base64");
  const original = Buffer.from(png), book = new Workbook({ compression: "store" });
  const sheet = book.addWorksheet("Links", { columns: [{ header: "Name" }], hyperlinks: [{ cell: "A2", target: "https://example.com/?x=1&y=2", tooltip: "Łódź <report> _x0041_" }] });
  await sheet.addRows([["Łódź"]]);
  sheet.addImage({ data: png, row: 4, column: 2, width: 180, height: 80, description: "Chart <🧪>" });
  png.fill(0);
  assert.throws(() => sheet.addHyperlink({ cell: "A2", target: "https://example.com" }), /Duplicate/);
  for (const target of ["javascript:alert(1)", "file:///tmp/report", "https://user:password@example.com", "https://example.com/\r\n"])
    assert.throws(() => sheet.addHyperlink({ cell: "A3", target }), TypeError);
  assert.throws(() => sheet.addImage({ data: new Uint8Array(), row: 1, column: 1, width: 100, height: 100 }), TypeError);
  const zip = await readZip(await book.toBlob());
  assert.deepEqual(zip.get("xl/media/sheet1-image1.png").bytes, original);
  const xml = zip.get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /hyperlink ref="A2" r:id="link1" tooltip="Łódź &lt;report&gt; _x005F_x0041_"/);
  assert.match(xml, /drawing r:id="drawing"/);
  const relationships = zip.get("xl/worksheets/_rels/sheet1.xml.rels").content;
  assert.match(relationships, /Target="https:\/\/example.com\/\?x=1&amp;y=2" TargetMode="External"/);
  const drawing = zip.get("xl/drawings/drawing1.xml").content;
  assert.match(drawing, /<xdr:col>1<\/xdr:col>/); assert.match(drawing, /<xdr:row>3<\/xdr:row>/);
  assert.match(drawing, /cx="1714500" cy="762000"/); assert.match(drawing, /descr="Chart &lt;🧪&gt;"/);
  assert.match(zip.get("xl/drawings/_rels/drawing1.xml.rels").content, /Target="..\/media\/sheet1-image1.png"/);
  const badBook = new Workbook(), bad = badBook.addWorksheet("Bad", { columns: [{ header: "V" }] });
  await bad.addRows([["row"]]); bad.addHyperlink({ cell: "A3", target: "https://example.com" });
  await assert.rejects(badBook.toBlob(), /within the exported rows/);
});
