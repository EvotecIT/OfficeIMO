import test from "node:test";
import assert from "node:assert/strict";
import { Workbook } from "../dist/xlsx/index.js";
import { readZip } from "./zip-reader.mjs";

const columns = [{ header: "Name", key: "name", groups: ["Report"] }, { header: "Amount", key: "amount", type: "number", format: "0.00", groups: ["Report"] }];
const highlight = { type: "cellIs", range: { column: "amount" }, operator: "lessThan", value: 0, style: { fill: { color: "FFC7CE" } } };

test("live rules resolve final data ranges without changing typed values or report layout", async () => {
  for (const streamed of [false, true]) {
    const chunks = [], book = new Workbook({ compression: "store", ...(streamed ? { sink: { write(bytes) { chunks.push(bytes.slice()); } } } : {}) });
    const rules = [structuredClone(highlight), { type: "expression", range: { column: 1, through: 2 }, formula: '=$B4<0', style: { font: { bold: false, color: "9C0006" } }, stopIfTrue: true },
      { type: "colorScale", range: { column: 2 }, stops: [{ threshold: { type: "min" }, color: "F8696B" }, { threshold: { type: "percentile", value: 50 }, color: "FFEB84" }, { threshold: { type: "max" }, color: "63BE7B" }] },
      { type: "dataBar", range: "B4:B5", color: "638EC6", showValue: false }];
    const sheet = book.addSheet("Live", { columns, title: { text: "Report" }, table: {}, freezeHeader: true, footer: { totals: { amount: "sum" } }, print: { repeatHeaders: true }, conditionalFormats: rules });
    rules[0].style.fill.color = "000000"; rules[0].range.column = "name"; rules[1].formula = '1=0'; rules[2].stops[1].threshold.value = 99;
    await sheet.addRows([["negative", -12.5]]); await sheet.addRows([["positive", 5]]);
    let blob; if (streamed) { await sheet.close(); await book.finish(); blob = new Blob(chunks); } else blob = await book.toBlob();
    const zip = await readZip(blob), xml = zip.get("xl/worksheets/sheet1.xml").content, styles = zip.get("xl/styles.xml").content;
    assert.match(xml, /sqref="B4:B5"><cfRule type="cellIs" priority="1" dxfId="0" operator="lessThan"><formula>0<\/formula>/);
    assert.match(xml, /sqref="A4:B5"><cfRule type="expression" priority="2" dxfId="1" stopIfTrue="1"><formula>\$B4&lt;0<\/formula>/);
    assert.match(xml, /<cfvo type="percentile" val="50"\/>/); assert.match(xml, /<dataBar showValue="0">/);
    assert.match(xml, /r="B4"[^>]*><v>-12.5<\/v>/); assert.match(xml, /r="B5"[^>]*><v>5<\/v>/);
    assert.match(zip.get("xl/tables/table1.xml").content, /ref="A3:B6"/);
    assert.match(styles, /<dxfs count="2">/); assert.match(styles, /<b val="0"\/>/); assert.match(styles, /FFFFC7CE/);
    assert.match(styles, /<fgColor rgb="FFFFC7CE"\/><bgColor rgb="FFFFC7CE"\/>/);
    assert.doesNotMatch(styles, /FF000000/); assert.ok(xml.indexOf("</mergeCells>") < xml.indexOf("<conditionalFormatting"));
  }
});

test("differential styles are minimal and cached across worksheets and input property order", async () => {
  const book = new Workbook({ limits: { maxDifferentialStyles: 1, maxConditionalFormats: 3 } });
  for (let i = 0; i < 3; i++) {
    const style = i % 2 ? { numberFormat: '"_x0041_"0.0', font: { color: "#123456", italic: false }, border: { bottom: { color: "abcdef", style: "thin" } } } :
      { border: { bottom: { style: "thin", color: "FFABCDEF" } }, font: { italic: false, color: "FF123456" }, numberFormat: '"_x0041_"0.0' };
    await book.addSheet("Sheet" + i, { columns: [{ header: "Value", format: "0.00" }], conditionalFormats: [{ ...highlight, range: "A2", style }] }).addRows([[12.5]]);
  }
  const zip = await readZip(await book.toBlob()), styles = zip.get("xl/styles.xml").content;
  assert.match(styles, /<dxfs count="1">/); const dxf = /<dxf>(.*?)<\/dxf>/.exec(styles)[1];
  assert.match(dxf, /<i val="0"\/>/); assert.match(dxf, /_x005F_x0041_/);
  assert.doesNotMatch(dxf, /<sz|<name|<left|<right|<top|<fill/);
  for (let i = 1; i <= 3; i++) assert.match(zip.get(`xl/worksheets/sheet${i}.xml`).content, /dxfId="0"/);
});

test("conditional metadata rejects unsupported or invalid contracts before sheet-name registration", async () => {
  const invalid = [
    { ...highlight, range: "B2:A2" }, { ...highlight, range: "C2" }, { ...highlight, range: { column: "missing" } },
    { ...highlight, range: { column: 0 } }, { ...highlight, value: NaN }, { ...highlight, operator: "between", values: [2, 1], value: undefined },
    { ...highlight, style: { font: { name: "Calibri" } } }, { ...highlight, style: { font: { bold: "yes" } } }, { ...highlight, style: {} },
    { ...highlight, style: { fill: { color: "blue" } } }, { ...highlight, stopIfTrue: 1 }, { ...highlight, type: "iconSet" },
    { type: "expression", range: "A1", formula: "=", style: highlight.style },
    { type: "expression", range: "A1", formula: 'A1="bad\u0001"', style: highlight.style },
    { type: "expression", range: "A1", formula: "1".repeat(8193), style: highlight.style },
    { type: "dataBar", range: "B2", color: "123456", stopIfTrue: true },
    { type: "dataBar", range: "B2", color: "123456", minimum: { type: "max" } },
    { type: "colorScale", range: "B2", stops: [{ threshold: { type: "percent", value: -1 }, color: "123456" }, { threshold: { type: "max" }, color: "abcdef" }] }
  ];
  for (const rule of invalid) {
    const book = new Workbook(); assert.throws(() => book.addSheet("Recover", { columns, conditionalFormats: [rule] }));
    assert.equal(book.addSheet("Recover", { columns }).name, "Recover");
  }
  const limited = new Workbook({ limits: { maxConditionalFormats: 1, maxDifferentialStyles: 1 } });
  assert.throws(() => limited.addSheet("Recover", { columns, conditionalFormats: [highlight, highlight] }), { code: "RESOURCE_LIMIT" });
  assert.throws(() => limited.addSheet("Recover", { columns, conditionalFormats: [highlight, { ...highlight, style: { font: { bold: true } } }] }), { code: "RESOURCE_LIMIT" });
  limited.addSheet("Recover", { columns, conditionalFormats: [highlight] });
  assert.throws(() => limited.addSheet("Overflow", { columns, conditionalFormats: [highlight] }), { code: "RESOURCE_LIMIT" });
  const styles = new Workbook({ limits: { maxDifferentialStyles: 1 } });
  assert.throws(() => styles.addSheet("Recover", { columns, conditionalFormats: [highlight, { ...highlight, style: { font: { bold: true } } }] }), { code: "RESOURCE_LIMIT" });
  assert.equal(styles.addSheet("Recover", { columns, conditionalFormats: [highlight] }).name, "Recover");
});

test("empty data-only rules do not color report headings or totals; literal ranges keep bounds", async () => {
  const empty = new Workbook(); await empty.addSheet("Empty", { columns, title: { text: "Empty" }, footer: { totals: { amount: "sum" } }, conditionalFormats: [highlight] }).addRows([]);
  const xml = (await readZip(await empty.toBlob())).get("xl/worksheets/sheet1.xml").content;
  assert.doesNotMatch(xml, /<conditionalFormatting/);
  for (const streamed of [false, true]) {
    const book = new Workbook({ ...(streamed ? { sink: { write() {} } } : {}) });
    const sheet = book.addSheet("Bounds", { columns, conditionalFormats: [{ ...highlight, range: "B99" }] }); await sheet.addRows([["one", 1]]);
    let original; try { await sheet.close(); } catch (error) { original = error; }
    assert.match(String(original), /exported rows/); await assert.rejects(book.finish(), error => error === original);
  }
});

test("literal formulas and thresholds escape XML while preserving formula semantics", async () => {
  const book = new Workbook(), formula = 'AND(A2="_x0041_ Łódź 🧪",B2<5)';
  await book.addSheet("Formula", { columns, conditionalFormats: [
    { type: "expression", range: "A2:B2", formula, style: highlight.style },
    { type: "dataBar", range: "B2", color: "638EC6", minimum: { type: "number", value: 0 }, maximum: { type: "formula", value: "=MAX($B$2:$B$2)" } },
    { type: "cellIs", range: "B2", operator: "between", values: [-1, 1], style: highlight.style }
  ] }).addRows([["_x0041_ Łódź 🧪", 0.5]]);
  const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /AND\(A2=&quot;_x0041_ Łódź 🧪&quot;,B2&lt;5\)/);
  assert.match(xml, /type="formula" val="MAX\(\$B\$2:\$B\$2\)"/); assert.match(xml, /<formula>-1<\/formula><formula>1<\/formula>/);
});
