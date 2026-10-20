import test from "node:test";
import assert from "node:assert/strict";
import { Workbook, Cell, FormulaCell, writeXlsx } from "../dist/xlsx/index.js";
import { writeCsv } from "../dist/csv/index.js";
import { writePdf } from "../dist/pdf/index.js";
import { readZip } from "./zip-reader.mjs";

test("explicit formula caches retain numeric styles, callbacks, totals and one authored budget charge", async () => {
  for (const streamed of [false, true]) {
    const chunks = [], rows = [], cells = [], formula = 'IF(A2="=SUM(1,2)",3,0)';
    const book = new Workbook({ compression: "store", limits: { maxFormulaCharacters: formula.length + 3 },
      ...(streamed ? { sink: { write(bytes) { chunks.push(bytes.slice()); } } } : {}) });
    const style = book.styles.add({ numberFormat: "0.000", font: { bold: true } });
    const sheet = book.addWorksheet("Calculated", { columns: [{ header: "Literal", key: "literal" }, { header: "Amount", key: "amount", type: "number", format: "0.0" }],
      autoSize: { sampleRows: 2 }, footer: { totals: { amount: "sum" } },
      rowStyle(context) { rows.push(context.values[1]); }, cellStyle(context) { if (context.columnIndex === 1) cells.push(context.value); } });
    await sheet.addRows([["=SUM(1,2)", new FormulaCell("=" + formula, 3, style)]]);
    await sheet.addRows([["ordinary", new FormulaCell("2+3", 5)]]);
    let blob;
    if (streamed) { await book.finish(); blob = new Blob(chunks); } else blob = await book.toBlob();
    const zip = await readZip(blob), xml = zip.get("xl/worksheets/sheet1.xml").content;
    assert.deepEqual(rows, [3, 5]); assert.deepEqual(cells, [3, 5]);
    assert.match(xml, /r="A2"[^>]*t="inlineStr"><is><t[^>]*>=SUM\(1,2\)<\/t>/);
    assert.ok(xml.includes('<c r="B2" s="' + style + '"><f>IF(A2=&quot;=SUM(1,2)&quot;,3,0)</f><v>3</v></c>'));
    assert.match(xml, /r="B3"[^>]*><f>2\+3<\/f><v>5<\/v>/);
    assert.match(xml, /r="B4"[^>]*><f>SUBTOTAL\(109,B2:B3\)<\/f><v>8<\/v>/);
    assert.match(zip.get("xl/workbook.xml").content, /<calcPr[^>]*calcMode="auto"/);
    assert.match(zip.get("xl/workbook.xml").content, /<calcPr[^>]*fullCalcOnLoad="1"/);
  }
});

test("column getters, custom writers and explicit footers preserve authored formula cells", async () => {
  let writes = 0, gets = 0;
  const book = new Workbook({ cellValueWriters: { calculated(value) { writes++; return new FormulaCell(value ? "1+1" : "0", value ? 2 : 0); } } });
  const sheet = book.addWorksheet("Sources", { columns: [
    { header: "Converted", key: "input", type: "calculated" },
    { header: "Projected", type: "number", value() { gets++; return new FormulaCell("4/2", 2); } }
  ], autoSize: { sampleRows: 2 }, footer: { values: [new FormulaCell("SUM(A2:A2)", 2), new FormulaCell("SUM(B2:B2)", 2)] } });
  await sheet.addRows([{ input: true }]);
  const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
  assert.equal(writes, 1); assert.equal(gets, 1);
  assert.match(xml, /r="A2"[^>]*><f>1\+1<\/f><v>2<\/v>/);
  assert.match(xml, /r="B2"[^>]*><f>4\/2<\/f><v>2<\/v>/);
  assert.match(xml, /r="A3"[^>]*><f>SUM\(A2:A2\)<\/f><v>2<\/v>/);
  assert.match(xml, /r="B3"[^>]*><f>SUM\(B2:B2\)<\/f><v>2<\/v>/);
});

test("formula cells are immutable, bounded XML-valid numeric caches with direct local references", async () => {
  const cell = new FormulaCell(" =SUM(1,2) ", 3, 0);
  assert.equal(cell.formula, "SUM(1,2)"); assert.equal(cell.cachedValue, 3); assert.ok(Object.isFrozen(cell));
  assert.throws(() => { cell.cachedValue = 99; }, TypeError);
  for (const formula of ["", "=", "x".repeat(8193), "SUM(1,\u0001)", 'IF(A1="\ud800",1,0)']) assert.throws(() => new FormulaCell(formula, 0));
  for (const value of [undefined, null, "3", true, NaN, Infinity, -Infinity]) assert.throws(() => new FormulaCell("1+2", value));
  for (const style of [-1, NaN, 1.5, Number.MAX_SAFE_INTEGER + 1]) assert.throws(() => new FormulaCell("0", 0, style));
  for (const formula of ["[Book.xlsx]Sheet1!A1", "'[Book.xlsx]Sheet 1'!A1", "SUM([1]Sheet!A1)", "cmd|' /C calc'!A0"])
    assert.throws(() => new FormulaCell(formula, 0), { code: "NOT_SUPPORTED" });
  for (const formula of ['COUNTIF(Services[Status],"Healthy")', "SUM(Services[[#This Row],[Score]])", "SUM([@[Error!Name]])", "'O''Brien | QA'!A1", 'IF(A1="[Book]Sheet!A1 | DDE",1,0)'])
    assert.equal(new FormulaCell(formula, 0).formula, formula);
  const forged = Object.create(FormulaCell.prototype);
  const book = new Workbook(), sheet = book.addWorksheet("Forged", { columns: [{ header: "Value" }] });
  await assert.rejects(sheet.addRows([[forged]]), /synchronous|strings/);
  const escaped = new Workbook();
  await escaped.addWorksheet("Escaped", { columns: [{ header: "Value" }] }).addRows([[new FormulaCell('IF(A2="</f><evil>&",1,0)', 1)]]);
  const xml = (await readZip(await escaped.toBlob())).get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /<f>IF\(A2=&quot;&lt;\/f&gt;&lt;evil&gt;&amp;&quot;,1,0\)<\/f><v>1<\/v>/);
  assert.doesNotMatch(xml, /<evil>/);
});

test("formulas remain advanced XLSX values and never enter portable or nested scalar contracts", async () => {
  const formula = new FormulaCell("1+1", 2), columns = [{ header: "Value" }];
  for (const write of [writeXlsx, writeCsv, writePdf]) await assert.rejects(write([[formula]], { columns }), /synchronous|strings/);
  const wrapped = new Workbook();
  await assert.rejects(wrapped.addWorksheet("Nested", { columns }).addRows([[new Cell(formula)]]), /synchronous|strings/);
  const structured = new Workbook();
  assert.throws(() => structured.addWorksheet("Header", { columns, headerRows: [[{ value: formula }]] }), /synchronous|strings/);
  const styled = new Workbook();
  await assert.rejects(styled.addWorksheet("Style", { columns }).addRows([[new FormulaCell("0", 0, 999)]]), /style index/);
});

test("inline dropdowns snapshot metadata and resolve final appended data rows around report regions", async () => {
  for (const streamed of [false, true]) {
    const chunks = [], rules = [{ type: "list", column: "status", items: ["Healthy", 'Needs "review"', "Offline"], allowBlank: false,
      errorStyle: "warning", errorTitle: 'Pick "_x0041_"', errorMessage: "Use a listed status & retry." }];
    const book = new Workbook({ compression: "store", ...(streamed ? { sink: { write(bytes) { chunks.push(bytes.slice()); } } } : {}) });
    const sheet = book.addWorksheet("Services", { columns: [
      { header: "Name", key: "name", groups: ["Report"] }, { header: "Status", key: "status", groups: ["Report"] },
      { header: "Score", key: "score", type: "number", groups: ["Report"] }
    ], title: { text: "Service report" }, table: {}, footer: { values: ["End"], totals: { score: "sum" } }, dataValidation: rules,
      conditionalFormats: [{ type: "cellIs", range: { column: "score" }, operator: "lessThan", value: 0, style: { font: { bold: true } } }],
      hyperlinks: [{ cell: "A4", target: "https://example.com/report" }] });
    rules[0].items[0] = "Mutated"; rules[0].column = "name"; rules[0].errorTitle = "Mutated";
    await sheet.addRows([["one", "Healthy", 1]]); await sheet.addRows([["two", "Offline", 2]]);
    let blob; if (streamed) { await book.finish(); blob = new Blob(chunks); } else blob = await book.toBlob();
    const xml = (await readZip(blob)).get("xl/worksheets/sheet1.xml").content;
    assert.match(xml, /<dataValidations count="1">/); assert.match(xml, /<dataValidation[^>]*sqref="B4:B5"/);
    assert.match(xml, /allowBlank="0"/); assert.match(xml, /showErrorMessage="1"/); assert.match(xml, /errorStyle="warning"/);
    assert.match(xml, /errorTitle="Pick &quot;_x005F_x0041_&quot;"/); assert.match(xml, /error="Use a listed status &amp; retry\."/);
    assert.match(xml, /<formula1>&quot;Healthy,Needs &quot;&quot;review&quot;&quot;,Offline&quot;<\/formula1>/);
    assert.doesNotMatch(xml, /Mutated/);
    assert.ok(xml.indexOf("</conditionalFormatting>") < xml.indexOf("<dataValidations"));
    assert.ok(xml.indexOf("</dataValidations>") < xml.indexOf("<hyperlinks>"));
  }
});

test("dropdown choices and column targets reject unsupported or ambiguous contracts without reserving a sheet", async () => {
  const valid = { type: "list", column: "status", items: ["Healthy", "Offline"] }, columns = [{ header: "Status", key: "status" }];
  const invalid = [
    { type: "custom" }, { column: "missing" }, { column: 0 }, { column: 2 }, { items: [] }, { items: [1] },
    { items: ["a,b"] }, { items: ["\u0001"] }, { items: ["x".repeat(254)] }, { items: ["x".repeat(251), '"'] },
    { errorTitle: "x".repeat(33) }, { errorMessage: "x".repeat(256) }, { errorMessage: "\ud800" },
    { allowBlank: "true" }, { showErrorMessage: 1 }, { errorStyle: "other" }, { formula: "A1:A2" }, { range: "A2:A9" }
  ];
  for (const patch of invalid) {
    const book = new Workbook({ invalidCharacterPolicy: "strip", limits: { maxDataValidations: 1 } });
    assert.throws(() => book.addWorksheet("Recover", { columns, dataValidation: [{ ...valid, ...patch }] }));
    assert.equal(book.addWorksheet("Recover", { columns, dataValidation: [valid] }).name, "Recover");
    await book.finish();
  }
  const book = new Workbook();
  assert.throws(() => book.addWorksheet("Duplicate", { columns, dataValidation: [valid, { ...valid, column: 1 }] }), /multiple validation/);
  assert.throws(() => book.addWorksheet("Ambiguous", { columns: [columns[0], columns[0]], dataValidation: [valid] }), /exactly one/);
});

test("dropdown formula length includes quotes and separators; empty sources do not validate report metadata", async () => {
  const book = new Workbook();
  const sheet = book.addWorksheet("Bounds", { columns: [{ header: "Status" }], dataValidation: [
    { type: "list", column: 1, items: ["x".repeat(253)], allowBlank: true, showErrorMessage: false, errorStyle: "information" }
  ] });
  await sheet.addRows([["x"]]);
  const empty = book.addWorksheet("Empty", { columns: [{ header: "Status", key: "status" }], title: { text: "Empty export" },
    footer: { values: ["No records"] }, dataValidation: [{ type: "list", column: 1, items: ["Healthy"] }] });
  await empty.addRows([]);
  const zip = await readZip(await book.toBlob()), xml = zip.get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /sqref="A2:A2"/); assert.match(xml, /showErrorMessage="0"/); assert.match(xml, /errorStyle="information"/);
  assert.doesNotMatch(zip.get("xl/worksheets/sheet2.xml").content, /<dataValidations/);
});

test("dropdown option and item getters are captured once before later caller mutation", async () => {
  const reads = new Map(), items = ["Healthy", "Offline"];
  const observedItems = new Proxy(items, { get(target, property, receiver) {
    if (["length", "0", "1"].includes(property)) reads.set("items." + property, (reads.get("items." + property) ?? 0) + 1);
    return Reflect.get(target, property, receiver);
  } });
  const input = { type: "list", column: "status", items: observedItems, errorTitle: "Select status" };
  const rule = new Proxy(input, { get(target, property, receiver) {
    if (property !== "then") reads.set(property, (reads.get(property) ?? 0) + 1);
    return Reflect.get(target, property, receiver);
  } });
  const book = new Workbook(), sheet = book.addWorksheet("Snapshot", { columns: [{ header: "Status", key: "status" }], dataValidation: [rule] });
  input.errorTitle = "Mutated"; items[0] = "Mutated";
  await sheet.addRows([["Healthy"]]);
  const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /Healthy,Offline/); assert.doesNotMatch(xml, /Mutated/);
  for (const [property, count] of reads) assert.equal(count, 1, String(property));
  assert.equal(reads.get("items.0"), 1); assert.equal(reads.get("items.length"), 1);
});

test("validation counts and text are cumulative; oversize preflight and cancellation do not read source rows", async () => {
  const rule = { type: "list", column: 1, items: ["a"] }, columns = [{ header: "Status" }];
  const counted = new Workbook({ limits: { maxDataValidations: 1 } });
  counted.addWorksheet("First", { columns, dataValidation: [rule] });
  assert.throws(() => counted.addWorksheet("Recover", { columns, dataValidation: [rule] }), { code: "RESOURCE_LIMIT" });
  assert.equal(counted.addWorksheet("Recover", { columns }).name, "Recover"); await counted.finish();
  const text = new Workbook({ limits: { maxValidationTextCharacters: 5 } });
  text.addWorksheet("First", { columns, dataValidation: [rule] });
  assert.throws(() => text.addWorksheet("Second", { columns, dataValidation: [rule] }), { code: "RESOURCE_LIMIT" }); await text.finish();
  let itemReads = 0;
  const items = ["a"]; Object.defineProperty(items, 0, { get() { itemReads++; return "a"; } });
  const bounded = new Workbook({ limits: { maxDataValidations: 0 } });
  assert.throws(() => bounded.addWorksheet("Over", { columns, dataValidation: [{ ...rule, items }] }), { code: "RESOURCE_LIMIT" });
  assert.equal(itemReads, 0);
  const controller = new AbortController(), aborted = new Workbook({ signal: controller.signal });
  const cancelling = ["a"]; Object.defineProperty(cancelling, 0, { get() { controller.abort(); return "a"; } });
  assert.throws(() => aborted.addWorksheet("Cancelled", { columns, dataValidation: [{ ...rule, items: cancelling }] }), { name: "AbortError" });
});

test("formula text limits remain cumulative across worksheets and failure does not complete an invalid workbook", async () => {
  const book = new Workbook({ limits: { maxFormulaCharacters: 3 } });
  await book.addWorksheet("First", { columns: [{ header: "Value" }] }).addRows([[new FormulaCell("1+1", 2)]]);
  const last = book.addWorksheet("Second", { columns: [{ header: "Value" }] });
  await assert.rejects(last.addRows([[new FormulaCell("0", 0)]]), { code: "RESOURCE_LIMIT" });
  await assert.rejects(book.finish(), { code: "RESOURCE_LIMIT" });
});
