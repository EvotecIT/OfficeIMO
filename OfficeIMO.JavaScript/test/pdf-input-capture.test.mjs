import test from "node:test";
import assert from "node:assert/strict";
import { writePdf, writePdfTo, writePdfDocument, pdfTable, ExportCell, createGroupedReport, writeGroupedPdf } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const options = { pageNumbers: false, compression: false };
const columns = [{ header: "Value" }];
const writers = [
  (rows, table) => writePdf(rows, { ...options, ...table }),
  (rows, table) => writePdfDocument([pdfTable(rows, table)], options)
];

test("ordinary PDF column presentations preserve captured defaults, overrides and footer geometry", async () => {
  for (const write of writers) for (const mode of ["formatValue", "formatRuns"]) {
    const reads = { alignment: 0, wrapText: 0, ownAlignment: 0 }, calls = [], referenceCalls = [];
    const defaults = [{ alignment: "right", wrapText: true }, { alignment: "center", wrapText: true }];
    const capturedColumns = defaults.map((style, index) => ({ header: index ? "Amount" : "Value", key: index ? "amount" : "value",
      ...(index ? { type: "number" } : {}),
      get alignment() { reads.alignment++; return style.alignment; },
      get wrapText() { reads.wrapText++; return style.wrapText; } }));
    const own = { background: "00ff00", bold: true, get alignment() { reads.ownAlignment++; return "left"; } };
    const values = Array.from({ length: 6 }, (_, index) => [index === 5 ? "R5C0\ncontinued" : "R" + index + "C0", index + 1]);
    values[2][0] = new ExportCell(values[2][0], { presentation: own, link: { target: "https://example.com/value" } });
    values[3][0] = new ExportCell("Not displayed", { text: "Literal" });
    values[4][1] = new ExportCell(values[4][1], { presentation: {} });
    const expectedColumns = [{ header: "Value", key: "value", alignment: "right", wrapText: true },
      { header: "Amount", key: "amount", type: "number", alignment: "center", wrapText: true }];
    const explicit = values.map((row, index) => row.map((value, column) => new ExportCell(value instanceof ExportCell ? value.value : value, {
      ...(value instanceof ExportCell && value.text !== undefined ? { text: value.text } : {}),
      ...(value instanceof ExportCell && value.link ? { link: value.link } : {}),
      presentation: { alignment: expectedColumns[column].alignment, wrapText: true,
        ...(index % 2 ? { background: "eeeeee" } : {}), ...(value instanceof ExportCell ? value.presentation : undefined) }
    })));
    const formatter = log => (value, context) => {
      log.push([context.rowIndex, context.columnIndex, value, context.column.header]);
      const text = "F:" + value;
      return mode === "formatRuns" ? [{ text }] : text;
    };
    const table = { columnWidths: [120, 120], alternateRowColor: "eeeeee", headerPresentation: { background: "ff0000" },
      footer: { values: ["Footer"], totals: { amount: "sum" } }, footerPresentation: { background: "0000ff", alignment: "right" } };
    function* source() {
      defaults[0].alignment = "left"; defaults[1].alignment = "right";
      defaults[0].wrapText = false; defaults[1].wrapText = false;
      yield* values;
    }
    const actual = await write(source(), { ...table, columns: capturedColumns, [mode]: formatter(calls) });
    const reference = await write(explicit, { ...table, columns: expectedColumns, [mode]: formatter(referenceCalls) });
    assert.deepEqual(Buffer.from(await actual.arrayBuffer()), Buffer.from(await reference.arrayBuffer()),
      "Captured column and row defaults emit the same complete PDF as explicit resolved cell presentations");
    assert.deepEqual(reads, { alignment: 2, wrapText: 2, ownAlignment: 1 });
    const expectedCalls = values.flatMap((row, rowIndex) => row.flatMap((value, columnIndex) =>
      value instanceof ExportCell && value.text !== undefined ? [] : [[rowIndex, columnIndex,
        value instanceof ExportCell ? value.value : value, expectedColumns[columnIndex].header]]));
    expectedCalls.push([6, 0, "Footer", "Value"], [6, 1, 21, "Amount"]);
    assert.deepEqual(calls, expectedCalls); assert.deepEqual(referenceCalls, expectedCalls);
    const pdf = await inspectPdf(actual), content = pdf.pages.map(page => page.content).join("\n");
    assert.equal([...content.matchAll(/q 0 0 1 rg /g)].length, 2, "Generated footer cells retain their own background");
    assert.equal([...content.matchAll(/q 1 0 0 rg /g)].length, 2, "Headings retain their own background");
    const left = text => Number(new RegExp("1 0 0 1 ([\\d.-]+) [\\d.-]+ Tm <" + Buffer.from(text).toString("hex") + "> Tj").exec(content)?.[1]);
    assert.equal(left("F:R2C0"), 40, "The authored cell alignment overrides the captured column");
    assert.ok(left("F:R0C0") > 40, "The original right-aligned column survives caller mutation");
    assert.ok([...pdf.objects.values()].some(object => /\/Subtype \/Link\b/.test(object.body)));
  }
});

test("PDF data defaults validate lazily after authored overrides and before formatting", async () => {
  for (const write of writers) {
    for (const [key, invalid, valid, error] of [["alignment", "justify", "left", /alignment supports/],
      ["wrapText", "yes", true, /wrapText must be boolean/]]) {
      const calls = [], table = { columns: [{ header: "Value", [key]: invalid }], includeHeader: false,
        formatValue(value, context) { calls.push(context.rowIndex); return String(value); } };
      await write([], table); assert.deepEqual(calls, [], "Unused column defaults are not eagerly validated");
      const authored = new ExportCell("Authored", { presentation: { [key]: valid } });
      const accepted = await inspectPdf(await write([[authored], [authored]], table));
      assert.equal(accepted.text, "Authored\nAuthored"); assert.deepEqual(calls, [0, 1]); calls.length = 0;
      let produced = 0, returned = 0;
      function* source() { try { for (const value of [authored, "Plain", "Unreached"]) { produced++; yield [value]; } } finally { returned++; } }
      await assert.rejects(write(source(), table), error);
      assert.deepEqual(calls, [0], "The first invalid default rejects before that cell's formatter");
      assert.equal(produced, 2); assert.equal(returned, 1);
    }
    const calls = [];
    await assert.rejects(write([["Plain"], ["Alternate"], [new ExportCell("Invalid own", { presentation: { alignment: "justify" } })]], {
      columns, includeHeader: false, alternateRowColor: "eeeeee", formatValue(value, context) { calls.push(context.rowIndex); return String(value); }
    }), /alignment supports/);
    assert.deepEqual(calls, [0, 1], "Earlier ordinary defaults never suppress a later authored presentation error");
  }
});

test("PDF cancellation from a later ordinary cell formatter closes the source and destination", async () => {
  for (const mode of ["formatValue", "formatRuns"]) {
    const controller = new AbortController(), reason = new Error("ordinary cell cancelled"), calls = [];
    let returned = 0, writes = 0, writesAtCancellation;
    function* source() { try { for (let row = 0; row < 5; row++) yield ["Row " + row]; } finally { returned++; } }
    const destination = new WritableStream({ write() { writes++; } });
    await assert.rejects(writePdfTo(source(), destination, { ...options, columns, includeHeader: false,
      alternateRowColor: "eeeeee", signal: controller.signal, [mode](value, context) {
        calls.push(context.rowIndex);
        if (context.rowIndex === 2) { writesAtCancellation = writes; controller.abort(reason); }
        return mode === "formatRuns" ? [{ text: value }] : value;
      } }), error => error === reason);
    assert.deepEqual(calls, [0, 1, 2]); assert.equal(returned, 1);
    assert.equal(writes, writesAtCancellation, "Cancellation must not flush later document bytes");
    assert.equal(destination.locked, false);
  }
});

test("PDF heading and footer metadata capture scalar getters once and own mutable Date values", async () => {
  for (const write of writers) {
    for (const location of ["heading", "footer"]) {
      let reads = 0;
      class Anchor {
        get value() { return ++reads === 1 ? "Captured metadata" : "Changed metadata"; }
      }
      const matrix = [[new Anchor()]];
      const pdf = await inspectPdf(await write([["Body"]], { columns,
        ...(location === "heading" ? { headerRows: matrix } : { footer: { rows: matrix } }) }));
      assert.equal(reads, 1);
      assert.ok(pdf.text.includes("Captured metadata")); assert.ok(!pdf.text.includes("Changed metadata"));
    }
    for (const wrapped of [false, true]) for (const structured of [false, true]) {
      const date = new Date("2026-01-01T00:00:00Z"), value = wrapped ? new ExportCell(date) : date;
      function* rows() { date.setUTCFullYear(2030); yield ["Body"]; }
      const pdf = await inspectPdf(await write(rows(), { columns,
        footer: structured ? { rows: [[{ value }]] } : { values: [value] } }));
      assert.ok(pdf.text.includes("2026-01-01T00:00:00.000Z")); assert.ok(!pdf.text.includes("2030-01-01"));
    }
    const date = new Date("2026-01-01T00:00:00Z");
    function* rows() { date.setUTCFullYear(2030); for (let i = 0; i < 90; i++) yield ["Body " + i]; }
    const repeated = await inspectPdf(await write(rows(), { columns, headerRows: [[{ value: date }]] }));
    assert.ok(repeated.pages.length > 1);
    assert.ok(repeated.pages.every(page => page.lines.includes("2026-01-01T00:00:00.000Z")));
  }
});

test("class-based PDF options retain prototype limits, metadata, formatting and page geometry", async () => {
  await assert.rejects(writePdf([["Body"]], { ...options, columns,
    limits: Object.create({ maxRows: 0 }) }), { code: "RESOURCE_LIMIT" });
  await assert.rejects(writePdfDocument([pdfTable([["Body"]], { columns })], { ...options,
    limits: Object.create({ maxRows: 0 }) }), { code: "RESOURCE_LIMIT" });
  let rowsRead = 0, limitReads = 0, footerReads = 0, formatReads = 0;
  const limits = Object.create({ get maxRows() { limitReads++; return 0; } });
  const source = { *[Symbol.iterator]() { rowsRead++; yield ["Body"]; } };
  class Configuration {
    get columns() { return columns; }
    get pageNumbers() { return false; }
    get compression() { return false; }
    get limits() { return limits; }
  }
  await assert.rejects(writePdf(source, new Configuration()), { code: "RESOURCE_LIMIT" });
  assert.equal(rowsRead, 1); assert.equal(limitReads, 1);
  rowsRead = 0; limitReads = 0;
  await assert.rejects(writePdfDocument([pdfTable(source, { columns })], new Configuration()), { code: "RESOURCE_LIMIT" });
  assert.equal(rowsRead, 1); assert.equal(limitReads, 1);

  class TableOptions {
    get columns() { return columns; }
    get formatValue() { formatReads++; return value => "Formatted " + value; }
    get footer() { footerReads++; return Object.create({ values: ["Footer"] }); }
  }
  class Geometry { get width() { return 240; } get height() { return 180; } }
  class Margins { get top() { return 10; } get right() { return 10; } get bottom() { return 10; } get left() { return 10; } }
  const pdf = await inspectPdf(await writePdfDocument([
    { kind: "section", title: "Inventory", pageSize: new Geometry(), margins: new Margins() },
    { ...pdfTable([["Body"]], new TableOptions()), caption: Object.create({ text: "Table caption" }) }
  ], options));
  assert.equal(formatReads, 1); assert.equal(footerReads, 1);
  assert.match(pdf.pages[0].body, /\/MediaBox \[0 0 180 240\]/);
  assert.ok(pdf.text.includes("Table caption")); assert.ok(pdf.text.includes("Formatted Body")); assert.ok(pdf.text.includes("Formatted Footer"));

  const grouped = createGroupedReport([["A", 1]], { columns: [{ header: "Group", key: "group" }, { header: "Amount", key: "amount", type: "number" }],
    groupBy: "group", totals: { amount: "sum" } });
  class GroupedOptions {
    get labelColumn() { return { header: "Summary" }; }
    get subtotalLabel() { return () => "Subtotal"; }
    get grandTotalLabel() { return () => "Total"; }
    get pageNumbers() { return false; }
    get limits() { return Object.create({ maxRows: 0 }); }
  }
  await assert.rejects(writeGroupedPdf(grouped, new GroupedOptions()), { code: "RESOURCE_LIMIT" });
});
