import test from "node:test";
import assert from "node:assert/strict";
import { writePdf, writePdfDocument, pdfTable, ExportCell, createGroupedReport, writeGroupedPdf } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

const options = { pageNumbers: false, compression: false };
const columns = [{ header: "Value" }];
const writers = [
  (rows, table) => writePdf(rows, { ...options, ...table }),
  (rows, table) => writePdfDocument([pdfTable(rows, table)], options)
];

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
