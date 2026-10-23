import test from "node:test";
import assert from "node:assert/strict";
import { writePdf } from "../dist/pdf/index.js";
import { writeXlsx } from "../dist/xlsx/index.js";
import { writeCsv } from "../dist/csv/index.js";
import { inspectPdf } from "./pdf-reader.mjs";
import { readZip } from "./zip-reader.mjs";

test("numeric positions and numeric-looking stable keys remain distinct in mixed band selections", async () => {
  const columns = [{ header: "Metric", key: "metric" }, { header: "Identifier", key: "0", rowHeader: true },
    { header: "First", key: "first" }, { header: "Last", key: "last" }];
  const options = { columns, columnWidths: [80, 80, 80, 80], pageSize: { width: 300, height: 210 },
    orientation: "landscape", margins: 20, pageNumbers: false, compression: false };
  const pdf = await inspectPdf(await writePdf([[7, "SVC-1", "A", "B"]], { ...options, wideTable: { mode: "bands", repeatColumns: ["0", 0] } }));
  assert.equal(pdf.pages.length, 2);
  for (const page of pdf.pages) {
    assert.deepEqual(page.lines.slice(0, 2), ["Metric", "Identifier"]);
    assert.ok(page.lines.includes("7") && page.lines.includes("SVC-1"));
  }
});

test("repeat selectors reject missing, ambiguous and duplicate selected identities before opening rows", async () => {
  let opened = 0;
  const columns = [{ header: "Not a key", key: "amount" }, { header: "Name", key: "service" },
    { header: "First", key: "shared" }, { header: "Last", key: "shared" }];
  function* rows() { opened++; yield [1, "SVC-1", "A", "B"]; }
  for (const [repeatColumns, reason] of [[ ["hidden"], /selected column key/ ], [ ["Not a key"], /selected column key/ ],
    [ ["shared"], /unambiguous/ ], [ ["service", "service"], /same column/ ], [ [1, "service"], /same column/ ],
    [ [1, 0], /sorted unique/ ], [ [4], /exceeds/ ]]) {
    await assert.rejects(writePdf(rows(), { columns, wideTable: { mode: "bands", repeatColumns } }), reason);
  }
  for (const writer of [writePdf, writeXlsx, writeCsv])
    await assert.rejects(writer(rows(), { columns: [{ header: "Name", rowHeader: "yes" }] }), /rowHeader must be boolean/);
  assert.equal(opened, 0);
});

test("row-header metadata preserves CSV values and XLSX cell content and presentation", async () => {
  const rows = [["SVC-1", 12.5]], plain = [{ header: "Name", key: "service" }, { header: "Amount", key: "amount", type: "number", format: "0.00" }];
  const marked = plain.map((column, index) => ({ ...column, rowHeader: index === 0 }));
  let observed;
  const csv = await writeCsv(rows, { columns: marked.map(column => ({ ...column, valueFormatter(value, context) {
    if (context.column.key === "service") observed = context.column.rowHeader;
    return value;
  } })) });
  assert.equal(observed, true);
  assert.equal(await csv.text(), await (await writeCsv(rows, { columns: plain })).text());
  const original = await readZip(await writeXlsx(rows, { columns: plain })), exported = await readZip(await writeXlsx(rows, { columns: marked }));
  for (const part of ["xl/worksheets/sheet1.xml", "xl/styles.xml"]) assert.equal(exported.get(part).content, original.get(part).content);
});
