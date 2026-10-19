import test from "node:test";
import assert from "node:assert/strict";
import vm from "node:vm";
import { writeCsv, writeCsvTo, writeXlsx, writeXlsxTo, writePdf, writePdfTo, Workbook, Cell, ExportCell } from "../dist/index.js";
import { readZip } from "./zip-reader.mjs";
import { inspectPdf } from "./pdf-reader.mjs";

test("iframe-origin Dates retain their typed value in each export and workbook metadata", async () => {
  const foreign = vm.runInNewContext('new Date("2026-10-19T12:00:00Z")');
  const columns = [{ header: "Date", type: "date" }];
  assert.equal(await (await writeCsv([[foreign]], { columns })).text(), "Date\r\n2026-10-19T12:00:00.000Z\r\n");
  const ordinary = await readZip(await writeXlsx([[new Date(foreign.getTime())]], { columns, dateMode: "utc" }));
  const workbook = new Workbook({ dateMode: "utc", created: foreign, modified: foreign });
  await workbook.addWorksheet("Dates", { columns, autoSize: { sampleRows: 3 } }).addRows([
    [foreign], [new Cell(foreign)], [new ExportCell(foreign)]
  ]);
  const parts = await readZip(await workbook.toBlob());
  const expected = /<v>([^<]+)<\/v>/.exec(ordinary.get("xl/worksheets/sheet1.xml").content)[1];
  const values = [...parts.get("xl/worksheets/sheet1.xml").content.matchAll(/<v>([^<]+)<\/v>/g)].map(match => match[1]);
  assert.deepEqual(values, [expected, expected, expected]);
  assert.match(parts.get("docProps/core.xml").content, /2026-10-19T12:00:00.000Z/);
  const pdf = await inspectPdf(await writePdf([[foreign]], { columns }));
  assert.match(pdf.text, /2026-10-19T12:00:00.000Z/);
});

test("Date-like objects cannot impersonate native Dates and Dates cannot act as domain rows", async () => {
  const columns = [{ header: "Date" }];
  const forged = { [Symbol.toStringTag]: "Date", getTime: () => 0 };
  for (const write of [writeCsv, writeXlsx, writePdf]) {
    await assert.rejects(write([[forged]], { columns }), /synchronous/);
    await assert.rejects(write([vm.runInNewContext("new Date()")], { columns }), /row must be an array or object/i);
  }
});

test("Blob and destination writers reject async progress callbacks without leaking their rejection", async () => {
  const columns = [{ header: "Value" }];
  for (const write of [writeCsv, writeXlsx, writePdf]) {
    await assert.rejects(write([[1]], { columns, onProgress: async () => { throw new Error("notification failed"); } }), /onProgress must be synchronous/);
  }
  for (const write of [writeCsvTo, writeXlsxTo, writePdfTo]) {
    await assert.rejects(write([[1]], { write() {} }, { columns, onProgress: async () => { throw new Error("notification failed"); } }), /onProgress must be synchronous/);
  }
  // Node's test runner also fails this test if any rejected callback escapes as unhandledRejection.
  await new Promise(resolve => setImmediate(resolve));
});

test("CSV row progress includes the complete row just accepted by its sink", async () => {
  const events = [], accepted = [];
  await writeCsvTo(Array.from({ length: 2000 }, () => ["x".repeat(33)]), { write(bytes) { accepted.push(bytes); } }, {
    columns: [{ header: "Value" }], includeHeader: false, onProgress: event => {
      events.push(event);
      const completeRows = (new TextDecoder().decode(Buffer.concat(accepted)).match(/\r\n/g) ?? []).length;
      assert.equal(event.rows, completeRows);
    }
  });
  assert.ok(events.some(event => event.phase === "rows"));
  assert.equal(events.at(-1).rows, 2000);
  assert.ok(new TextDecoder().decode(Buffer.concat(accepted)).endsWith("\r\n"));
});
