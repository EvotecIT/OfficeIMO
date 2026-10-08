import test from "node:test";
import assert from "node:assert/strict";
import { writeFile } from "node:fs/promises";
import { ExportCell } from "../dist/core/index.js";
import { writeCsv } from "../dist/csv/index.js";
import { writeXlsx, Workbook } from "../dist/xlsx/index.js";
import { writePdf } from "../dist/pdf/index.js";
import { readZip } from "./zip-reader.mjs";
import { inspectPdf } from "./pdf-reader.mjs";

const link = { target: "https://example.com/report?q=1&other=2#details", tooltip: "Report _x0041_ <value> 🧪" };
async function fixture(name, blob) {
  if (process.env.OFFICEIMO_CANOPY_FIXTURES) await writeFile(process.env.OFFICEIMO_CANOPY_FIXTURES + "/" + name, new Uint8Array(await blob.arrayBuffer()));
}
function links(pdf, page) {
  const references = /\/Annots \[([^\]]*)\]/.exec(page.body)?.[1] ?? "";
  return [...references.matchAll(/(\d+) 0 R/g)].map(match => pdf.objects.get(Number(match[1])).body);
}
function uri(body) { return Buffer.from(/\/URI <([0-9a-f]+)>/.exec(body)[1], "hex").toString("utf8"); }

test("portable links capture safe destinations without changing scalar or CSV value semantics", async () => {
  const requested = { ...link }, cell = new ExportCell(12.5, { text: "12.50 USD", link: requested });
  requested.target = "https://different.example";
  assert.equal(cell.link.target, link.target);
  assert.ok(Object.isFrozen(cell.link));
  assert.equal(await (await writeCsv([[cell]], { columns: [{ header: "Value" }], includeHeader: false })).text(), "12.5\r\n");
  assert.equal(await (await writeCsv([[cell]], { columns: [{ header: "Value" }], includeHeader: false, valueMode: "display" })).text(), "12.50 USD\r\n");
  for (const target of ["javascript:alert(1)", "data:text/html,evil", "file:///secret", "../report.html", "https://user:pass@example.com", "https://example.com/a b", "https://example.com/\n"]) {
    assert.throws(() => new ExportCell("value", { link: { target } }), TypeError);
  }
  assert.throws(() => new ExportCell("value", { link: { target: link.target, tooltip: 1 } }), TypeError);
});

test("streamed XLSX portable links preserve types, sampled dates, titles, styles and footer coordinates", async () => {
  const rows = [[new ExportCell(12.5, { text: "12.50 USD", link, presentation: { background: "E2F0D9" } }),
    new ExportCell(new Date("2026-10-08T00:00:00Z"), { link: { target: "mailto:report@example.com" } })]];
  const blob = await writeXlsx(rows, { dateMode: "utc", compression: "store", columns: [
    { header: "Amount", type: "number", format: "0.00", groups: ["Inventory"] }, { header: "Seen", type: "date", groups: ["Inventory"] }
  ], sheet: { title: { text: "Report" }, footer: { values: [new ExportCell("Footer", { link }), null] } } });
  const zip = await readZip(blob), sheet = zip.get("xl/worksheets/sheet1.xml").content;
  assert.match(sheet, /<c r="A4"[^>]*><v>12.5<\/v>/);
  assert.match(sheet, /<c r="B4"[^>]*><v>\d+<\/v>/);
  assert.match(sheet, /<hyperlink ref="A4" r:id="link1" tooltip="Report _x005F_x0041_ &lt;value&gt; 🧪"/);
  assert.match(sheet, /<hyperlink ref="B4" r:id="link2"/);
  assert.match(sheet, /<hyperlink ref="A5" r:id="link3"/);
  const relationships = zip.get("xl/worksheets/_rels/sheet1.xml.rels").content;
  assert.match(relationships, /Target="https:\/\/example.com\/report\?q=1&amp;other=2#details"/);
  assert.match(relationships, /Target="mailto:report@example.com"/);
  await fixture("portable-links.xlsx", blob);
});

test("XLSX portable links respect duplicate, merge, overflow and workbook link limits", async () => {
  const linked = new ExportCell("value", { link });
  await assert.rejects(writeXlsx([[linked]], { columns: [{ header: "Value" }], limits: { maxHyperlinks: 0 } }), /maxHyperlinks/);
  await assert.rejects(writeXlsx([[linked]], { columns: [{ header: "Value" }], sheet: { hyperlinks: [{ cell: "A2", target: link.target }] } }), /Duplicate hyperlink/);
  await assert.rejects(writeXlsx([[null, linked]], { columns: [{ header: "A" }, { header: "B" }], sheet: { mergedCells: ["A2:B2"] } }), /covered/);
  await assert.rejects(writeXlsx([[new ExportCell("x".repeat(32768), { link })]], { oversizedText: "preserve", columns: [{ header: "Value" }] }), /text-preservation/);
  const book = new Workbook({ limits: { maxHyperlinks: 1 } });
  const sheet = book.addWorksheet("Data", { columns: [{ header: "Value" }], autoSize: { sampleRows: 0 } });
  await sheet.addRows([[linked]]);
  assert.throws(() => sheet.addHyperlink({ cell: "A2", target: link.target }), /Duplicate/);
  assert.throws(() => sheet.addHyperlink({ cell: "A3", target: link.target }), /maxHyperlinks/);
  await book.toBlob();
});

test("XLSX custom value writers retain or deliberately replace a captured portable link", async () => {
  const blob = await writeXlsx([[new ExportCell(1, { link }), 2]], {
    columns: [{ header: "Kept", type: "kept" }, { header: "Replaced", type: "linked" }],
    cellValueWriters: { kept: value => value + 10, linked: value => new ExportCell(value + 20, { link: { target: "mailto:report@example.com" } }) }
  });
  const zip = await readZip(blob), sheet = zip.get("xl/worksheets/sheet1.xml").content;
  assert.match(sheet, /<c r="A2"[^>]*><v>11<\/v>/);
  assert.match(sheet, /<c r="B2"[^>]*><v>22<\/v>/);
  assert.match(sheet, /<hyperlink ref="A2" r:id="link1"/);
  assert.match(sheet, /<hyperlink ref="B2" r:id="link2"/);
});

test("PDF annotations cover each continued cell fragment and preserve URI and Unicode tooltip", async () => {
  const text = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".repeat(300);
  const blob = await writePdf([[new ExportCell(text, { link })]], { columns: [{ header: "Value" }], pageSize: "A5", columnWidths: [160], compression: false });
  const pdf = await inspectPdf(blob);
  assert.ok(pdf.pages.length > 3);
  for (const page of pdf.pages) {
    const annotations = links(pdf, page);
    assert.equal(annotations.length, 1);
    assert.equal(uri(annotations[0]), link.target);
    assert.match(annotations[0], /\/Contents <feff/);
    const coordinates = /\/Rect \[([^\]]+)\]/.exec(annotations[0])[1].split(" ").map(Number);
    assert.ok(coordinates[0] < coordinates[2] && coordinates[1] < coordinates[3]);
    assert.ok(coordinates[1] >= 36 && coordinates[3] <= 559.276);
  }
  const actual = pdf.pages.map(page => page.lines.slice(2).join("")).join("");
  assert.equal(actual, text);
  await fixture("portable-links.pdf", blob);
});

test("PDF structured headings and footers retain links, including repeated spans", async () => {
  const blob = await writePdf(Array.from({ length: 60 }, (_, i) => ["Row " + i, i]), {
    columns: [{ header: "Name" }, { header: "Number" }], pageSize: "A5", compression: false,
    headerRows: [[{ value: new ExportCell("Report", { link }), columnSpan: 2 }, null]],
    footer: { rows: [[{ value: new ExportCell("End", { link: { target: "mailto:report@example.com" } }), columnSpan: 2 }, null]] }
  });
  const pdf = await inspectPdf(blob);
  assert.ok(pdf.pages.length > 1);
  for (const page of pdf.pages) assert.equal(uri(links(pdf, page)[0]), link.target);
  assert.equal(uri(links(pdf, pdf.pages.at(-1)).at(-1)), "mailto:report@example.com");
  await fixture("portable-links-spans.pdf", blob);
});

test("PDF link budgets include page continuations and retained annotation bytes", async () => {
  const rows = [[new ExportCell("ABCDEFGHIJKLMNOPQRSTUVWXYZ".repeat(300), { link })]], options = { columns: [{ header: "Value" }], pageSize: "A5", columnWidths: [160] };
  await assert.rejects(writePdf(rows, { ...options, limits: { maxHyperlinks: 1 } }), /maxHyperlinks/);
  await assert.rejects(writePdf([[new ExportCell("Value", { link: { target: "https://example.com/" + "x".repeat(10000) } })]],
    { columns: [{ header: "Value" }], limits: { maxPageBytes: 5000 } }), /maxPageBytes/);
});

test("PDF link tooltips reject malformed Unicode in body cells and structured decorations", async () => {
  for (const tooltip of ["bad\ud800", "bad\udc00", "\ud800middle\udc00"]) {
    const value = new ExportCell("Value", { link: { target: link.target, tooltip } }), options = { columns: [{ header: "Value" }] };
    await assert.rejects(writePdf([[value]], options), /unpaired UTF-16/);
    await assert.rejects(writePdf([["Data"]], { ...options, headerRows: [[{ value }]] }), /unpaired UTF-16/);
    await assert.rejects(writePdf([["Data"]], { ...options, footer: { values: [value] } }), /unpaired UTF-16/);
  }
  const valid = await writePdf([[new ExportCell("Value", { link: { target: link.target, tooltip: "A 🧪 B" } })]], { columns: [{ header: "Value" }] });
  assert.match([...((await inspectPdf(valid)).objects.values())].map(o => o.body).join("\n"), /Contents <feff00410020d83eddea00200042>/);
});
