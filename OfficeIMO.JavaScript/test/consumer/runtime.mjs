import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { MessageChannel } from "node:worker_threads";
import { Workbook, writeXlsx, writeXlsxTo, writePdf, preflightPdfText, core, zip, xml, opc, xlsx, csv, pdf } from "@evotecit/officeimo";
import * as coreModule from "@evotecit/officeimo/core";
import * as zipModule from "@evotecit/officeimo/zip";
import * as xmlModule from "@evotecit/officeimo/xml";
import * as opcModule from "@evotecit/officeimo/opc";
import * as xlsxModule from "@evotecit/officeimo/xlsx";
import * as csvModule from "@evotecit/officeimo/csv";
import * as pdfModule from "@evotecit/officeimo/pdf";
import * as dataTablesModule from "@evotecit/officeimo/integrations/datatables";
import * as canopyModule from "@evotecit/officeimo/integrations/canopyx";
import { executeWorkerExport, handleWorkerExports } from "@evotecit/officeimo/execution";
import { createHarfBuzzTextShaper } from "@evotecit/officeimo/pdf/providers/harfbuzz";

for (const [namespace, module] of [[core, coreModule], [zip, zipModule], [xml, xmlModule], [opc, opcModule], [xlsx, xlsxModule], [csv, csvModule], [pdf, pdfModule]])
  for (const [key, value] of Object.entries(module)) assert.equal(namespace[key], value);
assert.equal(globalThis.document, undefined);
assert.equal(typeof createHarfBuzzTextShaper, "function", "The optional provider adapter imports without loading WASM or external packages");
assert.equal(typeof dataTablesModule.createDataTablesExport, 'function');
assert.equal(dataTablesModule.ExportCell, coreModule.ExportCell);
assert.equal(canopyModule.ExportCell, coreModule.ExportCell);
const canopyCapture = { request: { columns: [{ id: "Name", title: "Name", kind: "text" }], recordCount: 1, presentation: "text", values: "raw", timeZone: "utc", revision: "packed" },
  async *rows() { yield { id: "r", cells: { Name: { value: "Packed", text: "Packed" } } }; } };
assert.equal(await (await canopyModule.exportCanopy(canopyCapture, "csv", { unsupportedPresentation: "text" })).text(), "Name\r\nPacked\r\n");
assert.equal(writeXlsx, xlsx.writeXlsx); assert.equal(writeXlsxTo, xlsx.writeXlsxTo);
assert.equal(writePdf, pdf.writePdf);
assert.equal(preflightPdfText, pdf.preflightPdfText);
assert.deepEqual(preflightPdfText("Packed\r\n\tPDF"), { supported: true });
const textIssue = pdfModule.preflightPdfText("Packed Ł");
assert.equal(textIssue.supported, false); assert.equal(textIssue.code, "embedded-font-required"); assert.equal(textIssue.index, 7);
const packedPdf = await writePdf([{name:'Packed PDF'}],{columns:[{header:'Name',key:'name'}]});
assert.equal(packedPdf.type,'application/pdf');
await writeFile('packed-consumer.pdf',new Uint8Array(await packedPdf.arrayBuffer()));
const composed = await pdf.writePdfDocument([{ kind: "paragraph", text: "Packed document" },
  pdf.pdfTable([{ name: "Packed section" }], { columns: [{ header: "Name", key: "name" }] }),
  { kind: "image", image: { format: "rgb", width: 1, height: 1, data: Uint8Array.of(0, 128, 255) }, width: 20, height: 20 }]);
assert.equal(composed.type, "application/pdf");
await writeFile("packed-document.pdf", new Uint8Array(await composed.arrayBuffer()));
// Independently produced Pillow RGB PNG; no DOM or canvas participates in decoding.
const png = Uint8Array.from(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAwAAAAICAIAAABChommAAAAJklEQVR42mP8ICLCgApULqEJMDAxEAGop4hFheE6mtD1b5oD6yYA4aYFETss/rAAAAAASUVORK5CYII=", "base64"));
const richBlocks = [
  { kind: "paragraph", text: "Packed tagged report", headingLevel: 1 },
  { kind: "image", image: { format: "png", data: png }, width: 120, alternativeText: "Four colored quadrants" },
  { kind: "drawing", drawing: { version: 1, width: 120, height: 40, elements: [
    { kind: "path", fill: "008000", commands: [{ op: "move", x: 0, y: 0 }, { op: "line", x: 120, y: 40 }, { op: "line", x: 0, y: 40 }, { op: "close" }] }
  ] }, alternativeText: "Green triangle" },
  pdf.pdfTable([{ value: "Packed cell" }], { columns: [{ header: "Value", key: "value" }] })
];
const richOptions = { tagged: { language: "en-US" }, compression: false, pageNumbers: false };
const rich = await pdf.writePdfDocument(richBlocks, richOptions);
assert.equal(rich.type, "application/pdf");
assert.match(await rich.text(), /\/StructTreeRoot /);
await writeFile("packed-rich-report.pdf", new Uint8Array(await rich.arrayBuffer()));
await import("@evotecit/officeimo/pdf.js");
const classicRich = await globalThis.OfficeIMO.pdf.writePdfDocument(richBlocks, richOptions);
assert.deepEqual(Buffer.from(await classicRich.arrayBuffer()), Buffer.from(await rich.arrayBuffer()), "Packed standalone PDF script matches the ESM report");
const domainRows = [{ person: { name: "Łódź 🧪" }, ignored: { nested: true } }], columns = [{ header: "Name", value: row => row.person.name }];
const packed = await writeXlsx(domainRows, { columns, compression: "store", sheet: { name: "Packed" } });
await writeFile("packed-consumer.xlsx", new Uint8Array(await packed.arrayBuffer()));
assert.equal(await (await csv.writeCsv(domainRows, { columns })).text(), "Name\r\nŁódź 🧪\r\n");
const chunks = [], destination = new WritableStream({ write: chunk => { chunks.push(chunk); } });
const tableResult = await writeXlsxTo(domainRows, destination, { columns });
assert.deepEqual(tableResult, { rows: 1, columns: 1, bytes: new Blob(chunks).size }); assert.equal(destination.locked, false);
assert.equal(await (await csv.writeCsv([["=cmd"]], { columns: [{ header: "Name" }] })).text(), "Name\r\n'=cmd\r\n");
const sink = new core.BlobByteSink(), xmlWriter = new xml.XmlWriter(sink);
await xmlWriter.startElement("root"); await xmlWriter.text("Łódź"); await xmlWriter.endElement(); await xmlWriter.close();
const archive = new zip.ZipWriter(); await archive.add("test.xml", new Uint8Array(await sink.toBlob().arrayBuffer()));
assert.ok((await archive.toBlob()).size > 0);
const packageFile = new opc.OpcPackage(); packageFile.addPart({ uri: "/test.xml", contentType: "application/xml", data: "<test/>" });
assert.ok((await packageFile.toBlob()).size > 0);
const resolved = new core.ExportCell(12.5, { text: "=display", presentation: { background: "E2F0D9", bold: true } });
const streamedSink = new core.BlobByteSink();
const streamed = new Workbook({ sink: streamedSink, oversizedText: "preserve", dateMode: "utc", limits: { maxRows: 10, maxStyles: 32 } });
const streamedSheet = streamed.addWorksheet("Resolved", { columns: [{ header: "Amount", key: "amount", type: "number", groups: ["Metrics"], format: "0.00" }],
  conditionalFormats: [{ type: "cellIs", range: { column: "amount" }, operator: "greaterThan", value: 10, style: { fill: { color: "C6EFCE" } } }],
  autoSize: {}, footer: { totals: { amount: "sum" } }, print: { repeatHeaders: true } });
await streamedSheet.addRows([{ amount: resolved }]); await streamedSheet.close();
const streamedResult = await streamed.finish(); assert.equal(streamedResult.rows, 1);
const streamedBlob = streamedSink.toBlob(); assert.equal(streamedResult.bytes, streamedBlob.size);
await writeFile("packed-streamed.xlsx", new Uint8Array(await streamedBlob.arrayBuffer()));
assert.equal(await (await csv.writeCsv([[resolved]], { columns: [{ header: "Amount" }], valueMode: "display" })).text(), "Amount\r\n'=display\r\n");
const manifest = JSON.parse(await readFile("node_modules/@evotecit/officeimo/package.json", "utf8"));
assert.equal(manifest.dependencies, undefined);
assert.deepEqual(Object.keys(manifest.devDependencies), ["typescript"]);
// The optional installed entry must execute independently of a document or grid.
const { port1, port2 } = new MessageChannel(), workerSink = new core.BlobByteSink();
const handler = handleWorkerExports(port2, ({ rows, sink, signal }) => csvModule.writeCsvTo(rows, sink, {
  columns: [{ header: "Amount" }, { header: "Seen" }], signal
}));
try {
  const result = await executeWorkerExport(port1, [[new core.ExportCell(12.5, { text: "12.50 USD" }), new Date("2026-10-20T08:00:00Z")]], workerSink);
  assert.equal(result.rows, 1);
  assert.equal(await workerSink.toBlob().text(), "Amount,Seen\r\n12.5,2026-10-20T08:00:00.000Z\r\n");
  await import("@evotecit/officeimo/execution.js");
  assert.equal(typeof globalThis.OfficeIMO.executeWorkerExport, "function");
} finally { handler.dispose(); port1.close(); port2.close(); }
console.log("All packed subpaths imported and executed without runtime dependencies or a DOM.");
