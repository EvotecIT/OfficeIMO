import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { Workbook, writeXlsx, writeXlsxTo, core, zip, xml, opc, xlsx, csv } from "@evotecit/officeimo";
import * as coreModule from "@evotecit/officeimo/core";
import * as zipModule from "@evotecit/officeimo/zip";
import * as xmlModule from "@evotecit/officeimo/xml";
import * as opcModule from "@evotecit/officeimo/opc";
import * as xlsxModule from "@evotecit/officeimo/xlsx";
import * as csvModule from "@evotecit/officeimo/csv";
import * as dataTablesModule from "@evotecit/officeimo/integrations/datatables";

for (const [namespace, module] of [[core, coreModule], [zip, zipModule], [xml, xmlModule], [opc, opcModule], [xlsx, xlsxModule], [csv, csvModule]])
  for (const [key, value] of Object.entries(module)) assert.equal(namespace[key], value);
assert.equal(globalThis.document, undefined);
assert.equal(typeof dataTablesModule.createDataTablesExport, 'function');
assert.equal(dataTablesModule.ExportCell, coreModule.ExportCell);
assert.equal(writeXlsx, xlsx.writeXlsx); assert.equal(writeXlsxTo, xlsx.writeXlsxTo);
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
console.log("All packed subpaths imported and executed without runtime dependencies or a DOM.");
