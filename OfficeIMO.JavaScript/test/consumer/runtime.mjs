import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { createWorkbook, core, zip, xml, opc, xlsx, csv } from "@evotecit/officeimo";
import * as coreModule from "@evotecit/officeimo/core";
import * as zipModule from "@evotecit/officeimo/zip";
import * as xmlModule from "@evotecit/officeimo/xml";
import * as opcModule from "@evotecit/officeimo/opc";
import * as xlsxModule from "@evotecit/officeimo/xlsx";
import * as csvModule from "@evotecit/officeimo/csv";

for (const [namespace, module] of [[core, coreModule], [zip, zipModule], [xml, xmlModule], [opc, opcModule], [xlsx, xlsxModule], [csv, csvModule]])
  for (const [key, value] of Object.entries(module)) assert.equal(namespace[key], value);
assert.equal(globalThis.document, undefined);
const book = createWorkbook({ compression: "store" });
await book.addWorksheet("Packed", { columns: [{ header: "Name" }] }).addRows([["Łódź"]]);
await writeFile("packed-consumer.xlsx", new Uint8Array(await (await book.toBlob()).arrayBuffer()));
assert.equal(await (await csv.writeCsv([["=cmd"]], { columns: [{ header: "Name" }] })).text(), "Name\r\n'=cmd\r\n");
const sink = new core.BlobByteSink(), xmlWriter = new xml.XmlWriter(sink);
await xmlWriter.startElement("root"); await xmlWriter.text("Łódź"); await xmlWriter.endElement(); await xmlWriter.close();
const archive = new zip.ZipWriter(); await archive.add("test.xml", new Uint8Array(await sink.toBlob().arrayBuffer()));
assert.ok((await archive.toBlob()).size > 0);
const packageFile = new opc.OpcPackage(); packageFile.addPart({ uri: "/test.xml", contentType: "application/xml", data: "<test/>" });
assert.ok((await packageFile.toBlob()).size > 0);
const manifest = JSON.parse(await readFile("node_modules/@evotecit/officeimo/package.json", "utf8"));
assert.equal(manifest.dependencies, undefined);
assert.deepEqual(Object.keys(manifest.devDependencies), ["typescript"]);
console.log("All packed subpaths imported and executed without runtime dependencies or a DOM.");
