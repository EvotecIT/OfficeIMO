// Explicit evidence run, outside correctness gates. Node platform modules only.
import assert from "node:assert/strict";
import { readFile, mkdir, writeFile } from "node:fs/promises";
import { join, resolve } from "node:path";
import { gzipSync, brotliCompressSync, deflateRawSync, constants } from "node:zlib";
import { createWorkbook } from "../../OfficeIMO.JavaScript/dist/xlsx/index.js";
import { readZip } from "../../OfficeIMO.JavaScript/test/zip-reader.mjs";

if (process.argv.length !== 3) throw new Error("Usage: node Build/BrowserExports/measure.mjs <evidence-directory>");
const output = resolve(process.argv[2]);
await mkdir(output, { recursive: true });
const sizes = [];
for (const name of ["officeimo-xlsx", "officeimo-csv", "officeimo"]) {
  for (const extension of ["mjs", "js"]) {
    const file = name + "." + extension;
    const bytes = await readFile(new URL("../../OfficeIMO.JavaScript/bundles/" + file, import.meta.url));
    sizes.push({ file, bytes: bytes.length, gzip9: gzipSync(bytes, { level: 9 }).length,
      brotli11: brotliCompressSync(bytes, { params: { [constants.BROTLI_PARAM_QUALITY]: 11 } }).length });
  }
}

const strings = [];
for (const workload of ["repeated", "unique"]) {
  const rows = 10000, columns = 4;
  function* values() {
    for (let r = 0; r < rows; r++) yield Array.from({ length: columns }, (_, c) =>
      workload === "repeated" ? "Site " + r % 20 + " field " + c : "Row " + r + " field " + c);
  }
  const workbook = createWorkbook();
  await workbook.addSheet("Strings", { columns: Array.from({ length: columns }, () => ({ header: "V" })), includeHeader: false }).addRows(values());
  const xml = (await readZip(await workbook.toBlob())).get("xl/worksheets/sheet1.xml").content;
  // Compare equivalent XML payloads with the same deflater. This is a measurement
  // prototype, not a second workbook writer or a product shared-string dictionary.
  const dictionary = new Map(), nodes = [];
  let count = 0;
  const sharedSheet = xml.replace(/t="inlineStr"><is>(.*?)<\/is>/g, (_, node) => {
    if (!dictionary.has(node)) { dictionary.set(node, nodes.length); nodes.push(node); }
    const index = dictionary.get(node);
    assert.equal(nodes[index], node);
    count++;
    return 't="s"><v>' + index + '</v>';
  });
  assert.equal(count, rows * columns);
  const sharedTable = '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="' + count +
    '" uniqueCount="' + nodes.length + '">' + nodes.map(node => '<si>' + node + '</si>').join("") + '</sst>';
  const inlineCompressed = deflateRawSync(xml).length;
  const sharedCompressed = deflateRawSync(sharedSheet).length + deflateRawSync(sharedTable).length;
  strings.push({ workload, rows, columns, cells: count, uniqueStrings: nodes.length,
    inlineXmlBytes: Buffer.byteLength(xml), sharedXmlBytes: Buffer.byteLength(sharedSheet) + Buffer.byteLength(sharedTable),
    inlineDeflateBytes: inlineCompressed, sharedDeflateBytes: sharedCompressed,
    sharedSavingPercent: (inlineCompressed - sharedCompressed) / inlineCompressed * 100,
    dictionaryTextUtf8Bytes: nodes.reduce((sum, node) => sum + Buffer.byteLength(node), 0),
    dictionaryTextCodeUnits: nodes.reduce((sum, node) => sum + node.length, 0) });
}
const report = { recordedUtc: new Date().toISOString(), node: process.version, sizes, strings,
  method: "Actual inline worksheet XML versus an equivalent indexed worksheet plus SST; same Node deflateRaw defaults. ZIP metadata excluded. Dictionary counts are text only, not measured heap or Map overhead." };
await writeFile(join(output, "size-and-strings.json"), JSON.stringify(report, null, 2) + "\n");
console.log(JSON.stringify(report, null, 2));
