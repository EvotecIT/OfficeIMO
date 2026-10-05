import assert from "node:assert/strict";
import { createWorkbook, writeCsv } from "@evotecit/officeimo";
import { createWorkbook as xlsx } from "@evotecit/officeimo/xlsx";
import { writeCsv as csv } from "@evotecit/officeimo/csv";
import { readFile, writeFile } from "node:fs/promises";

assert.equal(typeof createWorkbook, "function");
assert.equal(typeof xlsx, "function");
assert.equal(typeof csv, "function");
assert.equal(globalThis.document, undefined); // Module imports are safe in a Node/SSR host.
const book = xlsx({ compression: "store" });
await book.addSheet("Packed", { columns: [{ header: "Name" }] }).addRows([["Łódź"]]);
await writeFile("packed-consumer.xlsx", Buffer.from(await book.toBlob().then(b => b.arrayBuffer())));
assert.equal(await (await writeCsv([["=cmd"]], { columns: [{ header: "Name" }] })).text(), "Name\r\n'=cmd\r\n");
const manifest = JSON.parse(await readFile("node_modules/@evotecit/officeimo/package.json", "utf8"));
assert.equal(manifest.dependencies, undefined);
assert.equal(manifest.devDependencies, undefined);
console.log("Packed package exports, SSR import and real writes passed.");
