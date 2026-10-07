import { mkdir, writeFile, readFile } from "node:fs/promises";
import { resolve, join } from "node:path";
import { fixtures, createFixture } from "../test/fixtures.mjs";
import { writeCsv } from "../dist/csv/index.js";

if (!process.argv[2]) throw new Error("Usage: node scripts/fixtures.mjs <output-directory>");
const output = resolve(process.argv[2]);
await mkdir(output, { recursive: true });
for (const spec of fixtures.cases) for (const compression of ["auto", "store"]) {
  const blob = await createFixture(spec, compression);
  await writeFile(join(output, spec.name + "-" + compression + ".xlsx"), new Uint8Array(await blob.arrayBuffer()));
}
const csv = JSON.parse(await readFile(new URL("../../OfficeIMO.TestAssets/CSV/browser-exports.json", import.meta.url), "utf8"));
for (const vector of csv.cases) {
  const rows = vector.rows.map(row => row.map(v => v?.kind === "date" ? new Date(v.value) : v));
  await writeFile(join(output, vector.name + ".csv"), new Uint8Array(await (await writeCsv(rows, vector)).arrayBuffer()));
}
console.log("Generated " + fixtures.cases.length * 2 + " shared XLSX fixtures and " + csv.cases.length + " shared CSV vectors.");
