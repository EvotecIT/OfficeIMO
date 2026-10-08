import test from "node:test";
import assert from "node:assert/strict";
import { fixtures, createFixture } from "./fixtures.mjs";
import { readZip } from "./zip-reader.mjs";

for (const spec of fixtures.cases) for (const compression of ["auto", "store"]) test("shared XLSX fixture: " + spec.name + "/" + compression, async () => {
  const zip = await readZip(await createFixture(spec, compression));
  assert.ok(zip.has("xl/workbook.xml")); assert.ok(zip.has("docProps/app.xml"));
  if (compression === "store") assert.ok([...zip.values()].every(e => e.method === 0));
  if (spec.parts) for (const part of spec.parts) assert.equal(zip.get(part.uri.slice(1)).content, part.data);
});
