import test from "node:test";
import assert from "node:assert/strict";
import { Workbook } from "../dist/xlsx/index.js";
import { writeCsv } from "../dist/csv/index.js";
import { readZip } from "./zip-reader.mjs";

test("CSV and XLSX project only own record properties, including prototype-shaped keys", async () => {
  const keys = ["constructor", "toString", "__proto__", "hasOwnProperty", "inherited"];
  let inheritedReads = 0;
  const prototype = Object.defineProperty({}, "inherited", { get() { inheritedReads++; throw new Error("Inherited getter was read."); } });
  const own = Object.fromEntries(keys.map((key, index) => [key, "own-" + index]));
  const nullPrototype = Object.assign(Object.create(null), own);
  const rows = [{}, Object.create(null), Object.create(prototype), own, nullPrototype];
  for (const columns of [keys.map(key => ({ header: "Column " + key, key })), keys.map(header => ({ header }))]) {
    const csv = await writeCsv(rows, { columns, includeHeader: false });
    assert.equal(await csv.text(), ",,,,\r\n".repeat(3) + "own-0,own-1,own-2,own-3,own-4\r\n".repeat(2));
    const book = new Workbook({ compression: "store" });
    await book.addWorksheet("Own properties", { columns, includeHeader: false }).addRows(rows);
    const xml = (await readZip(await book.toBlob())).get("xl/worksheets/sheet1.xml").content;
    for (let row = 1; row <= 3; row++) assert.match(xml, new RegExp('<row r="' + row + '">(?:<c r="[A-E]' + row + '" s="\\d+"/>){5}</row>'));
    for (let row = 4; row <= 5; row++) for (let index = 0; index < keys.length; index++)
      assert.match(xml, new RegExp('r="' + String.fromCharCode(65 + index) + row + '"[^>]*>.*?<t xml:space="preserve">own-' + index + '</t>'));
  }
  assert.equal(inheritedReads, 0);
});
