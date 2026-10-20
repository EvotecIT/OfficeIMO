import test from "node:test";
import assert from "node:assert/strict";
import { writeCsv, writePdf } from "../dist/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

test("CSV captures host boolean text, protects labels as formulas, and budgets rendered characters", async () => {
  const booleanText = { true: "Tak", false: "Nie" };
  function* rows() { booleanText.true = "changed"; yield [true]; yield [false]; }
  const options = { columns: [{ header: "Enabled" }], booleanText, includeHeader: false };
  assert.equal(await (await writeCsv(rows(), options)).text(), "Tak\r\nNie\r\n");
  const protectedCsv = await writeCsv([[true], [false]], { ...options, booleanText: { true: "=TRUE()", false: "-1" } });
  assert.equal(await protectedCsv.text(), "'=TRUE()\r\n'-1\r\n");
  await assert.rejects(writeCsv([[true]], { ...options, booleanText: { true: "long label", false: "" },
    limits: { maxTextCharacters: 3 } }), { code: "RESOURCE_LIMIT" });
});

test("PDF boolean display remains separate from typed input and explicit cell formatters", async () => {
  const columns = [{ header: "Enabled" }], booleanText = { true: "Ja", false: "Nein" };
  function* rows() { booleanText.true = "changed"; yield [true]; yield [false]; }
  const pdf = await inspectPdf(await writePdf(rows(), { columns, booleanText, compression: false, pageNumbers: false }));
  assert.ok(pdf.text.includes("Ja")); assert.ok(pdf.text.includes("Nein")); assert.ok(!pdf.text.includes("changed"));
  const values = [];
  const formatted = await inspectPdf(await writePdf([[true]], { columns, booleanText, compression: false, pageNumbers: false,
    formatValue(value) { values.push(value); return "explicit"; } }));
  assert.deepEqual(values, [true]); assert.ok(formatted.text.includes("explicit"));
});
