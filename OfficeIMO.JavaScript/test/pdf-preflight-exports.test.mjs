import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import vm from "node:vm";
import { preflightPdfText, pdf } from "../dist/index.js";

test("PDF preflight executes through root, standalone ESM and classic assets", async () => {
  assert.equal(preflightPdfText, pdf.preflightPdfText);
  for (const name of ["officeimo", "officeimo-pdf"]) {
    const standalone = await import("../bundles/" + name + ".mjs");
    assert.deepEqual(standalone.preflightPdfText("Report\r\n\t€"), { supported: true });
    const context = vm.createContext({ TextEncoder, Blob, Date, performance, setTimeout, DOMException });
    vm.runInContext(await readFile(new URL("../bundles/" + name + ".js", import.meta.url), "utf8"), context);
    assert.equal(context.OfficeIMO.preflightPdfText("Report").supported, true);
    const issue = context.OfficeIMO.pdf.preflightPdfText("City Ł");
    assert.equal(issue.supported, false); assert.equal(issue.code, "embedded-font-required"); assert.equal(issue.index, 5);
  }
});
