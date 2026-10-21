import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import * as api from "../dist/index.js";
import { reportColumns, serviceReliabilityDocument } from "../../OfficeIMO.JavaScript.Examples/report-engine-job.mjs";
import { inspectPdf } from "./pdf-reader.mjs";

const fonts = Object.fromEntries(await Promise.all(["regular", "bold"].map(async kind => [kind,
  new Uint8Array(await readFile(new URL("../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-" +
    (kind === "regular" ? "Regular" : "Bold") + ".ttf", import.meta.url)))])));
const input = { fonts, summary: { total: 1, Healthy: 1, Attention: 0, Offline: 0 } };

test("service PDF keeps semantic status and numeric roles after localized headings and reordering", async () => {
  const order = [4, 0, 2, 3, 5, 1], labels = ["Dostępność", "Status", "Stan", "Opóźnienie", "Sprawdzone", "Miejsce"];
  const columns = order.map((index, position) => ({ ...reportColumns[index], header: labels[position] }));
  const source = ["SVC-001", "Warsaw", "Healthy", 12.34, 0.995, new Date("2026-10-20T08:00:00Z")];
  let returned = 0;
  async function* rows() { try { yield order.map(index => source[index]); } finally { returned++; } }
  const report = await serviceReliabilityDocument(api, { input, rows: rows(), columns });
  const pdf = await inspectPdf(await api.writePdfDocument(report.blocks, { ...report.options, compression: false }));
  for (const text of [...labels, "SVC-001", "Healthy", "12.3", "99.50%", "2026-10-20 08:00"])
    assert.ok(pdf.text.includes(text), "Missing report text: " + text);
  assert.equal(returned, 1);
});

test("service PDF requires one semantic status field instead of accepting a misleading heading", async () => {
  for (const columns of [reportColumns.map(column => column.key === "status" ? { ...column, key: "other" } : column),
    [...reportColumns, { ...reportColumns[2], header: "Stan" }]]) {
    let consumed = false;
    await assert.rejects(serviceReliabilityDocument(api, { input, columns,
      rows: { *[Symbol.iterator]() { consumed = true; yield []; } } }), /one selected column with key 'status'/);
    assert.equal(consumed, false);
  }
});
