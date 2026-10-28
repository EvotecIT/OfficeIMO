import test from "node:test";
import assert from "node:assert/strict";
import { writeFile } from "node:fs/promises";
import { exportCanopy, createCanopyExport } from "../dist/integrations/canopyx/index.js";
import { writePdf } from "../dist/pdf/index.js";
import { inspectPdf } from "./pdf-reader.mjs";

// Ordinary 48-order, seven-column native grid shape: screen widths and wrap=false
// are defaults, not a PDF-specific request to reject rather than wrap content.
function ordersCapture() {
  const ids = ["order", "customer", "region", "status", "placed", "items", "amount"];
  const rows = Array.from({ length: 48 }, (_, i) => ({ id: "o" + i, cells: {
    order: { value: "SO-" + (1000 + i), text: "SO-" + (1000 + i) },
    customer: { value: "Customer " + String.fromCharCode(65 + i % 26), text: "Customer " + String.fromCharCode(65 + i % 26) },
    region: { value: ["North", "South", "East", "West"][i % 4], text: ["North", "South", "East", "West"][i % 4] },
    status: { value: "Shipped", text: "Shipped" },
    placed: { value: new Date(Date.UTC(2026, 0, i + 1)).toISOString(), text: new Date(Date.UTC(2026, 0, i + 1)).toISOString() },
    items: { value: 2 + i % 3, text: String(2 + i % 3) },
    amount: { value: 1234.5 + i, text: (1234.5 + i).toFixed(2) }
  } }));
  return { request: { columns: ids.map(id => ({ id, title: id, kind: id === "placed" ? "datetime" : ["items", "amount"].includes(id) ? "number" : "text",
    width: { unit: "css-px", preferred: 160, minimum: 80 }, wrap: false, alignment: ["items", "amount"].includes(id) ? "right" : "left" })),
    recordCount: rows.length, values: "raw", presentation: "semantic", timeZone: "utc", revision: "orders-48" },
    async *rows() { yield* rows; } };
}

test("Canopy default fitted PDF wraps all ordinary text date and numeric cells without losing rows", async () => {
  const capture = ordersCapture(), blob = await exportCanopy(capture, "pdf"), pdf = await inspectPdf(blob);
  const normalized = pdf.text.replace(/\s/g, "");
  for (let i = 0; i < 48; i++) {
    assert.ok(normalized.includes("SO-" + (1000 + i)));
    assert.ok(normalized.includes(new Date(Date.UTC(2026, 0, i + 1)).toISOString()));
    assert.ok(normalized.includes((1234.5 + i).toFixed(2)));
  }
  assert.ok(normalized.includes("CustomerA") && normalized.includes("Shipped"));
  assert.ok(createCanopyExport(capture, "pdf").columns.every(column => column.wrapText === true));
  assert.ok(createCanopyExport(capture, "xlsx").columns.every(column => column.wrapText === false));
  if (process.env.OFFICEIMO_CANOPY_FIXTURES)
    await writeFile(process.env.OFFICEIMO_CANOPY_FIXTURES + "/canopy-default-orders.pdf", new Uint8Array(await blob.arrayBuffer()));
});

test("Canopy explicit nonwrapping PDF columns and core writer keep strict overflow semantics", async () => {
  await assert.rejects(exportCanopy(ordersCapture(), "pdf", { columnOptions: { placed: { wrapText: false } } }), /Unwrapped PDF cell/);
  await assert.rejects(writePdf([["An ordinary long label"]], { columns: [{ header: "Name", wrapText: false }], columnWidths: [30] }), /Unwrapped PDF cell/);
  const result = await exportCanopy(ordersCapture(), "pdf", { columnOptions: { placed: { wrapText: false } },
    pdf: { columnWidths: [50, 90, 50, 50, 200, 40, 80], orientation: "landscape", pageNumbers: false } });
  assert.ok((await inspectPdf(result)).text.includes("2026-01-01T00:00:00.000Z"));
});
