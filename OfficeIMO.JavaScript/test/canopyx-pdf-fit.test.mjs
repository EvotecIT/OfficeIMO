import test from "node:test";
import assert from "node:assert/strict";
import { writeFile } from "node:fs/promises";
import { exportCanopy, createCanopyExport, writeCanopyTo } from "../dist/integrations/canopyx/index.js";
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
  await assert.rejects(exportCanopy(ordersCapture(), "pdf", { columnOptions: { placed: { wrapText: false } },
    pdf: { columnWidths: [50, 90, 50, 50, 30, 40, 80] } }), /Unwrapped PDF cell/);
  await assert.rejects(writePdf([["An ordinary long label"]], { columns: [{ header: "Name", wrapText: false }], columnWidths: [30] }), /Unwrapped PDF cell/);
  const result = await exportCanopy(ordersCapture(), "pdf", { columnOptions: { placed: { wrapText: false } },
    pdf: { columnWidths: [50, 90, 50, 50, 200, 40, 80], orientation: "landscape", pageNumbers: false } });
  assert.ok((await inspectPdf(result)).text.includes("2026-01-01T00:00:00.000Z"));
});

function painted(pdf) {
  return pdf.pages.flatMap(page => [...page.content.matchAll(/BT \/(F\d+) [\s\S]*?([-\d.]+) ([-\d.]+) Tm <([a-f0-9]*)> Tj ET/g)]
    .map(match => ({ font: match[1], x: Number(match[2]), y: Number(match[3]), text: Buffer.from(match[4], "hex").toString("latin1") })));
}

test("Canopy PDF fits measured headers and date values without mid-word breaks or equal column widths", async () => {
  const capture = ordersCapture(), rows = [];
  for await (const row of capture.rows()) rows.push(row);
  capture.request.columns = capture.request.columns.map(column => ({ ...column,
    title: column.id === "items" ? "Lines" : column.id[0].toUpperCase() + column.id.slice(1) }));
  rows.forEach(row => { row.cells.placed.text = "2026-01-01 00:00 UTC"; });
  capture.rows = async function* () { yield* rows; };
  const pdf = await inspectPdf(await exportCanopy(capture, "pdf", { pdf: { pageSize: { width: 600, height: 500 } } })), runs = painted(pdf);
  assert.ok(runs.some(run => run.text === "Amount"), "Amount is one whole header word");
  assert.equal(runs.filter(run => run.text === "2026-01-01 00:00 UTC").length, 48, "Every timestamp fits on one line");
  const header = name => runs.find(run => run.text === name);
  assert.ok(header("Placed").x - header("Status").x < header("Lines").x - header("Placed").x,
    "The date column receives more width than a short status column");
  assert.ok(pdf.pages.length <= 3, "Natural date widths do not double the height of every row");
});

test("Canopy PDF wraps text at whole words while preserving every character", async () => {
  const input = { request: { columns: [{ id: "label", title: "Description", kind: "text" }, { id: "date", title: "Placed", kind: "datetime" }],
    recordCount: 1, values: "raw", presentation: "text", timeZone: "utc", revision: "word-wrap" },
    async *rows() { yield { id: "one", cells: { label: { value: null, text: "A readable sentence with several whole words" }, date: { value: null, text: "2026-01-01 00:00 UTC" } } }; } };
  const pdf = await inspectPdf(await exportCanopy(input, "pdf", { pdf: { pageSize: { width: 310, height: 300 }, orientation: "portrait" } }));
  const lines = painted(pdf).filter(run => run.y < painted(pdf).find(run => run.text === "Description").y && !run.text.startsWith("Page"));
  const tokens = new Set("A readable sentence with several whole words 2026-01-01 00:00 UTC".split(" "));
  for (const line of lines) for (const word of line.text.trim().split(/\s+/)) assert.ok(tokens.has(word), "No word is split: " + word);
  assert.ok(pdf.text.replace(/\s/g, "").includes("Areadablesentencewithseveralwholewords"));
});

test("Canopy PDF content measurement bounds lookahead and closes the single source on destination failure", async () => {
  let read = 0, calls = 0, closed = 0;
  const input = ordersCapture(); input.request.recordCount = 500;
  const example = (await input.rows().next()).value;
  input.rows = async function* () { calls++; try { for (let i = 0; i < 500; i++) { read++; yield { ...example, id: "row-" + i }; } } finally { closed++; } };
  const failure = new Error("Destination rejected PDF");
  await assert.rejects(writeCanopyTo(input, "pdf", { write() { throw failure; } }), error => error === failure);
  assert.equal(calls, 1); assert.ok(read <= 128, "Measurement never buffers the whole source");
  await new Promise(resolve => setImmediate(resolve));
  assert.equal(closed, 1);
});

test("Canopy PDF freezes nested settings before awaiting its content sample", async () => {
  let entered, release;
  const waiting = new Promise(resolve => { entered = resolve; }), blocked = new Promise(resolve => { release = resolve; });
  const input = ordersCapture(), first = (await input.rows().next()).value; input.request.recordCount = 1;
  input.rows = async function* () { entered(); await blocked; yield first; };
  const widths = [60, 100, 50, 60, 160, 40, 70], labels = { prefix: "Original ", separator: " of " };
  const pending = exportCanopy(input, "pdf", { pdf: { columnWidths: widths, pageNumberLabels: labels } });
  await waiting; widths[0] = 1; labels.prefix = "Mutated "; release();
  const pdf = await inspectPdf(await pending);
  assert.ok(pdf.text.includes("Original 1 of ")); assert.ok(!pdf.text.includes("Mutated"));
});

test("Canopy PDF measures bare CR and CRLF line boundaries like the writer", async () => {
  for (const ending of ["\r", "\r\n"]) {
    const input = { request: { columns: [{ id: "label", title: "Region" + ending + "name", kind: "text" }],
      recordCount: 1, values: "raw", presentation: "text", timeZone: "utc", revision: "line-boundaries" },
      async *rows() { yield { id: "one", cells: { label: { value: null, text: "North" + ending + "South" } } }; } };
    const pdf = await inspectPdf(await exportCanopy(input, "pdf"));
    for (const line of ["Region", "name", "North", "South"]) assert.ok(pdf.pages[0].lines.includes(line));
  }
});

test("Canopy PDF distinguishes thousands marks from the sampled decimal convention", async () => {
  for (const [integer, fractional, large, small] of [["1,234", "12.5", "1,234.000", "12.500"], ["1.234", "12,5", "1.234,000", "12,500"]]) {
    const input = { request: { columns: [{ id: "amount", title: "Amount", kind: "number" }], recordCount: 2,
      values: "raw", presentation: "text", timeZone: "utc", revision: "grouped-decimals" },
      async *rows() { yield { id: "a", cells: { amount: { value: 1234, text: integer } } }; yield { id: "b", cells: { amount: { value: 12.5, text: fractional } } }; } };
    const pdf = await inspectPdf(await exportCanopy(input, "pdf", { columnOptions: { amount: { format: "#,##0.000" } } }));
    assert.ok(pdf.pages[0].lines.includes(large)); assert.ok(pdf.pages[0].lines.includes(small));
  }
});

test("flat PDF percentage labels choose the document culture without scaling their raw values", async () => {
  const input = { request: { columns: [
    { id: "amount", title: "Amount", kind: "number", format: { minimumFractionDigits: 2, maximumFractionDigits: 2 } },
    { id: "share", title: "Share", kind: "number", format: { minimumFractionDigits: 1, maximumFractionDigits: 1, suffix: "%" } }
  ], recordCount: 1, values: "raw", presentation: "text", timeZone: "utc", revision: "literal-percent" },
    async *rows() { yield { id: "a", cells: { amount: { value: 1234.5, text: "1234.50" }, share: { value: 18.6, text: "18,6%" } } }; } };
  const pdf = await inspectPdf(await exportCanopy(input, "pdf")), lines = pdf.pages.flatMap(page => page.lines);
  assert.ok(lines.includes("1234,50")); assert.ok(lines.includes("18,6%"));
});

test("flat PDF preserves quoted and backslash numeric affixes using the document separator", async () => {
  const affixes = [[ 'a"b', "" ], [ "", 'a"b' ], [ "a\\b", "" ], [ "", "a\\b" ], [ 'a"\\b', '\\z"' ]];
  const input = { request: { columns: [
    ...affixes.map(([prefix, suffix], i) => ({ id: "v" + i, title: "Value " + i, kind: "number",
      format: { minimumFractionDigits: 2, maximumFractionDigits: 2, prefix, suffix } })),
    { id: "share", title: "Share", kind: "number", format: { minimumFractionDigits: 1, maximumFractionDigits: 1, suffix: "%" } }
  ], recordCount: 1, values: "raw", presentation: "text", timeZone: "utc", revision: "escaped-affixes" },
    async *rows() { yield { id: "a", cells: { ...Object.fromEntries(affixes.map(([prefix, suffix], i) =>
      ["v" + i, { value: 12.5, text: prefix + "12.50" + suffix }])), share: { value: 18.6, text: "18,6%" } } }; } };
  const pdf = await inspectPdf(await exportCanopy(input, "pdf")), lines = pdf.pages.flatMap(page => page.lines);
  for (const [prefix, suffix] of affixes) assert.ok(lines.includes(prefix + "12,50" + suffix), prefix + "12,50" + suffix);
  assert.ok(lines.includes("18,6%"));
});

test("flat PDF undeclared numeric precision preserves captured integer rounding", async () => {
  for (const [value, text, expected] of [[12.0001, "12", "12"], [1234.25, "1,234", "1234"]]) {
    const input = { request: { columns: [{ id: "value", title: "Value", kind: "number" }], recordCount: 1,
      values: "raw", presentation: "text", timeZone: "utc", revision: "rounded-integer" },
      async *rows() { yield { id: "a", cells: { value: { value, text } } }; } };
    const pdf = await inspectPdf(await exportCanopy(input, "pdf"));
    assert.ok(pdf.pages.flatMap(page => page.lines).includes(expected), expected);
  }
});

test("flat PDF preserves undeclared decorated numeric labels without inferring scaling", async () => {
  for (const [value, text] of [[.186, "18.6%"], [.7, "Score 7"]]) {
    const input = { request: { columns: [{ id: "value", title: "Value", kind: "number" }], recordCount: 1,
      values: "raw", presentation: "text", timeZone: "utc", revision: "explicit-label" },
      async *rows() { yield { id: "a", cells: { value: { value, text } } }; } };
    const pdf = await inspectPdf(await exportCanopy(input, "pdf"));
    assert.ok(pdf.pages.flatMap(page => page.lines).includes(text), text);
  }
});
