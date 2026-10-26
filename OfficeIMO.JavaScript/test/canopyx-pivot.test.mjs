import test from "node:test";
import assert from "node:assert/strict";
import { readFile, writeFile } from "node:fs/promises";
import { exportCanopy, createCanopyExport } from "../dist/integrations/canopyx/index.js";
import { writeXlsx } from "../dist/xlsx/index.js";
import { readZip } from "./zip-reader.mjs";
import { inspectPdf } from "./pdf-reader.mjs";

// Frozen CanopyX projection: seven accepted records produce ten prefix/detail/total rows.
// Includes null axes, opaque keys, limited columns/Other, collapsed descendants and fractional shares.
const fixture = JSON.parse(await readFile(new URL("./fixtures/canopy-pivot.json", import.meta.url), "utf8"));
function capture(snapshot = fixture.snapshot) {
  const { rows, columns, query, matched, ...metadata } = structuredClone(snapshot);
  const { definition, rowAxis, columnAxis, columnGroups, distinctColumns, retainedColumns, columnsLimited } = metadata;
  return { request: { columns, recordCount: rows.length,
    scope: "filtered", values: "raw", presentation: "semantic", timeZone: "utc", revision: "inline", query, matched,
    pivot: { definition, rowAxis, columnAxis, columnGroups, distinctColumns, retainedColumns, columnsLimited } },
    async *rows({ signal } = {}) { for (const row of rows) { signal?.throwIfAborted(); yield row; } } };
}
async function retain(name, blob) {
  if (process.env.OFFICEIMO_CANOPY_FIXTURES) await writeFile(process.env.OFFICEIMO_CANOPY_FIXTURES + "/" + name, new Uint8Array(await blob.arrayBuffer()));
}

test("frozen pivot XLSX preserves prefix outlines, hidden columns, raw totals and physical row count without regrouping", async () => {
  const input = capture(), progress = [];
  const blob = await exportCanopy(input, "xlsx", { about: true, onProgress: p => progress.push(p) });
  const zip = await readZip(blob), xml = zip.get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /<outlinePr summaryBelow="0" summaryRight="0"/);
  assert.match(xml, /<row r="2"[^>]*collapsed="1"/);
  assert.match(xml, /<row r="3"[^>]*outlineLevel="1"[^>]*hidden="1"/);
  assert.match(xml, /<col min="4" max="4"[^>]*outlineLevel="1"[^>]*hidden="1"/);
  assert.match(xml, /<col min="2" max="2"[^>]*collapsed="1"/);
  assert.match(xml, /<c r="J11"[^>]*><v>215<\/v>/);
  assert.match(xml, /<c r="B11"[^>]*><v>115<\/v>/);
  assert.match(xml, /<c r="H11"[^>]*><v>105<\/v>/);
  assert.equal([...xml.matchAll(/<row /g)].length, 11);
  assert.doesNotMatch(xml, /<f>|<autoFilter|tablePart/);
  const about = zip.get("xl/worksheets/sheet2.xml").content;
  assert.match(about, /recordCount/); assert.match(about, /<v>10<\/v>/); assert.match(about, /<v>7<\/v>/);
  assert.equal(progress.at(-1).rows, 10);
  const source = createCanopyExport(input, "xlsx"), first = [], second = [];
  for await (const row of source.rows()) first.push(row.map(cell => cell.value));
  for await (const row of source.rows()) second.push(row.map(cell => cell.value));
  assert.deepEqual(first, second); assert.equal(first.length, 10);
  const csv = await (await exportCanopy(input, "csv")).text();
  assert.equal(csv.trimEnd().split("\r\n").length, 11);
  assert.ok(csv.includes("C1")); assert.ok(csv.includes("215"));
  await retain("canopy-pivot.xlsx", blob);
});

test("pivot shares retain fractional numeric cells and their per-cell native percent format", async () => {
  const blob = await exportCanopy(capture(fixture.percent), "xlsx");
  const zip = await readZip(blob), xml = zip.get("xl/worksheets/sheet1.xml").content, styles = zip.get("xl/styles.xml").content;
  const expected = fixture.percent.rows.flatMap(row => Object.values(row.cells)).filter(cell => cell.mode?.startsWith("percent") && cell.value !== null);
  assert.ok(expected.some(cell => cell.value > 0 && cell.value < 1));
  assert.match(styles, /formatCode="0\.00%"/);
  for (const cell of expected) assert.ok(xml.includes("<v>" + cell.value + "</v>"));
  assert.doesNotMatch(xml, /<f>/);
  await retain("canopy-pivot-percent.xlsx", blob);
});

test("explicit pivot column formats win over captured per-cell formats and retain the same numeric value", async () => {
  const input = capture(fixture.percent);
  // This host additionally supplies ordinary dimensions; both native minimal and enriched captures are supported.
  input.request.columns = input.request.columns.map(column => ({ ...column, width: { unit: "css-px", preferred: 160, minimum: 80 }, wrap: true, alignment: "left" }));
  const zip = await readZip(await exportCanopy(input, "xlsx", { columnOptions: { v0_0: { format: "0.0000%" } } }));
  const xml = zip.get("xl/worksheets/sheet1.xml").content, styles = zip.get("xl/styles.xml").content;
  const cell = /<c r="B2" s="(\d+)"[^>]*><v>([^<]+)<\/v>/.exec(xml);
  assert.ok(cell); assert.equal(Number(cell[2]), fixture.percent.rows[0].cells.v0_0.value);
  const formats = new Map([...styles.matchAll(/<numFmt numFmtId="(\d+)" formatCode="([^"]*)"/g)].map(m => [m[1], m[2]]));
  const xfs = [.../<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)[1].matchAll(/<xf\b[^>]*numFmtId="(\d+)"/g)];
  assert.equal(formats.get(xfs[Number(cell[1])][1]), "0.0000%");
});

test("pivot PDF includes hidden descendants and frozen display text while reporting Excel-only outline controls", async () => {
  const diagnostics = [], blob = await exportCanopy(capture(), "pdf", { onDiagnostic: d => diagnostics.push(d.code), pdf: { wideTable: { mode: "bands", repeatColumns: ["axis"] } } });
  const pdf = await inspectPdf(blob);
  for (const text of ["C1", "C3", "C5", "Other", "215"]) assert.ok(pdf.text.includes(text), text);
  assert.deepEqual(diagnostics, ["PIVOT_OUTLINES_XLSX_ONLY"]);
  await retain("canopy-pivot.pdf", blob);
});

test("explicit pivots reject malformed/ordinary-grouped metadata and cannot silently fall back to flat records", async () => {
  for (const change of [input => input.request.pivot = null, input => input.request.scope = "selected",
    input => input.request.grouping = {}, input => input.request.pivot.columnGroups[0].columnCount = 100000]) {
    const input = capture(); let read = false; input.rows = async function* () { read = true; };
    change(input); await assert.rejects(exportCanopy(input, "xlsx"), /pivot/i); assert.equal(read, false);
  }
  const input = capture(); await assert.rejects(exportCanopy(input, "xlsx", { xlsx: { sheet: { footer: { values: ["recompute"] } } } }), /owns outlines and totals/);
  const bad = capture(), rows = []; for await (const row of bad.rows()) rows.push(row);
  rows[0].cells.v0_0.numberFormat = undefined; bad.rows = async function* () { yield* rows; };
  await assert.rejects(exportCanopy(bad, "xlsx"), /mode and native numberFormat/);
  const controller = new AbortController(), aborted = capture(); let returned = false;
  aborted.rows = async function* ({ signal }) { try { yield rows[1]; controller.abort(); signal.throwIfAborted(); } finally { returned = true; } };
  await assert.rejects(exportCanopy(aborted, "xlsx", { signal: controller.signal }), { name: "AbortError" }); assert.equal(returned, true);
});

test("generic worksheet column ranges work without widths, capture caller settings and reject overlap/out-of-range levels", async () => {
  const columns = [{ header: "A" }, { header: "B" }, { header: "C" }], ranges = [{ firstColumn: 1, columnCount: 2, level: 1, hidden: true }];
  const pending = writeXlsx([[1, 2, 3]], { columns, sheet: { autoSize: { sampleRows: 0 }, outlineSummaryRight: false, columnOutlines: ranges } });
  ranges[0].hidden = false;
  const xml = (await readZip(await pending)).get("xl/worksheets/sheet1.xml").content;
  assert.match(xml, /outlineLevelCol="1"/); assert.match(xml, /<col min="2" max="2"[^>]*hidden="1"/);
  for (const columnOutlines of [[{ firstColumn: 2, columnCount: 2 }], [{ firstColumn: 0, columnCount: 1, level: 8 }],
    [{ firstColumn: 0, columnCount: 2 }, { firstColumn: 1, columnCount: 1 }]])
    await assert.rejects(writeXlsx([[1, 2, 3]], { columns, sheet: { columnOutlines } }), /outline/i);
});


