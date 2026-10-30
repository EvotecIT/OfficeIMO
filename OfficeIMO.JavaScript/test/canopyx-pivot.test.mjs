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

test("pivot PDF distinguishes summary rows and right-aligns consistently formatted numeric values", async () => {
  const snapshot = structuredClone(fixture.percent);
  snapshot.rows[0].cells.v0_0 = { value: .186, text: "18,6%", mode: "percentGrandTotal", numberFormat: "0.00%" };
  snapshot.rows[1].cells.v0_0 = { value: .093, text: "9,3%", mode: "percentGrandTotal", numberFormat: "0.00%" };
  const input = capture(snapshot), source = createCanopyExport(input, "pdf");
  assert.ok(source.columns.slice(1).every(column => column.alignment === "right"));
  const pdf = await inspectPdf(await exportCanopy(input, "pdf"));
  const runs = pdf.pages.flatMap(page => [...page.content.matchAll(/BT \/(F\d+) [\s\S]*?([-\d.]+) ([-\d.]+) Tm <([a-f0-9]*)> Tj ET/g)]
    .map(match => ({ font: match[1], x: Number(match[2]), text: Buffer.from(match[4], "hex").toString("latin1") })));
  assert.notEqual(runs.find(run => run.text === "South").font, runs.find(run => run.text === "C1").font,
    "Parent and member labels use distinct emphasis in the actual PDF");
  const large = runs.find(run => run.text === "18,60%"), small = runs.find(run => run.text === "9,30%");
  assert.ok(large && small, "The native percent format pads both decimal places with the captured comma separator");
  assert.ok(small.x > large.x, "A shorter number starts farther right in the same numeric column");
  assert.doesNotMatch(pdf.text, /(?:^|\n)(?:18,6|9,3)%\n/);
});

test("pivot PDF honors explicit grouped number formats without changing captured raw values", async () => {
  const snapshot = structuredClone(fixture.snapshot);
  snapshot.rows[0].cells.v0_1.value = 1010; snapshot.rows[0].cells.v0_1.text = "1010";
  const input = capture(snapshot), pdf = await inspectPdf(await exportCanopy(input, "pdf", { columnOptions: { v0_1: { format: "#,##0" } } }));
  assert.ok(pdf.text.includes("1,010"));
  assert.equal(snapshot.rows[0].cells.v0_1.value, 1010);
});

test("pivot PDF uses the percentage culture for every numeric column in the document", async () => {
  for (const [share, amount, expectedShare, expectedAmount, expectedCount] of [
    ["18,6%", "1234.50", "18,60%", "1234,50", "1,00"],
    ["18.6%", "1234,50", "18.60%", "1234.50", "1.00"]
  ]) {
    const snapshot = structuredClone(fixture.percent);
    for (const row of snapshot.rows) for (const cell of Object.values(row.cells)) {
      if (cell.mode?.startsWith("percent")) cell.text = cell.text.replace(/[.,]/g, share.includes(",") ? "," : ".");
    }
    snapshot.rows[0].cells.v0_0 = { value: 1234.5, text: amount, mode: "value", numberFormat: "0.00" };
    snapshot.rows[0].cells.v0_1 = { value: 1, text: "1", mode: "value", numberFormat: "0.00" };
    snapshot.rows[0].cells.v1_0 = { value: .186, text: share, mode: "percentGrandTotal", numberFormat: "0.00%" };
    const pdf = await inspectPdf(await exportCanopy(capture(snapshot), "pdf")), lines = pdf.pages.flatMap(page => page.lines);
    for (const expected of [expectedAmount, expectedShare, expectedCount]) assert.ok(lines.includes(expected), expected);
    const numeric = lines.filter(line => /^[+-]?\d+[.,]\d+%?$/.test(line));
    assert.ok(numeric.length > 3);
    assert.equal(new Set(numeric.map(line => /[.,]/.exec(line)[0])).size, 1,
      "Amounts, declared decimal counts and percentages use one decimal separator");
  }
});

test("pivot PDF keeps undeclared integer counts whole and preserves explicit decimal formats", async () => {
  const snapshot = structuredClone(fixture.percent);
  for (const row of snapshot.rows) for (const [id, cell] of Object.entries(row.cells)) {
    if (id.endsWith("_1")) { delete cell.numberFormat; cell.text = cell.value === null ? "" : cell.value + ",00"; }
    else if (cell.mode?.startsWith("percent")) cell.text = cell.text.replace(".", ",");
  }
  snapshot.rows[0].cells.v0_1 = { value: 731, text: "731,00", mode: "value" };
  snapshot.rows[0].cells.v4_1 = { value: 947, text: "947", mode: "value", numberFormat: "0.00" };
  const input = capture(snapshot), pdf = await inspectPdf(await exportCanopy(input, "pdf")), lines = pdf.pages.flatMap(page => page.lines);
  assert.ok(lines.includes("731")); assert.ok(lines.includes("947,00"));
  assert.ok(!lines.includes("731,00") && !lines.includes("731.00"));
  const zip = await readZip(await exportCanopy(input, "xlsx"));
  assert.match(zip.get("xl/worksheets/sheet1.xml").content, /<v>731<\/v>/);
  assert.match(zip.get("xl/styles.xml").content, /formatCode="0\.00"/);
});
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

test("empty pivot intersections need no numeric display mode or format", async () => {
  const snapshot = structuredClone(fixture.snapshot), empty = snapshot.rows[0].cells.v0_0;
  empty.value = null; empty.text = ""; delete empty.mode; delete empty.numberFormat;
  for (const format of ["xlsx", "pdf", "csv"]) {
    const blob = await exportCanopy(capture(snapshot), format);
    assert.ok(blob.size > 0, format);
  }
});

test("an empty explicit format clears a pivot percent format to General without changing values", async () => {
  const snapshot = structuredClone(fixture.percent); snapshot.rows[0].tone = "neutral";
  const zip = await readZip(await exportCanopy(capture(snapshot), "xlsx", {
    columnOptions: { v0_0: { format: "" } }, tones: { neutral: { bold: true } }
  }));
  const xml = zip.get("xl/worksheets/sheet1.xml").content, styles = zip.get("xl/styles.xml").content;
  const cell = /<c r="B2" s="(\d+)"[^>]*><v>([^<]+)<\/v>/.exec(xml);
  assert.ok(cell); assert.equal(Number(cell[2]), snapshot.rows[0].cells.v0_0.value);
  const xfs = [.../<cellXfs[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)[1].matchAll(/<xf\b[^>]*numFmtId="(\d+)"/g)];
  assert.equal(xfs[Number(cell[1])][1], "0");
});

test("pivot XLSX rejects value writers before enumerating captured rows", async () => {
  const input = capture(), rows = input.rows; let read = false, calls = 0;
  input.rows = async function* () { read = true; yield* rows(); };
  await assert.rejects(exportCanopy(input, "xlsx", { xlsx: { cellValueWriters: { number: () => { calls++; return 999; } } } }), /captured pivot.*cellValueWriters/i);
  assert.equal(read, false); assert.equal(calls, 0);
});

test("pivot PDF rejects recomputed footer totals before enumerating captured rows", async () => {
  const input = capture(), rows = input.rows; let read = false;
  input.rows = async function* () { read = true; yield* rows(); };
  await assert.rejects(exportCanopy(input, "pdf", { pdf: { footer: { totals: { v0_0: "sum" } } } }), /captured pivot.*totals/i);
  assert.equal(read, false);
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
  rows[0].cells.v0_0.numberFormat = undefined; rows[0].cells.v0_0.mode = "percentGrandTotal";
  bad.rows = async function* () { yield* rows; };
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


