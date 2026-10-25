// Synthetic contract-shaped regression data; independent consumer captures are qualified separately.
import test from "node:test";
import assert from "node:assert/strict";
import { createCanopyExport, exportCanopy, writeCanopyTo } from "../dist/integrations/canopyx/index.js";
import { readZip } from "./zip-reader.mjs";
import { inspectPdf } from "./pdf-reader.mjs";

const columns = [
  { id: "name", title: "Resource", kind: "text", width: { unit: "css-px", preferred: 640, minimum: 40 }, pinned: "left", wrap: false, alignment: "left" },
  { id: "region", title: "Region", kind: "text", width: { unit: "css-px", preferred: 296, minimum: 40 }, wrap: false, alignment: "left" },
  { id: "cost", title: "Monthly cost", kind: "number", width: { unit: "css-px", preferred: 150, minimum: 40 }, wrap: false, alignment: "right", format: { minimumFractionDigits: 2, maximumFractionDigits: 2, prefix: "$ " } },
  { id: "due", title: "Renewal", kind: "datetime", width: { unit: "css-px", preferred: 200, minimum: 40 }, wrap: false, alignment: "left" },
  { id: "usage", title: "Usage %", kind: "number", width: { unit: "css-px", preferred: 120, minimum: 40 }, wrap: false, alignment: "right" }
];
const formatRules = [
  { id: "usage-bar", column: "usage", style: { bar: { minimum: 0, maximum: 100, tone: "info" } } },
  { id: "cost-scale", column: "cost", style: { scale: { stops: [{ type: "min", tone: "success" }, { type: "percent", value: 50, tone: "warning" }, { type: "max", tone: "danger" }] } } },
  { id: "high-usage", when: { operator: "and", filters: [{ column: "usage", operator: "gt", values: [90] }], groups: [] }, target: "row", style: { tone: "warning" } }
];
const cell = (value, text = String(value), extra = {}) => ({ value, text, ...extra });
const records = Array.from({ length: 60 }, (_, index) => ({ id: "synthetic-resource-" + index, cells: {
  name: cell("Synthetic resource " + String(index).padStart(2, "0"), undefined, { tone: "neutral", link: { target: "https://example.com/synthetic/" + index } }),
  region: cell(index % 2 ? "South" : "North"), cost: cell(index + 0.25, "$ " + (index + 0.25).toFixed(2)),
  due: cell("2026-11-01T09:00:00Z", "2026-11-01 09:00 UTC"), usage: cell((index * 7 + 17) % 101)
} }));
const filtered = records.filter(row => row.cells.region.value === "North");
const all = [...filtered, ...records.filter(row => row.cells.region.value === "South")];
const scoped = { filtered, selected: [records[0], records[1]], all };
const syntheticTotal = 877.5; // Known sum for this fixture: 0.25 + 2.25 + ... + 58.25.
function capture(scope) {
  const rows = scoped[scope];
  return { request: {
    scope, columns, revision: "synthetic-v1", recordCount: rows.length, total: 60, matched: 30,
    selection: scope === "selected" ? rows.map(row => row.id) : [], presentation: "semantic", values: "raw", timeZone: "utc",
    query: { protocol: "canopyx/2", type: "gridQuery", view: "synthetic", search: scope === "filtered" ? "North" : "", filters: [], groupBy: [{ column: "region", direction: "asc" }] },
    formatRules, generatedAt: "2026-10-25T09:26:55.565Z",
    ...(scope === "filtered" ? { grouping: { levels: [{ column: "region", direction: "asc", title: "Region" }],
      aggregates: [{ id: "cost-total", column: "cost", operation: "sum", label: "Monthly cost" }], totals: { "cost-total": syntheticTotal },
      groups: [{ key: '["string","North"]', value: "North", text: "North", offset: 0, count: 30, aggregates: { "cost-total": syntheticTotal } }] } } : {})
  }, async *rows({ signal } = {}) { for (const row of rows) { signal?.throwIfAborted(); yield row; } } };
}
const rowXml = xml => [...xml.matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)].map(match => ({ row: Number(match[1]), xml: match[2] }));

test("captured scopes produce all 92 synthetic details, typed values, widths, pins, links and native rules", async () => {
  let exported = 0;
  for (const scope of ["filtered", "selected", "all"]) {
    const input = capture(scope), progress = [];
    const blob = await exportCanopy(input, "xlsx", { about: true, onProgress: value => progress.push(value), xlsx: { compression: "store" } });
    const parts = await readZip(blob), xml = parts.get("xl/worksheets/sheet1.xml").content;
    const details = rowXml(xml).filter(row => row.xml.includes("Synthetic resource"));
    const expected = scoped[scope];
    assert.deepEqual(details.map(row => /<t[^>]*>(Synthetic resource \d+)<\/t>/.exec(row.xml)?.[1]), expected.map(row => row.cells.name.value));
    const grouped = scope === "filtered", letters = grouped ? { cost: "D", due: "E", usage: "F" } : { cost: "C", due: "D", usage: "E" };
    for (let index = 0; index < details.length; index++) {
      const row = details[index], record = expected[index];
      assert.match(row.xml, new RegExp('<c r="' + letters.cost + row.row + '"[^>]*><v>' + String(record.cells.cost.value).replace('.', '\\.') + '<\\/v>'));
      assert.match(row.xml, new RegExp('<c r="' + letters.usage + row.row + '"[^>]*><v>' + record.cells.usage.value + '<\\/v>'));
      assert.match(row.xml, new RegExp('<c r="' + letters.due + row.row + '"[^>]*><v>\\d+(?:\\.\\d+)?<\\/v>'));
    }
    assert.match(xml, new RegExp('xSplit="' + (grouped ? 2 : 1) + '" ySplit="1"'));
    for (const width of [90.7109375, 41.5703125, 20.7109375, 27.85546875, 16.42578125]) assert.ok(xml.includes('width="' + width + '"'));
    assert.match(parts.get("xl/styles.xml").content, /formatCode="&quot;\$ &quot;0\.00"/);
    const first = details[0].row, last = details.at(-1).row;
    assert.match(xml, new RegExp('sqref="' + letters.usage + first + ':' + letters.usage + last + '"'));
    assert.match(xml, new RegExp('sqref="' + letters.cost + first + ':' + letters.cost + last + '"'));
    assert.match(xml, /type="dataBar" priority="1"/); assert.match(xml, /type="colorScale" priority="2"/);
    assert.match(xml, /type="expression"[^>]*priority="3"/);
    assert.ok(xml.includes('INDEX($' + letters.usage + ':$' + letters.usage + ',ROW())&gt;90'));
    const relationships = parts.get("xl/worksheets/_rels/sheet1.xml.rels").content;
    assert.equal((relationships.match(/TargetMode="External"/g) ?? []).length, expected.length);
    for (const record of expected) assert.ok(relationships.includes(record.cells.name.link.target));
    if (grouped) {
      assert.match(xml, /outlineLevel="1"/);
      assert.equal((xml.match(/SUBTOTAL\(9,D3:D32\)<\/f><v>877\.5<\/v>/g) ?? []).length, 2);
      assert.ok(!/sqref="[^"]*(?:D33|D34|F33|F34)/.test(xml));
    } else assert.ok(!xml.includes("SUBTOTAL("));
    const metadata = parts.get("xl/worksheets/sheet2.xml").content;
    for (const field of ["recordCount", "total", "matched", "generatedAt", "query", "selection"]) assert.ok(metadata.includes(field));
    assert.ok(metadata.includes(scope === "filtered" ? '&quot;search&quot;:&quot;North&quot;' : '&quot;search&quot;:&quot;&quot;'));
    if (scope === "selected") for (const record of expected) assert.ok(metadata.includes(record.id));
    assert.equal(progress.at(-1).rows, expected.length); assert.equal(progress.at(-1).totalRows, expected.length);
    assert.equal(progress.at(-1).phase, "complete"); assert.ok(progress.every(value => value.rows <= expected.length));
    exported += details.length;
  }
  assert.equal(exported, 92);
});

test("PDF keeps every synthetic detail and group context while diagnosing native XLSX-only effects", async () => {
  for (const scope of ["filtered", "selected", "all"]) {
    const input = capture(scope), diagnostics = [];
    await assert.rejects(exportCanopy(input, "pdf"), /PDF_NATIVE_RULE_NOT_RENDERED/);
    const pdf = await inspectPdf(await exportCanopy(input, "pdf", { unsupportedPresentation: "text", onDiagnostic: value => diagnostics.push(value), pdf: { compression: false } }));
    assert.deepEqual(pdf.pages.flatMap(page => page.lines).filter(line => line.startsWith("Synthetic resource ")), scoped[scope].map(row => row.cells.name.value));
    assert.deepEqual(diagnostics.map(value => value.code), formatRules.map(rule => "PDF_NATIVE_RULE_NOT_RENDERED:" + rule.id));
    for (const page of pdf.pages.filter(page => page.lines.some(line => line.startsWith("Synthetic resource ")))) {
      assert.ok(page.lines.includes("Resource")); assert.match(page.body, /\/MediaBox \[0 0 841\.89 595\.276\]/);
      if (scope === "filtered") assert.ok(page.lines.includes("North"));
    }
    assert.match([...pdf.objects.values()].map(object => object.body).join("\n"), /\/URI <68747470733a/);
  }
});

test("ordinary text capture and canonical neutral link need no lossy presentation escape hatch", async () => {
  const ordinary = { request: { columns: [{ id: "name", title: "Resource", kind: "text" }], recordCount: 1, values: "raw", presentation: "text", timeZone: "utc", revision: "text-v1" },
    async *rows() { yield records[0]; } };
  for (const format of ["xlsx", "pdf"]) await exportCanopy(ordinary, format);
  const source = createCanopyExport(ordinary, "xlsx", { tones: { neutral: { color: "123456" } } });
  const rows = []; for await (const row of source.rows()) rows.push(row);
  assert.equal(rows[0][0].presentation.color, "123456"); assert.equal(rows[0][0].link.target, records[0].cells.name.link.target);
  await assert.rejects(exportCanopy({ ...ordinary, async *rows() { yield { ...records[0], tone: "unknown-tone" }; } }, "xlsx"), /UNMAPPED_TONE/);
});

test("grouped export rejects contradictory scope ranges and expected totals before a successful artifact", async () => {
  for (const change of [
    input => { input.request.grouping.groups[0].count = 29; },
    input => { input.request.grouping.groups[0].offset = 1; },
    input => { input.request.grouping.groups[0].aggregates["cost-total"] = 1; },
    input => { input.request.grouping.totals["cost-total"] = 1; }
  ]) {
    const input = capture("filtered"); change(input); await assert.rejects(exportCanopy(input, "xlsx"), /group|aggregate/i);
  }
  const input = capture("selected");
  const result = await writeCanopyTo(input, "xlsx", { write() {} });
  assert.equal(result.rows, 2); // The accepted selected query still has groupBy; no grouping snapshot is invented.
});

test("grouped capture cancellation reaches an in-flight read and closes its original iterator", async () => {
  const input = capture("filtered"), controller = new AbortController(), reason = new Error("cancelled grouped capture");
  let started, closed = false, observed;
  const waiting = new Promise(resolve => { started = resolve; });
  input.rows = async function* ({ signal } = {}) {
    observed = signal;
    try {
      yield filtered[0]; started();
      await new Promise((_, reject) => signal.addEventListener("abort", () => reject(signal.reason), { once: true }));
    } finally { closed = true; }
  };
  const pending = exportCanopy(input, "xlsx", { signal: controller.signal });
  await waiting; controller.abort(reason);
  await assert.rejects(pending, error => error === reason); assert.equal(observed, controller.signal); assert.equal(closed, true);
});
