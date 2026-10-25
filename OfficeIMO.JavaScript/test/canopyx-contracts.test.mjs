import test from "node:test";
import assert from "node:assert/strict";
import { createCanopyExport, exportCanopy } from "../dist/integrations/canopyx/index.js";
import { readZip } from "./zip-reader.mjs";

const column = (id, kind = "number") => ({ id, title: id, kind, width: { unit: "css-px", preferred: 160, minimum: 40 }, wrap: false, alignment: kind === "number" ? "right" : "left" });
const record = (id, values) => ({ id, cells: Object.fromEntries(Object.entries(values).map(([key, value]) => [key, { value, text: value === null ? "" : String(value) }])) });
function capture(columns, rows, request = {}) {
  return { request: { columns, recordCount: rows.length, presentation: "semantic", values: "raw", timeZone: "utc", revision: "contract-fixture", ...request },
    async *rows({ signal } = {}) { for (const row of rows) { signal?.throwIfAborted(); yield row; } } };
}
const predicate = (column, operator, value) => ({ operator: "and", filters: [{ column, operator, values: [value] }], groups: [] });
const highlight = (id, column, when, extra = {}) => ({ id, column, target: "cell", when, style: { tone: "warning" }, ...extra });
const decode = text => text.replace(/&(amp|lt|gt|quot|apos);/g, (_, key) => ({ amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" })[key]);
async function worksheet(input, options = {}) {
  const parts = await readZip(await exportCanopy(input, "xlsx", { ...options, xlsx: { compression: "store", ...options.xlsx } }));
  return { parts, xml: parts.get("xl/worksheets/sheet1.xml").content };
}
function nativeRules(xml) {
  return [...xml.matchAll(/<conditionalFormatting\b[^>]*sqref="([^"]+)"[^>]*>([\s\S]*?)<\/conditionalFormatting>/g)].flatMap(block =>
    [...block[2].matchAll(/<cfRule\b([^>]*)>([\s\S]*?)<\/cfRule>/g)].map(rule => ({ range: block[1], attributes: rule[1],
      formula: decode(/<formula>([\s\S]*?)<\/formula>/.exec(rule[2])?.[1] ?? "") })));
}

test("native predicates preserve null guards, numeric operand kinds and typed UTC date serials", async () => {
  const columns = [column("amount"), column("due", "datetime"), column("mixed", "text")];
  const rows = [record("null", { amount: null, due: null, mixed: null }),
    record("zero", { amount: 0, due: "2026-11-01T09:00:00Z", mixed: false }),
    record("one", { amount: 1, due: "2026-11-01T10:00:00+01:00", mixed: "false" })];
  const formatRules = [highlight("not-one", "amount", predicate("amount", "ne", 1)),
    highlight("numeric-string", "amount", predicate("amount", "eq", "1.0")),
    highlight("same-instant", "due", predicate("due", "eq", "2026-11-01T09:00:00Z")),
    highlight("boolean-kind", "mixed", predicate("mixed", "eq", false)),
    highlight("boolean-ne", "mixed", predicate("mixed", "ne", false)),
    highlight("or-branches", "amount", { operator: "or", filters: [{ column: "amount", operator: "ne", values: [1] }, { column: "mixed", operator: "isNotNull", values: [] }], groups: [] })];
  const { xml } = await worksheet(capture(columns, rows, { formatRules }), { datetime: "typed", xlsx: { dateMode: "utc" } });
  const rules = nativeRules(xml);
  assert.equal(rules.length, 6);
  for (const rule of rules) assert.ok(rule.formula.includes("NOT(ISBLANK(INDEX("), "Null comparisons must remain false");
  assert.ok(rules[0].formula.includes("ISNUMBER(INDEX($A:$A,ROW()))"));
  assert.ok(rules[0].formula.includes("INDEX($A:$A,ROW())<>1"));
  assert.ok(rules[1].formula.includes("INDEX($A:$A,ROW())=1"));
  assert.ok(!rules[1].formula.includes('"1.0"'), "The numeric column's operand must not remain a text literal");
  assert.ok(rules[2].formula.includes("INDEX($B:$B,ROW())=46327.375"));
  assert.ok(!rules[2].formula.includes("2026-11-01"), "Typed native date values need a serial operand");
  for (const address of ["B3", "B4"]) assert.match(xml, new RegExp('<c r="' + address + '"[^>]*><v>46327\\.375<\\/v>'));
  assert.ok(rules[3].formula.includes("ISLOGICAL(INDEX($C:$C,ROW()))"));
  assert.ok(rules[4].formula.includes("OR(NOT(ISLOGICAL(INDEX($C:$C,ROW())))"));
  assert.ok(/^(?:OR\(|AND\(OR\()/.test(rules[5].formula), "Null guards remain inside the OR branches before any first-tone mask");
  assert.match(xml, /<c r="C3"[^>]*t="b"[^>]*><v>0<\/v>/);
  assert.match(xml, /<c r="C4"[^>]*t="inlineStr"[^>]*><is><t[^>]*>false<\/t>/);
  // This qualifies the emitted native formula/value contract, not execution by Excel.
});

test("unrepresentable predicate profiles report explicit loss instead of emitting a different comparison", async () => {
  const columns = [column("due", "datetime")], rows = [record("date", { due: "2026-11-01T09:00:00Z" })];
  const input = capture(columns, rows, { formatRules: [highlight("date", "due", predicate("due", "eq", "2026-11-01T09:00:00Z"))] });
  for (const options of [{}, { datetime: "text" }, { datetime: "typed", xlsx: { dateMode: "local" } }])
    await assert.rejects(exportCanopy(input, "xlsx", options), /DATETIME_RULE_REQUIRES_TYPED_UTC/);
  const submillisecond = capture(columns, rows, { formatRules: [highlight("date", "due", predicate("due", "eq", 0.5))] });
  await assert.rejects(exportCanopy(submillisecond, "xlsx", { datetime: "typed" }), /UNSUPPORTED_DATETIME_RULE_LITERAL/);
  const diagnostics = [];
  const { xml } = await worksheet(input, { unsupportedPresentation: "text", onDiagnostic: value => diagnostics.push(value.code) });
  assert.deepEqual(diagnostics, ["DATETIME_RULE_REQUIRES_TYPED_UTC:date"]); assert.equal(nativeRules(xml).length, 0);
  const numeric = capture([column("amount")], [record("number", { amount: 1000 })], {
    formatRules: [highlight("separators", "amount", predicate("amount", "eq", "1'000"))] });
  await assert.rejects(exportCanopy(numeric, "xlsx"), /UNSUPPORTED_NUMERIC_RULE_LITERAL/);
  const text = capture([column("name", "text")], [record("text", { name: "Ada" })], {
    formatRules: [highlight("name", "name", predicate("name", "eq", "ADA"))] });
  await assert.rejects(exportCanopy(text, "xlsx"), /UNQUALIFIED_NATIVE_TEXT_COMPARISON/);
});

test("native predicate profiles diagnose incompatible raw cells without changing flat capture values", async () => {
  const epoch = Date.parse("2026-11-01T09:00:00Z");
  for (const scenario of [
    { column: column("amount"), value: "1", operand: 1, type: "number" },
    { column: column("due", "datetime"), value: epoch, operand: "2026-11-01T09:00:00Z", type: "date" }
  ]) {
    const rows = [record("raw", { [scenario.column.id]: scenario.value })], options = { datetime: "typed", xlsx: { dateMode: "utc" } };
    const input = capture([scenario.column], rows), flat = createCanopyExport(input, "xlsx", options);
    const projected = []; for await (const row of flat.rows()) projected.push(row);
    assert.equal(projected[0][0].value, scenario.value, "Native predicate restrictions must not coerce the flat source");
    const styled = capture([scenario.column], rows, { formatRules: [highlight("raw-type", scenario.column.id, predicate(scenario.column.id, "eq", scenario.operand))] });
    const expected = new RegExp("UNSUPPORTED_NATIVE_RULE_CELL_TYPE:" + scenario.type);
    await assert.rejects(exportCanopy(styled, "xlsx", options), expected);
    const diagnostics = [];
    await assert.rejects(exportCanopy(styled, "xlsx", { ...options, unsupportedPresentation: "text", onDiagnostic: value => diagnostics.push(value.code) }), expected);
    assert.deepEqual(diagnostics, ["UNSUPPORTED_NATIVE_RULE_CELL_TYPE:" + scenario.type]);
  }
  const columns = [column("mixed", "text")], rows = [record("number", { mixed: 1 }), record("boolean", { mixed: true }), record("string", { mixed: "1" }), record("null", { mixed: null })];
  const { xml } = await worksheet(capture(columns, rows, { formatRules: [highlight("number-kind", "mixed", predicate("mixed", "eq", 1)), highlight("boolean-kind", "mixed", predicate("mixed", "eq", true))] }));
  const rules = nativeRules(xml);
  assert.ok(rules[0].formula.includes("ISNUMBER(INDEX($A:$A,ROW()))"));
  assert.ok(rules[1].formula.includes("ISLOGICAL(INDEX($A:$A,ROW()))"));
  assert.match(xml, /<c r="A2"[^>]*><v>1<\/v>/);
  assert.match(xml, /<c r="A3"[^>]*t="b"[^>]*><v>1<\/v>/);
  assert.match(xml, /<c r="A4"[^>]*t="inlineStr"[^>]*><is><t[^>]*>1<\/t>/);
});

function groupedCapture(values, operation, expected) {
  const rows = values.map((value, index) => record("detail-" + index, { region: "A", amount: value }));
  return capture([column("region", "text"), column("amount")], rows, { grouping: {
    levels: [{ column: "region", direction: "asc", title: "Region" }],
    aggregates: [{ id: "metric", column: "amount", operation, label: "Metric" }], totals: { metric: expected },
    groups: [{ key: '["string","A"]', value: "A", text: "A", offset: 0, count: rows.length, aggregates: { metric: expected } }]
  } });
}
test("producer-wire aggregates are verified exactly while native caches retain their own arithmetic", async () => {
  for (const scenario of [
    { operation: "sum", values: [0.1, 0.2, 0.3], expected: 0.6, native: 0.6000000000000001 },
    { operation: "average", values: [0.1, 0.2, 0.3], expected: 0.6 / 3, native: 0.2 },
    { operation: "sum", values: [null, null], expected: null, native: 0 },
    { operation: "count", values: [null, 2, 1, null], expected: 2, native: 2 },
    { operation: "minimum", values: [null, 2, 1, null], expected: 1, native: 1 },
    { operation: "maximum", values: [null, 2, 1, null], expected: 2, native: 2 }
  ]) {
    const { xml } = await worksheet(groupedCapture(scenario.values, scenario.operation, scenario.expected));
    const caches = [...xml.matchAll(/<f>[\s\S]*?<\/f><v>([^<]+)<\/v>/g)].map(match => Number(match[1]));
    assert.deepEqual(caches, [scenario.native, scenario.native], scenario.operation + " native subtotal/grand caches");
  }
  for (const part of ["group", "grand"]) {
    const input = groupedCapture([0.1, 0.2, 0.3], "sum", 0.6);
    if (part === "group") input.request.grouping.groups[0].aggregates.metric = 1.6;
    else input.request.grouping.totals.metric = 1.6;
    await assert.rejects(exportCanopy(input, "xlsx"), /captured aggregate differs/);
  }
  const invalidRange = groupedCapture([0.1, 0.2, 0.3], "sum", 0.6);
  invalidRange.request.grouping.groups[0].count = 2;
  await assert.rejects(exportCanopy(invalidRange, "xlsx"), /complete exported detail range/);
  const textCount = groupedCapture([1, 2], "count", 2);
  textCount.request.columns[1].kind = "text";
  await assert.rejects(exportCanopy(textCount, "xlsx"), /raw numeric column/);
});

test("cell stop masks later columns and mixed native visuals reject without a partial faithful profile", async () => {
  const columns = [column("a"), column("b")], rows = [record("stopped", { a: 1, b: 2 })];
  const stop = highlight("stop-a", "a", predicate("a", "gt", 0), { stopIfTrue: true });
  const later = highlight("later-b", "b", predicate("b", "gt", 0), { style: { tone: "info" } });
  const { xml } = await worksheet(capture(columns, rows, { formatRules: [stop, later] }));
  const rules = nativeRules(xml);
  assert.deepEqual(rules.map(rule => rule.range), ["A2:A2", "B2:B2"]);
  assert.ok(rules.every(rule => !rule.attributes.includes("stopIfTrue")));
  assert.ok(rules[1].formula.includes("NOT(OR(" + rules[0].formula + "))"));
  const disabled = { ...stop, enabled: false };
  const withoutStop = nativeRules((await worksheet(capture(columns, rows, { formatRules: [disabled, later] }))).xml);
  assert.equal(withoutStop.length, 1); assert.ok(!withoutStop[0].formula.includes("NOT(OR("));
  const mixed = capture(columns, rows, { formatRules: [stop, { id: "bar-b", column: "b", style: { bar: { minimum: 0, maximum: 10, tone: "info" } } }] });
  await assert.rejects(exportCanopy(mixed, "xlsx"), /UNSUPPORTED_GLOBAL_STOP_WITH_NATIVE_VISUALS/);
  const diagnostics = [];
  const lossy = await worksheet(mixed, { unsupportedPresentation: "text", onDiagnostic: value => diagnostics.push(value.code) });
  assert.deepEqual(diagnostics, ["UNSUPPORTED_GLOBAL_STOP_WITH_NATIVE_VISUALS"]);
  assert.equal(nativeRules(lossy.xml).length, 0);
});

test("native rule ranges follow zero, one and multiple data headers through flat, About and grouped routes", async () => {
  for (const grouped of [false, true]) for (const about of [false, true]) for (const headers of [0, 1, 2]) {
    const columns = [{ ...column("region", "text"), pinned: "left" }, column("name", "text"), column("amount")];
    const rows = [record("one", { region: "A", name: "detail-one", amount: 1 }), record("two", { region: "A", name: "detail-two", amount: 2 })];
    const formatRules = [{ id: "bar", column: "amount", style: { bar: { minimum: 0, maximum: 10, tone: "info" } } },
      { id: "row", target: "row", when: predicate("amount", "gt", 0), style: { tone: "warning" } }];
    const grouping = grouped ? { levels: [{ column: "region", direction: "asc", title: "Region" }],
      aggregates: [{ id: "metric", column: "amount", operation: "sum", label: "Metric" }], totals: { metric: 3 },
      groups: [{ key: '["string","A"]', value: "A", text: "A", offset: 0, count: 2, aggregates: { metric: 3 } }] } : undefined;
    const input = capture(columns, rows, { formatRules, ...(grouping === undefined ? {} : { grouping }) });
    const sheet = headers === 0 ? { includeHeader: false } : headers === 2 && !grouped ? {
      headerRows: [[{ value: "Captured", columnSpan: 3 }, null, null], [{ value: "region" }, { value: "name" }, { value: "amount" }]]
    } : {};
    const columnOptions = headers === 2 && grouped ? Object.fromEntries(columns.map(column => [column.id, { groups: ["Captured"] }])) : undefined;
    const { parts, xml } = await worksheet(input, { about, ...(columnOptions === undefined ? {} : { columnOptions }), xlsx: { sheet } });
    const details = [...xml.matchAll(/<row r="(\d+)"[^>]*>([\s\S]*?)<\/row>/g)].filter(row => /detail-(?:one|two)/.test(row[2])).map(row => Number(row[1]));
    const first = headers + (grouped ? 2 : 1), last = first + 1, amount = grouped ? "D" : "C", firstColumn = grouped ? "B" : "A";
    assert.deepEqual(details, [first, last]);
    const rules = nativeRules(xml);
    assert.deepEqual(rules.map(rule => rule.range), [amount + first + ":" + amount + last, firstColumn + first + ":" + amount + last]);
    assert.ok(rules[1].formula.includes("INDEX($" + amount + ":$" + amount + ",ROW())>0"));
    assert.match(xml, new RegExp('xSplit="' + (grouped ? 2 : 1) + '"'));
    if (headers) assert.match(xml, new RegExp('ySplit="' + headers + '"'));
    else assert.ok(!/<pane[^>]*ySplit=/.test(xml));
    assert.equal(parts.has("xl/worksheets/sheet2.xml"), about);
  }
  const explicit = capture([column("amount")], [record("one", { amount: 1 })]);
  await assert.rejects(exportCanopy(explicit, "xlsx", { xlsx: { sheet: { includeHeader: false, freezeHeader: true } } }), /requires a header row/);
});

test("a data sheet named About retains its detail progress while actual metadata rows remain excluded", async () => {
  const input = capture([column("name", "text")], [record("one", { name: "detail-one" })]);
  const named = [];
  await exportCanopy(input, "xlsx", { xlsx: { sheet: { name: "About" } }, onProgress: value => named.push(value) });
  const namedRows = named.filter(value => value.phase === "rows" && value.sheetName === "About");
  assert.ok(namedRows.length > 0); assert.equal(namedRows.at(-1).sheetRows, 1);
  assert.equal(named.at(-1).rows, 1); assert.equal(named.at(-1).totalRows, 1);
  const metadata = [];
  await exportCanopy(input, "xlsx", { about: true, onProgress: value => metadata.push(value) });
  const aboutRows = metadata.filter(value => value.phase === "rows" && value.sheetName === "About");
  assert.ok(aboutRows.length > 0); assert.ok(aboutRows.every(value => value.sheetRows === 0));
  assert.equal(metadata.at(-1).rows, 1); assert.equal(metadata.at(-1).totalRows, 1);
  assert.ok(metadata.filter(value => value.sheetName === "Data").some(value => value.sheetRows === 1));
});

test("native tone targets preserve cell precedence, first tone and declaration-order stops", async () => {
  const columns = [column("a"), column("b"), column("region", "text")], rows = [record("precedence", { a: 1, b: 2, region: "A" })];
  const active = predicate("a", "gt", 0);
  const rowWarning = { id: "row-warning", target: "row", when: active, style: { tone: "warning" } };
  const rowInfo = { id: "row-info", target: "row", when: active, style: { tone: "info" } };
  const cellWarning = highlight("cell-warning", "a", active), cellInfo = highlight("cell-info", "a", active, { style: { tone: "info" } });
  const scenarios = [
    { name: "row-before-cell", rules: [rowWarning, cellInfo], targets: ["cell", "row"], colors: ["DDEBF7", "FFF2CC"] },
    { name: "cell-before-row", rules: [cellInfo, rowWarning], targets: ["cell", "row"], colors: ["DDEBF7", "FFF2CC"] },
    { name: "first-row-tone", rules: [rowWarning, rowInfo], targets: ["row", "row"], colors: ["FFF2CC", "DDEBF7"], firstTone: true },
    { name: "first-cell-tone", rules: [cellWarning, cellInfo], targets: ["cell", "cell"], colors: ["FFF2CC", "DDEBF7"], firstTone: true },
    { name: "row-stops-later-cell", rules: [{ ...rowWarning, stopIfTrue: true }, cellInfo], targets: ["cell", "row"], colors: ["DDEBF7", "FFF2CC"], stop: true },
    { name: "ignored-tone-still-stops", rules: [rowWarning, { ...rowInfo, stopIfTrue: true }, cellInfo], targets: ["cell", "row", "row"], colors: ["DDEBF7", "FFF2CC", "DDEBF7"], ignoredStop: true }
  ];
  for (const grouped of [false, true]) for (const about of [false, true]) for (const scenario of scenarios) {
    const grouping = grouped ? { levels: [{ column: "region", direction: "asc", title: "Region" }],
      aggregates: [{ id: "metric", column: "b", operation: "sum", label: "Metric" }], totals: { metric: 2 },
      groups: [{ key: '["string","A"]', value: "A", text: "A", offset: 0, count: 1, aggregates: { metric: 2 } }] } : undefined;
    const { parts, xml } = await worksheet(capture(columns, rows, { formatRules: scenario.rules, ...(grouping === undefined ? {} : { grouping }) }), { about });
    const native = nativeRules(xml), body = grouped ? 3 : 2, cellRange = (grouped ? "B" : "A") + body + ":" + (grouped ? "B" : "A") + body;
    const rowRange = (grouped ? "B" : "A") + body + ":" + (grouped ? "D" : "C") + body;
    assert.deepEqual(native.map(rule => rule.range), scenario.targets.map(target => target === "cell" ? cellRange : rowRange), scenario.name);
    const differentials = [...parts.get("xl/styles.xml").content.matchAll(/<dxf\b[^>]*>([\s\S]*?)<\/dxf>/g)].map(match => match[1]);
    for (let index = 0; index < native.length; index++) {
      assert.ok(native[index].attributes.includes('priority="' + (index + 1) + '"'));
      const style = differentials[Number(/dxfId="(\d+)"/.exec(native[index].attributes)[1])];
      assert.ok(style.includes(scenario.colors[index]), scenario.name + " tone style at priority " + (index + 1));
      assert.ok(!native[index].attributes.includes("stopIfTrue"));
    }
    if (scenario.firstTone) assert.ok(native[1].formula.includes("NOT(OR(" + native[0].formula + "))"));
    if (scenario.stop || scenario.ignoredStop) assert.ok(native[0].formula.includes("NOT(OR(" + native[1].formula + "))"));
    if (scenario.ignoredStop) {
      assert.ok(native[2].formula.includes("NOT(OR(" + native[1].formula + "))"));
      assert.equal(native[0].formula, native[2].formula, "The stop uses the matched declaration, not its suppressed visual effect");
    }
  }
  // First-tone selection is atomic even when explicit tone mappings supply different native properties.
  const sparse = await worksheet(capture(columns, rows, { formatRules: [cellWarning, cellInfo] }), {
    tones: { warning: { color: "9C6500" }, info: { background: "DDEBF7" } }
  });
  const sparseRules = nativeRules(sparse.xml);
  assert.ok(sparseRules[1].formula.includes("NOT(OR(" + sparseRules[0].formula + "))"));
  const bar = { id: "bar-b", column: "b", style: { bar: { minimum: 0, maximum: 10, tone: "info" } } };
  const scale = { id: "scale-a", column: "a", style: { scale: { stops: [{ type: "number", value: 0, tone: "success" }, { type: "number", value: 10, tone: "danger" }] } } };
  const visual = nativeRules((await worksheet(capture(columns, rows, { formatRules: [rowWarning, bar, scale] }))).xml);
  assert.deepEqual(visual.map(rule => /type="([^"]+)"/.exec(rule.attributes)[1]), ["dataBar", "colorScale", "expression"]);
  const scaleTone = nativeRules((await worksheet(capture(columns, rows, { formatRules: [cellInfo, scale, rowWarning] }))).xml);
  assert.deepEqual(scaleTone.map(rule => /type="([^"]+)"/.exec(rule.attributes)[1]), ["colorScale", "expression", "expression"]);
  assert.deepEqual(scaleTone.map(rule => rule.range), ["A2:A2", "A2:A2", "A2:C2"]);
  const mixed = capture(columns, rows, { formatRules: [highlight("cell-b", "b", active), bar, rowWarning] });
  await assert.rejects(exportCanopy(mixed, "xlsx"), /UNSUPPORTED_CELL_TONE_WITH_DATA_BAR/);
  const diagnostics = [];
  const omitted = await worksheet(mixed, { unsupportedPresentation: "text", onDiagnostic: value => diagnostics.push(value.code) });
  assert.deepEqual(diagnostics, ["UNSUPPORTED_CELL_TONE_WITH_DATA_BAR"]); assert.equal(nativeRules(omitted.xml).length, 0);
});
