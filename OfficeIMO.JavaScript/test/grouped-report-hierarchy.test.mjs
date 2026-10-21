import test from "node:test";
import assert from "node:assert/strict";
import { createGroupedReport, groupedReportEvents } from "../dist/core/grouped-report.js";
import { ExportCell } from "../dist/core/index.js";

const columns = [{ key: "name", header: "Certificate" }, { key: "days", header: "Days", type: "number" }];
async function collect(source) { const result = []; for await (const event of source.events) result.push(event); return result; }
const hidden = key => ({ key, value: row => row[key] });

test("hierarchical totals use selected raw records at every depth and preserve hidden grouping fields", async () => {
  const selected = [
    { site: "Łódź", issuer: "A", state: "valid", name: "one", days: 10 },
    { site: "Łódź", issuer: "A", state: "valid", name: "two", days: null },
    { site: "Łódź", issuer: "A", state: "expired", name: "three", days: 20 },
    { site: "Łódź", issuer: "B", state: "valid", name: "four", days: 60 },
    { site: "東京", issuer: "A", state: "valid", name: "five", days: 30 },
    { site: "東京", issuer: "A", state: "valid", name: "six", days: 40 },
  ];
  const source = createGroupedReport(selected, { columns, groupBy: [hidden("site"), hidden("issuer"), hidden("state")], totals: { days: "average" } });
  const events = await collect(source), ends = events.filter(event => event.kind === "groupEnd");
  assert.deepEqual(source.groupBy, ["site", "issuer", "state"]);
  assert.deepEqual(source.columns.map(column => column.key), ["name", "days"]);
  assert.deepEqual(ends.map(event => [event.path.map(group => group.key), event.rows, event.totals[1]]), [
    [["Łódź", "A", "valid"], 2, 10], [["Łódź", "A", "expired"], 1, 20], [["Łódź", "A"], 3, 15],
    [["Łódź", "B", "valid"], 1, 60], [["Łódź", "B"], 1, 60], [["Łódź"], 4, 30],
    [["東京", "A", "valid"], 2, 35], [["東京", "A"], 2, 35], [["東京"], 2, 35],
  ]);
  assert.deepEqual(events.at(-1), { kind: "complete", rows: 6, groups: 9, totals: [null, 32] });
  assert.deepEqual(events.filter(event => event.kind === "row").map(event => event.values), selected.map(row => [row.name, row.days]));
  for (const event of events.filter(event => event.kind !== "complete")) {
    assert.ok(Object.isFrozen(event)); assert.ok(Object.isFrozen(event.path)); assert.ok(event.path.every(Object.isFrozen));
    assert.equal(event.path.at(-1).groupIndex, event.groupIndex);
    if (event.kind !== "row") assert.equal(event.path.length, event.level + 1);
  }
});

test("typed child identities recur only under another parent, including delimiter-containing values", async () => {
  const pairs = [["x", 1], ["x", "1"], ["x", true], ["x", null], ["y:0:string:x", 1], ["y:0:string:x", "1"]];
  const setup = { columns: [{ key: "parent", header: "Same" }, { key: "child", header: "Same" }], groupBy: ["parent", "child"] };
  const events = await collect(createGroupedReport(pairs, setup));
  assert.equal(events.at(-1).groups, 8);
  assert.deepEqual(events.filter(event => event.kind === "groupEnd" && event.level === 1).map(event => event.key), [1, "1", true, null, 1, "1"]);
  for (const rows of [[["x", "a"], ["x", "b"], ["x", "a"]], [["x", "a"], ["y", "a"], ["x", "b"]]])
    await assert.rejects(collect(createGroupedReport(rows, setup)), /closed group key.*same parent/);
});

test("four levels close deepest-first and all aggregate operations use every body row", async () => {
  const columns = ["a", "b", "c", "d"].map(key => ({ key, header: key }));
  for (const key of ["sum", "count", "min", "max", "average"]) columns.push({ key, header: key, type: "number" });
  const totals = Object.fromEntries(["sum", "count", "min", "max", "average"].map(operation => [operation, operation]));
  const events = await collect(createGroupedReport([
    ["a", "b", "c", "d", 2, 2, 2, 2, 2], ["a", "b", "c", "d", null, null, null, null, null],
    ["a", "b", "c", "e", 8, 8, 8, 8, 8],
  ], { columns, groupBy: ["a", "b", "c", "d"], totals }));
  assert.deepEqual(events.filter(event => event.kind === "groupEnd").map(event => event.level), [3, 3, 2, 1, 0]);
  assert.deepEqual(events.at(-1), { kind: "complete", rows: 3, groups: 5, totals: [null, null, null, null, 10, 2, 2, 8, 5] });
  for (const group of events.filter(event => event.kind === "groupEnd" && event.level < 3)) {
    assert.equal(group.rows, 3); assert.deepEqual(group.totals, events.at(-1).totals);
  }
});

test("hierarchy captures selectors, labels and positional input once and reuses displayed projections", async () => {
  let columnCalls = 0, keyReads = 0, valueReads = 0, labelReads = 0, groupByReads = 0, arrayReads = 0;
  const contexts = [], descriptor = Object.create({
    get key() { keyReads++; return "hidden"; },
    get value() { valueReads++; return (row, context) => { contexts.push(context); return row[1]; }; },
    get label() { labelReads++; return value => "H:" + value; },
  });
  const input = ["A", "B", 7];
  Object.defineProperty(input, 2, { get() { arrayReads++; return 7; } });
  const groupBy = ["shown", descriptor];
  const source = createGroupedReport([input], {
    columns: [{ key: "shown", header: "Shown", value: row => { columnCalls++; return new ExportCell(row[0], { text: "translated" }); } },
      { key: "n", header: "N", type: "number", value: row => row[2] }],
    get groupBy() { groupByReads++; return groupBy; }, totals: { n: "sum" },
  });
  groupBy.reverse(); descriptor.extra = "ignored";
  const events = await collect(source);
  assert.equal(keyReads, 1); assert.equal(valueReads, 1); assert.equal(labelReads, 1); assert.equal(groupByReads, 1);
  assert.equal(columnCalls, 1); assert.equal(arrayReads, 1);
  assert.deepEqual(contexts, [{ rowIndex: 0, level: 1, key: "hidden" }]);
  assert.deepEqual(events.filter(event => event.kind === "groupStart").map(event => [event.key, event.label]), [["A", "A"], ["B", "H:B"]]);
  assert.equal(events.at(-1).totals[1], 7);
});

test("event date keys and labels cannot mutate another event's hierarchy or recurrence", async () => {
  const original = new Date("2026-10-21T00:00:00Z"), expected = original.toISOString();
  const source = createGroupedReport([{ region: "A", instant: original, name: "one", days: 1 },
    { region: "A", instant: new Date(expected), name: "two", days: 2 }], {
    columns, groupBy: [hidden("region"), { ...hidden("instant"), label(value) { value.setUTCFullYear(2050); return "captured"; } }],
  });
  const iterator = source.events[Symbol.asyncIterator]();
  await iterator.next();
  const start = (await iterator.next()).value;
  start.key.setUTCFullYear(2040); original.setUTCFullYear(2030);
  const row = (await iterator.next()).value;
  assert.equal(row.path[1].key.toISOString(), expected);
  row.path[1].key.setUTCFullYear(2045);
  const rest = []; for await (const event of { [Symbol.asyncIterator]: () => iterator }) rest.push(event);
  assert.equal(rest.filter(event => event.kind === "groupStart").length, 0);
  const ending = rest.find(event => event.kind === "groupEnd" && event.level === 1);
  assert.equal(ending.key.toISOString(), expected); assert.equal(ending.rows, 2);
});

test("hierarchy configuration rejects contradictory selectors and out-of-profile depth before reading", () => {
  let read = 0;
  function* rows() { read++; yield { name: "one", days: 1 }; }
  for (const groupBy of [[], ["name", "name"], ["name", "days", "a", "b", "c"], ["missing"],
    [{ key: "name", value: row => row.name }], [{ key: "hidden", value: 1 }], [{ key: "hidden", value: row => row.name, label: 1 }]])
    assert.throws(() => createGroupedReport(rows(), { columns, groupBy }), TypeError);
  assert.throws(() => createGroupedReport(rows(), { columns, groupBy: ["name", "days"], groupLabel: String }), /one grouping level/);
  assert.throws(() => createGroupedReport(rows(), { columns, groupBy: [{ key: "name", label: String }], groupLabel: String }), /either/);
  assert.equal(read, 0);
});

test("hierarchy bounds count every node and close the source on invalid labels or getters", async () => {
  for (const option of [{ maxGroups: 2 }, { maxGroupKeyCharacters: 0 }, { limits: { maxRows: 0 } }, { limits: { maxTextCharacters: 20 } }]) {
    let returned = 0;
    function* rows() { try { yield { a: "A", b: "B", c: "C", name: "one", days: 1 }; } finally { returned++; } }
    await assert.rejects(collect(createGroupedReport(rows(), { columns, groupBy: [hidden("a"), hidden("b"), hidden("c")], ...option })), { code: "RESOURCE_LIMIT" });
    assert.equal(returned, 1);
  }
  for (const level of [{ key: "hidden", value: () => Promise.reject(new Error("async getter")) },
    { key: "hidden", value: () => "A", label: () => Promise.reject(new Error("async label")) },
    { key: "hidden", value: () => new ExportCell("A") }]) {
    let returned = 0;
    function* rows() { try { yield { name: "one", days: 1 }; } finally { returned++; } }
    await assert.rejects(collect(createGroupedReport(rows(), { columns, groupBy: [level] })), /synchronous/);
    assert.equal(returned, 1);
  }
});

test("hierarchy retains one boundary row, returns on disposal and propagates writer cancellation into a pending read", { timeout: 2000 }, async () => {
  let read = 0, returned = 0;
  function* rows() { try { for (const row of [["A", "x"], ["A", "y"], ["B", "x"]]) { read++; yield row; } } finally { returned++; } }
  const setup = { columns: [{ key: "a", header: "A" }, { key: "b", header: "B" }], groupBy: ["a", "b"] };
  const iterator = createGroupedReport(rows(), setup).events[Symbol.asyncIterator]();
  assert.equal((await iterator.next()).value.level, 0); assert.equal(read, 1);
  assert.equal((await iterator.next()).value.level, 1); assert.equal(read, 1);
  assert.equal((await iterator.next()).value.kind, "row"); assert.equal(read, 1);
  const boundary = (await iterator.next()).value;
  assert.equal(boundary.kind, "groupEnd"); assert.equal(boundary.level, 1); assert.equal(read, 2);
  await iterator.return(); assert.equal(returned, 1);
  let entered;
  const ready = new Promise(resolve => { entered = resolve; });
  const pendingRows = { [Symbol.asyncIterator]() { return {
    next() { entered(); return new Promise(() => {}); },
    return() { returned++; return Promise.resolve({ done: true }); },
  }; } };
  const controller = new AbortController(), reason = new Error("cancel hierarchy");
  const source = createGroupedReport(pendingRows, setup), operation = groupedReportEvents(source, controller.signal)[Symbol.asyncIterator]();
  const pending = operation.next(); await ready; controller.abort(reason);
  await assert.rejects(pending, error => error === reason); assert.equal(returned, 2);
  assert.throws(() => source.events[Symbol.asyncIterator](), /only once/);
});
