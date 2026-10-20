import test from "node:test";
import assert from "node:assert/strict";
import { createGroupedReport, assertGroupedReportSource, groupedReportEvents } from "../dist/core/grouped-report.js";
import { ExportCell } from "../dist/core/index.js";

const columns = [{ key: "site", header: "Lokalizacja" }, { key: "latency", header: "Opóźnienie", type: "number" },
  { key: "size", header: "Rozmiar", type: "number" }];
async function collect(source) { const events = []; for await (const event of source.events) events.push(event); return events; }

test("grouped reports share numeric rules and aggregate the complete selected rows rather than group averages", async () => {
  let calls = 0;
  const report = createGroupedReport([{ site: "Łódź", latency: 10, size: 1 }, { site: "Łódź", latency: null, size: null },
    { site: "東京", latency: 20, size: 2 }, { site: "東京", latency: new ExportCell(30, { text: "30 ms" }), size: 3 },
    { site: "東京", latency: 40, size: null }], { columns, groupBy: "site", totals: { latency: "average", size: "count" },
    groupLabel(value) { calls++; return "[" + value + "]"; } });
  const events = await collect(report), ends = events.filter(event => event.kind === "groupEnd"), final = events.at(-1);
  assert.deepEqual(ends.map(event => [event.label, event.rows, event.totals]), [["[Łódź]", 2, [null, 10, 1]], ["[東京]", 3, [null, 30, 2]]]);
  assert.deepEqual(final, { kind: "complete", rows: 5, groups: 2, totals: [null, 25, 3] });
  assert.equal(calls, 2); assert.ok(Object.isFrozen(report.columns));
  assert.deepEqual(events.filter(event => event.kind === "row").map(event => [event.rowIndex, event.groupIndex]), [[0, 0], [1, 0], [2, 1], [3, 1], [4, 1]]);
  assertGroupedReportSource(report);
  assert.throws(() => assertGroupedReportSource({ ...report }), TypeError);
  assert.throws(() => report.events[Symbol.asyncIterator](), /only once/);
});

test("group identity uses stable keys and exact typed values without heading or language inference", async () => {
  const pairs = [[undefined, 1], [null, 2], [1, 3], ["1", 4], [true, 5], ["TRUE", 6], [new Date("2026-01-01T00:00:00Z"), 7], [new Date("2026-01-01T01:00:00+01:00"), 8]];
  const report = createGroupedReport(pairs, { columns: [{ key: "group", header: "Status" }, { key: "value", header: "Status", type: "number" }], groupBy: "group", totals: { value: "sum" } });
  const events = await collect(report), ends = events.filter(event => event.kind === "groupEnd");
  assert.deepEqual(ends.map(event => [event.label, event.rows, event.totals[1]]), [["", 2, 3], ["1", 1, 3], ["1", 1, 4], ["true", 1, 5], ["TRUE", 1, 6], ["2026-01-01T00:00:00.000Z", 2, 15]]);
  assert.deepEqual(events.at(-1).totals, [null, 36]);
  for (const invalid of [NaN, Infinity, new Date(NaN)])
    await assert.rejects(collect(createGroupedReport([[invalid, 1]], { columns: [{ key: "group", header: "G" }, { key: "n", header: "N" }], groupBy: "group" })), TypeError);
  for (const changed of [[{ header: "site" }], [{ key: "same", header: "A" }, { key: "same", header: "B" }], [{ key: "", header: "A" }]])
    assert.throws(() => createGroupedReport([], { columns: changed, groupBy: "site" }), /stable keys/);
  assert.throws(() => createGroupedReport([], { columns, groupBy: "Lokalizacja" }), /stable column key/);
  assert.throws(() => createGroupedReport([], { columns, groupBy: "site", totals: { "Opóźnienie": "sum" } }), /unambiguous/);
});

test("grouped sources capture column metadata, selected getters, arrays and mutable dates once", async () => {
  let headers = 0, configured = 0, getters = 0, arrayReads = 0;
  const date = new Date("2026-10-20T12:00:00Z");
  const input = [{ site: "S", latency: date }, { site: "S", latency: new ExportCell(date, { text: "local date" }) }];
  const report = createGroupedReport(input, { get columns() { configured++; return [{ key: "site", get header() { headers++; return "部署"; } },
    { key: "latency", header: "Date", value(row) { getters++; return row.latency; } }]; }, groupBy: "site" });
  const iterator = report.events[Symbol.asyncIterator]();
  assert.equal((await iterator.next()).value.kind, "groupStart");
  date.setUTCFullYear(2040);
  const first = (await iterator.next()).value;
  assert.equal(first.values[1].toISOString(), "2026-10-20T12:00:00.000Z");
  assert.equal(headers, 1); assert.equal(configured, 1); assert.equal(getters, 1);
  assert.ok(report.columns.every(column => column.value === undefined));
  await iterator.return();
  const array = ["S", 0]; Object.defineProperty(array, 1, { get() { arrayReads++; return 7; } });
  const captured = await collect(createGroupedReport([array], { columns: columns.slice(0, 2), groupBy: "site", totals: { latency: "sum" } }));
  assert.equal(arrayReads, 1); assert.equal(captured.at(-1).totals[1], 7);
});

test("grouped iteration retains one boundary row and returns its input on break, recurrence and cancellation", async () => {
  let read = 0, returned = 0;
  function* input() { try { for (const row of [["A", 1], ["A", 3], ["B", 5], ["A", 7]]) { read++; yield row; } } finally { returned++; } }
  const make = () => createGroupedReport(input(), { columns: columns.slice(0, 2), groupBy: "site", totals: { latency: "sum" } });
  const iterator = make().events[Symbol.asyncIterator]();
  assert.equal((await iterator.next()).value.kind, "groupStart"); assert.equal(read, 1);
  assert.equal((await iterator.next()).value.kind, "row"); assert.equal(read, 1);
  assert.equal((await iterator.next()).value.kind, "row"); assert.equal(read, 2);
  const end = (await iterator.next()).value;
  assert.equal(end.kind, "groupEnd"); assert.equal(end.totals[1], 4); assert.equal(read, 3);
  await iterator.return(); assert.equal(returned, 1);
  await assert.rejects(collect(make()), /closed group key/); assert.equal(returned, 2);
  const controller = new AbortController(), reason = new Error("cancel grouped export");
  const aborted = createGroupedReport(input(), { columns: columns.slice(0, 2), groupBy: "site", signal: controller.signal }).events[Symbol.asyncIterator]();
  await aborted.next(); controller.abort(reason);
  await assert.rejects(aborted.next(), error => error === reason); assert.equal(returned, 3);
});

test("group ceilings and original iterator failures reject without partial completion", async () => {
  for (const options of [{ maxGroups: 1 }, { maxGroupKeyCharacters: 0 }, { limits: { maxRows: 1 } }, { limits: { maxCells: 3 } }]) {
    let returned = 0;
    function* source() { try { yield ["A", 1]; yield ["B", 2]; } finally { returned++; } }
    await assert.rejects(collect(createGroupedReport(source(), { columns: columns.slice(0, 2), groupBy: "site", ...options })), { code: "RESOURCE_LIMIT" });
    assert.equal(returned, 1);
  }
  const reason = new Error("upstream failed");
  async function* broken() { yield ["A", 1]; throw reason; }
  await assert.rejects(collect(createGroupedReport(broken(), { columns: columns.slice(0, 2), groupBy: "site" })), error => error === reason);
  const empty = await collect(createGroupedReport([], { columns, groupBy: "site", totals: { latency: "average", size: "count" }, maxGroups: 0 }));
  assert.deepEqual(empty, [{ kind: "complete", rows: 0, groups: 0, totals: [null, null, 0] }]);
  await assert.rejects(collect(createGroupedReport([["A", 1]], { columns: columns.slice(0, 2), groupBy: "site", groupLabel: () => Promise.reject(new Error("async label")) })), /synchronous text/);
});

test("numeric metrics reject implicit date, text, boolean or custom-type coercion in both output formats", async () => {
  for (const value of [new Date("2026-10-20T00:00:00Z"), "12", true, NaN, Infinity])
    await assert.rejects(collect(createGroupedReport([["A", value]], { columns: columns.slice(0, 2), groupBy: "site", totals: { latency: "sum" } })), /finite numbers or null/);
  for (const type of [undefined, "date", "money"])
    assert.throws(() => createGroupedReport([], { columns: [{ key: "site", header: "Site" }, { key: "metric", header: "Metric", ...(type ? { type } : {}) }], groupBy: "site", totals: { metric: "average" } }), /declared with type:number/);
});

test("group sources reject byte-output caps and capture supported row/cell/text ceilings once", async () => {
  let produced = 0;
  function* input() { produced++; yield ["A", 1]; }
  for (const limits of [{ maxOutputBytes: 0 }, Object.create({ maxOutputBytes: 0 })])
    assert.throws(() => createGroupedReport(input(), { columns: columns.slice(0, 2), groupBy: "site", limits }), /maxOutputBytes.*writer/);
  assert.equal(produced, 0, "Unsupported source options fail before producer activity");
  let reads = 0;
  const report = createGroupedReport([["A", 1], ["A", 2]], { columns: columns.slice(0, 2), groupBy: "site",
    limits: { get maxRows() { reads++; return reads === 1 ? 1 : 100; } } });
  assert.equal(reads, 1);
  await assert.rejects(collect(report), { code: "RESOURCE_LIMIT" });
  assert.equal(reads, 1, "The captured source ceiling cannot change during consumption");
});

test("a writer-only cancellation reaches a pending original producer and returns it once", { timeout: 2000 }, async () => {
  for (const abortConfigured of [false, true]) {
    const configured = new AbortController(), writer = new AbortController(), reason = new Error("pending grouped source");
    let returned = 0, entered;
    const ready = new Promise(resolve => { entered = resolve; });
    const rows = { [Symbol.asyncIterator]() { return {
      next() { entered(); return new Promise(() => {}); },
      return() { returned++; return Promise.resolve({ done: true }); }
    }; } };
    const report = createGroupedReport(rows, { columns, groupBy: "site", signal: configured.signal });
    const iterator = groupedReportEvents(report, writer.signal)[Symbol.asyncIterator](), pending = iterator.next();
    await ready; (abortConfigured ? configured : writer).abort(reason);
    await assert.rejects(pending, error => error === reason); assert.equal(returned, 1);
    assert.throws(() => report.events[Symbol.asyncIterator](), /only once/);
  }
});
