import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { writeCsv, writeCsvTo } from "../dist/csv/index.js";

const vectors = JSON.parse(await readFile(new URL("../../OfficeIMO.TestAssets/CSV/browser-exports.json", import.meta.url), "utf8"));
for (const vector of vectors.cases) test("shared CSV vector: " + vector.name, async () => {
  const rows = vector.rows.map(row => row.map(value => value?.kind === "date" ? new Date(value.value) : value));
  const blob = await writeCsv(rows, vector);
  const expected = Buffer.concat([vector.bom ? Buffer.from([239, 187, 191]) : Buffer.alloc(0), Buffer.from(vector.expected)]);
  assert.deepEqual(Buffer.from(await blob.arrayBuffer()), expected);
});

test("CSV snapshots projection, supports async objects and reports completion", async () => {
  const columns = [{ header: "Number", key: "value" }, { header: "When", key: "date" }], events = [];
  async function* rows() {
    columns[0].key = "other";
    yield { value: -4, date: new Date("2026-10-05T12:34:56.123Z") };
    yield { value: Infinity, date: new Date(NaN) };
  }
  const blob = await writeCsv(rows(), { columns, onProgress: p => events.push(p) });
  assert.equal(await blob.text(), "Number,When\r\n-4,2026-10-05T12:34:56.123Z\r\n,\r\n");
  assert.deepEqual(events.at(-1), { phase: "complete", rows: 2, bytes: blob.size });
});

test("CSV rejects unsupported values and malformed dialects", async () => {
  const columns = [{ header: "Value" }];
  await assert.rejects(writeCsv([[{}]], { columns }), TypeError);
  await assert.rejects(writeCsv([[1, 2]], { columns }), RangeError);
  await assert.rejects(writeCsv([], { columns, delimiter: '"' }), RangeError);
  await assert.rejects(writeCsv([], { columns, lineEnding: "bad" }), RangeError);
});

test("CSV cancels pending producer input promptly and returns its iterator", async () => {
  const controller = new AbortController();
  let returned = false;
  const rows = { [Symbol.asyncIterator]() { return {
    next: () => new Promise(() => {}), return: () => { returned = true; return { done: true }; }
  }; } };
  const writing = writeCsv(rows, { columns: [{ header: "V" }], signal: controller.signal });
  setTimeout(() => controller.abort(), 10);
  await assert.rejects(writing, { name: "AbortError" });
  assert.equal(returned, true);
});

test("CSV cancellation during encoding returns no partial Blob", async () => {
  const controller = new AbortController();
  function* rows() { for (let i = 0; i < 100000; i++) yield ["=unsafe " + i]; }
  const writing = writeCsv(rows(), { columns: [{ header: "V" }], signal: controller.signal });
  setTimeout(() => controller.abort(), 10);
  await assert.rejects(writing, { name: "AbortError" });
});
test("CSV preserves promised rows and the original consumer failure when producer return fails", async () => {
  const failure = new Error("formatter failure"); let returned = 0;
  const source = { [Symbol.iterator]() { let row = 0; return {
    next() { return row++ < 2 ? { value: Promise.resolve([row]), done: false } : { done: true }; },
    return() { returned++; throw Error("cleanup failure"); }
  }; } };
  const calls = [];
  await assert.rejects(writeCsv(source, { columns: [{ header: "Value", valueFormatter(value, context) {
    calls.push([value, context.rowIndex]); if (context.rowIndex === 1) throw failure; return value;
  } }] }), error => error === failure);
  assert.equal(returned, 1); assert.deepEqual(calls, [[1, 0], [2, 1]]);
});

test("CSV flushes inside a formatted row without repeating callbacks or losing its snapshot", async () => {
  const values = ["=unsafe", "Łódź 🧪".repeat(9000), '"quoted"', "last"], calls = [], snapshots = [];
  const columns = values.map((_, column) => ({ header: "Column " + column, valueFormatter(value, context) {
    calls.push([context.rowIndex, context.columnIndex]); snapshots.push(context.values); return value;
  } }));
  const blob = await writeCsv([values], { columns, includeHeader: false, quote: "all", limits: { maxCells: 4 } });
  const expected = '"\'=unsafe","' + values[1] + '","""quoted""","last"\r\n';
  assert.equal(await blob.text(), expected);
  assert.deepEqual(calls, [[0, 0], [0, 1], [0, 2], [0, 3]]);
  assert.ok(snapshots.every(snapshot => snapshot === snapshots[0] && Object.isFrozen(snapshot)));
  assert.deepEqual(snapshots[0], values);
});

test("CSV does not exhaust a synchronous producer while a destination waits and cancellation returns it", async () => {
  let produced = 0, returned = 0;
  const controller = new AbortController(), reason = new Error("destination cancelled");
  function* source() { try { for (let i = 0; i < 10; i++) { produced++; yield ["x".repeat(20000)]; } } finally { returned++; } }
  const destination = new WritableStream({ write() { setTimeout(() => controller.abort(reason), 0); return new Promise(() => {}); } });
  await assert.rejects(writeCsvTo(source(), destination, { columns: [{ header: "V" }], includeHeader: false, signal: controller.signal }), error => error === reason);
  assert.ok(produced > 0 && produced <= 2); assert.equal(returned, 1); assert.equal(destination.locked, false);
});
