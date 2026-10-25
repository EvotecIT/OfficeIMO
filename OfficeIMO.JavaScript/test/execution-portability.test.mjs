import test from "node:test";
import assert from "node:assert/strict";
import { MessageChannel } from "node:worker_threads";
import { runInNewContext } from "node:vm";
import { ExportCell } from "../dist/core/index.js";
import { executeWorkerExport, handleWorkerExports } from "../dist/execution/index.js";

async function session(operation, run) {
  const { port1, port2 } = new MessageChannel(), sent = [];
  const endpoint = {
    postMessage(message) { sent.push(message); port1.postMessage(message); },
    addEventListener(...args) { port1.addEventListener(...args); },
    removeEventListener(...args) { port1.removeEventListener(...args); },
    start() { port1.start(); }
  };
  const handler = handleWorkerExports(port2, operation);
  try { await run(endpoint, sent); }
  finally { handler.dispose(); port1.close(); port2.close(); }
}

test("worker rows reject nested presentation graphs before structured cloning", { timeout: 5000 }, async () => {
  const invalid = { background: { payload: "x".repeat(1000000) } };
  assert.throws(() => new ExportCell(0, { presentation: invalid }), /background must be text/);
  // Compatible branded cells can originate in another bundle; enforce the same boundary there.
  const foreign = { value: 0, presentation: invalid, [Symbol.for("@evotecit/officeimo/ExportCell")]: true };
  await session(async ({ rows }) => {
    let count = 0;
    for await (const row of rows) count++;
    return { rows: count, columns: 1, bytes: 0 };
  }, async (endpoint, sent) => {
    await assert.rejects(executeWorkerExport(endpoint, [[foreign]], { write() {} }, { maxBatchCharacters: 1 }), /background must be text/);
    assert.equal(sent.filter(message => message.kind === "rows").length, 0);
  });
});

test("portable presentation captures only declared scalars and budgets their text", { timeout: 5000 }, async () => {
  const cell = new ExportCell(1, { presentation: { bold: true, numberFormat: "0000", ignored: { payload: "x".repeat(1000000) } } });
  assert.deepEqual(cell.presentation, { numberFormat: "0000", bold: true });
  await session(async ({ rows }) => {
    for await (const row of rows) assert.deepEqual(row[0].presentation, cell.presentation);
    return { rows: 1, columns: 1, bytes: 0 };
  }, async (endpoint, sent) => {
    await assert.rejects(executeWorkerExport(endpoint, [[cell]], { write() {} }, { maxBatchCharacters: 3 }), error => error.code === "RESOURCE_LIMIT");
    assert.equal(sent.filter(message => message.kind === "rows").length, 0);
    assert.equal((await executeWorkerExport(endpoint, [[cell]], { write() {} }, { maxBatchCharacters: 4 })).rows, 1);
    assert.deepEqual(sent.find(message => message.kind === "rows").rows[0][0].options.presentation, cell.presentation);
  });
});

test("worker setup accepts plain records and byte arrays from another realm, rejecting instances", { timeout: 5000 }, async () => {
  const input = runInNewContext('({report:{title:"iframe",constructor:"label"},columns:[{header:"Usluga",key:"0",rowHeader:true}],wideTable:{mode:"bands",repeatColumns:[0,"0"]},font:new Uint8Array([3,7]),date:new Date("2026-10-20T00:00:00Z"),empty:Object.create(null)})');
  await session(async ({ input }) => {
    assert.equal(input.report.title, "iframe");
    assert.equal(input.report.constructor, "label");
    assert.deepEqual(input.columns, [{ header: "Usluga", key: "0", rowHeader: true }]);
    assert.deepEqual(input.wideTable.repeatColumns, [0, "0"]);
    assert.deepEqual([...input.font], [3, 7]);
    assert.equal(input.date.toISOString(), "2026-10-20T00:00:00.000Z");
    assert.deepEqual(Object.keys(input.empty), []);
    return { rows: 0, columns: 0, bytes: 0 };
  }, async (endpoint, sent) => {
    await executeWorkerExport(endpoint, [], { write() {} }, { input });
    const starts = sent.filter(message => message.kind === "start").length;
    for (const value of [runInNewContext('(new (class Report { constructor(){this.title="instance"} })())'), new (class Report {})()])
      await assert.rejects(executeWorkerExport(endpoint, [], { write() {} }, { input: value }), /plain records/);
    assert.equal(sent.filter(message => message.kind === "start").length, starts);
  });
});
