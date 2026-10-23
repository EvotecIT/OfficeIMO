import test from "node:test";
import assert from "node:assert/strict";
import { writeCsv, writeCsvTo } from "../dist/csv/index.js";
import { writePdf } from "../dist/pdf/index.js";
import { Workbook } from "../dist/xlsx/index.js";
import { ZipWriter } from "../dist/zip/index.js";

const columns = [{ header: "Value" }];
const rejectedReason = promise => promise.then(() => assert.fail("Expected cancellation"), reason => reason);

test("public writers preserve the exact pre-aborted signal reason, including null", async () => {
  for (const reason of [null, new Error("owned cancellation"), "cancelled", undefined]) {
    const controller = new AbortController(); controller.abort(reason);
    const signal = controller.signal, expected = signal.reason;
    assert.strictEqual(await rejectedReason(writeCsv([["one"]], { columns, signal })), expected);
    assert.strictEqual(await rejectedReason(writePdf([["one"]], { columns, signal })), expected);
    assert.strictEqual(await rejectedReason(Promise.resolve().then(() => {
      const workbook = new Workbook({ signal }); workbook.addWorksheet("Data", { columns });
    })), expected);
    assert.strictEqual(await rejectedReason(new ZipWriter(undefined, { signal }).add("data", new Uint8Array([1]))), expected);
  }
});

test("null cancellation returns a pending CSV producer and unlocks its caller-owned destination", { timeout: 5000 }, async () => {
  const controller = new AbortController(); let ready, returned = 0;
  const started = new Promise(resolve => { ready = resolve; });
  const rows = { [Symbol.asyncIterator]() { return {
    next() { ready(); return new Promise(() => {}); },
    return() { returned++; return { done: true }; }
  }; } };
  const destination = new WritableStream({ write() {} });
  const pending = writeCsvTo(rows, destination, { columns, signal: controller.signal });
  await started; controller.abort(null);
  assert.strictEqual(await rejectedReason(pending), null);
  assert.equal(returned, 1); assert.equal(destination.locked, false);
});

test("null cancellation settles a pending ZIP destination and compressor entry", { timeout: 5000 }, async () => {
  for (const compression of ["store", "auto"]) {
    const controller = new AbortController(); let ready, release;
    const started = new Promise(resolve => { ready = resolve; });
    const blocked = new Promise(resolve => { release = resolve; });
    let writes = 0;
    const zip = new ZipWriter({ write() { if (++writes > 1) { ready(); return blocked; } } }, { signal: controller.signal, compression });
    let adding;
    try {
      adding = zip.add("data", async sink => { await sink.write(new Uint8Array(65536)); });
      await started; controller.abort(null);
      assert.strictEqual(await rejectedReason(adding), null);
    } finally { release(); await zip.discard(null); }
  }
});
