import test from "node:test";
import assert from "node:assert/strict";
import { ZipWriter } from "../dist/zip/index.js";

for (const phase of ["header", "data", "descriptor", "directory", "end"]) {
  test("discard and configured cancellation release a pending ZIP " + phase + " acceptance", { timeout: 5000 }, async () => {
    for (const owner of ["discard", "signal"]) {
      const controller = new AbortController(), reason = new Error(owner + " " + phase), accepted = [];
      let entered, writes = 0;
      const held = new Promise(resolve => { entered = resolve; });
      const selected = { header: 1, data: 2, descriptor: 3, directory: 4, end: 5 }[phase];
      const zip = new ZipWriter({ write(bytes) {
        accepted.push(bytes.slice()); writes++;
        if (writes === selected) { entered(); return new Promise(() => {}); }
      } }, { compression: "store", signal: controller.signal });
      let pending;
      if (phase === "header") pending = zip.openEntry("report");
      else {
        const entry = await zip.openEntry("report");
        if (phase === "data") pending = entry.write(new Uint8Array([1]));
        else {
          await entry.write(new Uint8Array([1]));
          if (phase === "descriptor") pending = entry.close();
          else { await entry.close(); pending = zip.finish(); }
        }
      }
      await held;
      if (owner === "signal") controller.abort(reason);
      else {
        const discarded = zip.discard(reason);
        assert.equal(zip.discard(new Error("later reason")), discarded, "Discard is idempotent and retains its first failure");
        await discarded;
      }
      await assert.rejects(pending, error => error === reason);
      await zip.discard(reason);
      assert.equal(writes, selected); assert.equal(accepted.length, selected, "Caller-owned partial bytes remain available");
      await assert.rejects(zip.finish(), owner === "signal" ? error => error === reason : { code: "INVALID_STATE" });
    }
  });
}

test("explicit discard returns a held byte producer and stops a held producer callback", { timeout: 5000 }, async () => {
  for (const mode of ["iterator", "callback"]) {
    let entered, returned = 0;
    const held = new Promise(resolve => { entered = resolve; });
    const source = mode === "callback" ? () => { entered(); return new Promise(() => {}); } : {
      [Symbol.asyncIterator]() { return {
        next() { entered(); return new Promise(() => {}); },
        return() { returned++; return { done: true }; }
      }; }
    };
    const zip = new ZipWriter({ write() {} }, { compression: "store" }), reason = new Error("stop " + mode);
    const pending = zip.add("data", source);
    await held; await zip.discard(reason);
    await assert.rejects(pending, error => error === reason);
    await new Promise(resolve => setImmediate(resolve));
    assert.equal(returned, mode === "iterator" ? 1 : 0);
  }
});

test("ZIP successful and failed lifetimes release configured-signal listeners", async () => {
  for (const failure of [false, true]) {
    const controller = new AbortController(), signal = controller.signal, listeners = new Set();
    const add = signal.addEventListener.bind(signal), remove = signal.removeEventListener.bind(signal);
    signal.addEventListener = (type, callback, options) => { if (type === "abort") listeners.add(callback); return add(type, callback, options); };
    signal.removeEventListener = (type, callback, options) => { if (type === "abort") listeners.delete(callback); return remove(type, callback, options); };
    const reason = new Error("destination rejected"), zip = new ZipWriter({ write() { if (failure) throw reason; } }, { compression: "store", signal });
    if (failure) await assert.rejects(zip.add("data", new Uint8Array([1])), error => error === reason);
    else { await zip.add("data", new Uint8Array([1])); await zip.finish(); }
    assert.equal(listeners.size, 0);
  }
});
