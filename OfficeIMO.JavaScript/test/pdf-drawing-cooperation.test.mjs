import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { PdfFont, writePdfDocumentTo } from "../dist/pdf/index.js";

const regular = new PdfFont(new Uint8Array(await readFile(new URL("../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf", import.meta.url))));
const text = () => ({ kind: "text", text: "A".repeat(1000), x: 0, y: 20, font: "regular", fontSize: 10, color: "000000" });
const path = (count = 1000) => ({ kind: "path", stroke: "000000", commands: [
  { op: "move", x: 0, y: 0 }, ...Array.from({ length: count - 1 }, () => ({ op: "line", x: 1, y: 1 }))
] });

// Each child stays below its own former cooperation threshold. Only aggregate work crosses the budget.
for (const [name, child, siblings, abortYield] of [
  // The fifth text checkpoint requires both glyph preflight and paint to cooperate before output.
  ["font preflight and text painting", text, 2, 5],
  ["path commands", path, 2, 2],
  // Both traversal stages must cooperate before the outer block writer reaches its task boundary.
  ["nested element traversal", () => path(1), 200, 3]
]) {
  test("drawing cancellation cooperates across " + name, async () => {
    const now = Object.getOwnPropertyDescriptor(performance, "now"), scheduler = Object.getOwnPropertyDescriptor(globalThis, "scheduler");
    const abort = new AbortController(), reason = Error("aggregate drawing cancelled");
    let clock = 0, yields = 0, returned = 0, writes = 0;
    const drawing = { version: 1, width: 200, height: 100, elements: Array.from({ length: 2 }, () => ({
      kind: "group", elements: Array.from({ length: siblings }, () => child())
    })) };
    const source = { *[Symbol.iterator]() { try { yield { kind: "drawing", drawing }; } finally { returned++; } } };
    const destination = new WritableStream({ write() { writes++; } });
    try {
      Object.defineProperty(performance, "now", { configurable: true, value: () => clock += 40 });
      Object.defineProperty(globalThis, "scheduler", { configurable: true, value: { postTask: callback => new Promise(resolve => {
        setTimeout(() => { if (++yields === abortYield) abort.abort(reason); callback(); resolve(); }, 0);
      }) } });
      await assert.rejects(writePdfDocumentTo(source, destination, { fonts: { regular }, pageNumbers: false, compression: false, signal: abort.signal }), error => error === reason);
      assert.equal(writes, 0, "Cancellation during aggregate drawing work must precede destination output");
      assert.equal(yields, abortYield); assert.equal(returned, 1); assert.equal(destination.locked, false);
    } finally {
      if (now) Object.defineProperty(performance, "now", now); else delete performance.now;
      if (scheduler) Object.defineProperty(globalThis, "scheduler", scheduler); else delete globalThis.scheduler;
    }
  });
}
