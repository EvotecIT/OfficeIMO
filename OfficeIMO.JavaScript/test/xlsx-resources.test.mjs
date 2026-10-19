import test from "node:test";
import assert from "node:assert/strict";
import { Workbook } from "../dist/xlsx/index.js";

const png = new Uint8Array(Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jfKsAAAAASUVORK5CYII=", "base64"));

test("XLSX image admission rejects over-budget bytes before a defensive allocation", async () => {
  const book = new Workbook({ limits: { maxImageBytes: 0 } });
  const sheet = book.addWorksheet("Image", { columns: [{ header: "Value" }] });
  const NativeUint8Array = globalThis.Uint8Array;
  let copies = 0;
  try {
    // Observe allocation at the native boundary; the supplied PNG remains an ordinary Uint8Array.
    globalThis.Uint8Array = new Proxy(NativeUint8Array, { construct(target, args, newTarget) {
      if (args[0] === png) copies++;
      return Reflect.construct(target, args, newTarget);
    } });
    assert.throws(() => sheet.addImage({ data: png, row: 1, column: 1, width: 1, height: 1 }), { code: "RESOURCE_LIMIT" });
    assert.equal(copies, 0);
  } finally {
    globalThis.Uint8Array = NativeUint8Array;
    await book.discard(new Error("test complete"));
  }
});

test("the worksheet hyperlink ceiling includes external and text-preservation links", async () => {
  const book = new Workbook({ compression: "store", oversizedText: "preserve" });
  const links = Array.from({ length: 65529 }, (_, i) => ({ cell: "A" + (i + 1), target: "https://example.com/" + i }));
  const columns = [{ header: "External" }, { header: "Preserved" }];
  try {
    const sheet = book.addWorksheet("Links", { columns, hyperlinks: links });
    await sheet.addRows([[null, "x".repeat(32768)]]);
    assert.equal(sheet.hyperlinks.length, 65529);
    assert.throws(() => sheet.addHyperlink({ cell: "B1", target: "https://example.com/extra" }), /65,530 hyperlinks/);
    assert.equal(sheet.hyperlinks.length, 65529, "rejected admission must preserve existing links");
    const other = book.addWorksheet("Other", { columns });
    other.addHyperlink({ cell: "A1", target: "https://example.com/other" });
    assert.equal(other.hyperlinks.length, 1, "the native ceiling is per worksheet");
    assert.throws(() => book.addWorksheet("Too many", { columns, hyperlinks: [...links,
      { cell: "A65530", target: "https://example.com/65530" }, { cell: "A65531", target: "https://example.com/65531" }] }), /65,530 hyperlinks/);
    const full = book.addWorksheet("Full", { columns, hyperlinks: [...links, { cell: "A65530", target: "https://example.com/65530" }] });
    await assert.rejects(full.addRows([[null, "x".repeat(32768)]]), /65,530 hyperlinks/);
  } finally { await book.discard(new Error("test complete")); }
});

test("timer cancellation interrupts XLSX relationship finalization before remaining package output", async () => {
  const controller = new AbortController(), reason = new Error("cancel metadata");
  let bytes = 0, timer;
  const book = new Workbook({ compression: "store", signal: controller.signal, sink: { write(chunk) { bytes += chunk.length; } } });
  const links = Array.from({ length: 1024 }, (_, i) => ({ cell: "A" + (i + 1), target: "https://example.com/" + i }));
  const sheet = book.addWorksheet("Links", { columns: [{ header: "Value" }], hyperlinks: links });
  await sheet.addRows(Array.from({ length: 1023 }, (_, i) => [i]));
  await sheet.close();
  const before = bytes, original = Object.getOwnPropertyDescriptor(performance, "now"), clock = performance.now() + 1000;
  try {
    Object.defineProperty(performance, "now", { configurable: true, value: () => clock });
    timer = setTimeout(() => controller.abort(reason), 0);
    await assert.rejects(book.finish(), error => error === reason);
    assert.equal(bytes, before);
  } finally {
    clearTimeout(timer);
    if (original) Object.defineProperty(performance, "now", original); else delete performance.now;
    await book.discard(reason);
  }
});
