import assert from "node:assert/strict";
import { writeFile } from "node:fs/promises";

// Execute the compiled public recipe. Fetch is the external boundary; pagination
// must resolve against the API response URL rather than the worker module URL.
const originalSelf = globalThis.self, originalFetch = globalThis.fetch;
const requests = [], messages = [];
try {
  globalThis.self = { postMessage: message => messages.push(message) };
  globalThis.fetch = async url => {
    requests.push(url);
    assert.ok(url === "https://example.com/api/sales" || url === "https://example.com/api/sales?page=2");
    const page2 = url.endsWith("?page=2");
    const response = new Response(JSON.stringify({ rows: [{ customer: { name: page2 ? "second" : "Łódź 🧪" }, amount: page2 ? 7.5 : 12.5, seen: "2026-10-07T00:00:00Z" }], next: page2 ? null : "?page=2" }));
    Object.defineProperty(response, "url", { value: url });
    return response;
  };
  await import("./worker-runtime/worker.mjs");
  for (const format of ["csv", "xlsx"]) {
    messages.length = 0; requests.length = 0;
    await self.onmessage({ data: { format, url: "https://example.com/api/sales" } });
    assert.deepEqual(requests, ["https://example.com/api/sales", "https://example.com/api/sales?page=2"]);
    assert.equal(messages.some(message => message.error), false);
    const blob = messages.find(message => message.blob)?.blob; assert.ok(blob instanceof Blob);
    if (format === "csv") assert.equal(await blob.text(), "Customer,Amount,Seen\r\nŁódź 🧪,12.5,2026-10-07T00:00:00.000Z\r\nsecond,7.5,2026-10-07T00:00:00.000Z\r\n");
    else await writeFile("packed-worker.xlsx", new Uint8Array(await blob.arrayBuffer()));
  }
} finally { globalThis.self = originalSelf; globalThis.fetch = originalFetch; }
console.log("Compiled worker recipe exports both formats through response-relative paged links.");
