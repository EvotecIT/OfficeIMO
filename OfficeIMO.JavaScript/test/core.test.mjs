import test from "node:test";
import assert from "node:assert/strict";
import { BlobByteSink, ChunkedTextSink, writeBytes, detectFeatures, saveBlob, NotSupportedError } from "../dist/core/index.js";

test("text sinks preserve Unicode at chunk and append boundaries and await the destination", async () => {
  const chunks = [], sink = { async write(bytes) { await Promise.resolve(); chunks.push(bytes.slice()); } };
  const writer = new ChunkedTextSink(sink, undefined, 4);
  writer.append("abc\ud83e"); await writer.flush();
  writer.append("\uddeaשלום"); await writer.close();
  assert.equal(Buffer.concat(chunks).toString("utf8"), "abc🧪שלום");
  assert.ok(chunks.every(c => c.length <= 12));
  for (const bytes of [new Uint8Array([1, 2, 3]), Buffer.from([1, 2, 3])]) {
    const blob = new BlobByteSink();
    await writeBytes(bytes, blob); bytes.fill(0);
    assert.deepEqual(new Uint8Array(await blob.toBlob().arrayBuffer()), new Uint8Array([1, 2, 3]));
    assert.throws(() => blob.write(bytes), { code: "INVALID_STATE" });
  }
});

test("pending sink cancellation propagates the original reason", async () => {
  const controller = new AbortController(), error = new Error("stop");
  const writing = writeBytes(new Uint8Array([1]), { write: () => new Promise(() => {}) }, controller.signal);
  controller.abort(error); await assert.rejects(writing, e => e === error);
});

test("feature detection and typed errors work without a DOM", () => {
  assert.equal(detectFeatures().download, false);
  assert.equal(detectFeatures().blob, true);
  assert.throws(() => saveBlob(new Blob(), "file.csv"), { code: "PLATFORM_UNAVAILABLE" });
  assert.equal(new NotSupportedError("mergedCells").code, "NOT_SUPPORTED");
});
