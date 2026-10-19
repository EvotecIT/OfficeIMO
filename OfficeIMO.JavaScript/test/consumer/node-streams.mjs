import test from "node:test";
import assert from "node:assert/strict";
import { createWriteStream } from "node:fs";
import { readFile } from "node:fs/promises";
import { Writable } from "node:stream";
import { finished } from "node:stream/promises";
import { writeCsvTo, writeXlsxTo, writePdfTo } from "@evotecit/officeimo";

const rows = [{ customer: "Contoso", amount: 12.5 }];
const columns = [{ header: "Customer", key: "customer" }, { header: "Amount", key: "amount" }];
const formats = [
  ["csv", writeCsvTo, { columns }],
  ["xlsx", writeXlsxTo, { columns, compression: "store", created: new Date("2026-01-01T00:00:00Z") }],
  ["pdf", writePdfTo, { columns, compression: false }]
];

for (const [format, write, options] of formats) {
  await test(`packed ${format} awaits Node backpressure and leaves the destination open`, { timeout: 5000 }, async () => {
    const expected = [];
    await write(rows, { write: bytes => { expected.push(bytes.slice()); } }, options);
    const chunks = [];
    let release, started;
    const firstWrite = new Promise(resolve => { started = resolve; });
    const output = new Writable({
      highWaterMark: 1,
      write(chunk, _encoding, callback) {
        chunks.push(Buffer.from(chunk));
        if (chunks.length === 1) { release = callback; started(); }
        else callback();
      }
    });
    const destination = Writable.toWeb(output);
    let settled = false;
    const pending = write(rows, destination, options);
    void pending.then(() => { settled = true; }, () => { settled = true; });
    await firstWrite;
    await new Promise(setImmediate);
    assert.equal(settled, false, "The export must wait for the blocked Node write.");
    assert.equal(output.writableLength, chunks[0].length, "Later writes must not queue behind the blocked chunk.");
    release();
    const result = await pending;
    assert.deepEqual(Buffer.concat(chunks), Buffer.concat(expected));
    assert.deepEqual(result, { rows: 1, columns: 2, bytes: Buffer.concat(chunks).length });
    assert.equal(destination.locked, false);
    assert.equal(output.writableEnded, false);
    assert.equal(output.destroyed, false);
    output.end();
    await finished(output);
  });

  await test(`packed ${format} caller observes a Node destination failure`, { timeout: 5000 }, async () => {
    const failure = new Error("Node destination failed");
    const output = new Writable({ highWaterMark: 1, write(_chunk, _encoding, callback) { callback(failure); } });
    const destination = Writable.toWeb(output);
    const completion = finished(output);
    void completion.catch(() => {});
    // The adapter may accept a final chunk before Node reports its I/O error.
    const exportFile = async () => {
      try {
        const result = await write(rows, destination, options);
        output.end();
        await completion;
        return result;
      } catch (error) {
        output.destroy();
        await completion.catch(() => {});
        throw error;
      }
    };
    await assert.rejects(exportFile(), error => error === failure);
    assert.equal(destination.locked, false);
  });
}

await test("packed Node file export completes when the caller ends the stream", { timeout: 5000 }, async () => {
  const output = createWriteStream("packed-node-stream.xlsx");
  const completion = finished(output);
  void completion.catch(() => {});
  try {
    const result = await writeXlsxTo(rows, Writable.toWeb(output), formats[1][2]);
    assert.equal(output.writableEnded, false);
    output.end();
    await completion;
    const bytes = await readFile("packed-node-stream.xlsx");
    const expected = [];
    await writeXlsxTo(rows, { write: chunk => { expected.push(chunk.slice()); } }, formats[1][2]);
    assert.deepEqual(bytes, Buffer.concat(expected));
    assert.equal(bytes.length, result.bytes);
  } catch (error) {
    output.destroy();
    await completion.catch(() => {});
    throw error;
  }
});
