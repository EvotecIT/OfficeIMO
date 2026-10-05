import { checkAbort, withAbort, inputRows, pause } from "./iteration.js";
import { OfficeIMOError } from "./errors.js";

/** Writes must resolve only after the sink accepts bytes. Ownership stays with the caller. */
export interface ByteSink { write(bytes: Uint8Array): void | Promise<void>; }
export type ByteWriter = (bytes: Uint8Array) => void | Promise<void>;
export type ByteSource = Uint8Array | Iterable<Uint8Array> | AsyncIterable<Uint8Array>;

/** Collects output only; it does not retain source rows or XML strings. */
export class BlobByteSink implements ByteSink {
  private parts: Uint8Array<ArrayBuffer>[] = [];
  private state: "open" | "closed" | "discarded" = "open";
  private result: Blob | undefined;
  write(bytes: Uint8Array): void {
    if (this.state !== "open") throw new OfficeIMOError("INVALID_STATE", "Byte sink is closed.");
    // A caller may reuse its input buffer as soon as write resolves.
    this.parts.push(bytes.slice());
  }
  toBlob(type = "application/octet-stream"): Blob {
    if (this.state === "discarded") throw new OfficeIMOError("INVALID_STATE", "Byte sink was discarded.");
    this.state = "closed";
    if (!this.result) { this.result = new Blob(this.parts, { type }); this.parts = []; }
    return this.result.type === type ? this.result : new Blob([this.result], { type });
  }
  /** @internal Transfer buffered compressor chunks without a Blob read in the worker. */
  takeChunks(): readonly Uint8Array[] {
    if (this.state !== "open") throw new OfficeIMOError("INVALID_STATE", "Byte sink is closed.");
    const parts = this.parts; this.parts = []; this.state = "discarded"; return parts;
  }
  discard(): void { this.parts = []; this.result = undefined; this.state = "discarded"; }
}

/** Bounded UTF-8 batches, preserving surrogate pairs across append boundaries. */
export class ChunkedTextSink {
  private text = "";
  private deadline = performance.now() + 8;
  private readonly encoder = new TextEncoder();
  constructor(private readonly sink: ByteSink, private readonly signal?: AbortSignal, readonly chunkSize = 32768) {
    if (!Number.isInteger(chunkSize) || chunkSize < 2) throw new RangeError("Text chunk size must be at least 2.");
  }
  append(value: string): boolean {
    checkAbort(this.signal);
    this.text += value;
    return this.text.length >= this.chunkSize || performance.now() >= this.deadline;
  }
  /** Flush full text, keeping a trailing high surrogate until the next append. */
  async flush(final = false): Promise<void> {
    checkAbort(this.signal);
    while (this.text.length) {
      let end = Math.min(this.text.length, this.chunkSize);
      if (/[\ud800-\udbff]/.test(this.text[end - 1]!) && (end < this.text.length || !final)) end--;
      if (!end) break;
      const bytes = this.encoder.encode(this.text.slice(0, end));
      this.text = this.text.slice(end);
      await withAbort(Promise.resolve(this.sink.write(bytes)), this.signal);
      if (performance.now() >= this.deadline) { await pause(); this.deadline = performance.now() + 8; }
      checkAbort(this.signal);
    }
  }
  async write(text: string): Promise<void> { if (this.append(text)) await this.flush(); }
  async close(): Promise<void> { await this.flush(true); }
}

/** Feed a byte source into a caller-owned sink with backpressure and cancellation. */
export async function writeBytes(source: ByteSource, sink: ByteSink, signal?: AbortSignal): Promise<void> {
  for await (const bytes of inputRows(source instanceof Uint8Array ? [source] : source, signal)) {
    if (!(bytes instanceof Uint8Array)) throw new TypeError("Byte sources must yield Uint8Array chunks.");
    await withAbort(Promise.resolve(sink.write(bytes)), signal);
  }
}
