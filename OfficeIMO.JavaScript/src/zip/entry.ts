import { checkAbort, withAbort } from "../core/iteration.js";
import { OfficeIMOError } from "../core/errors.js";
import type { ByteSink } from "../core/sinks.js";

export type Compression = "auto" | "store";
export const zipLimit = 0xffffffff;
const crcTable = new Uint32Array(256).map((_, index) => {
  let value = index;
  for (let bit = 0; bit < 8; bit++) value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
  return value >>> 0;
});

/** Incremental CRC-32/ISO-HDLC, also usable by a future reader. */
export class Crc32 {
  private crc = 0xffffffff;
  update(bytes: Uint8Array): void { for (const byte of bytes) this.crc = crcTable[(this.crc ^ byte) & 255]! ^ (this.crc >>> 8); }
  get value(): number { return (this.crc ^ 0xffffffff) >>> 0; }
}

export function zipSize(value: number): number {
  if (!Number.isSafeInteger(value) || value < 0 || value >= zipLimit)
    throw new OfficeIMOError("ZIP64_REQUIRED", "Classic ZIP entries and archives must stay below 4 GiB; ZIP64 is not supported yet.");
  return value;
}

export interface EntryInfo { readonly crc: number; readonly size: number; readonly compressedSize: number; readonly method: 0 | 8; }
export interface PreparedEntry extends EntryInfo { readonly chunks: readonly Uint8Array[]; }

/** Internal compressor shared by buffered worksheets and the streaming ZIP surface. */
export class EntryWriter implements ByteSink {
  readonly method: 0 | 8;
  private readonly writer: WritableStreamDefaultWriter<BufferSource> | undefined;
  private readonly reader: ReadableStreamDefaultReader<Uint8Array> | undefined;
  private readonly drain: Promise<void>;
  private readonly crc = new Crc32();
  private size = 0;
  private compressedSize = 0;
  private failed = false;
  private failure: unknown;
  private closed = false;
  private readonly abort = () => { this.fail(this.signal?.reason ?? new DOMException("Export cancelled.", "AbortError")); };

  constructor(compression: Compression, private readonly sink: ByteSink, private readonly signal?: AbortSignal) {
    if (compression !== "auto" && compression !== "store") throw new RangeError("compression must be auto or store.");
    checkAbort(signal);
    let stream: CompressionStream | undefined;
    if (compression === "auto" && typeof CompressionStream === "function") {
      try { stream = new CompressionStream("deflate-raw"); } catch (error) { if (!(error instanceof TypeError)) throw error; }
    }
    this.method = stream ? 8 : 0;
    this.writer = stream?.writable.getWriter(); this.reader = stream?.readable.getReader();
    signal?.addEventListener("abort", this.abort, { once: true });
    this.drain = this.reader ? this.consume(this.reader) : Promise.resolve();
  }
  private async consume(reader: ReadableStreamDefaultReader<Uint8Array>): Promise<void> {
    try {
      while (true) {
        const { value, done } = await withAbort(reader.read(), this.signal);
        if (done) return;
        checkAbort(this.signal);
        this.compressedSize = zipSize(this.compressedSize + value.length);
        await withAbort(Promise.resolve(this.sink.write(value)), this.signal);
      }
    } catch (error) { this.fail(error); }
  }
  private fail(error: unknown): void {
    if (!this.failed) { this.failed = true; this.failure = error; }
    this.writer?.abort(error).catch(() => {});
    this.reader?.cancel(error).catch(() => {});
  }
  async write(bytes: Uint8Array): Promise<void> {
    checkAbort(this.signal);
    if (this.failed) throw this.failure;
    if (this.closed) throw new OfficeIMOError("INVALID_STATE", "ZIP entry is closed.");
    if (!(bytes instanceof Uint8Array)) throw new TypeError("ZIP chunks must be Uint8Array.");
    this.size = zipSize(this.size + bytes.length); this.crc.update(bytes);
    if (this.writer) await withAbort(this.writer.write(bytes.slice()), this.signal);
    else { await withAbort(Promise.resolve(this.sink.write(bytes)), this.signal); this.compressedSize = this.size; }
  }
  async close(): Promise<EntryInfo> {
    try {
      checkAbort(this.signal);
      if (this.failed) throw this.failure;
      if (this.closed) throw new OfficeIMOError("INVALID_STATE", "ZIP entry is closed.");
      this.closed = true;
      if (this.writer) await withAbort(this.writer.close(), this.signal);
      await this.drain;
      if (this.failed) throw this.failure;
      checkAbort(this.signal);
      return { crc: this.crc.value, size: this.size, compressedSize: this.compressedSize, method: this.method };
    } finally { this.cleanup(); }
  }
  async discard(error: unknown): Promise<void> { this.fail(error); await this.drain; this.closed = true; this.cleanup(); }
  private cleanup(): void {
    this.signal?.removeEventListener("abort", this.abort);
    this.writer?.releaseLock(); this.reader?.releaseLock();
  }
}
