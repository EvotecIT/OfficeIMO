import { checkAbort, withAbort, pause, taskYieldDue } from "../core/iteration.js";
import type { ByteSink } from "../core/sinks.js";
import { OfficeIMOError } from "../core/errors.js";
import { pdfXrefChunks, pdfXrefEntryBytes } from "./xref.js";

interface PdfCompactBatch {
  readonly id: number;
  readonly compressor: CompressionStream | undefined;
  readonly ids: number[];
  readonly bodyOffsets: number[];
  bodyBytes: number;
  indexBytes: number;
}

const compactObjectCount = 256, compactObjectBytes = 32768;
const compactIndexChunkEntries = 4096;

/** Lazy Uint16 chunks store compact membership; untouched zero slots denote ordinary objects. */
class PdfCompactIndices {
  private readonly chunks: (Uint16Array | undefined)[] = [];
  constructor(private count: number) {}
  get length(): number { return this.count; }
  reserve(): void { this.count++; }
  set(id: number, member: number): void {
    const chunk = this.chunks[Math.floor(id / compactIndexChunkEntries)] ??= new Uint16Array(compactIndexChunkEntries);
    // The current batch cap permits indices 0..255; adding one preserves index 255 as 256.
    chunk[id % compactIndexChunkEntries] = member + 1;
  }
  at(id: number): number | undefined {
    const encoded = this.chunks[Math.floor(id / compactIndexChunkEntries)]?.[id % compactIndexChunkEntries] ?? 0;
    return encoded ? encoded - 1 : undefined;
  }
}

export function pdfNumber(value: number): string {
  if (!Number.isFinite(value) || Math.abs(value) > 1e9) throw new RangeError("PDF coordinates and metrics must be finite and bounded.");
  return String(Math.round(value * 1000) / 1000);
}
export function unicodeHex(text: string): string {
  let hex = "";
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    if (code >= 0xd800 && code <= 0xdbff && !(text.charCodeAt(i + 1) >= 0xdc00 && text.charCodeAt(i + 1) <= 0xdfff) ||
      code >= 0xdc00 && code <= 0xdfff && !(text.charCodeAt(i - 1) >= 0xd800 && text.charCodeAt(i - 1) <= 0xdbff))
      throw new TypeError("PDF text contains an unpaired UTF-16 surrogate.");
    hex += code.toString(16).padStart(4, "0");
  }
  return hex;
}
/** @internal Forward-only PDF objects. Completed objects retain only their xref location. */
export class PdfObjects {
  private readonly offsets: (number | undefined)[] = [0];
  // A defined index changes the corresponding full-width offset into an object-stream number.
  // Membership chunks are allocated only when a compact member needs their ID range.
  private compactIndices: PdfCompactIndices | undefined;
  private compactBatch: PdfCompactBatch | undefined;
  private compactArena: Uint8Array<ArrayBuffer> | undefined;
  private compactUnavailable = false;
  private readonly encoder = new TextEncoder();
  private readonly buffer = new Uint8Array(32768);
  private buffered = 0;
  bytes = 0;
  constructor(private readonly sink: ByteSink, private readonly signal?: AbortSignal, private readonly maxBytes = Number.MAX_SAFE_INTEGER,
    private readonly compressObjects = false) {}
  reserve(): number { this.offsets.push(undefined); this.compactIndices?.reserve(); return this.offsets.length - 1; }
  async raw(bytes: Uint8Array): Promise<void> {
    checkAbort(this.signal);
    if (!Number.isSafeInteger(this.bytes + bytes.length) || this.bytes + bytes.length > this.maxBytes)
      throw new OfficeIMOError("RESOURCE_LIMIT", "maxOutputBytes exceeded.");
    this.bytes += bytes.length;
    for (let offset = 0; offset < bytes.length;) {
      const n = Math.min(this.buffer.length - this.buffered, bytes.length - offset);
      this.buffer.set(bytes.subarray(offset, offset + n), this.buffered); offset += n; this.buffered += n;
      if (this.buffered === this.buffer.length) await this.flushBytes();
    }
  }
  async text(text: string): Promise<void> { await this.raw(this.encoder.encode(text)); }
  async flush(): Promise<void> {
    await this.flushCompact();
    await this.flushBytes();
  }
  /** Publish completed page bytes without forcing a partial structure container to end. */
  async flushBytes(): Promise<void> {
    if (this.buffered) {
      const bytes = this.buffer.subarray(0, this.buffered);
      await withAbort(Promise.resolve(this.sink.write(bytes)), this.signal);
      // The sink borrows this view until acceptance; failed or cancelled writes must not reuse it.
      this.buffered = 0;
    }
    if (taskYieldDue()) await pause();
    checkAbort(this.signal);
  }
  /** @internal Append only when an existing compact batch can accept a member without async work. */
  tryCompactObject(id: number, body: string): boolean {
    checkAbort(this.signal); this.validateObject(id);
    const batch = this.compactBatch;
    if (!batch || !this.compressObjects || this.compactUnavailable || typeof CompressionStream !== "function" ||
      body.length > compactObjectBytes || batch.ids.length + 1 >= compactObjectCount) return false;
    const indexBytes = String(id).length + String(batch.bodyBytes).length + 2;
    const written = this.encodeCompactBody(body, batch.bodyBytes,
      compactObjectBytes - batch.indexBytes - indexBytes - batch.bodyBytes - 2);
    if (written === undefined) return false;
    this.storeCompactBody(batch, id, written); return true;
  }
  /** Small logical structure values share one bounded compressed container. Other callers keep
   * ordinary objects; unavailable deflate preserves the classic representation before packing starts. */
  async compactObject(id: number, body: string): Promise<void> {
    checkAbort(this.signal); this.validateObject(id);
    if (!this.compressObjects || this.compactUnavailable || body.length > compactObjectBytes)
      return this.object(id, body);
    if (typeof CompressionStream !== "function") {
      if (this.compactIndices) throw new Error("PDF object compression became unavailable.");
      this.compactUnavailable = true; return this.object(id, body);
    }
    const firstIndexBytes = String(id).length + 4;
    let batch = this.compactBatch;
    const indexBytes = String(id).length + String(batch?.bodyBytes ?? 0).length + 2;
    let written = batch && batch.ids.length < compactObjectCount ? this.encodeCompactBody(body, batch.bodyBytes,
      compactObjectBytes - batch.indexBytes - indexBytes - batch.bodyBytes - 2) : undefined;
    if (batch && written === undefined) {
      // Partial encodings only touch the uncommitted suffix. Flush the committed prefix and retry
      // the whole string so a multibyte scalar is never split between containers.
      await this.flushCompact(); batch = undefined;
      checkAbort(this.signal);
    }
    if (!batch) {
      written = this.encodeCompactBody(body, 0, compactObjectBytes - firstIndexBytes - 1);
      if (written === undefined) return this.object(id, body);
      let compressor: CompressionStream | undefined;
      // Probe before selecting the compact representation. Later batches need no idle native
      // codec while they accumulate, including while ordinary page streams are compressed.
      if (!this.compactIndices) {
        try { compressor = new CompressionStream("deflate"); }
        catch { this.compactUnavailable = true; return this.object(id, body); }
      }
      batch = { id: this.reserve(), compressor, ids: [], bodyOffsets: [], bodyBytes: 0, indexBytes: 1 };
      this.compactIndices ??= new PdfCompactIndices(this.offsets.length);
      this.compactBatch = batch;
    }
    this.storeCompactBody(batch, id, written!);
    if (batch.ids.length === compactObjectCount) await this.flushCompact();
  }
  private storeCompactBody(batch: PdfCompactBatch, id: number, written: number): void {
    this.offsets[id] = batch.id; this.compactIndices!.set(id, batch.ids.length);
    batch.indexBytes += String(id).length + String(batch.bodyBytes).length + 2;
    batch.ids.push(id); batch.bodyOffsets.push(batch.bodyBytes);
    this.compactArena![batch.bodyBytes + written] = 10; batch.bodyBytes += written + 1;
  }
  /** An incomplete encoding leaves only scratch bytes beyond the committed batch prefix. */
  private encodeCompactBody(body: string, offset: number, available: number): number | undefined {
    if (available < 0) return undefined;
    const destination = (this.compactArena ??= new Uint8Array(compactObjectBytes)).subarray(offset, offset + available);
    if (typeof this.encoder.encodeInto === "function") {
      const result = this.encoder.encodeInto(body, destination);
      return result.read === body.length ? result.written : undefined;
    }
    // Preserve TextEncoder implementations that only expose encode().
    const bytes = this.encoder.encode(body);
    if (bytes.length > available) return undefined;
    destination.set(bytes); return bytes.length;
  }
  private async flushCompact(): Promise<void> {
    const batch = this.compactBatch;
    if (!batch) return;
    this.compactBatch = undefined;
    let index = "";
    for (let member = 0; member < batch.ids.length; member++) index += batch.ids[member] + " " + batch.bodyOffsets[member] + " ";
    const header = this.encoder.encode(index + "\n"), count = batch.ids.length;
    // This writer awaits the codec's complete read before another member can reuse the arena.
    // Feed the small index separately instead of allocating/copying a full payload per batch.
    const body = this.compactArena!.subarray(0, batch.bodyBytes);
    batch.ids.length = 0; batch.bodyOffsets.length = 0;
    await this.compressedStream(batch.id, body, "/Type /ObjStm /N " + count + " /First " + header.length,
      batch.compressor ?? new CompressionStream("deflate"), header);
  }
  async object(id: number, body: string | Iterable<string>): Promise<void> {
    this.start(id);
    if (typeof body === "string") {
      await this.text(id + " 0 obj\n" + body + "\nendobj\n"); return;
    }
    // Logical parent arrays can be large. Coalesce small tokens without materializing the whole object.
    let chunk = id + " 0 obj\n";
    for (const part of body) {
      for (let offset = 0; offset < part.length;) {
        const end = Math.min(part.length, offset + 32768 - chunk.length);
        chunk += part.slice(offset, end); offset = end;
        if (chunk.length === 32768) {
          // Keep a trailing high surrogate with the next token, matching one-shot UTF-8 encoding.
          const code = chunk.charCodeAt(chunk.length - 1), length = code >= 0xd800 && code <= 0xdbff ? chunk.length - 1 : chunk.length;
          await this.text(chunk.slice(0, length)); chunk = chunk.slice(length);
        }
      }
    }
    await this.text(chunk + "\nendobj\n");
  }
  private validateObject(id: number): void {
    if (!Number.isSafeInteger(id) || id <= 0 || id >= this.offsets.length || this.offsets[id] !== undefined) throw new Error("Invalid PDF object state.");
  }
  private start(id: number): void {
    this.validateObject(id);
    this.offsets[id] = this.bytes;
  }
  async stream(id: number, source: Uint8Array | string, dictionary = "", compression = true, owned = false): Promise<void> {
    const bytes = typeof source === "string" ? this.encoder.encode(source) : source;
    let compressor: CompressionStream | undefined;
    if (compression && typeof CompressionStream === "function" && bytes.length) {
      try { compressor = new CompressionStream("deflate"); } catch { /* PDF streams can remain uncompressed. */ }
    }
    if (!compressor) {
      this.start(id);
      await this.text(id + " 0 obj\n<< /Length " + bytes.length + " " + dictionary + " >>\nstream\n");
      await this.raw(bytes); await this.text("\nendstream\nendobj\n"); return;
    }
    // Verified internal producers can transfer immutable ownership; other byte sources retain the defensive snapshot.
    const sourceBytes = typeof source === "string" || owned ? bytes as Uint8Array<ArrayBuffer> : new Uint8Array(bytes);
    await this.compressedStream(id, sourceBytes, dictionary, compressor);
  }
  private async compressedStream(id: number, sourceBytes: Uint8Array<ArrayBuffer>, dictionary: string, compressor: CompressionStream,
    prefix?: Uint8Array<ArrayBuffer>): Promise<void> {
    const input = new ReadableStream<BufferSource>({ start(controller) {
      if (prefix) controller.enqueue(prefix);
      controller.enqueue(sourceBytes); controller.close();
    } });
    const reader = input.pipeThrough(compressor).getReader(), lengthId = this.reserve();
    let length = 0, complete = false;
    try {
      this.start(id);
      await this.text(id + " 0 obj\n<< /Length " + lengthId + " 0 R /Filter /FlateDecode " + dictionary + " >>\nstream\n");
      while (true) {
        const next = await withAbort(reader.read(), this.signal);
        if (next.done) { complete = true; break; }
        length += next.value.length;
        await this.raw(next.value);
      }
      await this.text("\nendstream\nendobj\n");
      await this.object(lengthId, String(length));
    } finally {
      if (!complete) void reader.cancel(this.signal?.reason).catch(() => {});
      reader.releaseLock();
    }
  }
  async finish(root: number, info: number): Promise<void> {
    await this.flushCompact();
    for (let id = 1; id < this.offsets.length; id++) if (this.offsets[id] === undefined) throw new Error("PDF contains unwritten objects.");
    if (this.bytes > 9999999999) throw new OfficeIMOError("RESOURCE_LIMIT", "PDF exceeds classic cross-reference offset capacity.");
    if (this.compactIndices) return this.finishCompact(root, info);
    const start = this.bytes;
    await this.text("xref\n0 " + this.offsets.length + "\n0000000000 65535 f \n");
    for (let i = 1; i < this.offsets.length; i++) await this.text(String(this.offsets[i]).padStart(10, "0") + " 00000 n \n");
    await this.text("trailer\n<< /Size " + this.offsets.length + " /Root " + root + " 0 R /Info " + info + " 0 R >>\nstartxref\n" + start + "\n%%EOF\n");
    await this.flush();
  }
  /** Xref payloads have a known length and are emitted in bounded chunks, without another copy
   * of the document's random-access location arrays or a whole-xref compression buffer. */
  private async finishCompact(root: number, info: number): Promise<void> {
    const id = this.reserve(), start = this.bytes, length = this.offsets.length * pdfXrefEntryBytes;
    if (!Number.isSafeInteger(length)) throw new OfficeIMOError("RESOURCE_LIMIT", "PDF cross-reference stream is too large.");
    this.start(id);
    await this.text(id + " 0 obj\n<< /Type /XRef /Size " + this.offsets.length + " /W [1 8 2] /Root " + root + " 0 R /Info " + info +
      " 0 R /Length " + length + " >>\nstream\n");
    for (const chunk of pdfXrefChunks(this.offsets, this.compactIndices!)) await this.raw(chunk);
    await this.text("\nendstream\nendobj\nstartxref\n" + start + "\n%%EOF\n");
    await this.flushBytes();
  }
}
