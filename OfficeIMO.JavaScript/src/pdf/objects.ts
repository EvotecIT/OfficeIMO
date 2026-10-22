import { checkAbort, withAbort, pause, taskYieldDue } from "../core/iteration.js";
import type { ByteSink } from "../core/sinks.js";
import { OfficeIMOError } from "../core/errors.js";
import { pdfXrefChunks, pdfXrefEntryBytes } from "./xref.js";

interface PdfCompactBatch {
  readonly id: number;
  readonly compressor: CompressionStream;
  readonly entries: { readonly id: number; readonly bytes: Uint8Array<ArrayBuffer> }[];
  bodyBytes: number;
  indexBytes: number;
}

const compactObjectCount = 100, compactObjectBytes = 32768;

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
  // A defined index changes the corresponding offset into an object-stream number. Parallel numeric
  // arrays avoid retaining an allocated location object for every completed structure element.
  private compactIndices: (number | undefined)[] | undefined;
  private compactBatch: PdfCompactBatch | undefined;
  private compactUnavailable = false;
  private readonly encoder = new TextEncoder();
  private buffer = new Uint8Array(32768);
  private buffered = 0;
  bytes = 0;
  constructor(private readonly sink: ByteSink, private readonly signal?: AbortSignal, private readonly maxBytes = Number.MAX_SAFE_INTEGER,
    private readonly compressObjects = false) {}
  reserve(): number { this.offsets.push(undefined); this.compactIndices?.push(undefined); return this.offsets.length - 1; }
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
  private async flushBytes(): Promise<void> {
    if (this.buffered) {
      const bytes = this.buffer.subarray(0, this.buffered);
      await withAbort(Promise.resolve(this.sink.write(bytes)), this.signal);
      this.buffer = new Uint8Array(32768); this.buffered = 0;
    }
    if (taskYieldDue()) await pause();
    checkAbort(this.signal);
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
    const bytes = this.encoder.encode(body), firstIndexBytes = String(id).length + 4;
    if (firstIndexBytes + bytes.length + 1 > compactObjectBytes) return this.object(id, body);
    let batch = this.compactBatch;
    const indexBytes = String(id).length + String(batch?.bodyBytes ?? 0).length + 2;
    if (batch && (batch.entries.length === compactObjectCount || batch.indexBytes + indexBytes + batch.bodyBytes + bytes.length + 2 > compactObjectBytes)) {
      await this.flushCompact(); batch = undefined;
    }
    if (!batch) {
      let compressor: CompressionStream;
      try { compressor = new CompressionStream("deflate"); }
      catch (error) {
        if (this.compactIndices) throw error;
        this.compactUnavailable = true; return this.object(id, body);
      }
      batch = { id: this.reserve(), compressor, entries: [], bodyBytes: 0, indexBytes: 1 };
      this.compactIndices ??= this.offsets.map(() => undefined);
      this.compactBatch = batch;
    }
    this.offsets[id] = batch.id; this.compactIndices![id] = batch.entries.length;
    batch.indexBytes += String(id).length + String(batch.bodyBytes).length + 2;
    batch.entries.push({ id, bytes }); batch.bodyBytes += bytes.length + 1;
    if (batch.entries.length === compactObjectCount) await this.flushCompact();
  }
  private async flushCompact(): Promise<void> {
    const batch = this.compactBatch;
    if (!batch) return;
    this.compactBatch = undefined;
    let offset = 0, index = "";
    for (const entry of batch.entries) { index += entry.id + " " + offset + " "; offset += entry.bytes.length + 1; }
    const header = this.encoder.encode(index + "\n"), payload = new Uint8Array(header.length + batch.bodyBytes), count = batch.entries.length;
    payload.set(header); offset = header.length;
    for (const entry of batch.entries) { payload.set(entry.bytes, offset); offset += entry.bytes.length; payload[offset++] = 10; }
    batch.entries.length = 0;
    await this.compressedStream(batch.id, payload, "/Type /ObjStm /N " + count + " /First " + header.length, batch.compressor);
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
  private async compressedStream(id: number, sourceBytes: Uint8Array<ArrayBuffer>, dictionary: string, compressor: CompressionStream): Promise<void> {
    const input = new ReadableStream<BufferSource>({ start(controller) { controller.enqueue(sourceBytes); controller.close(); } });
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
