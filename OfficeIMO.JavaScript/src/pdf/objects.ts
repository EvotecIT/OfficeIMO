import { checkAbort, withAbort, pause, taskYieldDue } from "../core/iteration.js";
import type { ByteSink } from "../core/sinks.js";
import { OfficeIMOError } from "../core/errors.js";

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
/** @internal Forward-only PDF objects. Only xref offsets and page references survive a completed page. */
export class PdfObjects {
  private readonly offsets: (number | undefined)[] = [0];
  private readonly encoder = new TextEncoder();
  private buffer = new Uint8Array(32768);
  private buffered = 0;
  bytes = 0;
  constructor(private readonly sink: ByteSink, private readonly signal?: AbortSignal, private readonly maxBytes = Number.MAX_SAFE_INTEGER) {}
  reserve(): number { this.offsets.push(undefined); return this.offsets.length - 1; }
  async raw(bytes: Uint8Array): Promise<void> {
    checkAbort(this.signal);
    if (this.bytes + bytes.length > this.maxBytes) throw new OfficeIMOError("RESOURCE_LIMIT", "maxOutputBytes exceeded.");
    this.bytes += bytes.length;
    for (let offset = 0; offset < bytes.length;) {
      const n = Math.min(this.buffer.length - this.buffered, bytes.length - offset);
      this.buffer.set(bytes.subarray(offset, offset + n), this.buffered); offset += n; this.buffered += n;
      if (this.buffered === this.buffer.length) await this.flush();
    }
  }
  async text(text: string): Promise<void> { await this.raw(this.encoder.encode(text)); }
  async flush(): Promise<void> {
    if (this.buffered) {
      const bytes = this.buffer.subarray(0, this.buffered);
      await withAbort(Promise.resolve(this.sink.write(bytes)), this.signal);
      this.buffer = new Uint8Array(32768); this.buffered = 0;
    }
    if (taskYieldDue()) await pause();
    checkAbort(this.signal);
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
  private start(id: number): void {
    if (!Number.isInteger(id) || id <= 0 || id >= this.offsets.length || this.offsets[id] !== undefined) throw new Error("Invalid PDF object state.");
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
    if (this.offsets.slice(1).some(offset => offset === undefined)) throw new Error("PDF contains unwritten objects.");
    if (this.bytes > 9999999999) throw new OfficeIMOError("RESOURCE_LIMIT", "PDF exceeds classic cross-reference offset capacity.");
    const start = this.bytes;
    await this.text("xref\n0 " + this.offsets.length + "\n0000000000 65535 f \n");
    for (let i = 1; i < this.offsets.length; i++) await this.text(String(this.offsets[i]).padStart(10, "0") + " 00000 n \n");
    await this.text("trailer\n<< /Size " + this.offsets.length + " /Root " + root + " 0 R /Info " + info + " 0 R >>\nstartxref\n" + start + "\n%%EOF\n");
    await this.flush();
  }
}
