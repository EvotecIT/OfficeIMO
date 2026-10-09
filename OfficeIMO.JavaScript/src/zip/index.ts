export { Crc32 } from "./entry.js";
export type { Compression, EntryInfo } from "./entry.js";
import { EntryWriter, zipSize } from "./entry.js";
import type { Compression, PreparedEntry, EntryInfo } from "./entry.js";
import { checkAbort, withAbort, pause, taskYieldDue } from "../core/iteration.js";
import { OfficeIMOError } from "../core/errors.js";
import { BlobByteSink, writeBytes } from "../core/sinks.js";
import type { ByteSink, ByteSource } from "../core/sinks.js";

export interface ZipWriterOptions { readonly compression?: Compression; readonly signal?: AbortSignal; readonly maxOutputBytes?: number; }
/** An appendable entry. Await writes and close it before opening another entry. */
export interface ZipEntry extends ByteSink { close(): Promise<EntryInfo>; discard(error: unknown): Promise<void>; }
export type ZipEntrySource = ByteSource | ((sink: ByteSink) => void | Promise<void>);
const encoder = new TextEncoder();

export function validateEntryName(name: string): string {
  if (typeof name !== "string" || !name || /[\u0000-\u001f\\]/.test(name) || name.startsWith("/") ||
      /^[A-Za-z]:/.test(name) || name.split("/").some(s => !s || s === "." || s === ".."))
    throw new TypeError("ZIP entry names must be relative paths without empty, dot or parent segments.");
  for (let i = 0; i < name.length; i++) {
    const unit = name.charCodeAt(i);
    if (unit >= 0xd800 && unit <= 0xdbff) {
      const next = name.charCodeAt(++i);
      if (!(next >= 0xdc00 && next <= 0xdfff)) throw new TypeError("ZIP entry names must contain well-formed Unicode.");
    } else if (unit >= 0xdc00 && unit <= 0xdfff) throw new TypeError("ZIP entry names must contain well-formed Unicode.");
  }
  if (encoder.encode(name).length > 65535) throw new RangeError("ZIP entry name exceeds 65,535 UTF-8 bytes.");
  return name;
}

function header(signature: number, length: number, name: Uint8Array = new Uint8Array()): { bytes: Uint8Array; data: DataView } {
  const bytes = new Uint8Array(length + name.length), data = new DataView(bytes.buffer);
  data.setUint32(0, signature, true); bytes.set(name, length);
  return { bytes, data };
}

/** Sequential ZIP writer. Entry bytes use data descriptors; only central-directory metadata is retained. */
export class ZipWriter {
  private readonly sink: ByteSink;
  private readonly owned: BlobByteSink | undefined;
  private readonly names = new Set<string>();
  private readonly central: Uint8Array[] = [];
  private offset = 0;
  private active: ZipEntry | undefined;
  private state: "open" | "writing" | "finished" | "failed" = "open";
  constructor(sink?: ByteSink, private readonly options: ZipWriterOptions = {}) {
    if (options.maxOutputBytes !== undefined && (!Number.isSafeInteger(options.maxOutputBytes) || options.maxOutputBytes < 0)) throw new RangeError("maxOutputBytes must be a nonnegative safe integer.");
    if (options.compression !== undefined && options.compression !== "auto" && options.compression !== "store")
      throw new RangeError("compression must be auto or store.");
    this.owned = sink ? undefined : new BlobByteSink(); this.sink = sink ?? this.owned!;
  }
  private async emit(bytes: Uint8Array): Promise<void> {
    if (this.options.maxOutputBytes !== undefined && this.offset + bytes.length > this.options.maxOutputBytes)
      throw new OfficeIMOError("RESOURCE_LIMIT", "maxOutputBytes exceeded.");
    this.offset = zipSize(this.offset + bytes.length);
    await withAbort(Promise.resolve(this.sink.write(bytes)), this.options.signal);
  }
  private reserve(name: string): Uint8Array {
    checkAbort(this.options.signal);
    if (this.state !== "open") throw new OfficeIMOError("INVALID_STATE", "ZIP writer is busy, failed or finalized.");
    validateEntryName(name);
    if (this.names.has(name)) throw new TypeError("Duplicate ZIP entry: " + name);
    if (this.names.size >= 65534) throw new OfficeIMOError("ZIP64_REQUIRED", "ZIP64 is required for 65,535 entries.");
    this.names.add(name); this.state = "writing";
    return encoder.encode(name);
  }
  private local(name: Uint8Array, info: Pick<EntryInfo, "method">, descriptor: boolean): Uint8Array {
    const h = header(0x04034b50, 30, name);
    h.data.setUint16(4, 20, true); h.data.setUint16(6, descriptor ? 0x808 : 0x800, true);
    h.data.setUint16(8, info.method, true); h.data.setUint16(12, 33, true); h.data.setUint16(26, name.length, true);
    return h.bytes;
  }
  private record(name: Uint8Array, info: EntryInfo, offset: number, descriptor: boolean): void {
    const h = header(0x02014b50, 46, name);
    h.data.setUint16(4, 20, true); h.data.setUint16(6, 20, true); h.data.setUint16(8, descriptor ? 0x808 : 0x800, true);
    h.data.setUint16(10, info.method, true); h.data.setUint16(14, 33, true); h.data.setUint32(16, info.crc, true);
    h.data.setUint32(20, info.compressedSize, true); h.data.setUint32(24, info.size, true);
    h.data.setUint16(28, name.length, true); h.data.setUint32(42, offset, true); this.central.push(h.bytes);
  }
  async openEntry(name: string): Promise<ZipEntry> {
    const encoded = this.reserve(name), offset = this.offset;
    let entry: EntryWriter | undefined;
    try {
      entry = new EntryWriter(this.options.compression ?? "auto", { write: bytes => this.emit(bytes) }, this.options.signal);
      await this.emit(this.local(encoded, entry, true));
      const writer = entry;
      let busy = false, closed = false, completion: Promise<EntryInfo> | undefined;
      const discard = async (error: unknown) => { closed = true; this.state = "failed"; await writer.discard(error); this.owned?.discard(); this.active = undefined; };
      return this.active = {
        write: async bytes => {
          if (busy || closed || this.state !== "writing") throw new OfficeIMOError("INVALID_STATE", "Await the current ZIP entry operation.");
          busy = true;
          try { await writer.write(bytes); } catch (error) { await discard(error); throw error; } finally { busy = false; }
        },
        close: () => {
          if (completion) return completion;
          if (busy || closed || this.state !== "writing") throw new OfficeIMOError("INVALID_STATE", "ZIP entry is busy or closed.");
          closed = true;
          completion = (async () => {
            try {
              const info = await writer.close(), descriptor = header(0x08074b50, 16);
              descriptor.data.setUint32(4, info.crc, true); descriptor.data.setUint32(8, info.compressedSize, true);
              descriptor.data.setUint32(12, info.size, true); await this.emit(descriptor.bytes);
              this.record(encoded, info, offset, true); this.state = "open"; this.active = undefined; return info;
            } catch (error) { await discard(error); throw error; }
          })();
          return completion;
        }, discard
      };
    } catch (error) { this.state = "failed"; await entry?.discard(error); this.owned?.discard(); throw error; }
  }
  async add(name: string, source: ZipEntrySource): Promise<void> {
    const entry = await this.openEntry(name);
    try {
      if (typeof source === "function") await withAbort(Promise.resolve(source(entry)), this.options.signal);
      else await writeBytes(source, entry, this.options.signal);
      await entry.close();
    } catch (error) { await entry.discard(error); throw error; }
  }
  get bytesWritten(): number { return this.offset; }
  /** Drop owned output and release the active compressor. Caller-owned partial bytes remain with the caller. */
  async discard(error: unknown): Promise<void> { this.state = "failed"; await this.active?.discard(error); this.owned?.discard(); }
  /** @internal Buffered worksheet compression stays in the format owner. */
  async addPrepared(name: string, entry: PreparedEntry): Promise<void> {
    const encoded = this.reserve(name), offset = this.offset;
    try {
      const local = this.local(encoded, entry, false), data = new DataView(local.buffer);
      data.setUint32(14, entry.crc, true); data.setUint32(18, entry.compressedSize, true); data.setUint32(22, entry.size, true);
      await this.emit(local);
      for (const bytes of entry.chunks) { checkAbort(this.options.signal); await this.emit(bytes); }
      this.record(encoded, entry, offset, false); this.state = "open";
    } catch (error) { this.state = "failed"; this.owned?.discard(); throw error; }
  }
  /** Complete directory; does not close a caller-owned sink. A failed sink owns disposal of partial bytes. */
  async finish(): Promise<void> {
    checkAbort(this.options.signal);
    if (this.state !== "open") throw new OfficeIMOError("INVALID_STATE", "ZIP writer is busy, failed or finalized.");
    this.state = "writing";
    const offset = this.offset;
    try {
      for (let i = 0; i < this.central.length; i++) { await this.emit(this.central[i]!); if (i % 256 === 255 && taskYieldDue()) await pause(); }
      const end = header(0x06054b50, 22);
      end.data.setUint16(8, this.central.length, true); end.data.setUint16(10, this.central.length, true);
      end.data.setUint32(12, this.offset - offset, true); end.data.setUint32(16, offset, true);
      await this.emit(end.bytes); checkAbort(this.options.signal); this.state = "finished";
    } catch (error) { this.state = "failed"; this.owned?.discard(); throw error; }
  }
  async toBlob(type = "application/zip"): Promise<Blob> {
    if (!this.owned) throw new OfficeIMOError("INVALID_STATE", "This ZIP uses a caller-owned sink; use finish().");
    if (this.state !== "finished") await this.finish();
    return this.owned.toBlob(type);
  }
}
