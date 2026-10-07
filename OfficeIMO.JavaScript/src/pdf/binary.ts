/** @internal Bounded big-endian reader used for caller-provided TrueType programs. */
export class FontReader {
  readonly view: DataView;
  constructor(readonly bytes: Uint8Array) { this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength); }
  range(offset: number, length: number): void {
    if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(length) || offset < 0 || length < 0 || offset + length > this.bytes.length)
      throw new TypeError("Truncated or invalid TrueType table.");
  }
  u16(offset: number): number { this.range(offset, 2); return this.view.getUint16(offset); }
  i16(offset: number): number { this.range(offset, 2); return this.view.getInt16(offset); }
  u32(offset: number): number { this.range(offset, 4); return this.view.getUint32(offset); }
  tag(offset: number): string { this.range(offset, 4); return String.fromCharCode(...this.bytes.subarray(offset, offset + 4)); }
}
export interface FontTable { readonly offset: number; readonly length: number; }
export function checksum(bytes: Uint8Array): number {
  let sum = 0;
  for (let i = 0; i < bytes.length; i += 4) sum = (sum + (((bytes[i] ?? 0) * 0x1000000) + ((bytes[i + 1] ?? 0) << 16) + ((bytes[i + 2] ?? 0) << 8) + (bytes[i + 3] ?? 0))) >>> 0;
  return sum;
}
export const align4 = (value: number): number => Math.ceil(value / 4) * 4;
