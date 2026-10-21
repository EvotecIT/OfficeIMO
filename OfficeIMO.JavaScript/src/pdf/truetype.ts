import { NotSupportedError } from "../core/errors.js";
import { FontReader } from "./binary.js";
import type { FontTable } from "./binary.js";
import { subsetTrueType, subsetTrueTypeCooperatively } from "./truetype-subset.js";
import { checkAbort, pause, taskYieldDue } from "../core/iteration.js";

/** @internal Native TrueType outline profile shared with the C# PDF writer's glyph-preserving subset design. */
export class TrueTypeFont extends FontReader {
  readonly tables: Map<string, FontTable>;
  readonly glyphCount: number;
  readonly units: number;
  readonly ascent: number;
  readonly descent: number;
  readonly bbox: readonly number[];
  readonly widths: readonly number[];
  readonly offsets: readonly number[];
  readonly canSubset: boolean;
  private readonly cmaps: readonly FontTable[];
  constructor(bytes: Uint8Array, prepared?: ParsedFont) {
    super(bytes);
    const scanner = parseFont(this);
    let data = prepared;
    if (!data) while (true) { const next = scanner.next(); if (next.done) { data = next.value; break; } }
    this.tables = data.tables; this.glyphCount = data.glyphCount; this.units = data.units;
    this.ascent = data.ascent; this.descent = data.descent; this.bbox = data.bbox;
    this.widths = data.widths; this.offsets = data.offsets; this.canSubset = data.canSubset; this.cmaps = data.cmaps;
  }
  /** Parse an owned immutable snapshot while allowing cancellation and host events between bounded scans. */
  static async create(bytes: Uint8Array, signal?: AbortSignal): Promise<TrueTypeFont> {
    checkAbort(signal);
    const scanner = parseFont(new FontReader(bytes));
    while (true) {
      const next = scanner.next();
      if (next.done) { checkAbort(signal); return new TrueTypeFont(bytes, next.value); }
      if (taskYieldDue()) await pause();
      checkAbort(signal);
    }
  }
  table(tag: string, minimum = 0): FontTable {
    const table = this.tables.get(tag);
    if (!table || table.length < minimum) throw new TypeError("Missing or truncated TrueType table: " + tag);
    return table;
  }
  glyph(scalar: number): number {
    for (const map of this.cmaps) {
      const p = map.offset, type = this.u16(p);
      let glyph = 0;
      if (type === 12) {
        let left = 0, right = this.u32(p + 12) - 1;
        while (left <= right) {
          const middle = (left + right) >>> 1, q = p + 16 + middle * 12, start = this.u32(q), end = this.u32(q + 4);
          if (scalar < start) right = middle - 1;
          else if (scalar > end) left = middle + 1;
          else { glyph = this.u32(q + 8) + scalar - start; break; }
        }
      } else if (scalar <= 0xffff) {
        const n = this.u16(p + 6) / 2;
        let left = 0, right = n - 1;
        while (left < right) { const mid = (left + right) >>> 1; if (scalar > this.u16(p + 14 + mid * 2)) left = mid + 1; else right = mid; }
        const start = this.u16(p + 16 + n * 2 + left * 2);
        if (scalar >= start && scalar <= this.u16(p + 14 + left * 2)) {
          const delta = this.i16(p + 16 + n * 4 + left * 2), address = p + 16 + n * 6 + left * 2, range = this.u16(address);
          if (!range) glyph = (scalar + delta) & 0xffff;
          else {
            const q = address + range + (scalar - start) * 2;
            if (q + 2 > p + map.length) throw new TypeError("cmap glyph array exceeds its table.");
            glyph = this.u16(q); if (glyph) glyph = (glyph + delta) & 0xffff;
          }
        }
      }
      if (glyph >= this.glyphCount) throw new TypeError("cmap references an invalid glyph.");
      if (glyph) return glyph;
    }
    return 0;
  }
  subset(glyphs: ReadonlySet<number>, characters: ReadonlyMap<number, number>): Uint8Array { return this.canSubset ? subsetTrueType(this, glyphs, characters) : this.bytes; }
  async subsetCooperatively(glyphs: ReadonlySet<number>, characters: ReadonlyMap<number, number>, signal?: AbortSignal): Promise<Uint8Array> {
    checkAbort(signal); return this.canSubset ? subsetTrueTypeCooperatively(this, glyphs, characters, signal) : this.bytes;
  }
}

interface ParsedFont {
  readonly tables: Map<string, FontTable>; readonly glyphCount: number; readonly units: number;
  readonly ascent: number; readonly descent: number; readonly bbox: readonly number[]; readonly widths: readonly number[];
  readonly offsets: readonly number[]; readonly canSubset: boolean; readonly cmaps: readonly FontTable[];
}
/** One parser serves synchronous construction and cooperative preparation. */
function* parseFont(reader: FontReader): Generator<void, ParsedFont, undefined> {
  const tables = new Map<string, FontTable>();
  const table = (tag: string, minimum = 0): FontTable => {
    const result = tables.get(tag);
    if (!result || result.length < minimum) throw new TypeError("Missing or truncated TrueType table: " + tag);
    return result;
  };
  if (reader.u32(0) !== 0x10000) throw new NotSupportedError("PDF fonts require a static TrueType .ttf with glyf outlines; collections, CFF and WOFF are unsupported.");
  const count = reader.u16(4);
  if (!count || count > 256) throw new TypeError("Invalid TrueType table count.");
  reader.range(12, count * 16);
  for (let i = 0; i < count; i++) {
    const p = 12 + i * 16, tag = reader.tag(p), offset = reader.u32(p + 8), length = reader.u32(p + 12);
    reader.range(offset, length);
    if (tables.has(tag)) throw new TypeError("Duplicate TrueType table: " + tag);
    tables.set(tag, { offset, length });
  }
  if (tables.has("fvar")) throw new NotSupportedError("Instantiate a variable font as a static TrueType font before PDF export.");
  const head = table("head", 54), maxp = table("maxp", 6), hhea = table("hhea", 36), os2 = table("OS/2", 10);
  if (reader.u32(head.offset + 12) !== 0x5f0f3cf5) throw new TypeError("Invalid TrueType head magic.");
  const glyphCount = reader.u16(maxp.offset + 4); const units = reader.u16(head.offset + 18);
  if (!glyphCount || units < 16 || units > 16384) throw new TypeError("Invalid TrueType metrics.");
  const fsType = reader.u16(os2.offset + 8);
  if (fsType & 0x202) throw new NotSupportedError("The font's embedding permissions prohibit outline embedding.");
  const canSubset = !(fsType & 0x100);
  const scale = 1000 / units;
  const ascent = reader.i16(hhea.offset + 4) * scale; const descent = reader.i16(hhea.offset + 6) * scale;
  const bbox = [36, 38, 40, 42].map(p => reader.i16(head.offset + p) * scale);
  const metrics = reader.u16(hhea.offset + 34);
  if (!metrics || metrics > glyphCount) throw new TypeError("Invalid TrueType horizontal metric count.");
  const hmtx = table("hmtx", metrics * 4 + (glyphCount - metrics) * 2);
  const widths: number[] = [];
  for (let i = 0; i < glyphCount; i++) { widths.push(reader.u16(hmtx.offset + Math.min(i, metrics - 1) * 4) * scale); if ((i & 1023) === 1023) yield; }
  const format = reader.i16(head.offset + 50);
  if (format !== 0 && format !== 1) throw new TypeError("Unsupported TrueType loca index.");
  const loca = table("loca", (glyphCount + 1) * (format ? 4 : 2)), glyf = table("glyf");
  const offsets: number[] = [];
  for (let i = 0; i <= glyphCount; i++) { offsets.push(format ? reader.u32(loca.offset + i * 4) : reader.u16(loca.offset + i * 2) * 2); if ((i & 1023) === 1023) yield; }
  for (let i = 0; i <= glyphCount; i++) if (offsets[i]! > glyf.length || (i && offsets[i]! < offsets[i - 1]!)) throw new TypeError("Invalid TrueType glyph offsets.");
  const cmap = table("cmap", 4), maps: FontTable[] = [], validated = new Set<number>();
  const encodings = reader.u16(cmap.offset + 2);
  if (4 + encodings * 8 > cmap.length) throw new TypeError("Invalid TrueType cmap directory.");
  for (let i = 0; i < encodings; i++) {
    const p = cmap.offset + 4 + i * 8, platform = reader.u16(p), encoding = reader.u16(p + 2), relative = reader.u32(p + 4);
    if (platform !== 0 && !(platform === 3 && (encoding === 1 || encoding === 10))) continue;
    if (relative + 4 > cmap.length) throw new TypeError("Invalid cmap subtable offset.");
    const offset = cmap.offset + relative, type = reader.u16(offset);
    if (type !== 4 && type !== 12) continue;
    // Encoding records may share a Unicode subtable; validate its contents once.
    if (validated.has(offset)) continue;
    if (type === 12 && relative + 16 > cmap.length) throw new TypeError("Truncated cmap format 12.");
    const length = type === 12 ? reader.u32(offset + 4) : reader.u16(offset + 2);
    if (length < (type === 12 ? 16 : 16) || relative + length > cmap.length) throw new TypeError("Invalid cmap length.");
    if (type === 12) {
      const groups = reader.u32(offset + 12);
      if (16 + groups * 12 > length) throw new TypeError("Invalid cmap groups.");
      let previous = -1;
      for (let j = 0; j < groups; j++) {
        const q = offset + 16 + j * 12, start = reader.u32(q), end = reader.u32(q + 4), glyph = reader.u32(q + 8);
        if (start <= previous || start > end || end > 0x10ffff || glyph + end - start >= glyphCount) throw new TypeError("Invalid cmap scalar/glyph group.");
        previous = end; if ((j & 1023) === 1023) yield;
      }
    } else {
      const segments = reader.u16(offset + 6) / 2;
      if (!Number.isInteger(segments) || !segments || 16 + segments * 8 > length) throw new TypeError("Invalid cmap segments.");
      let previous = -1;
      for (let j = 0; j < segments; j++) {
        const end = reader.u16(offset + 14 + j * 2), start = reader.u16(offset + 16 + segments * 2 + j * 2);
        if (end <= previous || start > end) throw new TypeError("Invalid cmap segment ordering.");
        previous = end; if ((j & 1023) === 1023) yield;
      }
    }
    validated.add(offset); maps.push({ offset, length });
  }
  const cmaps = maps.sort((a, b) => reader.u16(b.offset) - reader.u16(a.offset));
  if (!maps.length) throw new NotSupportedError("TrueType font needs a Unicode cmap format 4 or 12.");
  return { tables, glyphCount, units, ascent, descent, bbox, widths, offsets, canSubset, cmaps };
}
