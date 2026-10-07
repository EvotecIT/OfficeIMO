import { NotSupportedError } from "../core/errors.js";
import { FontReader } from "./binary.js";
import type { FontTable } from "./binary.js";
import { subsetTrueType } from "./truetype-subset.js";

/** @internal Native TrueType outline profile shared with the C# PDF writer's glyph-preserving subset design. */
export class TrueTypeFont extends FontReader {
  readonly tables = new Map<string, FontTable>();
  readonly glyphCount: number;
  readonly units: number;
  readonly ascent: number;
  readonly descent: number;
  readonly bbox: readonly number[];
  readonly widths: readonly number[];
  readonly offsets: readonly number[];
  readonly canSubset: boolean;
  private readonly cmaps: readonly FontTable[];
  constructor(bytes: Uint8Array) {
    super(bytes);
    if (this.u32(0) !== 0x10000) throw new NotSupportedError("PDF fonts require a static TrueType .ttf with glyf outlines; collections, CFF and WOFF are unsupported.");
    const count = this.u16(4);
    if (!count || count > 256) throw new TypeError("Invalid TrueType table count.");
    this.range(12, count * 16);
    for (let i = 0; i < count; i++) {
      const p = 12 + i * 16, tag = this.tag(p), offset = this.u32(p + 8), length = this.u32(p + 12);
      this.range(offset, length);
      if (this.tables.has(tag)) throw new TypeError("Duplicate TrueType table: " + tag);
      this.tables.set(tag, { offset, length });
    }
    if (this.tables.has("fvar")) throw new NotSupportedError("Instantiate a variable font as a static TrueType font before PDF export.");
    const head = this.table("head", 54), maxp = this.table("maxp", 6), hhea = this.table("hhea", 36), os2 = this.table("OS/2", 10);
    if (this.u32(head.offset + 12) !== 0x5f0f3cf5) throw new TypeError("Invalid TrueType head magic.");
    this.glyphCount = this.u16(maxp.offset + 4); this.units = this.u16(head.offset + 18);
    if (!this.glyphCount || this.units < 16 || this.units > 16384) throw new TypeError("Invalid TrueType metrics.");
    const fsType = this.u16(os2.offset + 8);
    if (fsType & 0x202) throw new NotSupportedError("The font's embedding permissions prohibit outline embedding.");
    this.canSubset = !(fsType & 0x100);
    const scale = 1000 / this.units;
    this.ascent = this.i16(hhea.offset + 4) * scale; this.descent = this.i16(hhea.offset + 6) * scale;
    this.bbox = [36, 38, 40, 42].map(p => this.i16(head.offset + p) * scale);
    const metrics = this.u16(hhea.offset + 34);
    if (!metrics || metrics > this.glyphCount) throw new TypeError("Invalid TrueType horizontal metric count.");
    const hmtx = this.table("hmtx", metrics * 4 + (this.glyphCount - metrics) * 2);
    this.widths = Array.from({ length: this.glyphCount }, (_, i) => this.u16(hmtx.offset + Math.min(i, metrics - 1) * 4) * scale);
    const format = this.i16(head.offset + 50);
    if (format !== 0 && format !== 1) throw new TypeError("Unsupported TrueType loca index.");
    const loca = this.table("loca", (this.glyphCount + 1) * (format ? 4 : 2)), glyf = this.table("glyf");
    this.offsets = Array.from({ length: this.glyphCount + 1 }, (_, i) => format ? this.u32(loca.offset + i * 4) : this.u16(loca.offset + i * 2) * 2);
    for (let i = 0; i <= this.glyphCount; i++) if (this.offsets[i]! > glyf.length || (i && this.offsets[i]! < this.offsets[i - 1]!)) throw new TypeError("Invalid TrueType glyph offsets.");
    const cmap = this.table("cmap", 4), maps: FontTable[] = [];
    const encodings = this.u16(cmap.offset + 2);
    if (4 + encodings * 8 > cmap.length) throw new TypeError("Invalid TrueType cmap directory.");
    for (let i = 0; i < encodings; i++) {
      const p = cmap.offset + 4 + i * 8, platform = this.u16(p), encoding = this.u16(p + 2), relative = this.u32(p + 4);
      if (platform !== 0 && !(platform === 3 && (encoding === 1 || encoding === 10))) continue;
      if (relative + 4 > cmap.length) throw new TypeError("Invalid cmap subtable offset.");
      const offset = cmap.offset + relative, type = this.u16(offset);
      if (type !== 4 && type !== 12) continue;
      if (type === 12 && relative + 16 > cmap.length) throw new TypeError("Truncated cmap format 12.");
      const length = type === 12 ? this.u32(offset + 4) : this.u16(offset + 2);
      if (length < (type === 12 ? 16 : 16) || relative + length > cmap.length) throw new TypeError("Invalid cmap length.");
      if (type === 12) {
        const groups = this.u32(offset + 12);
        if (16 + groups * 12 > length) throw new TypeError("Invalid cmap groups.");
        let previous = -1;
        for (let j = 0; j < groups; j++) {
          const q = offset + 16 + j * 12, start = this.u32(q), end = this.u32(q + 4), glyph = this.u32(q + 8);
          if (start <= previous || start > end || end > 0x10ffff || glyph + end - start >= this.glyphCount) throw new TypeError("Invalid cmap scalar/glyph group.");
          previous = end;
        }
      } else {
        const segments = this.u16(offset + 6) / 2;
        if (!Number.isInteger(segments) || !segments || 16 + segments * 8 > length) throw new TypeError("Invalid cmap segments.");
        let previous = -1;
        for (let j = 0; j < segments; j++) {
          const end = this.u16(offset + 14 + j * 2), start = this.u16(offset + 16 + segments * 2 + j * 2);
          if (end <= previous || start > end) throw new TypeError("Invalid cmap segment ordering.");
          previous = end;
        }
      }
      if (!maps.some(m => m.offset === offset)) maps.push({ offset, length });
    }
    this.cmaps = maps.sort((a, b) => this.u16(b.offset) - this.u16(a.offset));
    if (!maps.length) throw new NotSupportedError("TrueType font needs a Unicode cmap format 4 or 12.");
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
}
