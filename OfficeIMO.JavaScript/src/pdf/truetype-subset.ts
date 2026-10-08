import type { TrueTypeFont } from "./truetype.js";
import { align4, checksum } from "./binary.js";

/** Preserve glyph IDs and complete composite dependencies, as in OfficeIMO.Pdf's native subsetter. */
export function subsetTrueType(font: TrueTypeFont, requested: ReadonlySet<number>, characters: ReadonlyMap<number, number>): Uint8Array {
  const glyphs = new Set([0, ...requested]), glyfTable = font.table("glyf"), queue = [...glyphs];
  const dependencies = new Map<number, number[]>();
  for (let i = 0; i < queue.length; i++) {
    const glyph = queue[i]!;
    if (!Number.isInteger(glyph) || glyph < 0 || glyph >= font.glyphCount) throw new TypeError("Invalid subset glyph.");
    const start = font.offsets[glyph]!, end = font.offsets[glyph + 1]!;
    if (start === end) continue;
    if (end - start < 10) throw new TypeError("Truncated TrueType glyph header.");
    if (font.i16(glyfTable.offset + start) >= 0) continue;
    let cursor = glyfTable.offset + start + 10, flags: number;
    const children: number[] = []; dependencies.set(glyph, children);
    const limit = glyfTable.offset + end;
    do {
      if (cursor + 4 > limit) throw new TypeError("Truncated TrueType composite.");
      flags = font.u16(cursor); const component = font.u16(cursor + 2);
      if (component >= font.glyphCount) throw new TypeError("Invalid TrueType composite dependency.");
      children.push(component);
      if (!glyphs.has(component)) { glyphs.add(component); queue.push(component); }
      cursor += 4 + (flags & 1 ? 4 : 2) + (flags & 8 ? 2 : flags & 64 ? 4 : flags & 128 ? 8 : 0);
      if (cursor > limit) throw new TypeError("Truncated TrueType composite transform.");
    } while (flags & 32);
    if (flags & 256) { if (cursor + 2 > limit || cursor + 2 + font.u16(cursor) > limit) throw new TypeError("Truncated composite instructions."); }
  }
  // A malformed caller-provided font must not export recursive composite glyphs.
  const visited = new Set<number>(), visiting = new Set<number>();
  function visit(glyph: number, depth: number): void {
    if (visiting.has(glyph)) throw new TypeError("Cyclic TrueType composite dependency.");
    if (depth > 64) throw new TypeError("TrueType composite nesting exceeds 64 levels.");
    if (visited.has(glyph)) return;
    visiting.add(glyph);
    for (const child of dependencies.get(glyph) ?? []) visit(child, depth + 1);
    visiting.delete(glyph); visited.add(glyph);
  }
  for (const glyph of glyphs) visit(glyph, 0);
  const offsets = new Uint32Array(font.glyphCount + 1);
  let length = 0;
  for (let glyph = 0; glyph < font.glyphCount; glyph++) { offsets[glyph] = length; if (glyphs.has(glyph)) length += align4(font.offsets[glyph + 1]! - font.offsets[glyph]!); }
  offsets[font.glyphCount] = length;
  const glyf = new Uint8Array(length);
  for (const glyph of glyphs) glyf.set(font.bytes.subarray(glyfTable.offset + font.offsets[glyph]!, glyfTable.offset + font.offsets[glyph + 1]!), offsets[glyph]);
  const loca = new Uint8Array(offsets.length * 4), locaView = new DataView(loca.buffer);
  for (let i = 0; i < offsets.length; i++) locaView.setUint32(i * 4, offsets[i]!);
  const retained = new Set(["OS/2", "cmap", "cvt ", "fpgm", "gasp", "glyf", "head", "hhea", "hmtx", "loca", "maxp", "name", "post", "prep"]);
  const cmap = subsetCmap(characters), hmtx = subsetMetrics(font, glyphs);
  const tables = [...font.tables].filter(([tag]) => retained.has(tag)).sort(([a], [b]) => a < b ? -1 : 1).map(([tag, table]) => {
    const bytes = tag === "glyf" ? glyf : tag === "loca" ? loca : tag === "cmap" ? cmap : tag === "hmtx" ? hmtx : font.bytes.slice(table.offset, table.offset + table.length);
    if (tag === "head") { const view = new DataView(bytes.buffer); view.setUint32(8, 0); view.setInt16(50, 1); }
    return { tag, bytes };
  });
  const size = 12 + tables.length * 16 + tables.reduce((n, t) => n + align4(t.bytes.length), 0), output = new Uint8Array(size), view = new DataView(output.buffer);
  view.setUint32(0, 0x10000); view.setUint16(4, tables.length);
  const power = 2 ** Math.floor(Math.log2(tables.length));
  view.setUint16(6, power * 16); view.setUint16(8, Math.log2(power)); view.setUint16(10, tables.length * 16 - power * 16);
  let cursor = 12 + tables.length * 16, head = 0;
  tables.forEach((table, i) => {
    const p = 12 + i * 16;
    for (let j = 0; j < 4; j++) output[p + j] = table.tag.charCodeAt(j);
    view.setUint32(p + 4, checksum(table.bytes)); view.setUint32(p + 8, cursor); view.setUint32(p + 12, table.bytes.length);
    output.set(table.bytes, cursor); if (table.tag === "head") head = cursor;
    cursor += align4(table.bytes.length);
  });
  view.setUint32(head + 8, (0xb1b0afba - checksum(output)) >>> 0);
  return output;
}

/** Keep original glyph IDs, advance widths and bearings for every selected/composite glyph.
 * Empty glyph slots have no outline; their unused metrics need not inflate the compressed subset. */
function subsetMetrics(font: TrueTypeFont, glyphs: ReadonlySet<number>): Uint8Array {
  const metrics = font.u16(font.table("hhea").offset + 34), table = font.table("hmtx");
  const result = new Uint8Array(metrics * 4 + (font.glyphCount - metrics) * 2);
  for (const glyph of glyphs) {
    const first = glyph < metrics ? glyph * 4 : metrics * 4 + (glyph - metrics) * 2, length = glyph < metrics ? 4 : 2;
    result.set(font.bytes.subarray(table.offset + first, table.offset + first + length), first);
    if (glyph >= metrics) {
      // Trailing short records share the final full record's advance width.
      const advance = (metrics - 1) * 4;
      result.set(font.bytes.subarray(table.offset + advance, table.offset + advance + 2), advance);
    }
  }
  return result;
}

/** One Unicode format-12 map covers BMP and supplementary scalars without retaining unused coverage. */
function subsetCmap(characters: ReadonlyMap<number, number>): Uint8Array {
  const entries = [...characters].sort(([a], [b]) => a - b), groups: { first: number; last: number; glyph: number }[] = [];
  for (const [scalar, glyph] of entries) {
    const previous = groups.at(-1);
    if (previous && scalar === previous.last + 1 && glyph === previous.glyph + scalar - previous.first) previous.last = scalar;
    else groups.push({ first: scalar, last: scalar, glyph });
  }
  const result = new Uint8Array(12 + 16 + groups.length * 12), view = new DataView(result.buffer);
  view.setUint16(2, 1); view.setUint16(4, 3); view.setUint16(6, 10); view.setUint32(8, 12);
  view.setUint16(12, 12); view.setUint32(16, result.length - 12); view.setUint32(24, groups.length);
  groups.forEach((group, i) => {
    const p = 28 + i * 12;
    view.setUint32(p, group.first); view.setUint32(p + 4, group.last); view.setUint32(p + 8, group.glyph);
  });
  return result;
}
