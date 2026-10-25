import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import { fontProgram } from "./font.js";
import type { PdfFont } from "./font.js";
import { resolveScalar, selectFont, standardCode, validateFonts } from "./font-support.js";
import { pdfNumber, unicodeHex } from "./objects.js";
import type { PdfObjects } from "./objects.js";
import type { PdfFonts } from "./types.js";
import type { TrueTypeFont } from "./truetype.js";
import { checkAbort, pause, taskYieldDue } from "../core/iteration.js";

const hexBytes = Object.freeze(Array.from({ length: 256 }, (_, byte) => byte.toString(16).padStart(2, "0")));

function cidHex(cid: number): string { return hexBytes[cid >>> 8]! + hexBytes[cid & 255]!; }

export class PdfFontResource {
  readonly name: string;
  readonly object: number;
  readonly program: TrueTypeFont | undefined;
  readonly syntheticBold: boolean;
  readonly syntheticItalic: boolean;
  private readonly mappings = new Map<number | string, { cid: number; hex: string; glyph: number; width: number; text: string }>();
  private readonly standardWidths = new Map<number, number>();
  constructor(readonly index: number, readonly bold: boolean, readonly italic: boolean, readonly font: PdfFont | undefined,
    objects: PdfObjects, syntheticBold: boolean, syntheticItalic: boolean, private readonly shared?: PdfFontResource, private readonly signal?: AbortSignal) {
    this.name = shared?.name ?? "F" + index; this.object = shared?.object ?? objects.reserve(); this.program = font ? fontProgram(font) : undefined;
    this.syntheticBold = syntheticBold; this.syntheticItalic = syntheticItalic;
  }
  width(scalar: number): number {
    if (this.shared) return this.shared.width(scalar);
    if (this.program) {
      const existing = this.mappings.get(scalar);
      if (existing) return existing.width;
    } else {
      const standard = this.standardWidths.get(scalar);
      if (standard !== undefined) return standard;
    }
    const resolved = resolveScalar(scalar, this.program, this.bold);
    if (!resolved.supported) {
      if (resolved.code === "unpaired-surrogate") throw new TypeError(resolved.message);
      throw new NotSupportedError(resolved.message);
    }
    if (!this.program) { this.standardWidths.set(scalar, resolved.width); return resolved.width; }
    if (this.mappings.size >= 65535) throw new OfficeIMOError("RESOURCE_LIMIT", "A PDF font supports at most 65,535 distinct Unicode scalars.");
    const cid = this.mappings.size + 1;
    const mapping = { cid, hex: cidHex(cid), glyph: resolved.glyph!, width: resolved.width, text: String.fromCodePoint(scalar) };
    this.mappings.set(scalar, mapping); return mapping.width;
  }
  lineHeight(size: number): number {
    return Math.max(1.3, this.program ? (this.program.ascent - this.program.descent) / 1000 + .1 : 1.3) * size;
  }
  encode(text: string): string {
    if (this.shared) return this.shared.encode(text);
    let hex = "";
    for (const char of text) {
      const cp = char.codePointAt(0)!;
      if (this.program) {
        let mapping = this.mappings.get(cp);
        // Layout normally registered the scalar already; validate only a newly painted scalar.
        if (!mapping) { this.width(cp); mapping = this.mappings.get(cp)!; }
        hex += mapping.hex;
      } else {
        if (!this.standardWidths.has(cp)) this.width(cp);
        hex += hexBytes[standardCode(cp)!]!;
      }
    }
    return "<" + hex + ">";
  }
  /** Register an already validated shaped glyph. A glyph may need distinct CIDs for different logical clusters. */
  encodeGlyph(glyph: number, text: string): string {
    if (this.shared) return this.shared.encodeGlyph(glyph, text);
    if (!this.program || !Number.isSafeInteger(glyph) || glyph <= 0 || glyph >= this.program.glyphCount) throw new TypeError("Invalid shaped TrueType glyph.");
    const key = glyph + ":" + text; let mapping = this.mappings.get(key);
    if (!mapping) {
      if (this.mappings.size >= 65535) throw new OfficeIMOError("RESOURCE_LIMIT", "A PDF font supports at most 65,535 distinct CID mappings.");
      const cid = this.mappings.size + 1;
      mapping = { cid, hex: cidHex(cid), glyph, width: this.program.widths[glyph]!, text }; this.mappings.set(key, mapping);
    }
    return "<" + mapping.hex + ">";
  }
  async write(objects: PdfObjects, compression: boolean): Promise<void> {
    checkAbort(this.signal);
    if (!this.program) {
      const name = "Helvetica" + (this.bold && this.italic ? "-BoldOblique" : this.bold ? "-Bold" : this.italic ? "-Oblique" : "");
      await objects.object(this.object, "<< /Type /Font /Subtype /Type1 /BaseFont /" + name + " /Encoding /WinAnsiEncoding >>"); return;
    }
    const program = this.program, entries = [...this.mappings], glyphs = new Set(entries.map(([, m]) => m.glyph));
    const fontName = (program.canSubset ? "OIMOAA+" : "") + "OfficeIMOFont" + this.index;
    const file = objects.reserve(), descriptor = objects.reserve(), cidFont = objects.reserve(), map = objects.reserve(), unicode = objects.reserve();
    const characters = new Map<number, number>();
    for (const [key, mapping] of entries) {
      if (typeof key === "number") characters.set(key, mapping.glyph);
      else for (const character of mapping.text) { const scalar = character.codePointAt(0)!, glyph = program.glyph(scalar); if (glyph) { characters.set(scalar, glyph); glyphs.add(glyph); } }
    }
    const subset = await program.subsetCooperatively(glyphs, characters, this.signal);
    await objects.stream(file, subset, "/Length1 " + subset.length, compression, true);
    await objects.object(descriptor, "<< /Type /FontDescriptor /FontName /" + fontName + " /Flags 32 /FontBBox [" + program.bbox.map(pdfNumber).join(" ") +
      "] /ItalicAngle 0 /Ascent " + pdfNumber(program.ascent) + " /Descent " + pdfNumber(program.descent) + " /CapHeight " + pdfNumber(program.ascent) + " /StemV 80 /FontFile2 " + file + " 0 R >>");
    const mapping = new Uint8Array((entries.length + 1) * 2), view = new DataView(mapping.buffer);
    for (let i = 0; i < entries.length; i++) {
      const entry = entries[i]![1]; view.setUint16(entry.cid * 2, entry.glyph);
      if ((i & 1023) === 1023) { if (taskYieldDue()) await pause(); checkAbort(this.signal); }
    }
    await objects.stream(map, mapping, "", compression, true);
    await objects.object(cidFont, "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /" + fontName +
      " /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor " + descriptor + " 0 R /CIDToGIDMap " + map + " 0 R" +
      (entries.length ? " /W [1 [" + entries.map(([, m]) => pdfNumber(m.width)).join(" ") + "]]" : "") + " >>");
    let cmap = "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /OfficeIMOUnicode def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n";
    const mapped = entries.filter(([, mapping]) => mapping.text.length);
    for (let i = 0; i < mapped.length; i += 100) {
      const batch = mapped.slice(i, i + 100); cmap += batch.length + " beginbfchar\n";
      for (const [, m] of batch) cmap += "<" + m.hex + "> <" + unicodeHex(m.text) + ">\n";
      cmap += "endbfchar\n";
      if (taskYieldDue()) await pause(); checkAbort(this.signal);
    }
    cmap += "endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n";
    await objects.stream(unicode, cmap, "", compression);
    await objects.object(this.object, "<< /Type /Font /Subtype /Type0 /BaseFont /" + fontName + " /Encoding /Identity-H /DescendantFonts [" + cidFont + " 0 R] /ToUnicode " + unicode + " 0 R >>");
  }
}
export class PdfFontResources {
  readonly resources: PdfFontResource[] = [];
  private readonly selections = new Map<number, PdfFontResource>();
  private readonly programs = new Map<PdfFont, PdfFontResource>();
  private readonly explicitSelections = new Map<PdfFont, Map<number, PdfFontResource>>();
  private readonly registered = new Set<PdfFont>();
  private fontBytes = 0;
  constructor(private readonly objects: PdfObjects, private readonly fonts?: PdfFonts, private readonly maxBytes = 16 * 1024 * 1024, private readonly signal?: AbortSignal) {
    validateFonts(fonts, maxBytes);
    for (const font of Object.values(fonts ?? {})) this.registerFont(font);
  }
  /** The same operation budget covers unused declared faces and every explicit run face, once per identity. */
  registerFont(font: PdfFont): void {
    if (this.registered.has(font)) return;
    const bytes = fontProgram(font).bytes.length;
    if (this.fontBytes + bytes > this.maxBytes) throw new OfficeIMOError("RESOURCE_LIMIT", "maxFontBytes exceeded.");
    this.fontBytes += bytes; this.registered.add(font);
  }
  select(bold = false, italic = false, explicit?: PdfFont): PdfFontResource {
    if (explicit) this.registerFont(explicit);
    let selections = this.selections;
    if (explicit) { const found = this.explicitSelections.get(explicit); selections = found ?? new Map(); if (!found) this.explicitSelections.set(explicit, selections); }
    const key = (bold ? 2 : 0) | (italic ? 1 : 0), existing = selections.get(key);
    if (existing) return existing;
    const { font, syntheticBold, syntheticItalic } = explicit ? { font: explicit, syntheticBold: bold, syntheticItalic: italic } : selectFont(this.fonts, bold, italic);
    const shared = font ? this.programs.get(font) : undefined;
    const resource = new PdfFontResource(this.resources.length + 1, bold, italic, font, this.objects, syntheticBold, syntheticItalic, shared, this.signal);
    if (!shared) { this.resources.push(resource); if (font) this.programs.set(font, resource); }
    selections.set(key, resource); return resource;
  }
  dictionary(): string { return "<< " + this.resources.map(f => "/" + f.name + " " + f.object + " 0 R").join(" ") + " >>"; }
}
