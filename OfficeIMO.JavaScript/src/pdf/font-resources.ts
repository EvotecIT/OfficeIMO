import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import { fontProgram } from "./font.js";
import type { PdfFont } from "./font.js";
import { resolveScalar, selectFont, standardCode, validateFonts } from "./font-support.js";
import { pdfNumber, unicodeHex } from "./objects.js";
import type { PdfObjects } from "./objects.js";
import type { PdfFonts } from "./types.js";
import type { TrueTypeFont } from "./truetype.js";

export class PdfFontResource {
  readonly name: string;
  readonly object: number;
  readonly program: TrueTypeFont | undefined;
  readonly syntheticBold: boolean;
  readonly syntheticItalic: boolean;
  private readonly mappings = new Map<number, { cid: number; glyph: number; width: number }>();
  private readonly standardWidths = new Map<number, number>();
  constructor(readonly index: number, readonly bold: boolean, readonly italic: boolean, font: PdfFont | undefined,
    objects: PdfObjects, syntheticBold: boolean, syntheticItalic: boolean, private readonly shared?: PdfFontResource) {
    this.name = shared?.name ?? "F" + index; this.object = shared?.object ?? objects.reserve(); this.program = font ? fontProgram(font) : undefined;
    this.syntheticBold = syntheticBold; this.syntheticItalic = syntheticItalic;
  }
  width(scalar: number): number {
    if (this.shared) return this.shared.width(scalar);
    const existing = this.mappings.get(scalar);
    if (existing) return existing.width;
    const standard = this.standardWidths.get(scalar);
    if (standard !== undefined) return standard;
    const resolved = resolveScalar(scalar, this.program, this.bold);
    if (!resolved.supported) {
      if (resolved.code === "unpaired-surrogate") throw new TypeError(resolved.message);
      throw new NotSupportedError(resolved.message);
    }
    if (!this.program) { this.standardWidths.set(scalar, resolved.width); return resolved.width; }
    if (this.mappings.size >= 65535) throw new OfficeIMOError("RESOURCE_LIMIT", "A PDF font supports at most 65,535 distinct Unicode scalars.");
    const mapping = { cid: this.mappings.size + 1, glyph: resolved.glyph!, width: resolved.width };
    this.mappings.set(scalar, mapping); return mapping.width;
  }
  lineHeight(size: number): number {
    return Math.max(1.3, this.program ? (this.program.ascent - this.program.descent) / 1000 + .1 : 1.3) * size;
  }
  encode(text: string): string {
    if (this.shared) return this.shared.encode(text);
    let hex = "";
    for (const char of text) {
      const cp = char.codePointAt(0)!; this.width(cp);
      hex += (this.program ? this.mappings.get(cp)!.cid : standardCode(cp)!).toString(16).padStart(this.program ? 4 : 2, "0");
    }
    return "<" + hex + ">";
  }
  async write(objects: PdfObjects, compression: boolean): Promise<void> {
    if (!this.program) {
      const name = "Helvetica" + (this.bold && this.italic ? "-BoldOblique" : this.bold ? "-Bold" : this.italic ? "-Oblique" : "");
      await objects.object(this.object, "<< /Type /Font /Subtype /Type1 /BaseFont /" + name + " /Encoding /WinAnsiEncoding >>"); return;
    }
    const program = this.program, entries = [...this.mappings], glyphs = new Set(entries.map(([, m]) => m.glyph));
    const fontName = (program.canSubset ? "OIMOAA+" : "") + "OfficeIMOFont" + this.index;
    const file = objects.reserve(), descriptor = objects.reserve(), cidFont = objects.reserve(), map = objects.reserve(), unicode = objects.reserve();
    const subset = program.subset(glyphs, new Map(entries.map(([scalar, mapping]) => [scalar, mapping.glyph])));
    await objects.stream(file, subset, "/Length1 " + subset.length, compression);
    await objects.object(descriptor, "<< /Type /FontDescriptor /FontName /" + fontName + " /Flags 32 /FontBBox [" + program.bbox.map(pdfNumber).join(" ") +
      "] /ItalicAngle 0 /Ascent " + pdfNumber(program.ascent) + " /Descent " + pdfNumber(program.descent) + " /CapHeight " + pdfNumber(program.ascent) + " /StemV 80 /FontFile2 " + file + " 0 R >>");
    const mapping = new Uint8Array((entries.length + 1) * 2), view = new DataView(mapping.buffer);
    for (const [, entry] of entries) view.setUint16(entry.cid * 2, entry.glyph);
    await objects.stream(map, mapping, "", compression);
    await objects.object(cidFont, "<< /Type /Font /Subtype /CIDFontType2 /BaseFont /" + fontName +
      " /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor " + descriptor + " 0 R /CIDToGIDMap " + map + " 0 R" +
      (entries.length ? " /W [1 [" + entries.map(([, m]) => pdfNumber(m.width)).join(" ") + "]]" : "") + " >>");
    let cmap = "/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /OfficeIMOUnicode def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n";
    for (let i = 0; i < entries.length; i += 100) {
      const batch = entries.slice(i, i + 100); cmap += batch.length + " beginbfchar\n";
      for (const [cp, m] of batch) cmap += "<" + m.cid.toString(16).padStart(4, "0") + "> <" + unicodeHex(String.fromCodePoint(cp)) + ">\n";
      cmap += "endbfchar\n";
    }
    cmap += "endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n";
    await objects.stream(unicode, cmap, "", compression);
    await objects.object(this.object, "<< /Type /Font /Subtype /Type0 /BaseFont /" + fontName + " /Encoding /Identity-H /DescendantFonts [" + cidFont + " 0 R] /ToUnicode " + unicode + " 0 R >>");
  }
}
export class PdfFontResources {
  readonly resources: PdfFontResource[] = [];
  private readonly selections = new Map<string, PdfFontResource>();
  private readonly programs = new Map<PdfFont, PdfFontResource>();
  constructor(private readonly objects: PdfObjects, private readonly fonts?: PdfFonts, maxBytes = 16 * 1024 * 1024) {
    validateFonts(fonts, maxBytes);
  }
  select(bold = false, italic = false): PdfFontResource {
    const key = String(bold) + String(italic), existing = this.selections.get(key);
    if (existing) return existing;
    const { font, syntheticBold, syntheticItalic } = selectFont(this.fonts, bold, italic);
    const shared = font ? this.programs.get(font) : undefined;
    const resource = new PdfFontResource(this.resources.length + 1, bold, italic, font, this.objects, syntheticBold, syntheticItalic, shared);
    if (!shared) { this.resources.push(resource); if (font) this.programs.set(font, resource); }
    this.selections.set(key, resource); return resource;
  }
  dictionary(): string { return "<< " + this.resources.map(f => "/" + f.name + " " + f.object + " 0 R").join(" ") + " >>"; }
}
