import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import { PdfFont, fontProgram } from "./font.js";
import { helveticaWidths } from "./standard-widths.js";
import { pdfNumber, unicodeHex } from "./objects.js";
import type { PdfObjects } from "./objects.js";
import type { PdfFonts } from "./types.js";
import type { TrueTypeFont } from "./truetype.js";

const winAnsiExtras = new Map([0x20ac,0,0x201a,0x192,0x201e,0x2026,0x2020,0x2021,0x2c6,0x2030,0x160,0x2039,0x152,0,0x17d,0,0,0x2018,0x2019,0x201c,0x201d,0x2022,0x2013,0x2014,0x2dc,0x2122,0x161,0x203a,0x153,0,0x17e,0x178].map((cp, i) => [cp, i + 128]));
function standardCode(scalar: number): number {
  if ((scalar >= 32 && scalar <= 126) || (scalar >= 160 && scalar <= 255)) return scalar;
  const code = scalar ? winAnsiExtras.get(scalar) : undefined;
  if (code !== undefined) return code;
  throw new NotSupportedError("Text U+" + scalar.toString(16).toUpperCase() + " needs an embedded Unicode TrueType font; supply fonts.regular.");
}
/** Glyph fallback must be explicit; unsupported shaping never produces a plausible-looking damaged report. */
export function validateScalar(scalar: number): void {
  if (scalar >= 0xd800 && scalar <= 0xdfff) throw new TypeError("PDF text contains an unpaired UTF-16 surrogate.");
  if (scalar < 32 || (scalar >= 127 && scalar < 160)) throw new NotSupportedError("PDF text contains an unsupported control character.");
  const character = String.fromCodePoint(scalar);
  if (/[\p{Mark}\p{Format}]/u.test(character) || !/[\p{Script=Latin}\p{Script=Greek}\p{Script=Cyrillic}\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Common}]/u.test(character) ||
      (scalar >= 0x1100 && scalar <= 0x11ff) || (scalar >= 0x1f3fb && scalar <= 0x1f3ff))
    throw new NotSupportedError("This PDF table writer supports Unicode scalar layout; complex shaping, combining sequences and bidirectional text require a shaping-capable writer.");
}
export class PdfFontResource {
  readonly name: string;
  readonly object: number;
  readonly program: TrueTypeFont | undefined;
  readonly syntheticBold: boolean;
  readonly syntheticItalic: boolean;
  private readonly mappings = new Map<number, { cid: number; glyph: number; width: number }>();
  constructor(readonly index: number, readonly bold: boolean, readonly italic: boolean, font: PdfFont | undefined,
    objects: PdfObjects, syntheticBold: boolean, syntheticItalic: boolean, private readonly shared?: PdfFontResource) {
    this.name = shared?.name ?? "F" + index; this.object = shared?.object ?? objects.reserve(); this.program = font ? fontProgram(font) : undefined;
    this.syntheticBold = syntheticBold; this.syntheticItalic = syntheticItalic;
  }
  width(scalar: number): number {
    if (this.shared) return this.shared.width(scalar);
    const existing = this.mappings.get(scalar);
    if (existing) return existing.width;
    validateScalar(scalar);
    if (!this.program) {
      standardCode(scalar);
      const metric = helveticaWidths.get(scalar === 160 ? 32 : scalar === 173 ? 45 : scalar);
      if (!metric) throw new NotSupportedError("Standard font has no metric for this character.");
      return metric[this.bold ? 1 : 0];
    }
    const glyph = this.program.glyph(scalar);
    if (!glyph) throw new NotSupportedError("The supplied font has no glyph for U+" + scalar.toString(16).toUpperCase() + ".");
    if (this.mappings.size >= 65535) throw new OfficeIMOError("RESOURCE_LIMIT", "A PDF font supports at most 65,535 distinct Unicode scalars.");
    const mapping = { cid: this.mappings.size + 1, glyph, width: this.program.widths[glyph]! };
    this.mappings.set(scalar, mapping); return mapping.width;
  }
  encode(text: string): string {
    if (this.shared) return this.shared.encode(text);
    let hex = "";
    for (const char of text) {
      const cp = char.codePointAt(0)!; this.width(cp);
      hex += (this.program ? this.mappings.get(cp)!.cid : standardCode(cp)).toString(16).padStart(this.program ? 4 : 2, "0");
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
    const subset = program.subset(glyphs);
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
    if (fonts) {
      if (!(fonts.regular instanceof PdfFont)) throw new TypeError("fonts.regular must be a PdfFont.");
      const unique = new Set<PdfFont>();
      for (const [key, font] of Object.entries(fonts)) {
        if (!["regular", "bold", "italic", "boldItalic"].includes(key) || !(font instanceof PdfFont)) throw new TypeError("Invalid PDF font family.");
        unique.add(font);
      }
      if ([...unique].reduce((n, f) => n + fontProgram(f).bytes.length, 0) > maxBytes) throw new OfficeIMOError("RESOURCE_LIMIT", "maxFontBytes exceeded.");
    }
  }
  select(bold = false, italic = false): PdfFontResource {
    const key = String(bold) + String(italic), existing = this.selections.get(key);
    if (existing) return existing;
    const exact = this.fonts && (bold && italic ? this.fonts.boldItalic : bold ? this.fonts.bold : italic ? this.fonts.italic : this.fonts.regular);
    const font = exact ?? this.fonts?.regular;
    const shared = font ? this.programs.get(font) : undefined;
    const resource = new PdfFontResource(this.resources.length + 1, bold, italic, font, this.objects, !!font && bold && !exact, !!font && italic && !exact, shared);
    if (!shared) { this.resources.push(resource); if (font) this.programs.set(font, resource); }
    this.selections.set(key, resource); return resource;
  }
  dictionary(): string { return "<< " + this.resources.map(f => "/" + f.name + " " + f.object + " 0 R").join(" ") + " >>"; }
}
