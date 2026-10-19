import { OfficeIMOError } from "../core/errors.js";
import { PdfFont, fontProgram } from "./font.js";
import { helveticaWidths } from "./standard-widths.js";
import type { TrueTypeFont } from "./truetype.js";
import type { PdfFonts, PdfTextIssueCode } from "./types.js";

const winAnsiExtras = new Map([0x20ac,0,0x201a,0x192,0x201e,0x2026,0x2020,0x2021,0x2c6,0x2030,0x160,0x2039,0x152,0,0x17d,0,0,0x2018,0x2019,0x201c,0x201d,0x2022,0x2013,0x2014,0x2dc,0x2122,0x161,0x203a,0x153,0,0x17e,0x178].map((cp, i) => [cp, i + 128]));

/** @internal Called after scalar support succeeds when encoding a standard-font run. */
export function standardCode(scalar: number): number | undefined {
  if ((scalar >= 32 && scalar <= 126) || (scalar >= 160 && scalar <= 255)) return scalar;
  return scalar ? winAnsiExtras.get(scalar) : undefined;
}

type ScalarResolution = { readonly supported: true; readonly width: number; readonly glyph?: number } |
  { readonly supported: false; readonly code: PdfTextIssueCode; readonly message: string };

/** @internal One encoding/glyph decision for the writer and public text preflight. */
export function resolveScalar(scalar: number, program: TrueTypeFont | undefined, bold: boolean): ScalarResolution {
  if (scalar >= 0xd800 && scalar <= 0xdfff)
    return { supported: false, code: "unpaired-surrogate", message: "PDF text contains an unpaired UTF-16 surrogate." };
  if (scalar < 32 || (scalar >= 127 && scalar < 160))
    return { supported: false, code: "unsupported-control", message: "PDF text contains an unsupported control character." };
  const character = String.fromCodePoint(scalar);
  if (/[\p{Mark}\p{Format}]/u.test(character) || !/[\p{Script=Latin}\p{Script=Greek}\p{Script=Cyrillic}\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Hangul}\p{Script=Common}]/u.test(character) ||
      (scalar >= 0x1100 && scalar <= 0x11ff) || (scalar >= 0x1f3fb && scalar <= 0x1f3ff))
    return { supported: false, code: "unsupported-layout", message: "This PDF table writer supports Unicode scalar layout; complex shaping, combining sequences and bidirectional text require a shaping-capable writer." };
  if (!program) {
    if (standardCode(scalar) === undefined)
      return { supported: false, code: "embedded-font-required", message: "Text U+" + scalar.toString(16).toUpperCase() + " needs an embedded Unicode TrueType font; supply fonts.regular." };
    const metric = helveticaWidths.get(scalar === 160 ? 32 : scalar === 173 ? 45 : scalar);
    if (!metric) return { supported: false, code: "missing-glyph", message: "Standard font has no metric for this character." };
    return { supported: true, width: metric[bold ? 1 : 0] };
  }
  let glyph = program.glyph(scalar);
  // Reuse only an inkless thin-space glyph; the source scalar keeps its own CID/ToUnicode.
  if (!glyph && scalar === 0x202f) {
    const thin = program.glyph(0x2009);
    if (thin && program.offsets[thin] === program.offsets[thin + 1]) glyph = thin;
  }
  if (!glyph) return { supported: false, code: "missing-glyph", message: "The supplied font has no glyph for U+" + scalar.toString(16).toUpperCase() + "." };
  return { supported: true, glyph, width: program.widths[glyph]! };
}

/** @internal Family validation is shared; preflight does not impose a document's byte budget. */
export function validateFonts(fonts: PdfFonts | undefined, maxBytes?: number): void {
  if (!fonts) return;
  if (!(fonts.regular instanceof PdfFont)) throw new TypeError("fonts.regular must be a PdfFont.");
  const unique = new Set<PdfFont>();
  for (const [key, font] of Object.entries(fonts)) {
    if (!["regular", "bold", "italic", "boldItalic"].includes(key) || !(font instanceof PdfFont)) throw new TypeError("Invalid PDF font family.");
    unique.add(font);
  }
  if (maxBytes !== undefined && [...unique].reduce((n, f) => n + fontProgram(f).bytes.length, 0) > maxBytes)
    throw new OfficeIMOError("RESOURCE_LIMIT", "maxFontBytes exceeded.");
}

/** @internal Missing emphasis faces select regular glyphs with the writer's synthetic styling. */
export function selectFont(fonts: PdfFonts | undefined, bold: boolean, italic: boolean): {
  readonly font: PdfFont | undefined; readonly syntheticBold: boolean; readonly syntheticItalic: boolean;
} {
  const exact = fonts && (bold && italic ? fonts.boldItalic : bold ? fonts.bold : italic ? fonts.italic : fonts.regular);
  const font = exact ?? fonts?.regular;
  return { font, syntheticBold: !!font && bold && !exact, syntheticItalic: !!font && italic && !exact };
}
