import { TrueTypeFont } from "./truetype.js";
const programs = new WeakMap<PdfFont, TrueTypeFont>();
const fontBrand = Symbol.for("@evotecit/officeimo/PdfFont");

/** @internal Font bytes stay private even in plain JavaScript consumers. */
export function fontProgram(font: PdfFont): TrueTypeFont {
  let program = programs.get(font);
  if (!program) {
    if (!(font instanceof PdfFont) || typeof font.toBytes !== "function") throw new TypeError("Invalid PdfFont instance.");
    // Standalone ESM assemblies have separate class/WeakMap identities. Import a
    // defensive copy once while preserving the same public reusable-font contract.
    const copied = new PdfFont(font.toBytes());
    program = programs.get(copied)!; programs.set(font, program);
  }
  return program;
}

/** Immutable, reusable static TrueType font. Bytes are copied; no fetch, DOM or installed-font lookup occurs. */
export class PdfFont {
  static [Symbol.hasInstance](value: unknown): boolean { return !!value && typeof value === "object" && (value as Record<symbol, unknown>)[fontBrand] === true; }
  private readonly brand = true;
  constructor(bytes: Uint8Array) {
    if (!(bytes instanceof Uint8Array)) throw new TypeError("PdfFont requires TrueType bytes as a Uint8Array.");
    if (bytes.length > 64 * 1024 * 1024) throw new RangeError("A TrueType font must not exceed 64 MiB.");
    programs.set(this, new TrueTypeFont(bytes.slice()));
    void this.brand;
    Object.defineProperty(this, fontBrand, { value: true });
    Object.freeze(this);
  }
  /** Return an independent copy, for storage or delivery to another worker. */
  toBytes(): Uint8Array { return fontProgram(this).bytes.slice(); }
}
