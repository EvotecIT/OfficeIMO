import { TrueTypeFont } from "./truetype.js";
import { beginTask, checkAbort } from "../core/iteration.js";
const programs = new WeakMap<PdfFont, TrueTypeFont>();
const fontBrand = Symbol.for("@evotecit/officeimo/PdfFont");

/** Intake limits for asynchronous font preparation; the source is copied before yielding. */
export interface PdfFontPreparationOptions {
  readonly signal?: AbortSignal;
  /** Defaults to 16 MiB; cannot exceed the constructor's 64 MiB hard ceiling. */
  readonly maxBytes?: number;
}

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
  /** Prepare the same immutable font with cooperative parsing and cancellation.
   * The bounded source snapshot is synchronous so later host mutation cannot change the prepared font. */
  static async create(bytes: Uint8Array, options: PdfFontPreparationOptions = {}): Promise<PdfFont> {
    checkAbort(options.signal);
    if (!(bytes instanceof Uint8Array)) throw new TypeError("PdfFont requires TrueType bytes as a Uint8Array.");
    const maximum = options.maxBytes ?? 16 * 1024 * 1024;
    if (!Number.isSafeInteger(maximum) || maximum < 0 || maximum > 64 * 1024 * 1024)
      throw new RangeError("Font preparation maxBytes must be a nonnegative safe integer at most 64 MiB.");
    if (bytes.length > maximum) throw new RangeError("TrueType font exceeds the preparation byte limit.");
    beginTask();
    const program = await TrueTypeFont.create(new Uint8Array(bytes), options.signal);
    checkAbort(options.signal);
    const font = Object.create(PdfFont.prototype) as PdfFont;
    programs.set(font, program);
    Object.defineProperty(font, "brand", { value: true });
    Object.defineProperty(font, fontBrand, { value: true });
    Object.freeze(font); return font;
  }
  /** Return an independent copy, for storage or delivery to another worker. */
  toBytes(): Uint8Array { return fontProgram(this).bytes.slice(); }
}
