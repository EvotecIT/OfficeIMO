import { OfficeIMOError } from "../core/errors.js";
import type { PdfParagraphStyle } from "./document-types.js";
import { positive, presentation } from "./presentation-settings.js";

const fields = ["fontSize", "presentation", "keepTogether", "keepWithNext", "spaceBefore", "spaceAfter"];
const presentationFields = ["bold", "italic", "color", "alignment"];

/** @internal Copy only paragraph appearance and keep rules, without semantics or an inheritance graph. */
export function captureParagraphStyle(input: PdfParagraphStyle): Readonly<PdfParagraphStyle> {
  if (input && typeof (input as { then?: unknown }).then === "function") {
    void Promise.resolve(input).catch(() => {});
    throw new TypeError("PDF paragraph styles must be synchronous objects.");
  }
  if (!input || typeof input !== "object" || Array.isArray(input) || Object.keys(input).some(key => !fields.includes(key)))
    throw new TypeError("PDF paragraph styles support fontSize, presentation, keepTogether, keepWithNext, spaceBefore and spaceAfter.");
  const { fontSize, presentation: inputPresentation, keepTogether, keepWithNext, spaceBefore, spaceAfter } = input;
  if (fontSize !== undefined) positive(fontSize, "Paragraph font size", 144);
  for (const [name, flag] of [["keepTogether", keepTogether], ["keepWithNext", keepWithNext]] as const)
    if (flag !== undefined && typeof flag !== "boolean") throw new TypeError("Paragraph " + name + " must be boolean.");
  for (const value of [spaceBefore, spaceAfter])
    if (value !== undefined && (!Number.isFinite(value) || value < 0 || value > 14400))
      throw new RangeError("Paragraph spacing must be from zero through 14,400 points.");
  if (inputPresentation !== undefined && (!inputPresentation || typeof inputPresentation !== "object" || Array.isArray(inputPresentation) ||
    Object.keys(inputPresentation).some(key => !presentationFields.includes(key))))
    throw new TypeError("PDF paragraph presentation supports bold, italic, color and alignment.");
  return Object.freeze({
    ...(fontSize === undefined ? {} : { fontSize }),
    ...(inputPresentation === undefined ? {} : { presentation: presentation(inputPresentation) }),
    ...(keepTogether === undefined ? {} : { keepTogether }), ...(keepWithNext === undefined ? {} : { keepWithNext }),
    ...(spaceBefore === undefined ? {} : { spaceBefore }), ...(spaceAfter === undefined ? {} : { spaceAfter })
  });
}

/** @internal Snapshot a bounded own-property registry before any report source is enumerated. */
export function captureParagraphStyles(input: Readonly<Record<string, PdfParagraphStyle>> | undefined, maximum: number): Readonly<Record<string, Readonly<PdfParagraphStyle>>> {
  if (!Number.isSafeInteger(maximum) || maximum < 0) throw new RangeError("maxParagraphStyles must be a nonnegative safe integer.");
  const result: Record<string, Readonly<PdfParagraphStyle>> = Object.create(null);
  if (input === undefined) return Object.freeze(result);
  if (input && typeof (input as { then?: unknown }).then === "function") {
    void Promise.resolve(input).catch(() => {});
    throw new TypeError("paragraphStyles must be a synchronous record.");
  }
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError("paragraphStyles must be a record of named paragraph styles.");
  const names = Object.keys(input);
  if (names.length > maximum) throw new OfficeIMOError("RESOURCE_LIMIT", "maxParagraphStyles exceeded.");
  for (const name of names) {
    if (!name.trim()) throw new TypeError("Paragraph style names must be nonempty text.");
    result[name] = captureParagraphStyle(input[name]!);
  }
  return Object.freeze(result);
}

/** @internal Explicit values win; member-wise presentation preserves the remaining named appearance. */
export function resolveParagraphStyle(name: string | undefined, overrides: PdfParagraphStyle,
  styles: Readonly<Record<string, Readonly<PdfParagraphStyle>>>): Readonly<PdfParagraphStyle> {
  let named: Readonly<PdfParagraphStyle> = {};
  if (name !== undefined) {
    if (typeof name !== "string" || !Object.prototype.hasOwnProperty.call(styles, name)) throw new TypeError("Unknown PDF paragraph style: " + String(name));
    named = styles[name]!;
  }
  const captured = captureParagraphStyle(overrides);
  return Object.freeze({ ...named, ...captured,
    ...(named.presentation || captured.presentation ? { presentation: Object.freeze({ ...named.presentation, ...captured.presentation }) } : {}) });
}
