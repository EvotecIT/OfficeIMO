import { fontProgram } from "./font.js";
import { resolveScalar, selectFont, validateFonts } from "./font-support.js";
import type { PdfTextPreflightOptions, PdfTextPreflightResult } from "./types.js";

/**
 * Check one text run against the writer's selected font and scalar-layout rules.
 * Returns the first text issue at its original UTF-16 index. This does not check
 * wrapping, page geometry, resource budgets or document accessibility.
 */
export function preflightPdfText(text: string, options: PdfTextPreflightOptions = {}): PdfTextPreflightResult {
  if (typeof text !== "string") throw new TypeError("PDF preflight text must be a string.");
  if (!options || typeof options !== "object") throw new TypeError("PDF preflight options must be an object.");
  for (const key of ["bold", "italic"] as const)
    if (options[key] !== undefined && typeof options[key] !== "boolean") throw new TypeError("PDF preflight " + key + " must be boolean.");
  validateFonts(options.fonts);
  const bold = options.bold ?? false, { font } = selectFont(options.fonts, bold, options.italic ?? false);
  const program = font ? fontProgram(font) : undefined, checked = new Set<number>();
  for (let index = 0; index < text.length;) {
    const scalar = text.codePointAt(index)!;
    // wrapText treats CR/LF as paragraph boundaries and expands tabs to spaces.
    const displayed = scalar === 9 ? 32 : scalar;
    if (scalar !== 10 && scalar !== 13 && !checked.has(displayed)) {
      const result = resolveScalar(displayed, program, bold);
      if (!result.supported) return { ...result, index, ...(result.code === "unpaired-surrogate" ? {} : { scalar }) };
      checked.add(displayed);
    }
    index += scalar > 0xffff ? 2 : 1;
  }
  return { supported: true };
}
