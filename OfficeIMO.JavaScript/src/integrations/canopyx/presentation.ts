import type { CanopyExportColumn, CanopyNumberFormat, CanopyDiagnostic } from "./types.js";

/** Calibri 11, 96-DPI CSS pixels, seven-pixel maximum digit and five-pixel cell padding.
 * Quantize to native 1/256-character units; this is a reproducible profile, not pixel-identical rendering. */
export function canopyXlsxWidth(pixels: number): number {
  const width = Math.floor((pixels < 12 ? pixels / 12 : (pixels - 5) / 7) * 256) / 256;
  if (!Number.isFinite(width) || width <= 0 || width > 255) throw new RangeError("Captured width exceeds the Calibri-11 XLSX width profile.");
  return width;
}

/** Escape literal prefixes/suffixes; no browser formatter or CSS becomes workbook code. */
export function canopyNumberFormat(input: CanopyNumberFormat, report: (diagnostic: CanopyDiagnostic) => void, id: string): string | undefined {
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError("Canopy numeric format must be an object.");
  if (Object.keys(input).some(key => !["minimumFractionDigits", "maximumFractionDigits", "prefix", "suffix"].includes(key))) {
    report({ code: "UNSUPPORTED_NUMBER_FORMAT", columnId: id }); return undefined;
  }
  const { minimumFractionDigits: minimum = 0, maximumFractionDigits: maximum = 3, prefix = "", suffix = "" } = input;
  if (!Number.isInteger(minimum) || !Number.isInteger(maximum) || minimum < 0 || maximum < minimum || maximum > 10 ||
      typeof prefix !== "string" || typeof suffix !== "string" || prefix.length > 16 || suffix.length > 16)
    throw new TypeError("Canopy numeric fraction bounds must be ordered integers from zero through ten, with literal affixes of at most sixteen characters.");
  const literal = (value: string) => value ? '"' + value.replace(/["\\]/g, character => "\\" + character) + '"' : "";
  return literal(prefix) + "0" + (maximum ? "." + "0".repeat(minimum) + "#".repeat(maximum - minimum) : "") + literal(suffix);
}

/** Only a contiguous leading left pin can be represented by an Excel freeze pane. */
export function canopyPins(columns: readonly CanopyExportColumn[], report: (diagnostic: CanopyDiagnostic) => void): number {
  let leading = 0, ended = false;
  for (const column of columns) {
    const pin = column.pinned;
    if (pin !== undefined && pin !== null && pin !== "left" && pin !== "right") throw new TypeError("Invalid Canopy pin.");
    if (!ended && pin === "left") leading++;
    else {
      ended = true;
      if (pin === "left" || pin === "right") report({ code: "UNREPRESENTABLE_PIN:" + pin, columnId: column.id });
    }
  }
  return leading;
}
