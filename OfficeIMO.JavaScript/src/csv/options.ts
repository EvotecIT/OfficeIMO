import { captureFields } from "../internal/capture.js";
import type { CsvOptions } from "./index.js";

/** @internal Share the format's capture policy with thin adapters before they compose options. */
export function captureCsvOptions(input: unknown): CsvOptions {
  if (!input || typeof input !== "object") throw new TypeError("CSV options are required.");
  return captureFields(input as CsvOptions, ["columns", "signal", "limits", "onProgress", "valueMode", "delimiter", "lineEnding", "bom", "includeHeader", "headerRows",
    "footer", "spanMode", "formulaInjectionProtection", "quote", "nullValue", "booleanText"]);
}
