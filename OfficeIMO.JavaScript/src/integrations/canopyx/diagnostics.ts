import type { CanopyDiagnostic, CanopyFormat, CanopySourceOptions } from "./types.js";

/** One policy for capture diagnostics and unsupported adapter mappings; callbacks stay synchronous. */
export function canopyDiagnostics(format: CanopyFormat, options: CanopySourceOptions): (diagnostic: CanopyDiagnostic) => void {
  const policy = options.unsupportedPresentation ?? (format === "csv" ? "text" : "reject"), notify = options.onDiagnostic;
  if (policy !== "reject" && policy !== "text") throw new TypeError("unsupportedPresentation must be reject or text.");
  if (notify !== undefined && typeof notify !== "function") throw new TypeError("onDiagnostic must be a synchronous function.");
  return diagnostic => {
    if (policy === "reject") throw new TypeError("Unsupported Canopy presentation: " + diagnostic.code + (diagnostic.columnId ? " (" + diagnostic.columnId + ")" : ""));
    const result: unknown = notify?.(Object.freeze(diagnostic));
    if (result && typeof (result as { then?: unknown }).then === "function") {
      void Promise.resolve(result).catch(() => {}); throw new TypeError("onDiagnostic must be synchronous.");
    }
  };
}
