import type { ExportProgress } from "./index.js";

/** @internal Notify synchronously; observe an accidental promise before rejecting its unsupported use. */
export function notifyProgress(callback: ((progress: ExportProgress) => void) | undefined, progress: ExportProgress): void {
  if (!callback) return;
  const result: unknown = callback(progress);
  if (result && typeof (result as { then?: unknown }).then === "function") {
    void Promise.resolve(result).catch(() => {});
    throw new TypeError("onProgress must be synchronous.");
  }
}
