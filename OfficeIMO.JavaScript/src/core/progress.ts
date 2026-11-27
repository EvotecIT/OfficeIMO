import type { ExportProgress } from "./index.js";

/** @internal Notify synchronously; observe an accidental promise before rejecting its unsupported use. */
export function notifyProgress(callback: ((progress: ExportProgress) => void) | undefined, progress: ExportProgress): void {
  notifySynchronously("onProgress", callback, progress);
}

/** @internal Run a host notification that must finish before the export continues. A returned promise is observed,
 * so its rejection is never unhandled, and the export fails instead of reporting success before the callback ends. */
export function notifySynchronously<T>(name: string, callback: ((value: T) => void) | undefined, value: T): void {
  if (!callback) return;
  const result: unknown = callback(value);
  if (result && typeof (result as { then?: unknown }).then === "function") {
    void Promise.resolve(result).catch(() => {});
    throw new TypeError(name + " must be synchronous.");
  }
}
