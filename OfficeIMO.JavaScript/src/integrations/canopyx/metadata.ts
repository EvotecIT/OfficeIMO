import type { CanopyCapture } from "./types.js";

/** Snapshot optional About metadata before enumeration. The accepted query is serialized, never evaluated. */
export function canopyAbout(request: CanopyCapture["request"]): readonly (readonly [string, string | number])[] {
  const result: [string, string | number][] = [["revision", request.revision], ["recordCount", request.recordCount === null ? "unknown" : request.recordCount]];
  if (request.scope !== undefined) {
    if (!["filtered", "selected", "all"].includes(request.scope)) throw new TypeError("Invalid Canopy export scope.");
    result.push(["scope", request.scope]);
  }
  for (const key of ["total", "matched"] as const) if (request[key] !== undefined) {
    const value = request[key]!;
    if (value === null) { result.push([key, "unknown"]); continue; }
    if (!Number.isSafeInteger(value) || value < 0) throw new TypeError("Canopy metadata counts must be nonnegative safe integers.");
    result.push([key, value]);
  }
  if (request.generatedAt !== undefined) {
    if (typeof request.generatedAt !== "string" || !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d{1,9})?Z$/.test(request.generatedAt) || !Number.isFinite(new Date(request.generatedAt).getTime()))
      throw new TypeError("Canopy generatedAt must be a captured UTC ISO instant.");
    result.push(["generatedAt", request.generatedAt]);
  }
  if (request.selection !== undefined) {
    if (!Array.isArray(request.selection) || request.selection.some(id => typeof id !== "string")) throw new TypeError("Canopy selection metadata must contain stable string IDs.");
    result.push(["selection", JSON.stringify(request.selection)]);
  }
  if (request.query !== undefined) {
    if (!request.query || typeof request.query !== "object" || Array.isArray(request.query)) throw new TypeError("Canopy accepted query must be a captured object.");
    const query = JSON.stringify(request.query);
    if (typeof query !== "string") throw new TypeError("Canopy accepted query must serialize as JSON metadata.");
    result.push(["query", query]);
  }
  if (result.reduce((sum, row) => sum + row[0].length + String(row[1]).length, 0) > 65536)
    throw new RangeError("Canopy About metadata exceeds 65,536 captured characters.");
  return Object.freeze(result.map(row => Object.freeze(row)));
}
