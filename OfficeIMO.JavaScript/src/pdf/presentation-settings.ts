import type { CellPresentation } from "../core/index.js";

/** Pure validation shared by document settings and named paragraph styles. */
export function positive(value: number, name: string, max = 14400): number {
  if (!Number.isFinite(value) || value <= 0 || value > max) throw new RangeError(name + " must be positive and at most " + max + ".");
  return value;
}
export function color(value: string): string {
  if (typeof value !== "string" || !/^#?[0-9a-f]{6}$/i.test(value)) throw new TypeError("PDF colors must be six-digit RGB hex strings.");
  const hex = value.replace(/^#/, "");
  return [0, 2, 4].map(i => String(Math.round(parseInt(hex.slice(i, i + 2), 16) / 255 * 10000) / 10000)).join(" ");
}
export function presentation(value: CellPresentation = {}): Readonly<CellPresentation> {
  if (!value || typeof value !== "object") throw new TypeError("PDF presentation must be an object.");
  const snapshot = { ...value };
  if (snapshot.background !== undefined) color(snapshot.background);
  if (snapshot.color !== undefined) color(snapshot.color);
  for (const key of ["bold", "italic", "wrapText"] as const) if (snapshot[key] !== undefined && typeof snapshot[key] !== "boolean") throw new TypeError("PDF presentation " + key + " must be boolean.");
  if (snapshot.alignment !== undefined && !["left", "center", "right"].includes(snapshot.alignment)) throw new TypeError("PDF table alignment supports left, center and right.");
  return Object.freeze(snapshot);
}
