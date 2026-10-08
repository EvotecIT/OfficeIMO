import { cleanXml, escapeXml } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import type { CellValue } from "../core/index.js";
import { ExportCell, assertExportValue, assertScalar } from "../core/presentation.js";

/** A typed value plus a workbook-local style index. */
export class Cell {
  constructor(readonly value: CellValue, readonly style?: number) {}
}
/** @internal Advanced worksheets admit cells registered on their owning workbook. */
export function assertXlsxValue(value: unknown): void {
  if (value instanceof Cell) assertScalar(value.value);
  else assertExportValue(value);
}
/** @internal Buffered samples and footer definitions capture mutable Date values. */
export function copyValue(value: unknown): unknown {
  if (value instanceof Date) return new Date(value);
  if (value instanceof Cell && value.value instanceof Date) return new Cell(new Date(value.value), value.style);
  if (value instanceof ExportCell && value.value instanceof Date) return new ExportCell(new Date(value.value), { ...(value.text === undefined ? {} : { text: value.text }), ...(value.presentation ? { presentation: value.presentation } : {}), ...(value.link ? { link: value.link } : {}) });
  return value;
}

export function cellText(value: unknown, policy: InvalidCharacterPolicy = "strip"): string {
  const text = cleanXml(value, policy);
  if (text.length > 32767) throw new RangeError("Excel cell text exceeds 32,767 UTF-16 code units.");
  return text;
}

export function inlineText(value: unknown, policy: InvalidCharacterPolicy): string {
  const text = cellText(value, policy), node = (part: string) => '<t xml:space="preserve">' + escapeXml(part, policy) + '</t>';
  // OOXML escapes are decoded per run. Splitting the initial underscore preserves literal text in both reader kinds.
  return /_x[0-9a-f]{4}_/i.test(text)
    ? text.split(/(?<=_)(?=x[0-9a-f]{4}_)/gi).map(part => '<r>' + node(part) + '</r>').join("") : node(text);
}

export function columnName(index: number): string {
  if (!Number.isInteger(index) || index < 1 || index > 16384) throw new RangeError("Column index must be from 1 through 16,384.");
  let name = "";
  while (index) { index--; name = String.fromCharCode(65 + index % 26) + name; index = Math.floor(index / 26); }
  return name;
}
function clipName(text: string, length: number): string { return text.slice(0, length).replace(/[\ud800-\udbff]$/, ""); }
export function sheetName(requested: string, names: Set<string>, policy: InvalidCharacterPolicy, register = true): string {
  if (typeof requested !== "string") throw new TypeError("Sheet name must be a string.");
  // Validate and deduplicate the decoded literal value, before ST_Xstring attribute encoding.
  // A requested _xHHHH_ token is literal text; the serializer protects its leading underscore.
  let base = cleanXml(requested, policy).replace(/[\[\]:*?/\\]/g, "_").trim().replace(/^'+|'+$/g, "").trim();
  if (!base) base = "Sheet";
  if (base.toLowerCase() === "history") base += "_";
  base = clipName(base, 31);
  let name = base, suffix = 2;
  while (names.has(name.toLowerCase())) { const tail = " (" + suffix++ + ")"; name = clipName(base, 31 - tail.length) + tail; }
  if (register) names.add(name.toLowerCase()); return name;
}

export function excelDate(date: Date, mode: "local" | "utc"): number | null {
  if (!Number.isFinite(date.getTime())) return null;
  const utc = mode === "utc", year = utc ? date.getUTCFullYear() : date.getFullYear();
  if (year < 1900 || year > 9999) throw new RangeError("Excel dates must be in years 1900 through 9999.");
  const wall = new Date(0);
  wall.setUTCFullYear(year, utc ? date.getUTCMonth() : date.getMonth(), utc ? date.getUTCDate() : date.getDate());
  wall.setUTCHours(utc ? date.getUTCHours() : date.getHours(), utc ? date.getUTCMinutes() : date.getMinutes(),
    utc ? date.getUTCSeconds() : date.getSeconds(), utc ? date.getUTCMilliseconds() : date.getMilliseconds());
  const time = wall.getTime(); return (time - Date.UTC(1899, 11, 31)) / 86400000 + (time >= Date.UTC(1900, 2, 1) ? 1 : 0);
}
