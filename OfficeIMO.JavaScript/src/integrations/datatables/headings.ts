import type { ExportValue } from "../../core/index.js";
import { ExportCell, assertScalar } from "../../core/presentation.js";
import { member } from "./api.js";

/** @internal */
export function value(input: unknown): ExportValue {
  if (input instanceof ExportCell) return input;
  assertScalar(input); return input;
}
/** @internal */
export function text(input: unknown): string {
  const scalar = value(input);
  return String(scalar instanceof ExportCell ? scalar.text ?? scalar.value ?? "" : scalar ?? "");
}
/** @internal Horizontal hierarchy maps to the existing shared Column.groups contract. */
export function headings(structure: unknown, leaf: readonly unknown[], mode: "grouped" | "leaf"): { rows: ExportValue[][]; groups: string[][] } {
  const groups = leaf.map(() => [] as string[]);
  if (mode === "leaf" || !Array.isArray(structure) || !structure.length) return { rows: [leaf.map(value)], groups };
  const rows = structure.map((source: unknown, level: number) => {
    if (!Array.isArray(source) || source.length !== leaf.length) throw new TypeError("Invalid DataTables heading structure.");
    const row = leaf.map(() => "" as ExportValue);
    for (let index = 0; index < source.length;) {
      const cell: unknown = source[index];
      const span = member(cell, "colspan"), height = member(cell, "rowspan");
      if (cell == null || height !== 1 || !Number.isInteger(span) || (span as number) < 1 || index + (span as number) > leaf.length)
        throw new TypeError("Grouped export requires horizontal header spans; select headings: 'leaf' for vertical spans.");
      if (level === structure.length - 1 && span !== 1) throw new TypeError("Grouped export requires one leaf header per column.");
      const title = text(member(cell, "title")); row[index] = title;
      if (span !== 1 && !title) throw new TypeError("Blank spanning headings require headings: 'leaf'.");
      for (let covered = index; covered < index + (span as number); covered++) {
        if (covered !== index && source[covered] != null) throw new TypeError("Invalid overlapping DataTables heading cells.");
        if (level < structure.length - 1) groups[covered]!.push(title);
      }
      index += span as number;
    }
    return row;
  });
  // The shared layout merges equal adjacent group paths. Reject distinct source cells
  // that would otherwise become one merge instead of changing their structure silently.
  for (let level = 0; level < rows.length - 1; level++) {
    const source = structure[level] as unknown[];
    for (let index = 1; index < source.length; index++) {
      if (source[index] != null && groups[index]![level] &&
        JSON.stringify(groups[index]!.slice(0, level + 1)) === JSON.stringify(groups[index - 1]!.slice(0, level + 1)))
        throw new TypeError("Adjacent separate headings with equal group paths require headings: 'leaf'.");
    }
  }
  return { rows, groups };
}
