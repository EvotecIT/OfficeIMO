import type { ExportValue, TableSpanRows } from "../core/index.js";
import { ExportCell } from "../core/presentation.js";
import { isDate } from "../core/dates.js";
import { tableSpans } from "../internal/table-spans.js";

/** Snapshot small report matrices; flattening is an explicit format decision, never inferred from labels. */
export function csvStructuredRows(rows: TableSpanRows, columns: number, mode: "repeat" | "anchors" | undefined): readonly (readonly ExportValue[])[] {
  const anchors = tableSpans(rows, columns), output: ExportValue[][] = anchors.map(() => Array(columns).fill(""));
  for (let row = 0; row < anchors.length; row++) for (const anchor of anchors[row]!) {
    if ((anchor.span > 1 || anchor.rowSpan > 1) && mode === undefined) throw new TypeError("CSV spanned headings/footers require an explicit spanMode: repeat or anchors.");
    const source = anchor.value, raw = source instanceof ExportCell ? source.value : source;
    const value = isDate(raw) ? source instanceof ExportCell ? new ExportCell(new Date(raw.getTime()), {
      ...(source.text === undefined ? {} : { text: source.text }), ...(source.presentation ? { presentation: source.presentation } : {}), ...(source.link ? { link: source.link } : {})
    }) : new Date(raw.getTime()) : source;
    for (let r = row; r < row + (mode === "repeat" ? anchor.rowSpan : 1); r++)
      for (let c = anchor.first; c < anchor.first + (mode === "repeat" ? anchor.span : 1); c++) output[r]![c] = value;
  }
  return output;
}
