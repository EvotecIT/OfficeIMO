import type { TableSpanRows, ExportValue } from "../core/index.js";
import { assertExportValue } from "../core/presentation.js";
export interface SpanAnchor { readonly first: number; readonly span: number; readonly rowSpan: number; readonly value: ExportValue; }
/** Validate one small, rectangular heading/footer matrix before any drawing. */
export function tableSpans(rows: TableSpanRows, columns: number): readonly (readonly SpanAnchor[])[] {
  if (!Array.isArray(rows) || !rows.length || rows.length > 16) throw new RangeError("Spanned headings/footers need from 1 through 16 rows.");
  const covered = Array.from({ length: rows.length }, () => new Uint8Array(columns));
  return rows.map((row, level) => {
    if (!Array.isArray(row) || row.length !== columns) throw new TypeError("Spanned rows must have the declared column count.");
    const anchors: SpanAnchor[] = [];
    for (let first = 0; first < columns; first++) {
      const cell = row[first];
      if (covered[level]![first]) { if (cell !== null) throw new TypeError("Covered heading/footer cells must be null."); continue; }
      if (!cell || typeof cell !== "object") throw new TypeError("Uncovered heading/footer cells need an anchor value.");
      const span = cell.columnSpan ?? 1, rowSpan = cell.rowSpan ?? 1;
      if (!Number.isInteger(span) || !Number.isInteger(rowSpan) || span < 1 || rowSpan < 1 || first + span > columns || level + rowSpan > rows.length) throw new RangeError("Heading/footer spans must stay inside the declared matrix.");
      assertExportValue(cell.value);
      for (let r = level; r < level + rowSpan; r++) for (let c = first; c < first + span; c++) {
        if (covered[r]![c]) throw new TypeError("Heading/footer spans overlap.");
        if ((r !== level || c !== first) && rows[r]?.[c] !== null) throw new TypeError("Covered heading/footer cells must be null.");
        covered[r]![c] = 1;
      }
      anchors.push(Object.freeze({ first, span, rowSpan, value: cell.value }));
    }
    return Object.freeze(anchors);
  });
}
