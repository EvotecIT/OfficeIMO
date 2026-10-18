using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    /// <summary>
    /// Retains ordinary row cuts when every active rowspan can share the source slice.
    /// The active sweep is bounded by the occupied columns; each span's content is
    /// inspected once rather than walking its visual tree at every later row.
    /// </summary>
    private IReadOnlyList<bool> ResolveRowBreakAvailability(
        IReadOnlyList<TableRowLayout> rows, double spacing, out IReadOnlyList<bool> hasSpanningCellsAfterRows) {
        var result = new bool[rows.Count];
        var spanning = new bool[rows.Count];
        var active = new List<TableSpanCutSafety>();
        var avoidedRows = new int[rows.Count + 1];
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            avoidedRows[rowIndex + 1] = avoidedRows[rowIndex] + (rows[rowIndex].Style.AvoidBreakInside ? 1 : 0);
        double rowY = 0D;
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            CheckCancellation();
            TableRowLayout row = rows[rowIndex];
            active.RemoveAll(span => span.LastRow <= rowIndex);
            foreach (TableCellLayout cell in row.Cells) {
                ChargeLayoutOperation("table rowspan break ownership");
                if (cell.RowSpan > 1) active.Add(CreateTableSpanCutSafety(cell,
                    rowIndex + cell.RowSpan - 1, rowY + cell.ContentOffsetY,
                    !row.IsHeader && !row.IsFooter
                        && avoidedRows[rowIndex + cell.RowSpan] == avoidedRows[rowIndex]));
            }
            rowY += row.Height + spacing;
            spanning[rowIndex] = active.Count > 0;
            bool safe = true;
            foreach (TableSpanCutSafety span in active) {
                ChargeLayoutOperation("table active rowspan cuts");
                if (!IsSafeTableSpanCut(span, rowY)) {
                    safe = false;
                    break;
                }
            }
            result[rowIndex] = safe;
        }

        for (int rowIndex = 0; rowIndex + 1 < rows.Count; rowIndex++) {
            TableRowLayout current = rows[rowIndex];
            TableRowLayout next = rows[rowIndex + 1];
            if (current.Style.AvoidBreakAfter || next.Style.AvoidBreakBefore) result[rowIndex] = false;
            if (current.GroupElement != null
                && ReferenceEquals(current.GroupElement, next.GroupElement)
                && current.GroupStyle?.AvoidBreakInside == true) result[rowIndex] = false;
        }
        hasSpanningCellsAfterRows = spanning;
        return result;
    }

    private TableSpanCutSafety CreateTableSpanCutSafety(
        TableCellLayout cell, int lastRow, double contentY, bool ordinaryBody) {
        var ranges = new List<HtmlRenderAvoidBreakRange>(cell.Inline.AvoidBreakRanges);
        bool supported = ordinaryBody && !cell.Style.AvoidBreakInside;
        if (!supported) return new TableSpanCutSafety(lastRow, contentY, cell.Inline.Height,
            cell.Inline.BreakOffsets, Array.Empty<HtmlRenderAvoidBreakRange>(), false);
        foreach (HtmlRenderVisual visual in cell.Inline.Visuals) Collect(visual);
        var merged = new List<HtmlRenderAvoidBreakRange>();
        foreach (HtmlRenderAvoidBreakRange range in ranges.OrderBy(range => range.Start).ThenBy(range => range.End)) {
            ChargeLayoutOperation("table rowspan cut ranges");
            // Adjacent line carriers share a legal endpoint. Merge overlapping
            // ranges only, so that endpoint remains available to the row cut.
            if (merged.Count > 0 && range.Start < merged[merged.Count - 1].End - 0.0001D) {
                HtmlRenderAvoidBreakRange previous = merged[merged.Count - 1];
                merged[merged.Count - 1] = previous.WithEnd(Math.Max(previous.End, range.End));
            } else {
                merged.Add(range);
            }
        }
        return new TableSpanCutSafety(lastRow, contentY, cell.Inline.Height,
            cell.Inline.BreakOffsets, merged, supported);

        void Collect(HtmlRenderVisual visual) {
            CheckCancellation();
            ChargeLayoutOperation("table rowspan content carriers");
            if (visual is HtmlRenderEffectGroup effect && !effect.Transform.Equals(OfficeTransform.Identity)) {
                // Transformed content needs its own painted-coordinate qualification.
                // Keep the previous cohesive-span boundary for this source cell.
                supported = false;
                return;
            }
            if (visual is HtmlRenderSemanticGroup table && IsTableFormattingGroup(table)) {
                // Flattening cell blocks does not make the nested table's repeat
                // reservations part of this table. Keep that nested table whole.
                ranges.Add(new HtmlRenderAvoidBreakRange(table.LayoutY,
                    Math.Max(table.LayoutY + table.LayoutHeight, MaximumScrollBottom(table.Visuals))));
                return;
            }
            if (GetGroupChildren(visual) is { } children) {
                foreach (HtmlRenderVisual child in children) Collect(child);
                return;
            }
            if (visual is HtmlRenderShape { IsAtomicReplacedPlaceholder: false }
                || visual is HtmlRenderAnchorFragment
                || visual is HtmlRenderBookmarkAnchor
                || visual is HtmlRenderNamedDestination
                || visual is HtmlRenderLayoutBox) return;
            // Text uses its used line carrier, not its glyph ink or CSS leading.
            // Images, drawings, form controls and other replaced leaves stay whole.
            ranges.Add(new HtmlRenderAvoidBreakRange(visual.LayoutY, visual.LayoutY + visual.LayoutHeight));
        }
    }

    private static bool IsSafeTableSpanCut(TableSpanCutSafety span, double rowEnd) {
        if (!span.Supported) return false;
        double relative = rowEnd - span.ContentY;
        if (relative > 0.0001D && relative < span.ContentHeight - 0.0001D) {
            int next = UpperBound(span.BreakOffsets, relative + 0.0001D);
            if (next == 0 || Math.Abs(span.BreakOffsets[next - 1] - relative) > 0.0001D) return false;
        }
        int low = 0;
        int high = span.UnsafeRanges.Count;
        while (low < high) {
            int middle = low + ((high - low) >> 1);
            if (span.UnsafeRanges[middle].Start < relative - 0.0001D) low = middle + 1;
            else high = middle;
        }
        return low == 0 || relative >= span.UnsafeRanges[low - 1].End - 0.0001D;
    }

    /// <summary>Projects the spanning cell's paragraph constraints into the shared table source coordinates.</summary>
    private void AppendSpanningCellLineBreakGroups(
        ICollection<HtmlRenderLineBreakGroup> groups, TableCellLayout cell, double contentY) {
        IReadOnlyList<HtmlRenderLineBreakGroup> source = cell.Inline.LineBreakGroups;
        if (source.Count == 0 && cell.Inline.LineBreakOffsets.Count > 0) {
            // Ordinary inline construction leaves default paragraph grouping to
            // its block owner. A table cell supplies that owner here.
            source = new[] { new HtmlRenderLineBreakGroup(
                cell.Inline.LineBreakOffsets.Where(offset => offset < cell.Inline.Height - 0.0001D),
                cell.Style.Orphans, cell.Style.Widows,
                hasImplicitFinalLine: cell.Inline.LineBreakOffsets.Any(offset => Math.Abs(offset - cell.Inline.Height) <= 0.0001D),
                end: cell.Inline.Height) };
        }
        foreach (HtmlRenderLineBreakGroup group in source) {
            ChargeLayoutOperation("table rowspan paragraph constraints");
            groups.Add(group.WithInteriorBreaks().Translate(contentY));
        }
    }

    private sealed record TableSpanCutSafety(
        int LastRow, double ContentY, double ContentHeight, IReadOnlyList<double> BreakOffsets,
        IReadOnlyList<HtmlRenderAvoidBreakRange> UnsafeRanges, bool Supported);

    private static void AddTableRowInternalBreakOffsets(ICollection<double> breakOffsets, TableRowLayout row, double rowY) {
        var candidates = new SortedSet<double>();
        foreach (TableCellLayout cell in row.Cells) {
            double contentTop = cell.ContentOffsetY;
            foreach (double offset in cell.Inline.BreakOffsets) {
                // These offsets already end actual laid-out lines. Subtracting
                // computed CSS leading can discard the penultimate line when
                // its selected font has a different used normal advance.
                if (offset < cell.Inline.Height - 0.0001D) candidates.Add(contentTop + offset);
            }
        }

        foreach (double candidate in candidates) {
            bool safeForEveryCell = true;
            foreach (TableCellLayout cell in row.Cells) {
                double contentTop = cell.ContentOffsetY;
                double relative = candidate - contentTop;
                if (relative <= 0.0001D) continue;
                if (relative >= cell.Inline.Height - 0.0001D) continue;
                if (!cell.Inline.BreakOffsets.Any(offset => Math.Abs(offset - relative) <= 0.0001D)) {
                    safeForEveryCell = false;
                    break;
                }
            }
            if (safeForEveryCell && candidate > 0.0001D && candidate < row.Height - 0.0001D) {
                breakOffsets.Add(rowY + candidate);
            }
        }
    }
}
