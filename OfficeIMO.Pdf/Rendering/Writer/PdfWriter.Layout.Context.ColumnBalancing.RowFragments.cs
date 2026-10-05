namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Preserves a fitting kept trailing group and row measurement for a kept leading block.</summary>
        private static ColumnBalanceUnit JoinColumnBalanceUnits(ColumnBalanceUnit first, ColumnBalanceUnit second) {
            // A fitting last row and its next block must remain atomic, just as in ordinary table flow.
            if (first.RowFragment != null) return new ColumnBalanceUnit(first.Height + second.SpacingBefore + second.Height, first.ContinuationHeight);
            if (first.RowFragment == null && second.RowFragment is { } leading)
                return new ColumnBalanceUnit(leading.WithSpacing(leading.Before + first.Height,
                    leading.MovedBefore + first.Height + first.ContinuationHeight, leading.After), first.SpacingBefore);
            return new ColumnBalanceUnit(first.Height + second.SpacingBefore + second.Height, first.ContinuationHeight, first.SpacingBefore);
        }

        private static bool PackColumnBalanceRowFragment(ColumnBalanceRowFragment fragment, double height,
            int columnCount, ref int columns, ref double used, double continuationPadding = 0D, double spacingBefore = 0D) {
            int start = fragment.StartLine;
            double before = fragment.Before + (used > continuationPadding + .001D ? spacingBefore : 0D);
            while (start < fragment.Rows.LineCounts[fragment.RowIndex]) {
                int remaining = fragment.Rows.LineCounts[fragment.RowIndex] - start;
                double whole = fragment.Measure(start, remaining);
                if (used + before + whole + fragment.After <= height + 0.001D) {
                    used += before + whole + fragment.After;
                    return true;
                }
                double available = height - used - before;
                int take = fragment.Fit(start, available, used + before > continuationPadding + 0.001D);
                // The last fragment must also leave room for a kept trailing block or footer.
                if (take == remaining)
                    take = fragment.Fit(start, available, used + before > continuationPadding + 0.001D, remaining - 1);
                if (take == 0) {
                    if (used <= continuationPadding + 0.001D) return false;
                    before = start == fragment.StartLine ? fragment.MovedBefore : fragment.SplitBefore;
                } else {
                    start += take;
                    before = fragment.SplitBefore;
                }
                if (++columns > columnCount) return false;
                used = continuationPadding;
            }
            return true;
        }

        /// <summary>A row cursor with the same per-cell height and paragraph-boundary rules used by table painting.</summary>
        private sealed class ColumnBalanceRowFragment {
            private readonly TableBlock table;
            private readonly PdfTableStyle style;
            private readonly int columnCount;
            private readonly double[] columnWidths;
            private readonly double gap;
            private readonly double fullFrameHeight;

            public ColumnBalanceRowFragment(TableBlock table, PdfTableStyle style, PreparedFlowTableRows rows,
                int columnCount, double[] columnWidths, double gap, int rowIndex, int startLine,
                double fullFrameHeight, double before, double movedBefore, double splitBefore, double after) {
                this.table = table; this.style = style; Rows = rows; this.columnCount = columnCount;
                this.columnWidths = columnWidths; this.gap = gap; RowIndex = rowIndex; StartLine = startLine;
                this.fullFrameHeight = fullFrameHeight; Before = before; MovedBefore = movedBefore;
                SplitBefore = splitBefore; After = after;
            }
            public PreparedFlowTableRows Rows { get; }
            public int RowIndex { get; }
            public int StartLine { get; }
            public double Before { get; }
            public double MovedBefore { get; }
            public double SplitBefore { get; }
            public double After { get; }
            public double Height => Before + Measure(StartLine, Rows.LineCounts[RowIndex] - StartLine) + After;
            public double Measure(int start, int count) => MeasurePreparedTableRowSegmentHeight(table, style, Rows,
                columnCount, columnWidths, gap, RowIndex, start, count, suppressCellObjects: false);
            public int Fit(int start, double available, bool requireDefaultFirstFragment, int? maximumLineCount = null) =>
                GetPreparedTableRowSegmentLineCountThatFits(table, style, Rows, columnCount, columnWidths, gap,
                    RowIndex, start, available, fullFrameHeight, canMoveToNextFrame: true,
                    requireDefaultFirstFragment, maximumLineCount);
            public ColumnBalanceRowFragment WithSpacing(double before, double movedBefore, double after) =>
                new(table, style, Rows, columnCount, columnWidths, gap, RowIndex, StartLine, fullFrameHeight,
                    before, movedBefore, SplitBefore, after);
        }
    }
}
