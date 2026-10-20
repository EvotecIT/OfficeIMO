namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private readonly record struct FlowTableCellTextFrame(TableCellContentFrame Content,
            int VisibleLineCount, double InnerWidth, double VerticalOffset,
            double VisibleTextHeight, double ContentTop, double FirstBaseline);

        /// <summary>Shares the exact admitted line and baseline calculation between table planning and painting.</summary>
        private FlowTableCellTextFrame PrepareFlowTableCellTextFrame(PdfTableStyle style, TableCellLayout cell,
            TableCellTextLayout lines, int rowIndex, int sourceStartLine, int requestedLineCount,
            double cellX, double cellTop, double cellWidth, double cellHeight, double rowSize,
            double rowLeading, bool rowUsesBold, double? verticalOffsetOverride = null) {
            int column = cell.Column;
            double padTop = GetTableCellPaddingTop(style, rowIndex, column);
            double padBottom = GetTableCellPaddingBottom(style, rowIndex, column);
            TableCellContentFrame content = GetTableCellContentFrame(cell, cellX, cellTop, cellWidth, cellHeight);
            double innerWidth = content.Width - GetTableCellPaddingLeft(style, rowIndex, column) -
                GetTableCellPaddingRight(style, rowIndex, column);
            double availableHeight = Math.Max(0D, content.Height - padTop - padBottom - (verticalOffsetOverride ?? 0D));
            int visibleCount = cell.TextRotation != 0 ? (sourceStartLine == 0 ? 1 : 0) :
                LimitTableCellLineCountToHeight(lines, sourceStartLine, requestedLineCount, rowLeading,
                    availableHeight, style.PreservePartialCellLines);
            double verticalOffset = verticalOffsetOverride ?? 0D;
            double visibleTextHeight = 0D;
            if (visibleCount > 0) {
                visibleTextHeight = MeasureTableCellTextHeight(lines, sourceStartLine, visibleCount, rowLeading);
                double contentHeight = MeasureTableCellContentHeight(cell, lines, sourceStartLine,
                    visibleCount, rowLeading, innerWidth);
                if (!verticalOffsetOverride.HasValue) {
                    double unused = Math.Max(0D, availableHeight - contentHeight);
                    PdfCellVerticalAlign alignment = GetTableCellVerticalAlignment(style, rowIndex, column);
                    if (alignment == PdfCellVerticalAlign.Middle) verticalOffset = unused / 2D;
                    else if (alignment == PdfCellVerticalAlign.Bottom) verticalOffset = unused;
                }
            }
            double contentTop = content.Top - padTop - verticalOffset;
            double baseline = contentTop - (sourceStartLine == 0 ? lines.TopSpacing : 0D) -
                GetAscenderForOptions(GetTableRowFont(currentOpts, rowUsesBold), rowSize, currentOpts) + style.RowBaselineOffset;
            return new FlowTableCellTextFrame(content, visibleCount, innerWidth, verticalOffset,
                visibleTextHeight, contentTop, baseline);
        }
    }
}
