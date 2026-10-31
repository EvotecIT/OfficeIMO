namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        private (double FlowAdvance, double PaintExtent) MeasureColumnSequence(IReadOnlyList<ColItem> items, double initialConsumed = 0D) {
            double advance = 0D;
            double paint = 0D;
            for (int index = 0; index < items.Count; index++) {
                ColItem item = items[index];
                double consumed = initialConsumed + advance;
                double height;
                double extent;
                if (item is ColGroupStart start) {
                    var group = MeasureColumnGroup(start.Group, consumed);
                    height = group.FlowAdvance;
                    extent = group.PaintExtent;
                    index += start.Group.ItemCount - 1;
                } else {
                    height = MeasureColItemFullHeight(item, consumed);
                    extent = item is ColTable table ? MeasureColumnTablePaintHeight(table, consumed) : height;
                }
                paint = Math.Max(paint, advance + extent);
                advance += height;
            }
            return (advance, Math.Max(paint, advance));
        }

        private static double MeasureColumnTablePaintHeight(ColTable table, double consumedBefore) =>
            ResolveColumnSpacingBefore(table.Style.SpacingBefore.ToPoints(), consumedBefore) + table.CaptionHeight
            + GetTableRowsHeight(table.RowHeights, 0, table.RowHeights.Length, GetTableCellSpacing(table.Style))
            + table.Style.SpacingAfter.ToPoints();

        private (double FlowAdvance, double PaintExtent) MeasureColumnGroup(ColumnGroup group, double consumedBefore) {
            double before = ResolveColumnSpacingBefore(group.Style?.SpacingBefore.ToPoints() ?? 0D, consumedBefore);
            double opening = before + (group.Style?.TopPadding ?? 0D);
            double closing = (group.Style?.BottomPadding ?? 0D) + (group.Style?.SpacingAfter.ToPoints() ?? 0D);
            var content = MeasureColumnSequence(group.Children, consumedBefore + opening);
            double advance = opening + content.FlowAdvance + closing;
            // Closing padding remains reserved while the children paint, even
            // when an overlay extends below the restored flow cursor.
            return (advance, opening + content.PaintExtent + closing);
        }
    }
}
