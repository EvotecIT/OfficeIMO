namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Uses the list renderer's wrapped line heights while leaving markers and item cursors in their original owner.</summary>
        private static List<ColumnBalanceUnit> MeasureListColumnBalanceUnits(PreparedListLayout prepared, int firstItem = 0,
            int consumedLines = 0, List<double>? continuingLineHeights = null, bool includeSpacingBefore = true) {
            var units = new List<ColumnBalanceUnit>();
            for (int item = firstItem; item < prepared.Items.Count; item++) {
                TableCellTextLayout layout = prepared.Items[item].TextLayout;
                List<double> heights = item == firstItem && continuingLineHeights != null ? continuingLineHeights : layout.LineHeights;
                int start = item == firstItem ? consumedLines : 0;
                for (int line = start; line < heights.Count; line++) {
                    double height = GetRichLineHeight(heights, line, prepared.Leading);
                    if (item == 0 && line == 0 && includeSpacingBefore) height += prepared.SpacingBefore;
                    if (line == heights.Count - 1) height += item == prepared.Items.Count - 1 ? prepared.SpacingAfter : prepared.ItemSpacing;
                    units.Add(new ColumnBalanceUnit(height));
                }
            }
            if (prepared.Style?.KeepTogether == true && units.Count > 0)
                return new() { new(units.Sum(unit => unit.Height)) };
            return units;
        }
    }
}
