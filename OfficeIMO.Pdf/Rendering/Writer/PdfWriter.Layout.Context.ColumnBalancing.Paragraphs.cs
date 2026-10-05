namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Keeps a paragraph whole when it fits the chosen column; permitted balancing splits longer paragraphs.</summary>
        private static bool PackColumnBalanceParagraph(ColumnBalanceParagraph paragraph, double height, int columnCount,
            double continuationPadding, ref int columns, ref double used) {
            double whole = paragraph.Units.Sum(unit => unit.Height);
            double before = used > continuationPadding + .001D && paragraph.Units.Count > 0
                ? paragraph.Units[0].SpacingBefore : 0D;
            if (whole > height - continuationPadding + .001D)
                return PackColumnBalanceUnits(paragraph.Units, height, columnCount, continuationPadding, ref columns, ref used);
            if (used + before + whole > height + .001D) { columns++; used = continuationPadding; before = 0D; }
            if (columns > columnCount) return false;
            used += before + whole;
            return true;
        }

        private sealed class ColumnBalanceParagraph {
            public ColumnBalanceParagraph(List<ColumnBalanceUnit> units) { Units = units; }
            public List<ColumnBalanceUnit> Units { get; }
        }
    }
}
