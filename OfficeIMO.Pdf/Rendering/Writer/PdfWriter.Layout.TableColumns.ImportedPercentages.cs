namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Fits each percentage preference before the fractions become normalized allocation weights.</summary>
    private static bool TryResolveImportedPercentageFrameWidth(PdfTableStyle style, double[]? preferredWidths,
        double[]? minimumWidths, int columns, double gap, out double width) {
        width = 0D;
        var fractions = style.ImportedAutomaticColumnWidthFractions;
        if (style.PreferredWidth.HasValue || fractions == null || fractions.Count != columns) return false;
        double minimumTotal = 0D;
        for (int column = 0; column < columns; column++) {
            double fraction = fractions[column];
            if (fraction <= 0D || double.IsNaN(fraction) || double.IsInfinity(fraction))
                throw new ArgumentException("Imported percentage column widths must be positive finite values.");
            double minimum = minimumWidths != null && column < minimumWidths.Length ? minimumWidths[column] : 0D;
            minimum = Math.Max(minimum, GetOptionalColumnWidth(style.ColumnMinWidthPoints, column,
                "Table minimum column widths must be positive finite values.") ?? 0D);
            double preferred = preferredWidths != null && column < preferredWidths.Length ? preferredWidths[column] : 0D;
            minimumTotal += minimum;
            width = Math.Max(width, Math.Max(preferred, minimum) / fraction);
        }
        width = Math.Max(width, minimumTotal) + Math.Max(0, columns - 1) * gap;
        return width > .001D;
    }

    /// <summary>Overcommitted imported percentages reserve later cell minima while consuming width in column order.</summary>
    private static bool TryDistributeImportedPercentageColumns(PdfTableStyle style, double[] widths, bool[] fixedColumns,
        double?[] minimums, double?[] maximums, double[]? paddingWidths, double available) {
        var fractions = style.ImportedAutomaticColumnWidthFractions;
        if (fractions == null || fractions.Count != widths.Length ||
            fixedColumns.Any(value => value) || maximums.Any(value => value.HasValue)) return false;
        if (TryCompressImportedColumnMinimums(style, widths, minimums, paddingWidths, available)) return true;
        if (fractions.Sum() <= 1D) return false;
        double reserved = minimums.Sum(value => value ?? 0D);
        double remaining = available;
        for (int column = 0; column < widths.Length; column++) {
            double minimum = minimums[column] ?? 0D;
            reserved -= minimum;
            widths[column] = column == widths.Length - 1 ? remaining :
                Math.Min(Math.Max(minimum, available * fractions[column]), Math.Max(minimum, remaining - reserved));
            remaining = Math.Max(0D, remaining - widths[column]);
        }
        return true;
    }
}
