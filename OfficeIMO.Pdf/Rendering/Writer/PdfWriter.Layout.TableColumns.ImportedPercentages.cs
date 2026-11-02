namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Fits each percentage preference before the fractions become normalized allocation weights.</summary>
    private static bool TryResolveImportedPercentageFrameWidth(PdfTableStyle style, double[]? preferredWidths,
        double[]? minimumWidths, int columns, double gap, double availableWidth, out double width) {
        width = 0D;
        double gapWidth = Math.Max(0, columns - 1) * gap;
        var fractions = style.ImportedColumnWidthFractions;
        if (!style.AutoFitUnspecifiedWidthToContent || style.PreferredWidth.HasValue ||
            fractions == null || fractions.Count != columns) return false;
        double minimumTotal = 0D, remainingPreference = 0D, preferredTotal = 0D, fractionTotal = 0D;
        bool hasRemainingColumns = false;
        for (int column = 0; column < columns; column++) {
            double minimum = minimumWidths != null && column < minimumWidths.Length ? minimumWidths[column] : 0D;
            minimum = Math.Max(minimum, GetOptionalColumnWidth(style.ColumnMinWidths?.ConvertAll(static length => length?.ToPoints()), column,
                "Table minimum column widths must be positive finite values.") ?? 0D);
            double preferred = preferredWidths != null && column < preferredWidths.Length ? preferredWidths[column] :
                GetOptionalColumnWidth(style.AutoFitColumnPreferredWidthPoints, column,
                    "Table preferred column widths must be positive finite values.") ?? 0D;
            minimumTotal += minimum;
            preferredTotal += preferred;
            if (!fractions[column].HasValue) {
                hasRemainingColumns = true;
                remainingPreference += Math.Max(preferred, minimum);
                continue;
            }
            double fraction = fractions[column]!.Value;
            if (fraction <= 0D || double.IsNaN(fraction) || double.IsInfinity(fraction))
                throw new ArgumentException("Imported percentage column widths must be positive finite values.");
            fractionTotal += fraction;
            width = Math.Max(width, Math.Max(preferred, minimum) / fraction);
        }
        if (hasRemainingColumns)
            // No finite content-derived frame can satisfy a remaining column
            // once percentages consume the frame. Use its real available extent.
            width = fractionTotal < 1D ? Math.Max(width, remainingPreference / (1D - fractionTotal)) : availableWidth - gapWidth;
        // Absolute preferences from other rows still provide the minimum
        // preferred extent of a fixed automatic table, even on percentage columns.
        if (!style.AutoFitColumns) width = Math.Max(width, preferredTotal);
        width = Math.Max(width, minimumTotal) + gapWidth;
        // Fixed automatic frames retain source twips. A fractional frame
        // below that boundary can split a word that fits in the source cell.
        if (!style.AutoFitColumns && width < double.MaxValue / 20D)
            width = Math.Ceiling(width * 20D - 1e-9D) / 20D;
        return width > .001D;
    }

    /// <summary>Columns held at their minimum leave the rest of the width to the others, shared by their percentages.</summary>
    private static void DistributeOvercommittedPercentages(double[] widths, System.Collections.Generic.List<double?> fractions,
        double?[] minimums, double available) {
        var held = new bool[widths.Length];
        for (bool changed = true; changed;) {
            changed = false;
            double remaining = available - Enumerable.Range(0, widths.Length).Where(column => held[column]).Sum(column => minimums[column] ?? 0D);
            double share = Enumerable.Range(0, widths.Length).Where(column => !held[column]).Sum(column => fractions[column]!.Value);
            for (int column = 0; column < widths.Length; column++) {
                if (held[column]) { widths[column] = minimums[column] ?? 0D; continue; }
                widths[column] = share > 0D ? Math.Max(0D, remaining) * fractions[column]!.Value / share : 0D;
                if (widths[column] < (minimums[column] ?? 0D)) { held[column] = true; changed = true; }
            }
        }
    }

    /// <summary>Overcommitted imported percentages reserve later cell minima while consuming width in column order.</summary>
    private static bool TryDistributeImportedPercentageColumns(PdfTableStyle style, double[] widths, bool[] fixedColumns,
        double?[] minimums, double?[] maximums, double[]? preferences, double[]? paddingWidths, double available) {
        var fractions = style.ImportedColumnWidthFractions;
        if (fractions == null || fractions.Count != widths.Length ||
            fixedColumns.Any(value => value) || maximums.Any(value => value.HasValue)) return false;
        if (TryCompressImportedColumnMinimums(style, widths, minimums, paddingWidths, available)) return true;
        bool hasRemainingColumns = fractions.Any(value => !value.HasValue);
        double fractionTotal = fractions.Sum(value => value ?? 0D);
        if (!hasRemainingColumns && fractionTotal <= 1D) return false;
        var remainingColumns = fractions.Select(value => !value.HasValue).ToArray();
        double remainingMinimum = minimums.Where((_, column) => remainingColumns[column]).Sum(value => value ?? 0D);
        if (!hasRemainingColumns) {
            DistributeOvercommittedPercentages(widths, fractions, minimums, available);
            return true;
        }
        double percentagePool = hasRemainingColumns && fractionTotal >= 1D ? available - remainingMinimum : available;
        double reserved = minimums.Sum(value => value ?? 0D);
        double remaining = available;
        for (int column = 0; column < widths.Length; column++) {
            if (!fractions[column].HasValue) continue;
            double minimum = minimums[column] ?? 0D;
            reserved -= minimum;
            widths[column] = !hasRemainingColumns && column == widths.Length - 1 ? remaining :
                Math.Min(Math.Max(minimum, percentagePool * fractions[column]!.Value), Math.Max(minimum, remaining - reserved));
            remaining = Math.Max(0D, remaining - widths[column]);
        }
        if (hasRemainingColumns) {
            preferences ??= Enumerable.Range(0, widths.Length).Select(column =>
                GetOptionalColumnWidth(style.AutoFitColumnPreferredWidthPoints, column,
                    "Table preferred column widths must be positive finite values.") ?? minimums[column] ?? 1D).ToArray();
            return DistributeImportedPreferredColumnSubset(style, widths, remainingColumns, minimums, maximums,
                preferences, paddingWidths, remaining);
        }
        return true;
    }
}
