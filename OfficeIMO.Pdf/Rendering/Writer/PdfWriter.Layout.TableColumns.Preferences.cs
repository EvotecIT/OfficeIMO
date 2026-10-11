namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Shrinks imported preferences through their available flexibility after reserving each content minimum.</summary>
    private static bool TryDistributeImportedPreferredColumns(PdfTableStyle style, double[] widths,
        bool[] fixedColumns, double?[] minimums, double?[] maximums, double[]? preferences,
        double[]? paddingWidths, double availableWidth) {
        if (style.AutoFitColumnPreferredWidthPoints == null || preferences == null ||
            fixedColumns.Any(value => value) || style.ColumnWidthWeights?.Count > 0) return false;

        var authored = new bool[widths.Length];
        int authoredCount = 0;
        double authoredPreference = 0D, authoredMinimum = 0D, automaticMinimum = 0D;
        for (int column = 0; column < widths.Length; column++) {
            authored[column] = column < style.AutoFitColumnPreferredWidthPoints.Count &&
                style.AutoFitColumnPreferredWidthPoints[column].HasValue;
            if (authored[column]) {
                authoredCount++;
                authoredPreference += preferences[column];
                authoredMinimum += minimums[column] ?? 0D;
            } else automaticMinimum += minimums[column] ?? 0D;
        }
        if (authoredMinimum + automaticMinimum > availableWidth + .001D)
            return TryCompressImportedColumnMinimums(style, widths, minimums, paddingWidths, availableWidth);
        if (authoredCount > 0 && authoredCount < widths.Length) {
            // Automatic columns absorb remaining space after the absolute
            // preferences fit. They reserve only their minimum while those
            // authored preferences are being compressed.
            double authoredWidth = Math.Max(authoredMinimum,
                Math.Min(authoredPreference, availableWidth - automaticMinimum));
            double allocated = DistributeImportedPreferenceGroup(widths, authored, minimums, maximums, preferences, authoredWidth);
            var automatic = authored.Select(value => !value).ToArray();
            double automaticAllocated = DistributeImportedPreferenceGroup(widths, automatic, minimums, maximums,
                preferences, Math.Max(automaticMinimum, availableWidth - allocated));
            if (allocated + automaticAllocated < availableWidth - .001D)
                DistributeImportedPreferenceGroup(widths, authored, minimums, maximums, preferences,
                    availableWidth - automaticAllocated);
            return true;
        }
        if (preferences.Sum() <= availableWidth + .001D) return false;
        DistributeImportedPreferenceGroup(widths, Enumerable.Repeat(true, widths.Length).ToArray(),
            minimums, maximums, preferences, availableWidth);
        return true;
    }

    /// <summary>Reserves cell padding while compressing text that cannot fit its unbreakable minimum.</summary>
    private static bool TryCompressImportedColumnMinimums(PdfTableStyle style, double[] widths,
        double?[] minimums, double[]? paddingWidths, double availableWidth) {
        double minimumTotal = minimums.Sum(value => value ?? 0D);
        if (minimumTotal <= availableWidth + .001D || !style.AutoFitWidthUsesContentMinimum || paddingWidths == null)
            return false;
        double reserved = 0D;
        for (int column = 0; column < widths.Length; column++)
            reserved += Math.Min(minimums[column] ?? 0D, paddingWidths[column]);
        if (reserved >= availableWidth - .001D) return false;
        double scale = (availableWidth - reserved) / (minimumTotal - reserved);
        for (int column = 0; column < widths.Length; column++) {
            double minimum = minimums[column] ?? 0D;
            double padding = Math.Min(minimum, paddingWidths[column]);
            widths[column] = padding + (minimum - padding) * scale;
        }
        return true;
    }

    private static double DistributeImportedPreferenceGroup(double[] widths, bool[] included,
        double?[] minimums, double?[] maximums, double[] preferences, double availableWidth) {
        var excluded = included.Select(value => !value).ToArray();
        double preferredTotal = 0D;
        for (int column = 0; column < widths.Length; column++)
            if (included[column]) preferredTotal += preferences[column];
        if (preferredTotal <= availableWidth + .001D) {
            DistributeFlexibleColumns(widths, preferences, excluded, minimums, maximums, availableWidth);
            return widths.Where((_, column) => included[column]).Sum();
        }

        var bases = new double[widths.Length];
        var flexibility = new double[widths.Length];
        var residualMinimums = new double?[widths.Length];
        var residualMaximums = new double?[widths.Length];
        var residualWidths = new double[widths.Length];
        double reserved = 0D;
        for (int column = 0; column < widths.Length; column++) {
            if (!included[column]) continue;
            bases[column] = minimums[column] ?? 0D;
            reserved += bases[column];
            flexibility[column] = Math.Max(0D, preferences[column] - bases[column]);
            residualMinimums[column] = 0D;
            if (maximums[column].HasValue)
                residualMaximums[column] = Math.Max(0D, maximums[column]!.Value - bases[column]);
        }
        DistributeFlexibleColumns(residualWidths, flexibility, excluded, residualMinimums, residualMaximums,
            Math.Max(0D, availableWidth - reserved));
        for (int column = 0; column < widths.Length; column++)
            if (included[column]) widths[column] = residualWidths[column] + bases[column];
        return widths.Where((_, column) => included[column]).Sum();
    }
}
