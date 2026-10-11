namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static void ApplyImportedAutomaticSpanMinimum(double[] minimums, double[] preferences,
        int startColumn, int endColumn, double requiredWidth) {
        double minimumTotal = 0D, preferredTotal = 0D, weightTotal = 0D;
        for (int column = startColumn; column < endColumn; column++) {
            minimumTotal += minimums[column];
            preferredTotal += preferences[column];
            weightTotal += minimums[column] + preferences[column];
        }
        if (requiredWidth <= minimumTotal || weightTotal <= 0D) return;
        if (requiredWidth <= preferredTotal) {
            GrowImportedAutomaticSpanWithinPreferences(minimums, preferences, startColumn, endColumn, requiredWidth);
        } else {
            // An unbreakable span wider than its preferences forces expansion.
            // Both the content minimum and the preference contribute to each share.
            for (int column = startColumn; column < endColumn; column++)
                minimums[column] = Math.Max(minimums[column],
                    requiredWidth * (minimums[column] + preferences[column]) / weightTotal);
        }
        for (int column = startColumn; column < endColumn; column++)
            preferences[column] = Math.Max(preferences[column], minimums[column]);
    }

    private static void GrowImportedAutomaticSpanWithinPreferences(double[] minimums, double[] preferences,
        int startColumn, int endColumn, double requiredWidth) {
        // Imported table widths use twips. Allocate a span's growth in column
        // order, then share the residual without crossing a preferred width.
        // The order matters when a later row overlaps this span.
        const double unitsPerPoint = 20D;
        int count = endColumn - startColumn;
        var targets = new double[count];
        var limits = new double[count];
        var weights = new double[count];
        var limited = new bool[count];
        double minimumTotal = 0D, preferredTotal = 0D;
        for (int index = 0; index < count; index++) {
            int column = startColumn + index;
            targets[index] = Math.Round(minimums[column] * unitsPerPoint);
            limits[index] = Math.Max(targets[index], Math.Round(preferences[column] * unitsPerPoint));
            minimumTotal += targets[index];
            preferredTotal += limits[index];
        }
        double remaining = Math.Min(preferredTotal, Math.Round(requiredWidth * unitsPerPoint)) - minimumTotal;
        if (remaining <= 0D) return;

        double limitedWeight = 0D, flexibleWeight = 0D;
        for (int index = 0; index < count; index++) {
            weights[index] = (targets[index] + limits[index]) / (minimumTotal + preferredTotal);
            double capacity = limits[index] - targets[index];
            limited[index] = remaining * weights[index] > capacity;
            if (limited[index]) {
                weights[index] = capacity / remaining;
                limitedWeight += weights[index];
            } else {
                flexibleWeight += weights[index];
            }
        }
        for (int index = 0; index < count; index++) {
            if (!limited[index] && flexibleWeight > 0D)
                weights[index] *= (1D - limitedWeight) / flexibleWeight;
            double growth = Math.Min(limits[index] - targets[index], Math.Floor(remaining * weights[index]));
            targets[index] += growth;
            remaining -= growth;
        }

        // Spend whole equal passes together; the final partial pass starts at
        // the first column. This also bounds work for very wide imported cells.
        while (remaining > 0D) {
            int flexibleCount = 0;
            double smallestCapacity = double.MaxValue;
            for (int index = 0; index < count; index++) {
                double capacity = limits[index] - targets[index];
                if (capacity <= 0D) continue;
                flexibleCount++;
                smallestCapacity = Math.Min(smallestCapacity, capacity);
            }
            if (flexibleCount == 0) break;
            double equalGrowth = Math.Min(smallestCapacity, Math.Floor(remaining / flexibleCount));
            for (int index = 0; index < count && remaining > 0D; index++) {
                if (targets[index] >= limits[index]) continue;
                double growth = equalGrowth > 0D ? equalGrowth : 1D;
                targets[index] += growth;
                remaining -= growth;
            }
        }
        for (int index = 0; index < count; index++) {
            int column = startColumn + index;
            minimums[column] = Math.Max(minimums[column], targets[index] / unitsPerPoint);
        }
    }

    private static void ApplyImportedAutomaticSpanPreference(double[] preferences,
        int startColumn, int endColumn, double requiredWidth) {
        double available = 0D;
        for (int column = startColumn; column < endColumn; column++)
            available += preferences[column];
        if (requiredWidth <= available || available <= 0D) return;
        for (int column = startColumn; column < endColumn; column++)
            preferences[column] *= requiredWidth / available;
    }
}
