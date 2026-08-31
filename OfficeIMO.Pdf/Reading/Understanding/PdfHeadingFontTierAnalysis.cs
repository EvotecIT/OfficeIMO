namespace OfficeIMO.Pdf;

/// <summary>Canonical document-wide heading font-tier ranking.</summary>
internal static class PdfHeadingFontTierAnalysis {
    internal static Dictionary<double, int> BuildLookup(
        IReadOnlyList<double> fontSizes,
        Action observeWork) {
        Guard.NotNull(fontSizes, nameof(fontSizes));
        Guard.NotNull(observeWork, nameof(observeWork));

        var sorted = new double[fontSizes.Count];
        for (int index = 0; index < fontSizes.Count; index++) {
            observeWork();
            sorted[index] = fontSizes[index];
        }
        SortDescending(sorted, observeWork);

        var tierByFontSize = new Dictionary<double, int>();
        int currentTier = 0;
        double currentTierFontSize = 0D;
        for (int index = 0; index < sorted.Length; index++) {
            observeWork();
            double fontSize = sorted[index];
            if (currentTier == 0 || Math.Abs(currentTierFontSize - fontSize) > 0.5D) {
                currentTier++;
                currentTierFontSize = fontSize;
            }
            tierByFontSize[fontSize] = currentTier;
        }
        return tierByFontSize;
    }

    private static void SortDescending(double[] values, Action observeWork) {
        if (values.Length < 2) return;
        var target = new double[values.Length];
        double[] source = values;
        for (int width = 1; width < source.Length;) {
            for (int left = 0; left < source.Length; left += width * 2) {
                int middle = Math.Min(left + width, source.Length);
                int right = Math.Min(left + (width * 2), source.Length);
                int first = left;
                int second = middle;
                int output = left;
                while (first < middle && second < right) {
                    observeWork();
                    target[output++] = source[first] >= source[second]
                        ? source[first++]
                        : source[second++];
                }
                while (first < middle) {
                    observeWork();
                    target[output++] = source[first++];
                }
                while (second < right) {
                    observeWork();
                    target[output++] = source[second++];
                }
            }
            double[] swap = source;
            source = target;
            target = swap;
            if (width > source.Length / 2) break;
            width *= 2;
        }
        if (!ReferenceEquals(source, values)) {
            for (int index = 0; index < source.Length; index++) {
                observeWork();
                values[index] = source[index];
            }
        }
    }
}
