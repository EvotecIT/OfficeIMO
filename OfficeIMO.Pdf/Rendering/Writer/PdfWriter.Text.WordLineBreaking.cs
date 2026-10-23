namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>
    /// Chooses between a legal expanded prefix and a feasible compressed whole-word candidate
    /// in the same fixed frame. Uses the actual preferred source-space resources, excluding
    /// the request-sized residual expansion tier. Does not measure, shape or consume runs.
    /// </summary>
    private static bool PreferWordExpansionBreak(IReadOnlyList<RichSeg> preceding,
        IReadOnlyList<RichSeg> candidate, double precedingNaturalWidth, double candidateNaturalWidth,
        double availableWidth) {
        double expansion = availableWidth - precedingNaturalWidth;
        double compression = candidateNaturalWidth - availableWidth;
        if (!IsFiniteWordBreakMetric(expansion) || !IsFiniteWordBreakMetric(compression) ||
            expansion < 0D || compression <= 0D || !preceding.Any(segment => segment.Text.Length > 0))
            return false;
        if (expansion == 0D) return true;

        double preferred = 0D;
        for (int index = WordCompressionStart(preceding); index < preceding.Count; index++)
            preferred += WordPreferredExpansionCapacity(preceding[index]);
        double capacity = MeasureWordCompressionCapacity(candidate);
        if (!IsFiniteWordBreakMetric(preferred) || !IsFiniteWordBreakMetric(capacity) ||
            preferred <= 0D || capacity <= 0D) return false;

        double expansionFraction = expansion / preferred;
        double compressionFraction = compression / capacity;
        return IsFiniteWordBreakMetric(expansionFraction) && IsFiniteWordBreakMetric(compressionFraction) &&
            expansionFraction <= compressionFraction;
    }

    private static bool IsFiniteWordBreakMetric(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
