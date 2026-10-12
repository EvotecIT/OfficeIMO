namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves the native decimal marker profile qualified against Pages export.</summary>
internal static class IWorkDecimalListLayout {
    internal static bool TryResolve(IWorkTextParagraph paragraph, out double markerFontSizePoints,
        out double textOffsetPoints) {
        markerFontSizePoints = 0;
        textOffsetPoints = 0;
        if (paragraph.ListLayout is not { } layout || !paragraph.ListNumber.HasValue
            || paragraph.ListMarkerKind != IWorkListMarkerKind.Number
            || !IWorkDecimalListProfile.TryGetMarkerStyle(paragraph.ListLabel, out _)
            || paragraph.ListLevel is < 0 or > IWorkDecimalListProfile.MaximumLevel
            || paragraph.Style.FirstLineIndentPoints.GetValueOrDefault() != 0
            || paragraph.Style.LeftIndentPoints.GetValueOrDefault() != 0
            || paragraph.Style.RightIndentPoints.GetValueOrDefault() != 0
            || paragraph.Style.TextStyle.FontSizePoints is not double size
            || !IsFinite(size) || size <= 0 || size != Math.Round(size)) return false;
        double? firstRunSize = null;
        bool uniformSize = true, wholePointSizes = true;
        foreach (IWorkTextRun run in paragraph.Runs) {
            double runSize = run.Style.FontSizePoints ?? size;
            if (!IsFinite(runSize) || runSize <= 0) return false;
            uniformSize &= !firstRunSize.HasValue || firstRunSize.Value == runSize;
            wholePointSizes &= runSize == Math.Round(runSize);
            firstRunSize ??= runSize;
        }
        double bodySize = firstRunSize ?? size;
        double scaledBody = layout.MarkerScale * bodySize;
        double uniformMarkerSize = bodySize == size ? Math.Round(scaledBody) : size;
        if (uniformSize && IsFinite(scaledBody)
            && uniformMarkerSize > 0
            && Math.Abs(scaledBody - uniformMarkerSize) <= Math.Max(1e-6d, uniformMarkerSize * 1e-7d)) {
            // Native uniform items use either a marker-size or body-size paragraph default.
            // Resolve the relative scale to a whole-point marker without changing the em indent.
            markerFontSizePoints = uniformMarkerSize;
        } else if (layout.MarkerScale == 1 && wholePointSizes) {
            // Native mixed-size items take their marker size from the first character.
            // Later larger runs do not change the marker or paragraph-default em indent.
            markerFontSizePoints = bodySize;
        } else return false;
        double offset = layout.TextIndentEm * size;
        if (!IsFinite(offset) || !IsFinite(layout.MarkerIndentPoints)) return false;
        textOffsetPoints = offset;
        return true;
    }

    internal static string? MarkerFontName(IWorkTextParagraph paragraph) =>
        paragraph.ListFontName ?? paragraph.Style.TextStyle.FontName;

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
