namespace OfficeIMO.IWork.Internal;

/// <summary>Resolves the native decimal marker profile qualified against Pages export.</summary>
internal static class IWorkDecimalListLayout {
    internal static bool TryResolve(IWorkTextParagraph paragraph, out double markerFontSizePoints,
        out double textOffsetPoints) {
        markerFontSizePoints = 0;
        textOffsetPoints = 0;
        if (paragraph.ListLayout is not { } layout || !paragraph.ListNumber.HasValue
            || paragraph.ListMarkerKind != IWorkListMarkerKind.Number
            || paragraph.ListLabel != "1." || paragraph.ListLevel is < 0 or > 1
            || paragraph.Style.FirstLineIndentPoints.GetValueOrDefault() != 0
            || paragraph.Style.LeftIndentPoints.GetValueOrDefault() != 0
            || paragraph.Style.RightIndentPoints.GetValueOrDefault() != 0
            || paragraph.Style.TextStyle.FontSizePoints is not double size
            || !IsFinite(size) || size <= 0 || size != Math.Round(size)) return false;
        double? bodyFontSize = null;
        foreach (IWorkTextRun run in paragraph.Runs) {
            double runSize = run.Style.FontSizePoints ?? size;
            if (!IsFinite(runSize) || runSize <= 0
                || bodyFontSize.HasValue && bodyFontSize.Value != runSize) return false;
            bodyFontSize = runSize;
        }
        // Pages retains the paragraph's marker font size while rich runs may use a
        // different body size. Its relative marker scale relates those two sizes;
        // the text-indent em remains relative to the paragraph size.
        double scaledBody = layout.MarkerScale * (bodyFontSize ?? size);
        if (!IsFinite(scaledBody) || Math.Abs(scaledBody - size) > Math.Max(1e-6d, size * 1e-7d)) return false;
        double offset = layout.TextIndentEm * size;
        if (!IsFinite(offset) || !IsFinite(layout.MarkerIndentPoints)) return false;
        markerFontSizePoints = size;
        textOffsetPoints = offset;
        return true;
    }

    internal static string? MarkerFontName(IWorkTextParagraph paragraph) =>
        paragraph.ListFontName ?? paragraph.Style.TextStyle.FontName;

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
