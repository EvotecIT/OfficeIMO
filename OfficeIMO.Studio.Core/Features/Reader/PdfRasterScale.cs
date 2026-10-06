namespace OfficeIMO.Studio.Features.Reader;

/// <summary>
/// Composes the raster fallback scale, in device pixels per PDF drawing unit, from the presented size and the
/// display's render scaling. Scales are whole percentages: at or above <see cref="StepPercent"/> they round up to
/// a step so a bitmap never has fewer pixels than its target, then fall back to the largest step within the
/// per-page pixel budget.
/// </summary>
internal static class PdfRasterScale {
    internal const int StepPercent = 25;
    internal const int MaximumPercent = 400;
    private const double PercentTolerance = 0.5D;

    /// <summary>Display render scaling, treating unknown or invalid values as 1:1.</summary>
    internal static double NormalizeRenderScaling(double renderScaling) =>
        double.IsFinite(renderScaling) && renderScaling > 0D ? renderScaling : 1D;

    /// <summary>
    /// Returns the scale for presenting a <paramref name="unitWidth"/> × <paramref name="unitHeight"/> drawing at
    /// <paramref name="displayUnitsPerDrawingUnit"/> device-independent pixels per drawing unit.
    /// </summary>
    internal static double Compose(
        double displayUnitsPerDrawingUnit,
        double renderScaling,
        double unitWidth,
        double unitHeight,
        long maximumPixels) {
        if (!double.IsFinite(displayUnitsPerDrawingUnit) || displayUnitsPerDrawingUnit <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(displayUnitsPerDrawingUnit));
        }
        if (maximumPixels <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPixels));

        int desired = ToPercent(displayUnitsPerDrawingUnit * NormalizeRenderScaling(renderScaling));
        return Math.Min(desired, GetBudgetPercent(unitWidth, unitHeight, maximumPixels)) / 100D;
    }

    /// <summary>The cache bucket for a requested scale; already composed scales map to themselves.</summary>
    internal static int ToPercent(double scale) {
        if (!double.IsFinite(scale) || scale <= 0D) throw new ArgumentOutOfRangeException(nameof(scale));
        double percent = scale * 100D;
        int bucket = percent < StepPercent
            ? (int)Math.Ceiling(percent - PercentTolerance)
            : (int)Math.Ceiling((percent - PercentTolerance) / StepPercent) * StepPercent;
        return Math.Clamp(bucket, 1, MaximumPercent);
    }

    private static int GetBudgetPercent(double unitWidth, double unitHeight, long maximumPixels) {
        double width = Math.Max(1D, unitWidth);
        double height = Math.Max(1D, unitHeight);
        int percent = Math.Min(MaximumPercent, (int)Math.Floor(Math.Sqrt(maximumPixels / (width * height)) * 100D));
        if (percent >= StepPercent) percent -= percent % StepPercent;
        // The renderer sizes the bitmap with ceilings; step down until that exact size fits.
        while (percent > 1 && PixelCount(width, height, percent) > maximumPixels) {
            percent = percent > StepPercent ? percent - StepPercent : percent - 1;
        }
        return Math.Max(1, percent);
    }

    private static long PixelCount(double width, double height, int percent) {
        double scale = percent / 100D;
        return (long)Math.Ceiling(width * scale) * (long)Math.Ceiling(height * scale);
    }
}
