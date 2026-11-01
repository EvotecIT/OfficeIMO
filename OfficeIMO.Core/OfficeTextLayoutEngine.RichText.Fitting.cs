using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextLayoutEngine {
    // Drawing frames require height-aware fitting, including wrapped text and hard lines.
    // The public rich-text layout overloads retain their existing width-only shrink policy.
    private static IReadOnlyList<OfficeRichTextRun> FitRichTextRunsToFrame(
        IReadOnlyList<OfficeRichTextRun> runs, double width, double height, double lineHeightFactor,
        OfficeRichTextMeasurement measure, bool wrap,
        double minimumFontSize, OfficeTextParagraphIndent paragraphIndent, CancellationToken cancellationToken,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint,
        Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?>? measureSegmentPaint,
        out double appliedScale) {
        double availableHeight = NormalizeNonNegative(height);
        bool Fits(IReadOnlyList<OfficeRichTextRun> candidate, double scale) {
            OfficeRichTextBlockLayout measured = LayoutRichTextBlockCore(candidate, width,
                double.MaxValue, lineHeightFactor, measure, wrap, OfficeTextOverflowBehavior.Clip,
                paragraphIndent.Scale(scale), inputTruncated: false, cancellationToken);
            return !measured.Clipped && measured.Width <= width + 0.01D
                && OfficeDrawingTextLayout.RequiredFrameHeight(measured, measurePaint, measureSegmentPaint) <= availableHeight + .000001D;
        }

        double maxFontSize = ResolveMaxRichTextFontSize(runs);
        double minFontSize = Math.Min(maxFontSize, Math.Max(1D, NormalizePositive(minimumFontSize, 1D)));
        double minimumScale = minFontSize / Math.Max(maxFontSize, 1D);
        bool FitsScale(double scale) => Fits(scale == 1D ? runs : ScaleRichTextRuns(runs, scale, cancellationToken), scale);
        appliedScale = measure.HasFeatureMeasurement
            ? ResolveRichTextFrameFitScale(runs, minimumScale, FitsScale, cancellationToken)
            : ResolveFrameFitScale(minimumScale, FitsScale, cancellationToken);
        return appliedScale == 1D ? runs : ScaleRichTextRuns(runs, appliedScale, cancellationToken);
    }

    // Kerning activation can reduce width as size increases, so a single
    // monotonic search can miss a larger fitting interval. Search the intervals
    // between run-policy changes from largest to smallest; keep the old search
    // exactly when no policy changes in the requested range.
    internal static double ResolveRichTextFrameFitScale(IEnumerable<OfficeRichTextRun> runs, double minimumScale,
        Func<double, bool> fits, CancellationToken cancellationToken) {
        var changes = new SortedSet<double>();
        foreach (OfficeRichTextRun run in runs) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!(run.AutomaticPairKerningThreshold is double threshold) || threshold <= 0D || run.FontSize <= 0D) continue;
            double boundary = threshold / run.FontSize;
            if (boundary <= minimumScale || boundary >= 1D || double.IsNaN(boundary)) continue;
            // Locate the first double whose scaled size actually enables the
            // feature, rather than assuming division/multiplication round-trip.
            if (run.FontSize * boundary < threshold) boundary = NextScale(boundary, 1);
            double prior = NextScale(boundary, -1);
            if (run.FontSize * prior >= threshold) boundary = prior;
            if (boundary > minimumScale && boundary < 1D) changes.Add(boundary);
        }
        if (changes.Count == 0) return ResolveFrameFitScale(minimumScale, fits, cancellationToken);
        if (fits(1D)) return 1D;
        changes.Add(minimumScale); changes.Add(1D);
        var boundaries = new List<double>(changes);
        for (int index = boundaries.Count - 2; index >= 0; index--) {
            cancellationToken.ThrowIfCancellationRequested();
            double lower = boundaries[index];
            double upper = NextScale(boundaries[index + 1], -1);
            if (upper < lower) continue;
            if (fits(upper)) return upper;
            if (fits(lower)) return ResolveFrameFitScale(lower / upper, scale => fits(scale * upper), cancellationToken) * upper;
        }
        return minimumScale;

        static double NextScale(double value, long direction) =>
            BitConverter.Int64BitsToDouble(BitConverter.DoubleToInt64Bits(value) + direction);
    }

    // Both inline and paragraph layout use the same bounded frame-fitting search.
    internal static double ResolveFrameFitScale(double minimumScale, Func<double, bool> fits,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (fits(1D)) return 1D;
        double low = minimumScale;
        if (!fits(low)) return low;
        double high = 1D;
        for (int iteration = 0; iteration < 12; iteration++) {
            cancellationToken.ThrowIfCancellationRequested();
            double candidateScale = (low + high) / 2D;
            if (fits(candidateScale)) {
                low = candidateScale;
            } else high = candidateScale;
        }
        return low;
    }
}
