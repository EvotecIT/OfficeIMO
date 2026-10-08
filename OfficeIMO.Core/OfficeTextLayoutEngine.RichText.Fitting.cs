using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextLayoutEngine {
    // Drawing frames require height-aware fitting, including wrapped text and hard lines.
    // The public rich-text layout overloads retain their existing width-only shrink policy.
    private static IReadOnlyList<OfficeRichTextRun> FitRichTextRunsToFrame(
        IReadOnlyList<OfficeRichTextRun> runs, double width, double height, double lineHeightFactor,
        Func<string?, double, string?, OfficeFontStyle, double> measure, bool wrap,
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
        appliedScale = ResolveFrameFitScale(minFontSize / Math.Max(maxFontSize, 1D),
            scale => Fits(scale == 1D ? runs : ScaleRichTextRuns(runs, scale, cancellationToken), scale), cancellationToken);
        return appliedScale == 1D ? runs : ScaleRichTextRuns(runs, appliedScale, cancellationToken);
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
