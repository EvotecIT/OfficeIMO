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
        double minimumFontSize, OfficeTextParagraphIndent paragraphIndent, CancellationToken cancellationToken) {
        double availableHeight = NormalizeNonNegative(height);
        bool Fits(IReadOnlyList<OfficeRichTextRun> candidate) {
            OfficeRichTextBlockLayout measured = LayoutRichTextBlockCore(candidate, width,
                double.MaxValue, lineHeightFactor, measure, wrap, OfficeTextOverflowBehavior.Clip,
                paragraphIndent, inputTruncated: false, cancellationToken);
            return !measured.Clipped && measured.Width <= width + 0.01D
                && Math.Max(measured.Height, OfficeDrawingTextLayout.PaintedHeight(measured)) <= availableHeight + 0.01D;
        }

        if (Fits(runs)) return runs;
        double maxFontSize = ResolveMaxRichTextFontSize(runs);
        double minFontSize = Math.Min(maxFontSize, Math.Max(1D, NormalizePositive(minimumFontSize, 1D)));
        double low = minFontSize / Math.Max(maxFontSize, 1D);
        IReadOnlyList<OfficeRichTextRun> best = ScaleRichTextRuns(runs, low, cancellationToken);
        if (!Fits(best)) return best;

        double high = 1D;
        for (int iteration = 0; iteration < 12; iteration++) {
            cancellationToken.ThrowIfCancellationRequested();
            double candidateScale = (low + high) / 2D;
            IReadOnlyList<OfficeRichTextRun> candidate = ScaleRichTextRuns(runs, candidateScale, cancellationToken);
            if (Fits(candidate)) {
                low = candidateScale;
                best = candidate;
            } else high = candidateScale;
        }
        return best;
    }
}
