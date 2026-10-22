namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Wraps a heading and retains its complete source policy before physical line slicing.</summary>
    private static (List<List<RichSeg>> Lines, List<double> LineHeights) WrapRichHeadingRuns(
        IEnumerable<PdfTextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont,
        double lineHeight, double? firstLineWidthPts, double tabStopWidth, PdfOptions? options,
        PdfHeadingStyle? style, PdfAlign alignment) {
        var wrapped = WrapRichRunsWithSpacing(runs, maxWidthPts, fontSize, baseFont, lineHeight,
            firstLineWidthPts, tabStopWidth, options, style?.LineSpacing,
            wordJustificationMode: style?.UseWordJustificationSpacing == true && alignment == PdfAlign.Justify ? style.JustificationMode : null);
        for (int lineIndex = 0; lineIndex < wrapped.Lines.Count; lineIndex++) {
            List<RichSeg> line = wrapped.Lines[lineIndex];
            var policy = new RichLineJustificationPolicy(style?.JustificationMode ?? PdfJustificationMode.Words,
                style?.JustifyExplicitLineBreaks == true, style?.JustifyLastLine == true,
                lineIndex == wrapped.Lines.Count - 1, line.Any(segment => segment.EndsWithHardBreak),
                style?.UseWordJustificationSpacing == true);
            for (int segment = 0; segment < line.Count; segment++)
                line[segment] = line[segment].WithJustificationPolicy(policy);
        }
        return wrapped;
    }
}
