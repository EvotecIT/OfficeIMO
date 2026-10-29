namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Wraps a heading and retains its complete source policy before physical line slicing.</summary>
    private static (List<List<RichSeg>> Lines, List<double> LineHeights) WrapRichHeadingRuns(
        IEnumerable<PdfTextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont,
        double lineHeight, double? firstLineWidthPts, double tabStopWidth, PdfOptions? options,
        PdfHeadingStyle? style, PdfAlign alignment,
        Func<int, double, double, double, RichTextRowFrame>? rowLayout = null,
        System.Threading.CancellationToken cancellationToken = default) {
        var wrapped = WrapRichRunsCoreWithFirstLineOrigin(runs, maxWidthPts, fontSize, baseFont, lineHeight,
            firstLineWidthPts, null, tabStopWidth, options, style?.TabStops.ToArray(), lineSpacing: style?.LineSpacing,
            wordJustificationMode: style?.UseWordJustificationSpacing == true && alignment == PdfAlign.Justify ? style.JustificationMode : null,
            rowLayout: rowLayout, cancellationToken: cancellationToken);
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
