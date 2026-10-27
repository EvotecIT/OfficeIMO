namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Intrinsic measurements answer how much fixed content demands. A rule
    // demands no minimum width; finite layout supplies its actual ink width.
    private enum RichWrapPurpose { Layout, IntrinsicMeasurement }

    private static (List<List<RichSeg>> Lines, List<double> LineHeights) WrapRichRunsIntrinsic(
        IEnumerable<PdfTextRun> runs, double maxWidthPts, double fontSize, PdfStandardFont baseFont,
        double lineHeight, double? firstLineWidthPts, double tabStopWidth, PdfCreateOptions? options) =>
        WrapRichRunsCoreWithFirstLineOrigin(runs, maxWidthPts, fontSize, baseFont, lineHeight,
            firstLineWidthPts, null, tabStopWidth, options, purpose: RichWrapPurpose.IntrinsicMeasurement);

    private static bool UsesInlineRunTextMetrics(PdfInlineElement? element) =>
        element is PdfInlineRule or PdfInlineBox { UsesRunTextMetrics: true } or
            PdfInlineFormWidget { SourceRun: not null };

    private static PdfInlineBox ResolveInlineRule(PdfInlineRule rule, double availableWidth, RichWrapPurpose purpose) {
        double width = purpose == RichWrapPurpose.IntrinsicMeasurement ? 0D : Math.Max(0D, availableWidth);
        if (rule.MaximumWidth is double maximum) width = Math.Min(width, maximum);
        return PdfInlineBox.FromRule(rule, width);
    }
}
