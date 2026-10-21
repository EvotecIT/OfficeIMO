namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // Imported multiple spacing adds distance to the next line. A page or
    // column needs the last line's font box, including the paragraph mark and
    // inline objects, without reserving that trailing distance below it.
    private static double ResolveImportedLineFrameFitHeight(RichLine line,
        PdfStandardFont baseFont, double fontSize, PdfOptions? options, PdfLineSpacing spacing) {
        double descent = 0D;
        bool hasDescent = false;
        foreach (RichSeg segment in line) {
            if (segment.Text.Length == 0 && segment.LeadingSpace && segment.InlineElement == null) continue;
            if (segment.InlineElement is PdfInlineFormWidget { SourceRun: not null } widget) {
                if (segment.ContributesLineBoxDescent) {
                    GetInlineFormWidgetLineMetrics(widget, baseFont, fontSize, options, spacing,
                        out _, out double runDescent);
                    descent = Math.Max(descent, runDescent);
                    hasDescent = true;
                }
            } else if (segment.InlineElement is { } inline) {
                descent = Math.Max(descent, -inline.BaselineOffset);
                hasDescent = true;
            } else if (segment.ContributesLineBoxDescent) {
                GetRichRunLineMetrics(segment.Font, segment.NamedFont, segment.FontSize, segment.Baseline,
                    options, spacing, out _, out double runDescent);
                descent = Math.Max(descent, runDescent);
                hasDescent = true;
            }
        }
        if (!hasDescent)
            GetRichRunLineMetrics(baseFont, null, fontSize, PdfTextBaseline.Normal, options, spacing, out _, out descent);
        return Math.Max(0D, line.BaselineOffset + descent);
    }

    private static double GetRichLineFrameFitHeight(IReadOnlyList<RichSeg> line, double advance, double topGap = 0D) =>
        line is RichLine { FrameFitHeight: double required }
            ? Math.Min(advance, Math.Max(0D, topGap) + required) : advance;

    private static double MeasureRichLinesFrameFitHeight(IReadOnlyList<List<RichSeg>> lines,
        IReadOnlyList<double> heights, double fallbackLeading, IReadOnlyList<double>? topGaps = null) {
        double height = MeasureRichLinesHeight(heights, lines.Count, fallbackLeading);
        if (lines.Count == 0) return height;
        int last = lines.Count - 1;
        double advance = GetRichLineHeight(heights, last, fallbackLeading);
        return height - advance + GetRichLineFrameFitHeight(lines[last], advance, topGaps?[last] ?? 0D);
    }
}
