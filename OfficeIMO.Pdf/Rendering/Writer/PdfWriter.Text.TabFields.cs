namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static bool AlignsCompleteTabField(PdfTabAlignment alignment) =>
        alignment == PdfTabAlignment.Center || alignment == PdfTabAlignment.Right;

    private static double MeasureRichTabField(IReadOnlyList<PdfTextRun> runs, int runIndex, int textOffset,
        PdfStandardFont baseFont, double fontSize, PdfOptions? options, PdfTextWhitespaceMode whitespaceMode) {
        var field = new List<PdfTextRun>();
        for (int index = runIndex; index < runs.Count; index++) {
            PdfTextRun run = runs[index];
            if (run.InlineElement != null) {
                field.Add(run);
                continue;
            }

            string text = (run.Text ?? string.Empty).Replace("\r\n", "\n").Replace('\r', '\n');
            int start = index == runIndex ? textOffset : 0;
            int end = start;
            while (end < text.Length && text[end] != '\t' && text[end] != '\n') end++;
            if (end > start) field.Add(CreateStyledTextRun(text.Substring(start, end - start), run, run.Font));
            if (end < text.Length) break;
        }

        // The slice contains no tabs or hard breaks, so this cannot recurse into
        // another field measurement. Reuse the actual token/space/kerning layout
        // rather than measuring a differently normalized copy of its text.
        var wrapped = WrapRichRunsCoreWithFirstLineOrigin(field, double.MaxValue / 4D, fontSize, baseFont,
            fontSize * 1.2D, null, null, DefaultParagraphTabStopWidth, options, textWhitespaceMode: whitespaceMode);
        return wrapped.Lines.Count == 0 ? 0D : MeasureRichLineWidth(wrapped.Lines[0], options);
    }

    private static double MeasureWrappedTabField(IReadOnlyList<RichSeg> line, int start) {
        double width = 0D;
        for (int index = start; index < line.Count; index++) {
            RichSeg segment = line[index];
            if (index > start && segment.LeadingIsTab) break;
            width += GetRichSegmentWidth(segment);
            if (index > start && segment.LeadingSpace) width += segment.LeadingAdvance;
        }
        return width;
    }
}
