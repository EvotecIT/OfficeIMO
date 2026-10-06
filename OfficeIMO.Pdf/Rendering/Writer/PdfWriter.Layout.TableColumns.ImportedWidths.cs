namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static double MeasureImportedTableMinimumTextWidth(TableCellLayout cell,
        PdfStandardFont font, double size, PdfOptions? options, double runFontSizeScale = 1D, double minimumShrinkFontSize = 0D) {
        // Resolve actual run fonts and sizes, then measure break opportunities
        // explicitly: a wide layout never needs to discover internal breaks.
        TableCellTextLayout layout = CreateTableCellTextLayout(cell, TableCellNoWrapWidth,
            font, size, size * 1.25D, options, runFontSizeScale, minimumShrinkFontSize);
        double minimum = 0D;
        for (int lineIndex = 0; lineIndex < layout.Lines.Count; lineIndex++) {
            var line = layout.Lines[lineIndex];
            double indents = layout.LineWidths != null
                ? Math.Max(0D, TableCellNoWrapWidth - layout.LineWidths[lineIndex]) : 0D;
            if (cell.NoWrap) {
                minimum = Math.Max(minimum, MeasureRichLineWidth(line, options) + indents);
                continue;
            }
            var word = new System.Collections.Generic.List<RichSeg>();
            foreach (RichSeg segment in line) {
                if (segment.LeadingSpace || segment.LeadingIsTab) {
                    minimum = Math.Max(minimum, MeasureImportedTableWordMinimum(word, options) + indents);
                    word.Clear();
                }
                word.Add(segment);
                if (segment.EndsWithHardBreak || segment.EndsWithTextSeparator) {
                    minimum = Math.Max(minimum, MeasureImportedTableWordMinimum(word, options) + indents);
                    word.Clear();
                }
            }
            minimum = Math.Max(minimum, MeasureImportedTableWordMinimum(word, options) + indents);
        }
        return minimum;
    }

    private static double MeasureImportedTableWordMinimum(System.Collections.Generic.IReadOnlyList<RichSeg> segments, PdfOptions? options) {
        if (segments.Count == 0) return 0D;
        string text = string.Concat(segments.Select(segment => segment.Text));
        if (text.Length == 0) return segments.Sum(GetRichSegmentWidth);
        // Word permits hyphen and multilingual breaks but keeps a slash token
        // whole when finding a table minimum. Generic technical-token wrapping
        // remains independent of this imported-grid policy.
        int[] hyphens = GetValidHyphenationBreakpoints(text, options);
        int[] points = GetValidSoftLineBreakpoints(text, options)
            .Concat(hyphens)
            .Concat(OfficeIMO.Drawing.OfficeTextLineBreaks.GetBreakPositions(text)
                .Where(point => text[point - 1] != '/'))
            .Distinct().OrderBy(point => point).Concat(new[] { text.Length }).ToArray();
        double minimum = 0D;
        int start = 0;
        foreach (int end in points) {
            int segmentStart = 0;
            double width = 0D;
            RichSeg? lastSegment = null;
            foreach (RichSeg segment in segments) {
                int left = Math.Max(start, segmentStart);
                int right = Math.Min(end, segmentStart + segment.Text.Length);
                if (right > left) {
                    width += left == segmentStart && right == segmentStart + segment.Text.Length
                        ? segment.MeasuredWidth
                        : MeasureRichText(segment.Text.Substring(left - segmentStart, right - left),
                            segment.Font, segment.NamedFont, segment.FontSize, segment.Baseline, options, segment.FeatureSettings);
                    lastSegment = segment;
                } else if (segment.Text.Length == 0 && segmentStart >= start && segmentStart < end) {
                    width += GetRichSegmentWidth(segment);
                }
                segmentStart += segment.Text.Length;
            }
            if (end < text.Length && Array.IndexOf(hyphens, end) >= 0 && lastSegment != null)
                width += MeasureRichText("-", lastSegment.Font, lastSegment.NamedFont,
                    lastSegment.FontSize, lastSegment.Baseline, options, lastSegment.FeatureSettings);
            minimum = Math.Max(minimum, width);
            start = end;
        }
        return minimum;
    }

    private static double ResolveImportedTableCellWrapWidth(TableCellLayout cell, double innerWidth,
        PdfStandardFont font, double size, PdfOptions? options, double runFontSizeScale,
        double minimumShrinkFontSize, bool wrapOversizedNoWrap) {
        bool noWrap = cell.NoWrap;
        // Word can enlarge an automatic grid for a no-wrap cell, but falls
        // back to wrapping when its text cannot fit inside the page frame.
        if (noWrap && wrapOversizedNoWrap &&
            MeasureImportedTableMinimumTextWidth(cell, font, size, options, runFontSizeScale, minimumShrinkFontSize) > innerWidth + .001D)
            noWrap = false;
        return GetTableCellWrapWidth(innerWidth, noWrap);
    }
}
