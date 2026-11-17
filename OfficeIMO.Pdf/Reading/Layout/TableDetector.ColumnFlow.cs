namespace OfficeIMO.Pdf;

internal static partial class TableDetector {
    private const double MinimumProseColumnWidthInFontSizes = 12D;
    private const double MinimumProseLineFill = 0.8D;

    /// <summary>
    /// True when a text-only candidate is a horizontal slice of a multi-column page: every column holds wrapped
    /// prose lines that fill the column width, the lines sit at body-text leading, the columns reach the page's text
    /// margins, and the gutters between them continue past the candidate through other lines of the page. Such a
    /// slice read as a table interleaves the page columns line by line, so it stays in the paragraph flow instead.
    /// </summary>
    private static bool IsPageColumnFlowProse(PdfStructuredTable table, IReadOnlyList<TextLayoutEngine.TextLine> pageLines) {
        int columnCount = table.Columns.Count;
        // Positioned-cell candidates carry their own narrative checks; this covers the band-split recovery paths.
        if (columnCount < 2 || columnCount > 3 || table.HasRuledBounds || table.SourceRuns.Count == 0 ||
            !table.Kind.StartsWith("band-", StringComparison.Ordinal)) return false;

        // Visual lines of the candidate, keyed by baseline.
        var lines = new List<(double Y, List<PdfTextSpan> Spans)>();
        var fontSizes = new List<double>(table.SourceRuns.Count);
        foreach (PdfTextSpan run in table.SourceRuns.OrderByDescending(static run => run.Y)) {
            if (string.IsNullOrWhiteSpace(run.Text)) continue;
            fontSizes.Add(run.FontSize);
            if (lines.Count > 0 && Math.Abs(lines[lines.Count - 1].Y - run.Y) <= 2D) {
                lines[lines.Count - 1].Spans.Add(run);
            } else {
                lines.Add((run.Y, new List<PdfTextSpan> { run }));
            }
        }
        // Body text size; footnote marks and inline math must not stretch the expected leading.
        double fontSize = fontSizes.Count == 0 ? 0D : Median(fontSizes);
        if (lines.Count < 2 || fontSize <= 0D) return false;

        // Body text keeps one leading; table rows are usually spaced wider or unevenly.
        for (int index = 1; index < lines.Count; index++) {
            double gap = lines[index - 1].Y - lines[index].Y;
            if (gap > fontSize * 2.5D) return false;
        }

        var textLefts = new double[columnCount];
        var textRights = new double[columnCount];
        for (int column = 0; column < columnCount; column++) {
            var extents = new List<(double Left, double Right, string Text)>(lines.Count);
            foreach ((double _, List<PdfTextSpan> spans) in lines) {
                if (TryGetColumnExtent(spans, table.Columns, column, out double left, out double right, out string text)) {
                    extents.Add((left, right, text));
                }
            }
            if (extents.Count < 2) return false;
            double columnLeft = extents.Min(static extent => extent.Left);
            double columnRight = extents.Max(static extent => extent.Right);
            double width = columnRight - columnLeft;
            if (width < fontSize * MinimumProseColumnWidthInFontSizes) return false;
            // Paragraph-final lines and hanging indents are short, so most lines carry words and at least half fill the column.
            int filledLines = extents.Count(extent => extent.Right - extent.Left >= width * MinimumProseLineFill);
            int wordLines = extents.Count(extent => LooksLikeProseLine(extent.Text));
            if (wordLines * 3 < extents.Count * 2 || filledLines * 2 < extents.Count) return false;
            textLefts[column] = columnLeft;
            textRights[column] = columnRight;
        }

        // Page columns reach the page's text margins; a table inside one page column does not.
        if (!SpansPageTextMargins(pageLines, textLefts[0], textRights[columnCount - 1], fontSize)) return false;
        return table.Rows.Count >= 6 || HasContinuingPageGutters(table, pageLines, textLefts, textRights, fontSize);
    }

    private static bool TryGetColumnExtent(List<PdfTextSpan> spans, List<PdfStructuredTableColumn> columns, int columnIndex,
        out double left, out double right, out string text) {
        left = double.MaxValue;
        right = double.MinValue;
        var builder = new StringBuilder();
        foreach (PdfTextSpan span in spans.OrderBy(static span => span.X)) {
            if (GetBestOverlappingColumn(span, columns) != columnIndex) continue;
            left = Math.Min(left, span.X);
            right = Math.Max(right, span.X + Math.Max(0D, span.Advance));
            if (builder.Length > 0) builder.Append(' ');
            builder.Append(span.Text.Trim());
        }
        text = builder.ToString();
        return right > left;
    }

    private static bool SpansPageTextMargins(IReadOnlyList<TextLayoutEngine.TextLine> pageLines, double left, double right, double fontSize) {
        // Margins come from body-sized lines so running headers, page numbers and wide figures do not move them.
        var lefts = new List<double>();
        var rights = new List<double>();
        foreach (TextLayoutEngine.TextLine line in pageLines) {
            foreach (PdfTextSpan span in line.Spans) {
                if (string.IsNullOrWhiteSpace(span.Text) || Math.Abs(span.FontSize - fontSize) > fontSize * 0.15D) continue;
                lefts.Add(span.X);
                rights.Add(span.X + Math.Max(0D, span.Advance));
            }
        }
        if (lefts.Count == 0) return false;
        lefts.Sort();
        rights.Sort();
        // Ignore the outermost few percent: stray marginal marks should not define the text block.
        double pageLeft = lefts[lefts.Count / 50];
        double pageRight = rights[rights.Count - 1 - rights.Count / 50];
        // An indented abstract leaves one page margin, so one outer edge on the text block is enough.
        return left <= pageLeft + fontSize * 1.5D || right >= pageRight - fontSize * 1.5D;
    }

    // Long runs can extend past a split position, so a run belongs to the column it overlaps most.
    private static int GetBestOverlappingColumn(PdfTextSpan span, List<PdfStructuredTableColumn> columns) {
        double spanLeft = span.X, spanRight = span.X + Math.Max(0D, span.Advance);
        int best = -1;
        double bestScore = double.NegativeInfinity;
        for (int index = 0; index < columns.Count; index++) {
            double from = Math.Min(columns[index].From, columns[index].To);
            double to = Math.Max(columns[index].From, columns[index].To);
            // Overlap when positive; otherwise the negative distance to the column.
            double score = Math.Min(spanRight, to) - Math.Max(spanLeft, from);
            if (score > bestScore) {
                bestScore = score;
                best = index;
            }
        }
        return best;
    }

    // Wrapped prose has several words, most of them containing letters. Numeric cells, codes and short labels do not.
    private static bool LooksLikeProseLine(string text) {
        int tokens = 0, wordTokens = 0;
        foreach (string token in text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
            tokens++;
            int letters = 0;
            foreach (char ch in token) {
                if (char.IsLetter(ch)) letters++;
            }
            if (letters >= 2) wordTokens++;
        }
        return tokens >= 4 && wordTokens * 2 >= tokens;
    }

    // The page's column gutters run above or below the candidate: other lines either stay inside one column or
    // leave the gutter empty, and they also fill a column with prose.
    private static bool HasContinuingPageGutters(PdfStructuredTable table, IReadOnlyList<TextLayoutEngine.TextLine> pageLines,
        double[] textLefts, double[] textRights, double fontSize) {
        double top = Math.Max(table.YTop, table.YBottom) + 1D;
        double bottom = Math.Min(table.YTop, table.YBottom) - 1D;
        var owned = new HashSet<PdfTextSpan>(table.SourceRuns);
        int continuingLines = 0;
        foreach (TextLayoutEngine.TextLine line in pageLines) {
            if (line.Y <= top && line.Y >= bottom) continue;
            bool crossesGutter = false, fillsColumn = false;
            for (int column = 0; column < textLefts.Length && !crossesGutter; column++) {
                double left = double.MaxValue, right = double.MinValue;
                foreach (PdfTextSpan span in line.Spans) {
                    if (string.IsNullOrWhiteSpace(span.Text) || owned.Contains(span)) continue;
                    double spanLeft = span.X, spanRight = span.X + Math.Max(0D, span.Advance);
                    if (column + 1 < textLefts.Length && spanLeft < textRights[column] - fontSize && spanRight > textLefts[column + 1] + fontSize) {
                        crossesGutter = true;
                        break;
                    }
                    double center = (spanLeft + spanRight) / 2D;
                    if (center < textLefts[column] - fontSize || center > textRights[column] + fontSize) continue;
                    left = Math.Min(left, spanLeft);
                    right = Math.Max(right, spanRight);
                }
                if (right > left && right - left >= (textRights[column] - textLefts[column]) * MinimumProseLineFill) fillsColumn = true;
            }
            if (!crossesGutter && fillsColumn) continuingLines++;
        }
        return continuingLines >= 2;
    }
}
