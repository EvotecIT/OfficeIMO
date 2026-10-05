namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    /// <summary>Rewraps unpainted cell text while keeping paragraph formatting and the row's consumed-line cursor.</summary>
    private static TableCellTextLayout ContinueTableCellTextLayout(TableCellLayout cell, TableCellTextLayout previous,
        int consumedLines, double innerWidth, PdfStandardFont font, double size, double leading, PdfOptions options,
        double additionalFontSizeScale = 1D, double minimumShrinkFontSize = 0D) {
        double wrapWidth = GetTableCellWrapWidth(innerWidth, cell.NoWrap);
        TableCellTextLayout remainder;
        if (consumedLines >= previous.Lines.Count) {
            remainder = new TableCellTextLayout(new(), new());
        } else if (previous.ParagraphRanges is { } ranges) {
            var paragraphs = new List<PdfTableCellParagraph>();
            foreach (TableCellParagraphRange range in ranges) {
                int end = range.StartLine + range.LineCount;
                if (end <= consumedLines) continue;
                PdfTableCellParagraph paragraph = range.Paragraph;
                bool continued = consumedLines > range.StartLine;
                int start = Math.Max(consumedLines, range.StartLine);
                var runs = continued ? BuildTextRunsFromWrappedLines(previous.Lines, start, end - start) : paragraph.Runs;
                paragraphs.Add(new PdfTableCellParagraph(runs, paragraph.SpacingAfter, paragraph.Align,
                    continued ? 0D : paragraph.SpacingBefore, paragraph.LeftIndent, paragraph.RightIndent,
                    continued ? 0D : paragraph.FirstLineIndent, paragraph.LineHeight, paragraph.DefaultTabStopWidth,
                    paragraph.TabStops, paragraph.FontSize, paragraph.LineSpacing,
                    paragraph.WidowControl, paragraph.KeepTogether, paragraph.KeepWithNext));
            }
            remainder = CreateTableCellParagraphTextLayout(ScaleTableCellParagraphsForShrink(paragraphs, additionalFontSizeScale, minimumShrinkFontSize),
                wrapWidth, innerWidth, font, size, leading, options);
        } else {
            var remainingRuns = BuildTextRunsFromWrappedLines(previous.Lines, consumedLines, previous.Lines.Count - consumedLines);
            var wrapped = WrapRichRunsCore(ScaleTableRunsForShrink(remainingRuns, additionalFontSizeScale, minimumShrinkFontSize),
                wrapWidth, size, font, leading, null, DefaultParagraphTabStopWidth, options);
            remainder = new TableCellTextLayout(wrapped.Lines, wrapped.LineHeights);
        }

        // Keep the consumed prefix as zero-height placeholders. Drawing and object-placement
        // paths retain a nonzero cursor, so bookmarks, images, and form fields are not repeated.
        var lines = Enumerable.Range(0, consumedLines).Select(_ => new List<RichSeg>()).ToList();
        lines.AddRange(remainder.Lines);
        List<double> PrefixHeights(List<double> values) => Enumerable.Repeat(0D, consumedLines).Concat(values).ToList();
        List<T>? Prefix<T>(List<T>? values, T empty) => values == null ? null : Enumerable.Repeat(empty, consumedLines).Concat(values).ToList();
        var paragraphRanges = remainder.ParagraphRanges?.Select(range => range with { StartLine = range.StartLine + consumedLines }).ToList();
        return new TableCellTextLayout(lines, PrefixHeights(remainder.LineHeights), Prefix(remainder.LineAlignments, (PdfAlign?)null),
            Prefix(remainder.LineXOffsets, 0D), Prefix(remainder.LineWidths, 0D), 0D, PrefixHeights(remainder.LineBoxHeights), paragraphRanges);
    }
}
