using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static TableCellTextLayout CreateTableCellTextLayout(TableCellLayout cell, double innerWidth, PdfStandardFont baseFont, double fontSize, double leading, PdfOptions? options, double runFontSizeScale = 1D, double minimumShrinkFontSize = 0D) {
        double wrapWidth = GetTableCellWrapWidth(innerWidth, cell.NoWrap);
        if (cell.Paragraphs.Count > 0) {
            return CreateTableCellParagraphTextLayout(ScaleTableCellParagraphsForShrink(cell.Paragraphs, runFontSizeScale, minimumShrinkFontSize), wrapWidth, innerWidth, baseFont, fontSize, leading, options);
        }

        var wrap = WrapRichRunsCore(ScaleTableRunsForShrink(cell.Runs, runFontSizeScale, minimumShrinkFontSize), wrapWidth, fontSize, baseFont, leading, null, DefaultParagraphTabStopWidth, options);
        if (wrap.Lines.Count == 0) {
            wrap.Lines.Add(new System.Collections.Generic.List<RichSeg>());
        }

        while (wrap.LineHeights.Count < wrap.Lines.Count) {
            wrap.LineHeights.Add(leading);
        }

        return new TableCellTextLayout(wrap.Lines, wrap.LineHeights);
    }

    private static System.Collections.Generic.IReadOnlyList<PdfTextRun> ScaleTableRunsForShrink(System.Collections.Generic.IReadOnlyList<PdfTextRun> runs, double runFontSizeScale, double minimumShrinkFontSize) {
        if (runFontSizeScale >= 0.999D) {
            return runs;
        }

        var scaledRuns = new System.Collections.Generic.List<PdfTextRun>(runs.Count);
        foreach (PdfTextRun run in runs) {
            if (run.InlineElement != null) {
                scaledRuns.Add(run);
                continue;
            }

            double? scaledFontSize = ScaleTableFontSizeForShrink(run.FontSize, runFontSizeScale, minimumShrinkFontSize);

            scaledRuns.Add(new PdfTextRun(
                run.Text,
                run.Bold,
                run.Underline,
                run.Color,
                run.Italic,
                run.Strike,
                scaledFontSize,
                run.Font,
                run.LinkUri,
                run.LinkContents,
                run.Baseline,
                run.LinkDestinationName,
                run.TabLeader,
                run.TabAlignment,
                run.BackgroundColor,
                run.FontFamily,
                run.UnderlineStyle,
                run.StrikeStyle,
                run.DecorationColor)
                .WithFeatureSettings(run.FeatureSettings)
                .WithTextDirection(run.TextDirection));
        }

        return scaledRuns.AsReadOnly();
    }

    private static double? ScaleTableFontSizeForShrink(double? fontSize, double scale, double minimum) {
        double floor = minimum > 0D ? minimum : 0.001D;
        return !fontSize.HasValue || fontSize.Value <= floor ? fontSize : Math.Max(floor, fontSize.Value * scale);
    }

    private static System.Collections.Generic.IReadOnlyList<PdfTableCellParagraph> ScaleTableCellParagraphsForShrink(System.Collections.Generic.IReadOnlyList<PdfTableCellParagraph> paragraphs, double runFontSizeScale, double minimumShrinkFontSize) {
        if (runFontSizeScale >= 0.999D) {
            return paragraphs;
        }

        var scaledParagraphs = new System.Collections.Generic.List<PdfTableCellParagraph>(paragraphs.Count);
        foreach (PdfTableCellParagraph paragraph in paragraphs) {
            scaledParagraphs.Add(new PdfTableCellParagraph(
                ScaleTableRunsForShrink(paragraph.Runs, runFontSizeScale, minimumShrinkFontSize),
                paragraph.SpacingAfter,
                paragraph.Align,
                paragraph.SpacingBefore,
                paragraph.LeftIndent,
                paragraph.RightIndent,
                paragraph.FirstLineIndent,
                paragraph.LineHeight,
                paragraph.DefaultTabStopWidth,
                paragraph.TabStops,
                ScaleTableFontSizeForShrink(paragraph.FontSize, runFontSizeScale, minimumShrinkFontSize),
                paragraph.LineSpacing, paragraph.WidowControl, paragraph.KeepTogether, paragraph.KeepWithNext));
        }

        return scaledParagraphs.AsReadOnly();
    }

    private static double GetTableCellWrapWidth(double innerWidth, bool noWrap) =>
        noWrap ? TableCellNoWrapWidth : innerWidth;

    private static TableCellTextLayout CreateTableCellParagraphTextLayout(System.Collections.Generic.IReadOnlyList<PdfTableCellParagraph> paragraphs, double wrapWidth, double cellInnerWidth, PdfStandardFont baseFont, double fontSize, double leading, PdfOptions? options) {
        var lines = new System.Collections.Generic.List<System.Collections.Generic.List<RichSeg>>();
        var lineHeights = new System.Collections.Generic.List<double>();
        var lineAlignments = new System.Collections.Generic.List<PdfAlign?>();
        var lineXOffsets = new System.Collections.Generic.List<double>();
        var lineWidths = new System.Collections.Generic.List<double>();
        var lineBoxHeights = new System.Collections.Generic.List<double>();
        var paragraphRanges = new System.Collections.Generic.List<TableCellParagraphRange>();
        double topSpacing = 0D;
        for (int paragraphIndex = 0; paragraphIndex < paragraphs.Count; paragraphIndex++) {
            PdfTableCellParagraph paragraph = paragraphs[paragraphIndex];
            PdfParagraphStyle paragraphStyle = CreateTableCellParagraphStyle(paragraph, cellInnerWidth);
            double paragraphFontSize = paragraphStyle.FontSize ?? fontSize;
            double paragraphLeading = paragraphStyle.LineHeight.HasValue || paragraphStyle.LineSpacing != null
                ? GetParagraphLeading(paragraphStyle, paragraphFontSize) : leading;
            var paragraphFrame = GetParagraphTextFrame(paragraphStyle, 0D, wrapWidth);
            var alignmentFrame = wrapWidth > cellInnerWidth
                ? GetParagraphTextFrame(paragraphStyle, 0D, cellInnerWidth)
                : paragraphFrame;
            var wrap = WrapRichRunsCoreWithFirstLineOrigin(
                paragraph.Runs,
                paragraphFrame.Width,
                paragraphFontSize,
                baseFont,
                paragraphLeading,
                paragraphFrame.FirstLineWidth,
                paragraphFrame.FirstLineX - paragraphFrame.X,
                GetParagraphTabStopWidth(paragraphStyle),
                options,
                paragraphStyle.TabStops.ToArray(), lineSpacing: paragraphStyle.LineSpacing);
            if (wrap.Lines.Count == 0) {
                wrap.Lines.Add(new System.Collections.Generic.List<RichSeg>());
            }

            while (wrap.LineHeights.Count < wrap.Lines.Count) {
                wrap.LineHeights.Add(paragraphLeading);
            }

            int firstNewLineIndex = lines.Count;
            if (paragraph.SpacingBefore > 0D) {
                if (lineHeights.Count > 0) {
                    lineHeights[lineHeights.Count - 1] += paragraph.SpacingBefore;
                } else {
                    topSpacing = paragraph.SpacingBefore;
                }
            }

            lines.AddRange(wrap.Lines);
            paragraphRanges.Add(new TableCellParagraphRange(paragraph, firstNewLineIndex, wrap.Lines.Count));
            lineHeights.AddRange(wrap.LineHeights);
            lineBoxHeights.AddRange(wrap.LineHeights);
            for (int lineIndex = firstNewLineIndex; lineIndex < lines.Count; lineIndex++) {
                lineAlignments.Add(paragraph.Align);
                bool firstParagraphLine = lineIndex == firstNewLineIndex;
                lineXOffsets.Add(firstParagraphLine ? alignmentFrame.FirstLineX : alignmentFrame.X);
                lineWidths.Add(firstParagraphLine ? alignmentFrame.FirstLineWidth : alignmentFrame.Width);
            }

            if (paragraphIndex < paragraphs.Count - 1 && lines.Count > firstNewLineIndex) {
                MarkRichLineHardBreak(lines[lines.Count - 1]);
            }

            if (paragraph.SpacingAfter > 0D && lineHeights.Count > firstNewLineIndex) {
                int lastParagraphLineIndex = lineHeights.Count - 1;
                lineHeights[lastParagraphLineIndex] += paragraph.SpacingAfter;
            }
        }

        if (lines.Count == 0) {
            lines.Add(new System.Collections.Generic.List<RichSeg>());
            lineHeights.Add(leading);
            lineBoxHeights.Add(leading);
            lineAlignments.Add(null);
            lineXOffsets.Add(0D);
            lineWidths.Add(wrapWidth);
        }

        return new TableCellTextLayout(lines, lineHeights, lineAlignments, lineXOffsets, lineWidths, topSpacing, lineBoxHeights, paragraphRanges);
    }

    private static PdfParagraphStyle CreateTableCellParagraphStyle(PdfTableCellParagraph paragraph, double availableWidth) {
        const double minimumTextWidth = 0.001D;
        const double maximumSafeIndentMagnitude = double.MaxValue / 8D;
        double safeWidth = double.IsNaN(availableWidth) || double.IsInfinity(availableWidth)
            ? minimumTextWidth
            : System.Math.Min(maximumSafeIndentMagnitude, System.Math.Max(minimumTextWidth, availableWidth));
        double leftIndent = System.Math.Max(-maximumSafeIndentMagnitude, System.Math.Min(paragraph.LeftIndent, safeWidth - minimumTextWidth));
        double rightIndent = System.Math.Max(-maximumSafeIndentMagnitude, System.Math.Min(paragraph.RightIndent, safeWidth - leftIndent - minimumTextWidth));
        double textWidth = System.Math.Max(minimumTextWidth, safeWidth - leftIndent - rightIndent);
        double firstLineIndent = System.Math.Max(-maximumSafeIndentMagnitude, System.Math.Min(paragraph.FirstLineIndent, textWidth - minimumTextWidth));
        var style = new PdfParagraphStyle {
            FontSize = paragraph.FontSize,
            LineSpacing = paragraph.LineSpacing,
            LineHeight = paragraph.LineHeight,
            LeftIndent = leftIndent,
            RightIndent = rightIndent,
            FirstLineIndent = firstLineIndent,
            DefaultTabStopWidth = paragraph.DefaultTabStopWidth
        };

        foreach (PdfTabStop tabStop in paragraph.TabStops) {
            style.TabStops.Add(tabStop.Clone());
        }

        return style;
    }

    private static TableCellTextLayout CreateListItemTextLayout(PdfListItem item, double innerWidth, PdfStandardFont baseFont, double fontSize, double leading, PdfOptions? options, PdfLineSpacing? lineSpacing) {
        var wrap = WrapRichRunsWithSpacing(item.Runs, innerWidth, fontSize, baseFont, leading, null, DefaultParagraphTabStopWidth, options, lineSpacing);
        if (wrap.Lines.Count == 0) {
            wrap.Lines.Add(new System.Collections.Generic.List<RichSeg>());
        }

        while (wrap.LineHeights.Count < wrap.Lines.Count) {
            wrap.LineHeights.Add(leading);
        }

        return new TableCellTextLayout(wrap.Lines, wrap.LineHeights);
    }

    private static double GetRichLineHeight(System.Collections.Generic.IReadOnlyList<double> heights, int lineIndex, double fallbackLeading) =>
        lineIndex >= 0 && lineIndex < heights.Count ? heights[lineIndex] : fallbackLeading;

    private static int LimitTableCellLineCountToHeight(TableCellTextLayout lines, int startLine, int requestedLineCount, double fallbackLeading, double availableHeight) {
        int maximumLineCount = System.Math.Max(0, System.Math.Min(requestedLineCount, lines.LineCount - startLine));
        double consumedHeight = startLine == 0 ? lines.TopSpacing : 0D;
        int visibleLineCount = 0;
        for (int offset = 0; offset < maximumLineCount; offset++) {
            double lineHeight = GetRichLineHeight(lines.LineHeights, startLine + offset, fallbackLeading);
            // Paragraph spacing advances the next line; it does not make this line's box taller.
            double lineBoxHeight = GetRichLineHeight(lines.LineBoxHeights, startLine + offset, fallbackLeading);
            if (consumedHeight + lineBoxHeight > availableHeight + 0.001D) {
                break;
            }

            consumedHeight += lineHeight;
            visibleLineCount++;
        }

        return visibleLineCount;
    }

    private static double MeasureRichLinesHeight(System.Collections.Generic.IReadOnlyList<double> heights, int lineCount, double fallbackLeading) {
        double height = 0D;
        for (int index = 0; index < lineCount; index++) {
            height += GetRichLineHeight(heights, index, fallbackLeading);
        }

        return height;
    }

    private static double MeasureTableCellTextHeight(TableCellTextLayout layout, int startLine, int lineCount, double fallbackLeading) {
        int available = System.Math.Max(0, layout.Lines.Count - startLine);
        int visible = System.Math.Max(0, System.Math.Min(lineCount, available));
        if (visible == 0) {
            return fallbackLeading;
        }

        double height = startLine == 0 ? layout.TopSpacing : 0D;
        int measuredStart = System.Math.Min(startLine, layout.LineHeights.Count);
        int measuredEnd = System.Math.Min(startLine + visible, layout.LineHeights.Count);
        height += layout.LineHeightPrefix[measuredEnd] - layout.LineHeightPrefix[measuredStart];
        height += (visible - (measuredEnd - measuredStart)) * fallbackLeading;
        return height;
    }

}
