using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

internal static partial class OfficeDrawingTextLayout {
    /// <summary>Creates a backdrop around placed paragraph lines without including alignment offsets or outer insets.</summary>
    internal static OfficeTextBlockBackgroundBounds CreateParagraphBackgroundBounds(OfficeRichTextBlockLayout layout,
        double left, double top, double paddingX, double paddingY) {
        double start = double.PositiveInfinity, end = double.NegativeInfinity;
        foreach (OfficeRichTextLine line in layout.Lines) {
            if (line.Width <= 0D) continue;
            start = Math.Min(start, line.OffsetX);
            end = Math.Max(end, line.OffsetX + line.Width);
        }
        if (double.IsPositiveInfinity(start)) start = end = 0D;
        return new OfficeTextBlockBackgroundBounds(left + start - paddingX, top - paddingY,
            end - start + 2D * paddingX, layout.Height + 2D * paddingY);
    }

    // Paragraph layout remains here, alongside the single shared wrapping and styled font measurement owner.
    // Offsets and measured whitespace advances let SVG, raster and PDF consume the same line model.
    private static OfficeRichTextBlockLayout CreateParagraphs(OfficeDrawingRichText text, double width, double height,
        Func<string?, double, string?, OfficeFontStyle, double> measure, double scale,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint,
        bool reportUnwrappedWidthOverflow = true,
        Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?>? measureSegmentPaint = null) =>
        CreateParagraphsWithArea(text.Paragraphs, width, height, measure, text.TextAreaAlignment, text.WrapText, scale, measurePaint,
            shrinkToFit: text.ShrinkToFit, minimumFontSize: 6D * scale,
            reportUnwrappedWidthOverflow: reportUnwrappedWidthOverflow, measureSegmentPaint: measureSegmentPaint);

    /// <summary>Places paragraph runs with shared measured wrapping and paragraph-relative line offsets.</summary>
    internal static OfficeRichTextBlockLayout CreateParagraphs(IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        double width, double height, Func<string?, double, string?, OfficeFontStyle, double> measure,
        bool wrap = true, double scale = 1D,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint = null,
        CancellationToken cancellationToken = default, bool shrinkToFit = false, double minimumFontSize = 1D) =>
        CreateParagraphsWithArea(paragraphs, width, height, measure, OfficeTextAreaAlignment.FullWidth,
            wrap, scale, measurePaint, cancellationToken, shrinkToFit, minimumFontSize);

    internal static OfficeRichTextBlockLayout CreateParagraphsWithMetrics(IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        double width, double height, OfficeDrawingTextMetrics metrics, bool wrap = true, double scale = 1D,
        CancellationToken cancellationToken = default, bool shrinkToFit = false, double minimumFontSize = 1D) =>
        CreateParagraphsWithArea(paragraphs, width, height, metrics.MeasureText, OfficeTextAreaAlignment.FullWidth,
            wrap, scale, metrics.MeasurePaintBounds, cancellationToken, shrinkToFit, minimumFontSize,
            measureSegmentPaint: metrics.MeasureSegmentPaintBounds);

    // The friend-callable method above retains its CLR signature across Core package upgrades.
    private static OfficeRichTextBlockLayout CreateParagraphsWithArea(IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        double width, double height, Func<string?, double, string?, OfficeFontStyle, double> measure,
        OfficeTextAreaAlignment areaAlignment, bool wrap = true, double scale = 1D,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint = null,
        CancellationToken cancellationToken = default, bool shrinkToFit = false, double minimumFontSize = 1D,
        bool reportUnwrappedWidthOverflow = true,
        Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?>? measureSegmentPaint = null) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!Enum.IsDefined(typeof(OfficeTextAreaAlignment), areaAlignment)) throw new ArgumentOutOfRangeException(nameof(areaAlignment));
        double fontScale = 1D;
        if (shrinkToFit) {
            double maximumSize = 1D;
            foreach (OfficeRichTextParagraph paragraph in paragraphs) {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (OfficeRichTextRun run in paragraph.Runs) maximumSize = Math.Max(maximumSize, run.FontSize * scale);
                if (paragraph.Label != null) maximumSize = Math.Max(maximumSize, paragraph.Label.Run.FontSize * scale);
            }
            double minimumScale = Math.Min(maximumSize, Math.Max(1D, minimumFontSize)) / maximumSize;
            fontScale = OfficeTextLayoutEngine.ResolveFrameFitScale(minimumScale, candidate => {
                OfficeRichTextBlockLayout measured = CreateParagraphsCore(paragraphs, width, double.MaxValue,
                    measure, wrap, scale, candidate, measurePaint, cancellationToken, reportLeaderPaintClipping: false,
                    areaAlignment: areaAlignment == OfficeTextAreaAlignment.FullWidth ? areaAlignment : OfficeTextAreaAlignment.Left);
                return !measured.Clipped && measured.Width <= width + .01D
                    && RequiredFrameHeight(measured, measurePaint, measureSegmentPaint) <= height + .000001D;
            }, cancellationToken);
        }
        OfficeRichTextBlockLayout layout = CreateParagraphsCore(paragraphs, width, height, measure, wrap,
            scale, fontScale, measurePaint, cancellationToken, areaAlignment: areaAlignment,
            reportUnwrappedWidthOverflow: reportUnwrappedWidthOverflow);
        return shrinkToFit ? IncludePaintedHeight(layout, height, measurePaint, measureSegmentPaint) : layout;
    }

    private static OfficeRichTextBlockLayout CreateParagraphsCore(IReadOnlyList<OfficeRichTextParagraph> paragraphs,
        double width, double height, Func<string?, double, string?, OfficeFontStyle, double> measure,
        bool wrap, double scale, double fontScale,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint,
        CancellationToken cancellationToken, bool reportLeaderPaintClipping = true,
        OfficeTextAreaAlignment areaAlignment = OfficeTextAreaAlignment.FullWidth,
        bool reportUnwrappedWidthOverflow = true) {
        if (areaAlignment != OfficeTextAreaAlignment.FullWidth)
            return CreateIntrinsicParagraphs(paragraphs, width, height, measure, wrap, scale, fontScale,
                measurePaint, cancellationToken, reportLeaderPaintClipping, areaAlignment, reportUnwrappedWidthOverflow);
        var lines = new List<OfficeRichTextLine>();
        // Truncated decorative paint cannot be recovered by shrinking the body.
        // Fit probes still enforce the budget, but fit against body geometry.
        var leaderBudget = new OfficeTextTabLeaderLayout.Budget(reportLeaderPaintClipping);
        double usedHeight = 0, usedWidth = 0, maximumLineHeight = 1;
        bool clipped = false;
        foreach (OfficeRichTextParagraph paragraph in paragraphs) {
            cancellationToken.ThrowIfCancellationRequested();
            MeasuredParagraph? measured = MeasureParagraph(paragraph, width, measure, wrap, scale, fontScale,
                measurePaint, cancellationToken, leaderBudget, ref clipped, reportUnwrappedWidthOverflow);
            if (measured == null) continue;
            OfficeTextPadding margin = measured.Margins;
            ParagraphLabelLayout? label = measured.Label;
            OfficeRichTextBlockLayout paragraphLayout = measured.Layout;
            if (!AddGap(margin.Top)) { clipped = true; break; }
            bool paragraphTruncated = false;
            int lineCount = Math.Max(paragraphLayout.Lines.Count, label == null ? 0 : 1);
            for (int i = 0; i < lineCount; i++) {
                cancellationToken.ThrowIfCancellationRequested();
                OfficeRichTextLine line = measured.Line(i);
                double lineHeight = paragraph.LineHeight * scale ?? OfficeTextBlockRenderer.ResolveRichTextRenderLineHeight(line, paragraphLayout.LineHeight);
                if (i == 0 && label != null && !paragraph.LineHeight.HasValue) lineHeight = Math.Max(lineHeight, label.Height);
                if (lines.Count >= OfficeTextLayoutEngine.MaximumLayoutLines || usedHeight + lineHeight > height + .000001D) {
                    clipped = true; paragraphTruncated = true; break;
                }
                OfficeRichTextLine placed = PlaceParagraphLine(paragraph, measured, line, i, lineHeight, width,
                    wrap, paragraph.Alignment, measure);
                lines.Add(placed); usedHeight += lineHeight;
                maximumLineHeight = Math.Max(maximumLineHeight, lineHeight);
                usedWidth = Math.Max(usedWidth, placed.OffsetX + placed.Width + margin.Right);
            }
            if (paragraphTruncated || !AddGap(margin.Bottom)) { clipped = true; break; }
        }
        return new OfficeRichTextBlockLayout(lines, maximumLineHeight, usedWidth, usedHeight, clipped);

        bool AddGap(double gap) {
            if (gap == 0) return true;
            if (usedHeight + gap > height || lines.Count >= OfficeTextLayoutEngine.MaximumLayoutLines) return false;
            lines.Add(new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>(), gap)); usedHeight += gap; return true;
        }
    }

    private sealed class MeasuredParagraph {
        internal MeasuredParagraph(OfficeRichTextBlockLayout layout, OfficeTextPadding margins,
            OfficeTextParagraphIndent indent, ParagraphLabelLayout? label) {
            Layout = layout; Margins = margins; Indent = indent; Label = label;
        }
        internal OfficeRichTextBlockLayout Layout { get; }
        internal OfficeTextPadding Margins { get; }
        internal OfficeTextParagraphIndent Indent { get; }
        internal ParagraphLabelLayout? Label { get; }
        internal int LineCount => Math.Max(Layout.Lines.Count, Label == null ? 0 : 1);
        internal OfficeRichTextLine Line(int index) => Layout.Lines.Count == 0
            ? new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>(), offsetX: Indent.FirstLineOffset) : Layout.Lines[index];
    }

    // Measure once at the authored width. Area placement must not reflow tabs or soft wraps.
    private static MeasuredParagraph? MeasureParagraph(OfficeRichTextParagraph paragraph, double width,
        Func<string?, double, string?, OfficeFontStyle, double> measure, bool wrap, double scale, double fontScale,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint,
        CancellationToken cancellationToken, OfficeTextTabLeaderLayout.Budget leaderBudget, ref bool clipped,
        bool reportUnwrappedWidthOverflow) {
        OfficeTextPadding margin = paragraph.Margins.Scale(scale);
        OfficeTextParagraphIndent indent = paragraph.Indent.Scale(scale);
        double originalLeft = margin.Left;
        ParagraphLabelLayout? label = ResolveParagraphLabel(paragraph, width, scale, fontScale, measure, measurePaint,
            cancellationToken, ref margin, ref indent, ref clipped);
        double availableWidth = width - margin.Horizontal;
        if (availableWidth <= 0) { clipped = true; return null; }
        var runs = new List<OfficeRichTextRun>(paragraph.Runs.Count);
        foreach (OfficeRichTextRun run in paragraph.Runs) {
            cancellationToken.ThrowIfCancellationRequested();
            double size = fontScale == 1D ? run.FontSize * scale : Math.Max(1D, run.FontSize * scale * fontScale);
            runs.Add(new OfficeRichTextRun(run.Text, size, run.Color, run.Bold, run.Italic,
                run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle,
                run.StrikethroughStyle, run.Baseline, run.ParagraphIndent?.Scale(scale)) { LinkUri = run.LinkUri });
        }
        OfficeRichTextBlockLayout layout = OfficeTextLayoutEngine.LayoutStyledRichTextBlock(runs,
            availableWidth, double.MaxValue, paragraph.LineHeightFactor ?? 1.2D, measure, wrap,
            overflowBehavior: OfficeTextOverflowBehavior.Clip, paragraphIndent: indent, measurePaint: measurePaint,
            hardBreakStartsParagraph: false, tabStops: paragraph.TabStops?.Scale(scale, originalLeft - margin.Left, fontScale),
            cancellationToken: cancellationToken, leaderBudget: leaderBudget);
        clipped |= layout.Clipped && (reportUnwrappedWidthOverflow || !layout.OnlyUnwrappedWidthOverflow);
        return new MeasuredParagraph(layout, margin, indent, label);
    }

    private static OfficeRichTextLine PlaceParagraphLine(OfficeRichTextParagraph paragraph, MeasuredParagraph measured,
        OfficeRichTextLine line, int index, double lineHeight, double width, bool wrap, OfficeTextAlignment alignment,
        Func<string?, double, string?, OfficeFontStyle, double> measure) {
        double lineWidth = Math.Max(0, width - measured.Margins.Horizontal - line.OffsetX);
        IReadOnlyList<OfficeRichTextSegment> segments = line.Segments;
        bool tabbedLine = paragraph.TabStops != null && ContainsTabAdvance(line);
        if (!tabbedLine && OfficeTextBlockRenderer.ShouldJustifyRichTextLine(line, index, measured.Layout.Lines.Count, lineWidth, alignment))
            segments = JustifiedSegments(line, lineWidth, measure);
        double alignmentOffset = wrap ? OfficeTextPlacement.ResolveLineLeft(0, lineWidth, line.Width, alignment) :
            OfficeTextPlacement.ResolveLeftFromAnchor(OfficeTextPlacement.ResolveAnchorX(0, lineWidth, alignment), line.Width, alignment);
        double offset = measured.Margins.Left + line.OffsetX + (tabbedLine && !paragraph.TabStops!.AlignWithParagraph ? 0 : alignmentOffset);
        if (alignment == OfficeTextAlignment.Justify) offset = measured.Margins.Left + line.OffsetX;
        var placed = new OfficeRichTextLine(segments, lineHeight, offset);
        return index == 0 && measured.Label != null ? AddParagraphLabel(placed, measured.Label, lineHeight) : placed;
    }

    private static bool ContainsTabAdvance(OfficeRichTextLine line) {
        foreach (OfficeRichTextSegment segment in line.Segments) if (segment.Text.Length == 0) return true;
        return false;
    }

    private static IReadOnlyList<OfficeRichTextSegment> JustifiedSegments(OfficeRichTextLine line, double width,
        Func<string?, double, string?, OfficeFontStyle, double> measure) {
        var tokens = OfficeTextBlockRenderer.CreateRichTextRenderTokens(line, measure);
        int gaps = OfficeTextBlockRenderer.CountJustifiableRichTextGaps(tokens);
        double extra = gaps == 0 ? 0 : Math.Max(0, width - line.Width) / gaps;
        var result = new List<OfficeRichTextSegment>(tokens.Count);
        bool hasWord = false;
        for (int i = 0; i < tokens.Count; i++) {
            var token = tokens[i]; OfficeRichTextSegment source = token.Segment;
            double advance = token.Width;
            if (token.IsWhitespace && hasWord && OfficeTextBlockRenderer.HasWordAfter(tokens, i + 1)) advance += extra;
            result.Add(new OfficeRichTextSegment(token.Text, advance, source.FontSize, source.Color, source.Bold,
                source.Italic, source.Underline, source.FontFamily, source.Strikethrough, source.BackgroundColor,
                source.UnderlineStyle, source.StrikethroughStyle, source.Baseline) { LinkUri = source.LinkUri });
            hasWord |= !token.IsWhitespace;
        }
        return result;
    }
}
