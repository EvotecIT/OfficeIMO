using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextLayoutEngine {
    /// <summary>
    /// Lays out styled rich text runs into a bounded text block with optional wrapping and height clipping.
    /// </summary>
    /// <param name="runs">Styled text runs.</param>
    /// <param name="maxWidth">Maximum block width.</param>
    /// <param name="maxHeight">Maximum block height.</param>
    /// <param name="lineHeightFactor">Multiplier used with the largest run font size to derive line height.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="wrap">Whether soft wrapping is enabled. Hard line breaks are always honored.</param>
    /// <param name="shrinkToFit">Whether non-wrapped rich text should proportionally shrink run font sizes to fit the requested width.</param>
    /// <param name="minimumFontSize">Minimum font size for any run when <paramref name="shrinkToFit"/> is enabled.</param>
    /// <param name="paragraphIndent">Optional first-line and continuation-line offsets applied while laying out wrapped rich text.</param>
    /// <returns>Measured rich text block with visible lines and clipping state.</returns>
    public static OfficeRichTextBlockLayout LayoutRichTextBlock(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, double> measure,
        bool wrap,
        bool shrinkToFit = false,
        double minimumFontSize = 1D,
        OfficeTextParagraphIndent? paragraphIndent = null) =>
        LayoutRichTextBlock(
            runs,
            maxWidth,
            maxHeight,
            lineHeightFactor,
            measure,
            wrap,
            shrinkToFit,
            minimumFontSize,
            OfficeTextOverflowBehavior.Ellipsis,
            paragraphIndent);

    /// <summary>
    /// Lays out styled rich text runs into a bounded text block with optional wrapping, overflow policy, and height clipping.
    /// </summary>
    /// <param name="runs">Styled text runs.</param>
    /// <param name="maxWidth">Maximum block width.</param>
    /// <param name="maxHeight">Maximum block height.</param>
    /// <param name="lineHeightFactor">Multiplier used with the largest run font size to derive line height.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double)"/>.</param>
    /// <param name="wrap">Whether soft wrapping is enabled. Hard line breaks are always honored.</param>
    /// <param name="shrinkToFit">Whether non-wrapped rich text should proportionally shrink run font sizes to fit the requested width.</param>
    /// <param name="minimumFontSize">Minimum font size for any run when <paramref name="shrinkToFit"/> is enabled.</param>
    /// <param name="overflowBehavior">How overflowing rich text should be represented in the returned layout.</param>
    /// <param name="paragraphIndent">Optional first-line and continuation-line offsets applied while laying out wrapped rich text.</param>
    /// <returns>Measured rich text block with visible lines and clipping state.</returns>
    public static OfficeRichTextBlockLayout LayoutRichTextBlock(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, double> measure,
        bool wrap,
        bool shrinkToFit,
        double minimumFontSize,
        OfficeTextOverflowBehavior overflowBehavior,
        OfficeTextParagraphIndent? paragraphIndent = null) {
        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        return LayoutRichTextBlock(
            runs,
            maxWidth,
            maxHeight,
            lineHeightFactor,
            (text, fontSize, _) => measure(text, fontSize),
            wrap,
            shrinkToFit,
            minimumFontSize,
            overflowBehavior,
            paragraphIndent);
    }

    /// <summary>
    /// Lays out styled rich text runs into a bounded text block with optional wrapping and height clipping.
    /// </summary>
    /// <param name="runs">Styled text runs.</param>
    /// <param name="maxWidth">Maximum block width.</param>
    /// <param name="maxHeight">Maximum block height.</param>
    /// <param name="lineHeightFactor">Multiplier used with the largest run font size to derive line height.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double, string?)"/>.</param>
    /// <param name="wrap">Whether soft wrapping is enabled. Hard line breaks are always honored.</param>
    /// <param name="shrinkToFit">Whether non-wrapped rich text should proportionally shrink run font sizes to fit the requested width.</param>
    /// <param name="minimumFontSize">Minimum font size for any run when <paramref name="shrinkToFit"/> is enabled.</param>
    /// <param name="paragraphIndent">Optional first-line and continuation-line offsets applied while laying out wrapped rich text.</param>
    /// <returns>Measured rich text block with visible lines and clipping state.</returns>
    public static OfficeRichTextBlockLayout LayoutRichTextBlock(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, string?, double> measure,
        bool wrap,
        bool shrinkToFit = false,
        double minimumFontSize = 1D,
        OfficeTextParagraphIndent? paragraphIndent = null) =>
        LayoutRichTextBlock(
            runs,
            maxWidth,
            maxHeight,
            lineHeightFactor,
            measure,
            wrap,
            shrinkToFit,
            minimumFontSize,
            OfficeTextOverflowBehavior.Ellipsis,
            paragraphIndent);

    /// <summary>
    /// Lays out styled rich text runs into a bounded text block with optional wrapping, overflow policy, and height clipping.
    /// </summary>
    /// <param name="runs">Styled text runs.</param>
    /// <param name="maxWidth">Maximum block width.</param>
    /// <param name="maxHeight">Maximum block height.</param>
    /// <param name="lineHeightFactor">Multiplier used with the largest run font size to derive line height.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double, string?)"/>.</param>
    /// <param name="wrap">Whether soft wrapping is enabled. Hard line breaks are always honored.</param>
    /// <param name="shrinkToFit">Whether non-wrapped rich text should proportionally shrink run font sizes to fit the requested width.</param>
    /// <param name="minimumFontSize">Minimum font size for any run when <paramref name="shrinkToFit"/> is enabled.</param>
    /// <param name="overflowBehavior">How overflowing rich text should be represented in the returned layout.</param>
    /// <param name="paragraphIndent">Optional first-line and continuation-line offsets applied while laying out wrapped rich text.</param>
    /// <returns>Measured rich text block with visible lines and clipping state.</returns>
    public static OfficeRichTextBlockLayout LayoutRichTextBlock(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, string?, double> measure,
        bool wrap,
        bool shrinkToFit,
        double minimumFontSize,
        OfficeTextOverflowBehavior overflowBehavior,
        OfficeTextParagraphIndent? paragraphIndent = null) {
        return LayoutRichTextBlock(
            runs,
            maxWidth,
            maxHeight,
            lineHeightFactor,
            measure,
            wrap,
            shrinkToFit,
            minimumFontSize,
            overflowBehavior,
            paragraphIndent,
            CancellationToken.None);
    }

    /// <summary>
    /// Lays out styled rich text runs into a bounded text block with cooperative cancellation.
    /// </summary>
    /// <param name="runs">Styled text runs.</param>
    /// <param name="maxWidth">Maximum block width.</param>
    /// <param name="maxHeight">Maximum block height.</param>
    /// <param name="lineHeightFactor">Multiplier used with the largest run font size to derive line height.</param>
    /// <param name="measure">Measurement delegate matching <see cref="OfficeRasterCanvas.MeasureText(string?, double, string?)"/>.</param>
    /// <param name="wrap">Whether soft wrapping is enabled. Hard line breaks are always honored.</param>
    /// <param name="shrinkToFit">Whether non-wrapped rich text should proportionally shrink run font sizes to fit the requested width.</param>
    /// <param name="minimumFontSize">Minimum font size for any run when <paramref name="shrinkToFit"/> is enabled.</param>
    /// <param name="overflowBehavior">How overflowing rich text should be represented in the returned layout.</param>
    /// <param name="paragraphIndent">Optional first-line and continuation-line offsets applied while laying out wrapped rich text.</param>
    /// <param name="cancellationToken">Token checked while normalizing, tokenizing, measuring, and wrapping rich text.</param>
    /// <returns>Measured rich text block with visible lines and clipping state.</returns>
    public static OfficeRichTextBlockLayout LayoutRichTextBlock(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, string?, double> measure,
        bool wrap,
        bool shrinkToFit,
        double minimumFontSize,
        OfficeTextOverflowBehavior overflowBehavior,
        OfficeTextParagraphIndent? paragraphIndent,
        CancellationToken cancellationToken) {
        if (measure == null) throw new ArgumentNullException(nameof(measure));
        return LayoutStyledRichTextBlock(runs, maxWidth, maxHeight, lineHeightFactor,
            (text, size, family, _) => measure(text, size, family), wrap, shrinkToFit,
            minimumFontSize, overflowBehavior, paragraphIndent, cancellationToken);
    }

    // Renderer adapters use the same wrapping algorithm while measuring the actual styled face.
    internal static OfficeRichTextBlockLayout LayoutStyledRichTextBlock(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, string?, OfficeFontStyle, double> measure,
        bool wrap,
        bool shrinkToFit = false,
        double minimumFontSize = 1D,
        OfficeTextOverflowBehavior overflowBehavior = OfficeTextOverflowBehavior.Ellipsis,
        OfficeTextParagraphIndent? paragraphIndent = null,
        CancellationToken cancellationToken = default,
        bool shrinkToHeight = false,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint = null,
        bool hardBreakStartsParagraph = true, OfficeTextTabStops? tabStops = null,
        OfficeTextTabLeaderLayout.Budget? leaderBudget = null) =>
        LayoutStyledRichTextBlockWithPaint(runs, maxWidth, maxHeight, lineHeightFactor, measure, wrap, shrinkToFit,
            minimumFontSize, overflowBehavior, paragraphIndent, cancellationToken, shrinkToHeight, measurePaint,
            hardBreakStartsParagraph, tabStops, leaderBudget);

    internal static OfficeRichTextBlockLayout LayoutStyledRichTextBlockWithPaint(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, string?, OfficeFontStyle, double> measure,
        bool wrap,
        bool shrinkToFit = false,
        double minimumFontSize = 1D,
        OfficeTextOverflowBehavior overflowBehavior = OfficeTextOverflowBehavior.Ellipsis,
        OfficeTextParagraphIndent? paragraphIndent = null,
        CancellationToken cancellationToken = default,
        bool shrinkToHeight = false,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint = null,
        bool hardBreakStartsParagraph = true, OfficeTextTabStops? tabStops = null,
        OfficeTextTabLeaderLayout.Budget? leaderBudget = null,
        Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?>? measureSegmentPaint = null) =>
        LayoutRichTextBlockWithMetrics(runs, maxWidth, maxHeight, lineHeightFactor, new OfficeRichTextMeasurement(measure),
            wrap, shrinkToFit, minimumFontSize, overflowBehavior, paragraphIndent, cancellationToken, shrinkToHeight,
            measurePaint, hardBreakStartsParagraph, tabStops, leaderBudget, measureSegmentPaint);

    internal static OfficeRichTextBlockLayout LayoutRichTextBlockWithMetrics(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        OfficeRichTextMeasurement measure,
        bool wrap,
        bool shrinkToFit = false,
        double minimumFontSize = 1D,
        OfficeTextOverflowBehavior overflowBehavior = OfficeTextOverflowBehavior.Ellipsis,
        OfficeTextParagraphIndent? paragraphIndent = null,
        CancellationToken cancellationToken = default,
        bool shrinkToHeight = false,
        Func<string?, double, string?, OfficeFontStyle, OfficeTextPaintBounds>? measurePaint = null,
        bool hardBreakStartsParagraph = true, OfficeTextTabStops? tabStops = null,
        OfficeTextTabLeaderLayout.Budget? leaderBudget = null,
        Func<OfficeRichTextSegment, double, OfficeTextPaintBounds?>? measureSegmentPaint = null) {
        cancellationToken.ThrowIfCancellationRequested();
        if (runs == null) {
            throw new ArgumentNullException(nameof(runs));
        }

        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        IReadOnlyList<OfficeRichTextRun> normalizedRuns =
            NormalizeRichTextRuns(runs, out bool inputTruncated, cancellationToken);
        double width = NormalizeNonNegative(maxWidth);
        OfficeTextParagraphIndent effectiveParagraphIndent = paragraphIndent ?? OfficeTextParagraphIndent.Empty;
        if (shrinkToFit && shrinkToHeight) {
            normalizedRuns = FitRichTextRunsToFrame(normalizedRuns, width, maxHeight,
                lineHeightFactor, measure, wrap, minimumFontSize,
                effectiveParagraphIndent, cancellationToken, measurePaint, measureSegmentPaint, out double appliedScale);
            effectiveParagraphIndent = effectiveParagraphIndent.Scale(appliedScale);
        } else if (shrinkToFit && !wrap) {
            double unwrappedWidth = MeasureMaxUnwrappedRichTextWidth(
                normalizedRuns,
                measure,
                effectiveParagraphIndent,
                cancellationToken);
            if (unwrappedWidth > width) {
                double maxFontSize = ResolveMaxRichTextFontSize(normalizedRuns);
                double minFontSize = Math.Min(maxFontSize, Math.Max(1D, NormalizePositive(minimumFontSize, 1D)));
                double scale = Math.Max(minFontSize / Math.Max(maxFontSize, 1D), width / Math.Max(unwrappedWidth, 1D));
                normalizedRuns = ScaleRichTextRuns(normalizedRuns, scale, cancellationToken);
                effectiveParagraphIndent = effectiveParagraphIndent.Scale(scale);
            }
        }

        return LayoutRichTextBlockCore(
            normalizedRuns,
            width,
            maxHeight,
            lineHeightFactor,
            measure,
            wrap,
            overflowBehavior,
            effectiveParagraphIndent,
            inputTruncated,
            cancellationToken,
            hardBreakStartsParagraph, tabStops, leaderBudget);
    }

    private static OfficeRichTextBlockLayout LayoutRichTextBlockCore(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        OfficeRichTextMeasurement measure,
        bool wrap,
        OfficeTextOverflowBehavior overflowBehavior,
        OfficeTextParagraphIndent paragraphIndent,
        bool inputTruncated,
        CancellationToken cancellationToken,
        bool hardBreakStartsParagraph = true, OfficeTextTabStops? tabStops = null,
        OfficeTextTabLeaderLayout.Budget? leaderBudget = null) {
        cancellationToken.ThrowIfCancellationRequested();
        double width = NormalizeNonNegative(maxWidth);
        double height = NormalizeNonNegative(maxHeight);
        double maxFontSize = ResolveMaxEffectiveRichTextFontSize(runs);
        double lineFactor = NormalizePositive(lineHeightFactor, 1.2D);
        double lineHeight = Math.Max(1D, Math.Ceiling(maxFontSize * lineFactor));
        var lines = new List<OfficeRichTextLine>();
        var builder = new RichTextLineBuilder(measure);
        OfficeTextParagraphIndent currentParagraphIndent = runs.Count > 0
            ? runs[0].ParagraphIndent ?? paragraphIndent
            : paragraphIndent;
        builder.SetOffset(ResolveLineOffset(currentParagraphIndent, firstVisualLine: true));
        bool atParagraphStart = true;
        bool clipped = inputTruncated;
        bool processingStopped = false;

        IEnumerable<RichTextToken> sourceTokens = CreateRichTextTokens(runs, cancellationToken, preserveTabs: tabStops != null);
        if (measure.UsesPublisherContext) sourceTokens = CombineContextWords(sourceTokens, cancellationToken);
        IReadOnlyList<RichTextToken>? tabTokens = tabStops == null ? null : new List<RichTextToken>(sourceTokens);
        int tokenIndex = 0;
        leaderBudget ??= new OfficeTextTabLeaderLayout.Budget();
        foreach (RichTextToken token in (IEnumerable<RichTextToken>?)tabTokens ?? sourceTokens) {
            int currentToken = tokenIndex++;
            cancellationToken.ThrowIfCancellationRequested();
            if (lines.Count >= MaximumLayoutLines) {
                clipped = true;
                processingStopped = true;
                break;
            }
            if (token.HardBreak) {
                builder.SourceTextEnd = token.SourceEnd;
                if (!AddRichTextLine(lines, builder)) {
                    clipped = true;
                    processingStopped = true;
                    break;
                }
                currentParagraphIndent = token.Run.ParagraphIndent ?? paragraphIndent;
                builder.SetOffset(ResolveLineOffset(currentParagraphIndent, firstVisualLine: hardBreakStartsParagraph));
                atParagraphStart = hardBreakStartsParagraph;
                continue;
            }

            if (atParagraphStart && token.Run.ParagraphIndent is { } runIndent) {
                currentParagraphIndent = runIndent;
                builder.SetOffset(ResolveLineOffset(currentParagraphIndent, firstVisualLine: true));
            }

            if (token.IsTab && tabStops != null) {
                double advance = ResolveTabAdvance(tabStops, builder.OffsetX + builder.Width, tabTokens!, currentToken, measure, cancellationToken, out OfficeTextTabStop? selectedStop);
                double room = Math.Max(0, width - builder.OffsetX - builder.Width);
                if (double.IsInfinity(advance) || advance > room) {
                    clipped = true;
                    if (wrap && !builder.IsEmpty) {
                        if (!AddRichTextLine(lines, builder)) { processingStopped = true; break; }
                        builder.SetOffset(ResolveLineOffset(currentParagraphIndent, firstVisualLine: false));
                        advance = ResolveTabAdvance(tabStops, builder.OffsetX, tabTokens!, currentToken, measure, cancellationToken, out selectedStop);
                        room = Math.Max(0, width - builder.OffsetX);
                    }
                    if (double.IsInfinity(advance) || advance > room) advance = 0;
                }
                bool leaderLimited = builder.AddTabAdvance(token.Run, advance, selectedStop, leaderBudget, cancellationToken);
                builder.SourceTextEnd = token.SourceEnd;
                clipped |= leaderLimited && leaderBudget.ReportClipping;
                atParagraphStart = false;
                continue;
            }

            if (token.IsWhitespace && builder.IsEmpty && tabStops == null) {
                builder.SourceTextEnd = token.SourceEnd;
                continue;
            }

            atParagraphStart = false;

            double tokenWidth = MeasureRichToken(token, measure);
            double availableWidth = Math.Max(0D, width - builder.OffsetX);
            if (wrap && builder.MeasureWidthAfter(token) > availableWidth && !builder.IsEmpty) {
                if (!AddRichTextLine(lines, builder)) {
                    clipped = true;
                    processingStopped = true;
                    break;
                }
                builder.SetOffset(ResolveLineOffset(currentParagraphIndent, firstVisualLine: false));
                if (token.IsWhitespace) {
                    lines[lines.Count - 1].SourceTextEnd = token.SourceEnd;
                    builder.SourceTextEnd = token.SourceEnd;
                    continue;
                }

                availableWidth = Math.Max(0D, width - builder.OffsetX);
            }

            if (wrap && tokenWidth > availableWidth && builder.IsEmpty) {
                if (!AddBrokenRichTextToken(
                    lines,
                    builder,
                    token,
                    width,
                    measure,
                    currentParagraphIndent,
                    cancellationToken)) {
                    clipped = true;
                    processingStopped = true;
                    break;
                }
            } else {
                builder.Add(token);
                builder.SourceTextEnd = token.SourceEnd;
            }
        }

        if (!processingStopped) {
            if (!AddRichTextLine(lines, builder)) clipped = true;
            else if (!inputTruncated) lines[lines.Count - 1].CompletesSource = true;
        }
        if (lines.Count == 0) {
            lines.Add(new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>()));
        }

        ApplyRichTextLineHeights(lines, lineFactor, maxFontSize);
        lineHeight = ResolveMaximumRichTextLineHeight(lines, lineHeight);

        bool clippedByOtherCauses = clipped;
        bool unwrappedWidthOverflow = false;
        if (!wrap) {
            foreach (OfficeRichTextLine line in lines) {
                if (line.OffsetX + line.Width > width + 0.01D) {
                    unwrappedWidthOverflow = true;
                    break;
                }
            }
        }
        if (unwrappedWidthOverflow) {
            if (overflowBehavior == OfficeTextOverflowBehavior.Ellipsis && lines[0].OffsetX + lines[0].Width > width + 0.01D) {
                lines[0] = TrimRichTextLineToWidthWithEllipsis(lines[0], Math.Max(0D, width - lines[0].OffsetX), measure);
            }

            clipped = true;
        }

        if (ClipRichTextLinesToHeight(lines, height, width, measure, overflowBehavior)) {
            clipped = true;
            clippedByOtherCauses = true;
        }

        double blockWidth = MeasureMaxRichTextLineWidth(lines);
        double blockHeight = MeasureRichTextBlockHeight(lines, lineHeight);
        return new OfficeRichTextBlockLayout(lines, lineHeight, blockWidth, blockHeight, clipped) {
            OnlyUnwrappedWidthOverflow = unwrappedWidthOverflow && !clippedByOtherCauses &&
                overflowBehavior == OfficeTextOverflowBehavior.Clip
        };
    }

    private static bool AddBrokenRichTextToken(
        List<OfficeRichTextLine> lines,
        RichTextLineBuilder builder,
        RichTextToken token,
        double maxWidth,
        OfficeRichTextMeasurement measure,
        OfficeTextParagraphIndent paragraphIndent,
        CancellationToken cancellationToken) {
        int sourceEnd = token.SourceStart;
        foreach (RichTextToken part in token.Parts ?? new[] { token })
        foreach (string textElement in OfficeTextElements.Enumerate(part.Text)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (lines.Count >= MaximumLayoutLines) {
                builder.Clear();
                return false;
            }
            double availableWidth = Math.Max(0D, maxWidth - builder.OffsetX);
            if (builder.MeasureWidthAfter(part.Run, textElement) > availableWidth && !builder.IsEmpty) {
                if (!AddRichTextLine(lines, builder)) {
                    return false;
                }
                builder.SetOffset(ResolveLineOffset(paragraphIndent, firstVisualLine: false));
            }

            builder.Add(part.Run, textElement);
            sourceEnd += textElement.Length;
            builder.SourceTextEnd = sourceEnd;
        }

        return true;
    }

    private static bool AddRichTextLine(List<OfficeRichTextLine> lines, RichTextLineBuilder builder) {
        if (lines.Count >= MaximumLayoutLines) {
            builder.Clear();
            return false;
        }
        if (builder.IsEmpty) {
            lines.Add(new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>(), offsetX: builder.OffsetX) { SourceTextEnd = builder.SourceTextEnd });
            return true;
        }

        lines.Add(builder.ToLine());
        builder.Clear();
        return true;
    }

    private static double Measure(string? text, double fontSize, string? fontFamily, Func<string?, double, string?, double> measure) =>
        string.IsNullOrEmpty(text) ? 0D : Math.Max(0D, measure(text, NormalizePositive(fontSize, 1D), fontFamily));

    private static OfficeRichTextSegment CreateRichTextSegment(string text, OfficeRichTextSegment style, Func<string?, double, string?, double> measure) =>
        CopyRichTextSegment(style, text, Measure(text, EffectiveFontSize(style), style.FontFamily, measure));

    private static double Measure(string? text, double fontSize, string? fontFamily, OfficeFontStyle fontStyle, OfficeRichTextMeasurement measure) =>
        measure.Measure(text, fontSize, fontFamily, fontStyle);

    private static double Measure(string? text, double fontSize, string? fontFamily, OfficeFontStyle fontStyle, Func<string?, double, string?, OfficeFontStyle, double> measure) =>
        string.IsNullOrEmpty(text) ? 0D : Math.Max(0D, measure(text, NormalizePositive(fontSize, 1D), fontFamily, fontStyle));

    private static double MeasureMaxRichTextLineWidth(IReadOnlyList<OfficeRichTextLine> lines) {
        double max = 0D;
        for (int i = 0; i < lines.Count; i++) {
            max = Math.Max(max, lines[i].OffsetX + lines[i].Width);
        }

        return max;
    }

    private static void ApplyRichTextLineHeights(List<OfficeRichTextLine> lines, double lineHeightFactor, double fallbackFontSize) {
        for (int i = 0; i < lines.Count; i++) {
            OfficeRichTextLine line = lines[i];
            lines[i] = line.WithPlacement(line.Segments,
                ResolveRichTextLineHeight(line, lineHeightFactor, fallbackFontSize), line.OffsetX);
        }
    }

    private static double ResolveRichTextLineHeight(OfficeRichTextLine line, double lineHeightFactor, double fallbackFontSize) {
        if (line.LineHeight > 0D) {
            return line.LineHeight;
        }

        ResolveRichTextVerticalExtents(line.Segments, out double top, out double bottom);
        double extent = bottom - top;
        if (extent <= 0D) extent = Math.Max(1D, fallbackFontSize);
        return Math.Max(1D, Math.Ceiling(extent * lineHeightFactor));
    }

    private static double ResolveMaximumRichTextLineHeight(IReadOnlyList<OfficeRichTextLine> lines, double fallbackLineHeight) {
        double maximum = Math.Max(1D, fallbackLineHeight);
        for (int i = 0; i < lines.Count; i++) {
            maximum = Math.Max(maximum, lines[i].LineHeight);
        }
        return maximum;
    }

    internal static void ResolveRichTextVerticalExtents(
        IReadOnlyList<OfficeRichTextSegment> segments,
        out double top,
        out double bottom) {
        top = 0D;
        bottom = 0D;
        bool hasSegment = false;
        for (int i = 0; i < segments.Count; i++) {
            OfficeRichTextSegment segment = segments[i];
            double renderedFontSize = EffectiveFontSize(segment);
            double baselineOffset = segment.Baseline == OfficeTextBaseline.Superscript
                ? -(segment.FontSize * 0.30D)
                : segment.Baseline == OfficeTextBaseline.Subscript ? segment.FontSize * 0.15D : 0D;
            double segmentTop = baselineOffset - (renderedFontSize * 0.84D);
            double segmentBottom = segmentTop + renderedFontSize;
            if (!hasSegment) {
                top = segmentTop;
                bottom = segmentBottom;
                hasSegment = true;
            } else {
                top = Math.Min(top, segmentTop);
                bottom = Math.Max(bottom, segmentBottom);
            }
        }
    }

    private static double MeasureRichTextBlockHeight(IReadOnlyList<OfficeRichTextLine> lines, double fallbackLineHeight) {
        double height = 0D;
        for (int i = 0; i < lines.Count; i++) {
            height += lines[i].LineHeight > 0D ? lines[i].LineHeight : fallbackLineHeight;
        }

        return height;
    }


}
