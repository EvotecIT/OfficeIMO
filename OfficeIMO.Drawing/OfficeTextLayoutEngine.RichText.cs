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
        cancellationToken.ThrowIfCancellationRequested();
        if (runs == null) {
            throw new ArgumentNullException(nameof(runs));
        }

        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        IReadOnlyList<OfficeRichTextRun> normalizedRuns =
            NormalizeRichTextRuns(runs, cancellationToken);
        double width = NormalizeNonNegative(maxWidth);
        if (shrinkToFit && !wrap) {
            double unwrappedWidth = MeasureMaxUnwrappedRichTextWidth(
                normalizedRuns,
                measure,
                cancellationToken);
            if (unwrappedWidth > width) {
                double maxFontSize = ResolveMaxRichTextFontSize(normalizedRuns);
                double minFontSize = Math.Min(maxFontSize, Math.Max(1D, NormalizePositive(minimumFontSize, 1D)));
                double scale = Math.Max(minFontSize / Math.Max(maxFontSize, 1D), width / Math.Max(unwrappedWidth, 1D));
                normalizedRuns = ScaleRichTextRuns(normalizedRuns, scale, cancellationToken);
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
            paragraphIndent ?? OfficeTextParagraphIndent.Empty,
            cancellationToken);
    }

    private static OfficeRichTextBlockLayout LayoutRichTextBlockCore(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, string?, double> measure,
        bool wrap,
        OfficeTextOverflowBehavior overflowBehavior,
        OfficeTextParagraphIndent paragraphIndent,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        double width = NormalizeNonNegative(maxWidth);
        double height = NormalizeNonNegative(maxHeight);
        double maxFontSize = ResolveMaxRichTextFontSize(runs);
        double lineFactor = NormalizePositive(lineHeightFactor, 1.2D);
        double lineHeight = Math.Max(1D, Math.Ceiling(maxFontSize * lineFactor));
        var lines = new List<OfficeRichTextLine>();
        var builder = new RichTextLineBuilder(measure);
        builder.SetOffset(ResolveLineOffset(paragraphIndent, firstVisualLine: true));
        bool clipped = false;

        foreach (RichTextToken token in CreateRichTextTokens(runs, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (token.HardBreak) {
                AddRichTextLine(lines, builder);
                builder.SetOffset(ResolveLineOffset(paragraphIndent, firstVisualLine: true));
                continue;
            }

            if (token.IsWhitespace && builder.IsEmpty) {
                continue;
            }

            double tokenWidth = Measure(token.Text, token.Run.FontSize, token.Run.FontFamily, measure);
            double availableWidth = Math.Max(0D, width - builder.OffsetX);
            if (wrap && builder.Width + tokenWidth > availableWidth && !builder.IsEmpty) {
                AddRichTextLine(lines, builder);
                builder.SetOffset(ResolveLineOffset(paragraphIndent, firstVisualLine: false));
                if (token.IsWhitespace) {
                    continue;
                }

                availableWidth = Math.Max(0D, width - builder.OffsetX);
            }

            if (wrap && tokenWidth > availableWidth && builder.IsEmpty) {
                AddBrokenRichTextToken(
                    lines,
                    builder,
                    token,
                    width,
                    measure,
                    paragraphIndent,
                    cancellationToken);
            } else {
                builder.Add(token.Run, token.Text);
            }
        }

        AddRichTextLine(lines, builder);
        if (lines.Count == 0) {
            lines.Add(new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>()));
        }

        ApplyRichTextLineHeights(lines, lineFactor, maxFontSize);

        if (!wrap && lines.Count > 0 && lines[0].OffsetX + lines[0].Width > width + 0.01D) {
            if (overflowBehavior == OfficeTextOverflowBehavior.Ellipsis) {
                lines[0] = TrimRichTextLineToWidthWithEllipsis(lines[0], Math.Max(0D, width - lines[0].OffsetX), measure);
            }

            clipped = true;
        }

        if (ClipRichTextLinesToHeight(lines, height, width, measure, overflowBehavior)) {
            clipped = true;
        }

        double blockWidth = MeasureMaxRichTextLineWidth(lines);
        double blockHeight = MeasureRichTextBlockHeight(lines, lineHeight);
        return new OfficeRichTextBlockLayout(lines, lineHeight, blockWidth, blockHeight, clipped);
    }

    private static IReadOnlyList<OfficeRichTextRun> NormalizeRichTextRuns(
        IReadOnlyList<OfficeRichTextRun> runs,
        CancellationToken cancellationToken = default) {
        int runCapacity = Math.Min(runs.Count, MaximumLayoutTextRuns);
        var normalized = new List<OfficeRichTextRun>(runCapacity);
        int remainingCharacters = MaximumLayoutTextCharacters;
        for (int i = 0; i < runs.Count && i < MaximumLayoutTextRuns && remainingCharacters > 0; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeRichTextRun run = runs[i];
            string text = run.Text ?? string.Empty;
            if (text.Length > remainingCharacters) {
                int length = remainingCharacters;
                if (length > 0 && char.IsHighSurrogate(text[length - 1])) length--;
                text = text.Substring(0, length) + "...";
            }
            normalized.Add(new OfficeRichTextRun(
                text,
                NormalizePositive(run.FontSize, 1D),
                run.Color,
                run.Bold,
                run.Italic,
                run.Underline,
                run.FontFamily,
                run.Strikethrough,
                run.BackgroundColor));
            remainingCharacters -= Math.Min(remainingCharacters, run.Text?.Length ?? 0);
        }

        return normalized;
    }

    private static IReadOnlyList<OfficeRichTextRun> ScaleRichTextRuns(
        IReadOnlyList<OfficeRichTextRun> runs,
        double scale,
        CancellationToken cancellationToken = default) {
        double factor = Math.Max(0D, scale);
        var scaled = new List<OfficeRichTextRun>(runs.Count);
        for (int i = 0; i < runs.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeRichTextRun run = runs[i];
            scaled.Add(new OfficeRichTextRun(
                run.Text,
                Math.Max(1D, run.FontSize * factor),
                run.Color,
                run.Bold,
                run.Italic,
                run.Underline,
                run.FontFamily,
                run.Strikethrough,
                run.BackgroundColor));
        }

        return scaled;
    }

    private static double MeasureMaxUnwrappedRichTextWidth(
        IReadOnlyList<OfficeRichTextRun> runs,
        Func<string?, double, string?, double> measure,
        CancellationToken cancellationToken) {
        double current = 0D;
        double max = 0D;
        foreach (RichTextToken token in CreateRichTextTokens(runs, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            if (token.HardBreak) {
                max = Math.Max(max, current);
                current = 0D;
                continue;
            }

            current += Measure(token.Text, token.Run.FontSize, token.Run.FontFamily, measure);
        }

        return Math.Max(max, current);
    }

    private static double ResolveMaxRichTextFontSize(IReadOnlyList<OfficeRichTextRun> runs) {
        double max = 1D;
        for (int i = 0; i < runs.Count; i++) {
            max = Math.Max(max, NormalizePositive(runs[i].FontSize, 1D));
        }

        return max;
    }

    private static IEnumerable<RichTextToken> CreateRichTextTokens(
        IReadOnlyList<OfficeRichTextRun> runs,
        CancellationToken cancellationToken) {
        for (int i = 0; i < runs.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeRichTextRun run = runs[i];
            string normalized = ExpandTabs(run.Text.Replace("\r\n", "\n").Replace('\r', '\n'));
            cancellationToken.ThrowIfCancellationRequested();
            var word = new StringBuilder();
            for (int c = 0; c < normalized.Length; c++) {
                cancellationToken.ThrowIfCancellationRequested();
                char value = normalized[c];
                if (value == '\n') {
                    foreach (RichTextToken token in FlushRichTextWord(run, word)) {
                        yield return token;
                    }

                    yield return RichTextToken.CreateHardBreak(run);
                    continue;
                }

                if (char.IsWhiteSpace(value)) {
                    foreach (RichTextToken token in FlushRichTextWord(run, word)) {
                        yield return token;
                    }

                    yield return RichTextToken.CreateText(run, " ", isWhitespace: true);
                    continue;
                }

                word.Append(value);
            }

            foreach (RichTextToken token in FlushRichTextWord(run, word)) {
                yield return token;
            }
        }
    }

    private static IEnumerable<RichTextToken> FlushRichTextWord(OfficeRichTextRun run, StringBuilder word) {
        if (word.Length == 0) {
            yield break;
        }

        yield return RichTextToken.CreateText(run, word.ToString(), isWhitespace: false);
        word.Clear();
    }

    private static void AddBrokenRichTextToken(
        List<OfficeRichTextLine> lines,
        RichTextLineBuilder builder,
        RichTextToken token,
        double maxWidth,
        Func<string?, double, string?, double> measure,
        OfficeTextParagraphIndent paragraphIndent,
        CancellationToken cancellationToken) {
        foreach (string textElement in OfficeTextElements.Enumerate(token.Text)) {
            cancellationToken.ThrowIfCancellationRequested();
            double width = Measure(textElement, token.Run.FontSize, token.Run.FontFamily, measure);
            double availableWidth = Math.Max(0D, maxWidth - builder.OffsetX);
            if (builder.Width + width > availableWidth && !builder.IsEmpty) {
                AddRichTextLine(lines, builder);
                builder.SetOffset(ResolveLineOffset(paragraphIndent, firstVisualLine: false));
            }

            builder.Add(token.Run, textElement);
        }
    }

    private static void AddRichTextLine(List<OfficeRichTextLine> lines, RichTextLineBuilder builder) {
        if (builder.IsEmpty) {
            lines.Add(new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>(), offsetX: builder.OffsetX));
            return;
        }

        lines.Add(builder.ToLine());
        builder.Clear();
    }

    private static OfficeRichTextLine TrimRichTextLineToWidthWithEllipsis(OfficeRichTextLine line, double maxWidth, Func<string?, double, string?, double> measure) {
        if (line.Segments.Count == 0) {
            return line;
        }

        var segments = new List<OfficeRichTextSegment>(line.Segments.Count);
        for (int i = 0; i < line.Segments.Count; i++) {
            segments.Add(line.Segments[i]);
        }

        OfficeRichTextSegment ellipsisStyle = segments[segments.Count - 1];
        double width = NormalizeNonNegative(maxWidth);
        while (segments.Count > 0) {
            OfficeRichTextLine candidate = CreateRichTextLineWithEllipsis(segments, ellipsisStyle, measure, line.LineHeight);
            if (candidate.Width <= width) {
                return new OfficeRichTextLine(candidate.Segments, candidate.LineHeight, line.OffsetX);
            }

            int last = segments.Count - 1;
            OfficeRichTextSegment segment = segments[last];
            string text = OfficeTextElements.RemoveLast(segment.Text);
            if (text.Length == 0) {
                segments.RemoveAt(last);
            } else {
                segments[last] = new OfficeRichTextSegment(text, Measure(text, segment.FontSize, segment.FontFamily, measure), segment.FontSize, segment.Color, segment.Bold, segment.Italic, segment.Underline, segment.FontFamily, segment.Strikethrough, segment.BackgroundColor);
            }
        }

        const string ellipsis = "...";
        return Measure(ellipsis, ellipsisStyle.FontSize, ellipsisStyle.FontFamily, measure) <= width
            ? new OfficeRichTextLine(new[] { CreateRichTextSegment(ellipsis, ellipsisStyle, measure) }, line.LineHeight, line.OffsetX)
            : new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>(), line.LineHeight, line.OffsetX);
    }

    private static OfficeRichTextLine CreateRichTextLineWithEllipsis(List<OfficeRichTextSegment> segments, OfficeRichTextSegment ellipsisStyle, Func<string?, double, string?, double> measure, double lineHeight) {
        var measured = new List<OfficeRichTextSegment>(segments.Count);
        for (int i = 0; i < segments.Count; i++) {
            OfficeRichTextSegment segment = segments[i];
            string text = i == segments.Count - 1 ? segment.Text + "..." : segment.Text;
            measured.Add(new OfficeRichTextSegment(text, Measure(text, segment.FontSize, segment.FontFamily, measure), segment.FontSize, segment.Color, segment.Bold, segment.Italic, segment.Underline, segment.FontFamily, segment.Strikethrough, segment.BackgroundColor));
        }

        if (measured.Count == 0) {
            const string ellipsis = "...";
            measured.Add(CreateRichTextSegment(ellipsis, ellipsisStyle, measure));
        }

        return new OfficeRichTextLine(measured, lineHeight);
    }

    private static OfficeRichTextLine CreateRichTextLine(List<OfficeRichTextSegment> segments, Func<string?, double, string?, double> measure) {
        var measured = new List<OfficeRichTextSegment>(segments.Count);
        for (int i = 0; i < segments.Count; i++) {
            OfficeRichTextSegment segment = segments[i];
            measured.Add(new OfficeRichTextSegment(segment.Text, Measure(segment.Text, segment.FontSize, segment.FontFamily, measure), segment.FontSize, segment.Color, segment.Bold, segment.Italic, segment.Underline, segment.FontFamily, segment.Strikethrough, segment.BackgroundColor));
        }

        return new OfficeRichTextLine(measured);
    }

    private static OfficeRichTextSegment CreateRichTextSegment(string text, OfficeRichTextSegment style, Func<string?, double, string?, double> measure) =>
        new OfficeRichTextSegment(text, Measure(text, style.FontSize, style.FontFamily, measure), style.FontSize, style.Color, style.Bold, style.Italic, style.Underline, style.FontFamily, style.Strikethrough, style.BackgroundColor);

    private static double Measure(string? text, double fontSize, string? fontFamily, Func<string?, double, string?, double> measure) =>
        string.IsNullOrEmpty(text) ? 0D : Math.Max(0D, measure(text, NormalizePositive(fontSize, 1D), fontFamily));

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
            lines[i] = new OfficeRichTextLine(
                line.Segments,
                ResolveRichTextLineHeight(line, lineHeightFactor, fallbackFontSize),
                line.OffsetX);
        }
    }

    private static double ResolveRichTextLineHeight(OfficeRichTextLine line, double lineHeightFactor, double fallbackFontSize) {
        if (line.LineHeight > 0D) {
            return line.LineHeight;
        }

        double fontSize = line.FontSize > 0D ? line.FontSize : Math.Max(1D, fallbackFontSize);
        return Math.Max(1D, Math.Ceiling(fontSize * lineHeightFactor));
    }

    private static bool ClipRichTextLinesToHeight(
        List<OfficeRichTextLine> lines,
        double maxHeight,
        double maxWidth,
        Func<string?, double, string?, double> measure,
        OfficeTextOverflowBehavior overflowBehavior) {
        if (lines.Count == 0) {
            return false;
        }

        double height = NormalizeNonNegative(maxHeight);
        double used = 0D;
        int visibleCount = 0;
        for (int i = 0; i < lines.Count; i++) {
            double lineHeight = Math.Max(1D, lines[i].LineHeight);
            if (visibleCount > 0 && used + lineHeight > height + 0.01D) {
                break;
            }

            used += lineHeight;
            visibleCount++;
        }

        visibleCount = Math.Max(1, visibleCount);
        if (visibleCount >= lines.Count) {
            return false;
        }

        lines.RemoveRange(visibleCount, lines.Count - visibleCount);
        if (overflowBehavior == OfficeTextOverflowBehavior.Ellipsis) {
            OfficeRichTextLine last = lines[lines.Count - 1];
            lines[lines.Count - 1] = TrimRichTextLineToWidthWithEllipsis(last, Math.Max(0D, maxWidth - last.OffsetX), measure);
        }

        return true;
    }

    private static double MeasureRichTextBlockHeight(IReadOnlyList<OfficeRichTextLine> lines, double fallbackLineHeight) {
        double height = 0D;
        for (int i = 0; i < lines.Count; i++) {
            height += lines[i].LineHeight > 0D ? lines[i].LineHeight : fallbackLineHeight;
        }

        return height;
    }

    private readonly struct RichTextToken {
        private RichTextToken(OfficeRichTextRun run, string text, bool hardBreak, bool isWhitespace) {
            Run = run;
            Text = text;
            HardBreak = hardBreak;
            IsWhitespace = isWhitespace;
        }

        internal OfficeRichTextRun Run { get; }

        internal string Text { get; }

        internal bool HardBreak { get; }

        internal bool IsWhitespace { get; }

        internal static RichTextToken CreateText(OfficeRichTextRun run, string text, bool isWhitespace) =>
            new RichTextToken(run, text, hardBreak: false, isWhitespace);

        internal static RichTextToken CreateHardBreak(OfficeRichTextRun run) =>
            new RichTextToken(run, string.Empty, hardBreak: true, isWhitespace: false);
    }

    private sealed class RichTextLineBuilder {
        private readonly Func<string?, double, string?, double> _measure;
        private readonly List<OfficeRichTextSegment> _segments = new List<OfficeRichTextSegment>();

        internal RichTextLineBuilder(Func<string?, double, string?, double> measure) {
            _measure = measure;
        }

        internal bool IsEmpty => _segments.Count == 0;

        internal double Width { get; private set; }

        internal double OffsetX { get; private set; }

        internal void SetOffset(double offsetX) {
            if (IsEmpty) {
                OffsetX = offsetX > 0D && !double.IsNaN(offsetX) && !double.IsInfinity(offsetX) ? offsetX : 0D;
            }
        }

        internal void Add(OfficeRichTextRun run, string text) {
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            double measured = Measure(text, run.FontSize, run.FontFamily, _measure);
            if (_segments.Count > 0 && CanMerge(_segments[_segments.Count - 1], run)) {
                OfficeRichTextSegment previous = _segments[_segments.Count - 1];
                string mergedText = previous.Text + text;
                _segments[_segments.Count - 1] = new OfficeRichTextSegment(mergedText, Measure(mergedText, run.FontSize, run.FontFamily, _measure), run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor);
            } else {
                _segments.Add(new OfficeRichTextSegment(text, measured, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor));
            }

            Width += measured;
        }

        internal OfficeRichTextLine ToLine() =>
            new OfficeRichTextLine(new List<OfficeRichTextSegment>(_segments), offsetX: OffsetX);

        internal void Clear() {
            _segments.Clear();
            Width = 0D;
        }

        private static bool CanMerge(OfficeRichTextSegment segment, OfficeRichTextRun run) =>
            segment.FontSize == run.FontSize &&
            segment.Color.Equals(run.Color) &&
            segment.Bold == run.Bold &&
            segment.Italic == run.Italic &&
            segment.Underline == run.Underline &&
            segment.Strikethrough == run.Strikethrough &&
            Nullable.Equals(segment.BackgroundColor, run.BackgroundColor) &&
            string.Equals(segment.FontFamily, run.FontFamily, StringComparison.Ordinal);
    }
}
