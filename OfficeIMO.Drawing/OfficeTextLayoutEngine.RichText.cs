using System;
using System.Collections.Generic;
using System.Text;

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
    /// <returns>Measured rich text block with visible lines and clipping state.</returns>
    public static OfficeRichTextBlockLayout LayoutRichTextBlock(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, double> measure,
        bool wrap,
        bool shrinkToFit = false,
        double minimumFontSize = 1D) {
        if (runs == null) {
            throw new ArgumentNullException(nameof(runs));
        }

        if (measure == null) {
            throw new ArgumentNullException(nameof(measure));
        }

        IReadOnlyList<OfficeRichTextRun> normalizedRuns = NormalizeRichTextRuns(runs);
        double width = NormalizeNonNegative(maxWidth);
        if (shrinkToFit && !wrap) {
            double unwrappedWidth = MeasureMaxUnwrappedRichTextWidth(normalizedRuns, measure);
            if (unwrappedWidth > width) {
                double maxFontSize = ResolveMaxRichTextFontSize(normalizedRuns);
                double minFontSize = Math.Min(maxFontSize, Math.Max(1D, NormalizePositive(minimumFontSize, 1D)));
                double scale = Math.Max(minFontSize / Math.Max(maxFontSize, 1D), width / Math.Max(unwrappedWidth, 1D));
                normalizedRuns = ScaleRichTextRuns(normalizedRuns, scale);
            }
        }

        return LayoutRichTextBlockCore(normalizedRuns, width, maxHeight, lineHeightFactor, measure, wrap);
    }

    private static OfficeRichTextBlockLayout LayoutRichTextBlockCore(
        IReadOnlyList<OfficeRichTextRun> runs,
        double maxWidth,
        double maxHeight,
        double lineHeightFactor,
        Func<string?, double, double> measure,
        bool wrap) {
        double width = NormalizeNonNegative(maxWidth);
        double height = NormalizeNonNegative(maxHeight);
        double maxFontSize = ResolveMaxRichTextFontSize(runs);
        double lineFactor = NormalizePositive(lineHeightFactor, 1.2D);
        double lineHeight = Math.Max(1D, Math.Ceiling(maxFontSize * lineFactor));
        var lines = new List<OfficeRichTextLine>();
        var builder = new RichTextLineBuilder(measure);
        bool clipped = false;

        foreach (RichTextToken token in CreateRichTextTokens(runs)) {
            if (token.HardBreak) {
                AddRichTextLine(lines, builder);
                continue;
            }

            if (token.IsWhitespace && builder.IsEmpty) {
                continue;
            }

            double tokenWidth = Measure(token.Text, token.Run.FontSize, measure);
            if (wrap && builder.Width + tokenWidth > width && !builder.IsEmpty) {
                AddRichTextLine(lines, builder);
                if (token.IsWhitespace) {
                    continue;
                }
            }

            if (wrap && tokenWidth > width && builder.IsEmpty) {
                AddBrokenRichTextToken(lines, builder, token, width, measure);
            } else {
                builder.Add(token.Run, token.Text);
            }
        }

        AddRichTextLine(lines, builder);
        if (lines.Count == 0) {
            lines.Add(new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>()));
        }

        if (!wrap && lines.Count > 0 && lines[0].Width > width + 0.01D) {
            lines[0] = TrimRichTextLineToWidthWithEllipsis(lines[0], width, measure);
            clipped = true;
        }

        int maxLines = Math.Max(1, (int)Math.Floor(height / lineHeight));
        if (lines.Count > maxLines) {
            clipped = true;
            lines.RemoveRange(maxLines, lines.Count - maxLines);
            if (lines.Count > 0) {
                lines[lines.Count - 1] = TrimRichTextLineToWidthWithEllipsis(lines[lines.Count - 1], width, measure);
            }
        }

        double blockWidth = MeasureMaxRichTextLineWidth(lines);
        double blockHeight = lines.Count * lineHeight;
        return new OfficeRichTextBlockLayout(lines, lineHeight, blockWidth, blockHeight, clipped);
    }

    private static IReadOnlyList<OfficeRichTextRun> NormalizeRichTextRuns(IReadOnlyList<OfficeRichTextRun> runs) {
        var normalized = new List<OfficeRichTextRun>(runs.Count);
        for (int i = 0; i < runs.Count; i++) {
            OfficeRichTextRun run = runs[i];
            normalized.Add(new OfficeRichTextRun(
                run.Text,
                NormalizePositive(run.FontSize, 1D),
                run.Color,
                run.Bold,
                run.Italic,
                run.Underline,
                run.FontFamily,
                run.Strikethrough));
        }

        return normalized;
    }

    private static IReadOnlyList<OfficeRichTextRun> ScaleRichTextRuns(IReadOnlyList<OfficeRichTextRun> runs, double scale) {
        double factor = Math.Max(0D, scale);
        var scaled = new List<OfficeRichTextRun>(runs.Count);
        for (int i = 0; i < runs.Count; i++) {
            OfficeRichTextRun run = runs[i];
            scaled.Add(new OfficeRichTextRun(
                run.Text,
                Math.Max(1D, run.FontSize * factor),
                run.Color,
                run.Bold,
                run.Italic,
                run.Underline,
                run.FontFamily,
                run.Strikethrough));
        }

        return scaled;
    }

    private static double MeasureMaxUnwrappedRichTextWidth(IReadOnlyList<OfficeRichTextRun> runs, Func<string?, double, double> measure) {
        double current = 0D;
        double max = 0D;
        foreach (RichTextToken token in CreateRichTextTokens(runs)) {
            if (token.HardBreak) {
                max = Math.Max(max, current);
                current = 0D;
                continue;
            }

            current += Measure(token.Text, token.Run.FontSize, measure);
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

    private static IEnumerable<RichTextToken> CreateRichTextTokens(IReadOnlyList<OfficeRichTextRun> runs) {
        for (int i = 0; i < runs.Count; i++) {
            OfficeRichTextRun run = runs[i];
            string normalized = run.Text.Replace("\r\n", "\n").Replace('\r', '\n');
            var word = new StringBuilder();
            for (int c = 0; c < normalized.Length; c++) {
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
        Func<string?, double, double> measure) {
        for (int i = 0; i < token.Text.Length; i++) {
            string character = token.Text[i].ToString();
            double width = Measure(character, token.Run.FontSize, measure);
            if (builder.Width + width > maxWidth && !builder.IsEmpty) {
                AddRichTextLine(lines, builder);
            }

            builder.Add(token.Run, character);
        }
    }

    private static void AddRichTextLine(List<OfficeRichTextLine> lines, RichTextLineBuilder builder) {
        if (builder.IsEmpty) {
            lines.Add(new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>()));
            return;
        }

        lines.Add(builder.ToLine());
        builder.Clear();
    }

    private static OfficeRichTextLine TrimRichTextLineToWidthWithEllipsis(OfficeRichTextLine line, double maxWidth, Func<string?, double, double> measure) {
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
            OfficeRichTextLine candidate = CreateRichTextLineWithEllipsis(segments, ellipsisStyle, measure);
            if (candidate.Width <= width) {
                return candidate;
            }

            int last = segments.Count - 1;
            OfficeRichTextSegment segment = segments[last];
            if (segment.Text.Length <= 1) {
                segments.RemoveAt(last);
            } else {
                string text = segment.Text.Substring(0, segment.Text.Length - 1);
                segments[last] = new OfficeRichTextSegment(text, Measure(text, segment.FontSize, measure), segment.FontSize, segment.Color, segment.Bold, segment.Italic, segment.Underline, segment.FontFamily, segment.Strikethrough);
            }
        }

        const string ellipsis = "...";
        return Measure(ellipsis, ellipsisStyle.FontSize, measure) <= width
            ? new OfficeRichTextLine(new[] { CreateRichTextSegment(ellipsis, ellipsisStyle, measure) })
            : new OfficeRichTextLine(Array.Empty<OfficeRichTextSegment>());
    }

    private static OfficeRichTextLine CreateRichTextLineWithEllipsis(List<OfficeRichTextSegment> segments, OfficeRichTextSegment ellipsisStyle, Func<string?, double, double> measure) {
        var measured = new List<OfficeRichTextSegment>(segments.Count);
        for (int i = 0; i < segments.Count; i++) {
            OfficeRichTextSegment segment = segments[i];
            string text = i == segments.Count - 1 ? segment.Text + "..." : segment.Text;
            measured.Add(new OfficeRichTextSegment(text, Measure(text, segment.FontSize, measure), segment.FontSize, segment.Color, segment.Bold, segment.Italic, segment.Underline, segment.FontFamily, segment.Strikethrough));
        }

        if (measured.Count == 0) {
            const string ellipsis = "...";
            measured.Add(CreateRichTextSegment(ellipsis, ellipsisStyle, measure));
        }

        return new OfficeRichTextLine(measured);
    }

    private static OfficeRichTextLine CreateRichTextLine(List<OfficeRichTextSegment> segments, Func<string?, double, double> measure) {
        var measured = new List<OfficeRichTextSegment>(segments.Count);
        for (int i = 0; i < segments.Count; i++) {
            OfficeRichTextSegment segment = segments[i];
            measured.Add(new OfficeRichTextSegment(segment.Text, Measure(segment.Text, segment.FontSize, measure), segment.FontSize, segment.Color, segment.Bold, segment.Italic, segment.Underline, segment.FontFamily, segment.Strikethrough));
        }

        return new OfficeRichTextLine(measured);
    }

    private static OfficeRichTextSegment CreateRichTextSegment(string text, OfficeRichTextSegment style, Func<string?, double, double> measure) =>
        new OfficeRichTextSegment(text, Measure(text, style.FontSize, measure), style.FontSize, style.Color, style.Bold, style.Italic, style.Underline, style.FontFamily, style.Strikethrough);

    private static double MeasureMaxRichTextLineWidth(IReadOnlyList<OfficeRichTextLine> lines) {
        double max = 0D;
        for (int i = 0; i < lines.Count; i++) {
            max = Math.Max(max, lines[i].Width);
        }

        return max;
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
        private readonly Func<string?, double, double> _measure;
        private readonly List<OfficeRichTextSegment> _segments = new List<OfficeRichTextSegment>();

        internal RichTextLineBuilder(Func<string?, double, double> measure) {
            _measure = measure;
        }

        internal bool IsEmpty => _segments.Count == 0;

        internal double Width { get; private set; }

        internal void Add(OfficeRichTextRun run, string text) {
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            double measured = Measure(text, run.FontSize, _measure);
            if (_segments.Count > 0 && CanMerge(_segments[_segments.Count - 1], run)) {
                OfficeRichTextSegment previous = _segments[_segments.Count - 1];
                string mergedText = previous.Text + text;
                _segments[_segments.Count - 1] = new OfficeRichTextSegment(mergedText, Measure(mergedText, run.FontSize, _measure), run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily, run.Strikethrough);
            } else {
                _segments.Add(new OfficeRichTextSegment(text, measured, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily, run.Strikethrough));
            }

            Width += measured;
        }

        internal OfficeRichTextLine ToLine() =>
            new OfficeRichTextLine(new List<OfficeRichTextSegment>(_segments));

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
            string.Equals(segment.FontFamily, run.FontFamily, StringComparison.Ordinal);
    }
}
