using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextLayoutEngine {
    private static IReadOnlyList<OfficeRichTextRun> NormalizeRichTextRuns(
        IReadOnlyList<OfficeRichTextRun> runs,
        out bool truncated,
        CancellationToken cancellationToken = default) {
        int runCapacity = Math.Min(runs.Count, MaximumLayoutTextRuns);
        var normalized = new List<OfficeRichTextRun>(runCapacity);
        int remainingCharacters = MaximumLayoutTextCharacters;
        int processedRuns = 0;
        truncated = runs.Count > MaximumLayoutTextRuns;
        for (int i = 0; i < runs.Count && i < MaximumLayoutTextRuns && remainingCharacters > 0; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            OfficeRichTextRun run = runs[i];
            string text = run.Text ?? string.Empty;
            if (text.Length > remainingCharacters) {
                truncated = true;
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
                run.BackgroundColor,
                run.UnderlineStyle,
                run.StrikethroughStyle,
                run.Baseline) { LinkUri = run.LinkUri });
            remainingCharacters -= Math.Min(remainingCharacters, run.Text?.Length ?? 0);
            processedRuns = i + 1;
        }

        truncated |= processedRuns < runs.Count;

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
                run.BackgroundColor,
                run.UnderlineStyle,
                run.StrikethroughStyle,
                run.Baseline) { LinkUri = run.LinkUri });
        }

        return scaled;
    }

    private static double MeasureMaxUnwrappedRichTextWidth(
        IReadOnlyList<OfficeRichTextRun> runs,
        Func<string?, double, string?, OfficeFontStyle, double> measure,
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

            current += Measure(token.Text, token.Run.EffectiveFontSize, token.Run.FontFamily, token.Run.FontStyle, measure);
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

    private static double ResolveMaxEffectiveRichTextFontSize(IReadOnlyList<OfficeRichTextRun> runs) {
        double max = 1D;
        for (int i = 0; i < runs.Count; i++) {
            max = Math.Max(max, NormalizePositive(runs[i].EffectiveFontSize, 1D));
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
        private readonly Func<string?, double, string?, OfficeFontStyle, double> _measure;
        private readonly List<OfficeRichTextSegment> _segments = new List<OfficeRichTextSegment>();

        internal RichTextLineBuilder(Func<string?, double, string?, OfficeFontStyle, double> measure) {
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

        internal double MeasureWidthAfter(OfficeRichTextRun run, string text) {
            if (_segments.Count > 0 && CanMerge(_segments[_segments.Count - 1], run)) {
                OfficeRichTextSegment previous = _segments[_segments.Count - 1];
                return Width - previous.Width + Measure(previous.Text + text, run.EffectiveFontSize, run.FontFamily, run.FontStyle, _measure);
            }
            return Width + Measure(text, run.EffectiveFontSize, run.FontFamily, run.FontStyle, _measure);
        }

        internal void Add(OfficeRichTextRun run, string text) {
            if (string.IsNullOrEmpty(text)) {
                return;
            }

            double measured = Measure(text, run.EffectiveFontSize, run.FontFamily, run.FontStyle, _measure);
            if (_segments.Count > 0 && CanMerge(_segments[_segments.Count - 1], run)) {
                OfficeRichTextSegment previous = _segments[_segments.Count - 1];
                string mergedText = previous.Text + text;
                double mergedWidth = Measure(mergedText, run.EffectiveFontSize, run.FontFamily, run.FontStyle, _measure);
                measured = mergedWidth - previous.Width;
                _segments[_segments.Count - 1] = CreateSegment(run, mergedText, mergedWidth);
            } else {
                _segments.Add(CreateSegment(run, text, measured));
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
            segment.LinkUri == run.LinkUri &&
            segment.FontSize == run.FontSize &&
            segment.Color.Equals(run.Color) &&
            segment.Bold == run.Bold &&
            segment.Italic == run.Italic &&
            segment.Underline == run.Underline &&
            segment.Strikethrough == run.Strikethrough &&
            segment.UnderlineStyle == run.UnderlineStyle &&
            segment.StrikethroughStyle == run.StrikethroughStyle &&
            segment.Baseline == run.Baseline &&
            Nullable.Equals(segment.BackgroundColor, run.BackgroundColor) &&
            string.Equals(segment.FontFamily, run.FontFamily, StringComparison.Ordinal);

        private static OfficeRichTextSegment CreateSegment(OfficeRichTextRun run, string text, double width) =>
            new OfficeRichTextSegment(text, width, run.FontSize, run.Color, run.Bold, run.Italic, run.Underline, run.FontFamily, run.Strikethrough, run.BackgroundColor, run.UnderlineStyle, run.StrikethroughStyle, run.Baseline) { LinkUri = run.LinkUri };
    }
}
