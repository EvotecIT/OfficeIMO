using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextLayoutEngine {
    // Paint/style boundaries do not turn part of a native word into a legal wrap point.
    private static IEnumerable<RichTextToken> CombineContextWords(IEnumerable<RichTextToken> tokens, CancellationToken token) {
        var word = new List<RichTextToken>();
        foreach (var item in tokens) {
            token.ThrowIfCancellationRequested();
            if (!item.HardBreak && !item.IsWhitespace && !item.IsTab) { word.Add(item); continue; }
            if (word.Count != 0) {
                yield return word.Count == 1 ? word[0] : RichTextToken.CreateWord(word.ToArray());
                word.Clear();
            }
            yield return item;
        }
        if (word.Count != 0) yield return word.Count == 1 ? word[0] : RichTextToken.CreateWord(word.ToArray());
    }

    private static double MeasureRichToken(RichTextToken token, OfficeRichTextMeasurement measure) {
        if (token.Parts == null) return measure.Measure(token.Text, token.Run);
        var builder = new RichTextLineBuilder(measure);
        builder.Add(token);
        return builder.Width;
    }

    private sealed partial class RichTextLineBuilder {
        private ContextPlan? _pendingContext;

        internal void Add(RichTextToken token) {
            if (token.Parts == null) Add(token.Run, token.Text);
            else foreach (var part in token.Parts) Add(part.Run, part.Text);
        }

        internal double MeasureWidthAfter(RichTextToken token) {
            if (token.Parts == null) return MeasureWidthAfter(token.Run, token.Text);
            var candidate = new RichTextLineBuilder(_measure) { Width = Width };
            candidate._segments.AddRange(_segments);
            candidate.Add(token);
            return candidate.Width;
        }

        internal double MeasurePrefix(int length) {
            double width = 0;
            foreach (var segment in _segments) {
                if (length >= segment.Text.Length) {
                    width += segment.Width; length -= segment.Text.Length;
                } else {
                    if (length > 0) width += segment.PreparedShaping != null
                        && segment.PreparedShaping.TrySlice(0, length, out var prefix) && prefix != null
                            ? prefix.Measure(segment.FontSize)
                            : _measure.Measure(segment.Text.Substring(0, length), segment);
                    break;
                }
                if (length == 0) break;
            }
            return width;
        }

        internal double MeasureFittingWidthAfter(OfficeRichTextRun run, string text,
            OfficeRichTextRun? followingRun, string? followingText) {
            if (!_measure.UsesPublisherContext || followingRun == null || followingText == null)
                return MeasureWidthAfter(run, text);
            var candidate = new RichTextLineBuilder(_measure) { Width = Width };
            candidate._segments.AddRange(_segments);
            candidate.Add(run, text);
            int length = 0;
            foreach (var segment in candidate._segments) length += segment.Text.Length;
            candidate.Add(followingRun, followingText);
            return candidate.MeasurePrefix(length);
        }

        internal void AddFitted(OfficeRichTextRun run, string text, double fittingWidth) {
            Add(run, text);
            if (_measure.UsesPublisherContext) _fittingWidth = fittingWidth;
        }

        private ContextPlan ContextCandidate(OfficeRichTextRun run, string text) {
            if (_pendingContext is { } pending && ReferenceEquals(pending.Run, run) && pending.Text == text) return pending;
            int start = _segments.Count;
            while (start > 0 && SameContext(_segments[start - 1], run)) start--;
            var group = new List<OfficeRichTextSegment>();
            for (int i = start; i < _segments.Count; i++) group.Add(_segments[i]);
            if (group.Count != 0 && CanMerge(group[group.Count - 1], run)) {
                var last = group[group.Count - 1];
                group[group.Count - 1] = CreateSegment(run, last.Text + text, 0);
            } else group.Add(CreateSegment(run, text, 0));
            var combined = new System.Text.StringBuilder();
            foreach (var part in group) combined.Append(part.Text);
            string value = combined.ToString();
            bool qualified = true;
            foreach (char character in value) if (character < ' ' || character > '~') { qualified = false; break; }
            OfficePreparedTextShaping? prepared = qualified && SimpleStyle(run) ? _measure.Prepare(value, group[0]) : null;
            var shaped = new List<OfficeRichTextSegment>(group.Count);
            int offset = 0;
            foreach (var part in group) {
                if (prepared == null || !prepared.TrySlice(offset, part.Text.Length, out var slice) || slice == null) {
                    prepared = null; break;
                }
                // Justification splits at whitespace; those paint boundaries must also retain whole clusters.
                int paintStart = 0;
                for (int index = 1; index <= part.Text.Length; index++) {
                    if (index < part.Text.Length && char.IsWhiteSpace(part.Text[index]) == char.IsWhiteSpace(part.Text[index - 1])) continue;
                    if (!slice.TrySlice(paintStart, index - paintStart, out _)) { prepared = null; break; }
                    paintStart = index;
                }
                if (prepared == null) break;
                shaped.Add(CopyRichTextSegment(part, part.Text, slice.Measure(part.FontSize)));
                shaped[shaped.Count - 1].PreparedShaping = slice;
                offset += part.Text.Length;
            }
            if (prepared == null) {
                shaped.Clear();
                foreach (var part in group) {
                    var unprepared = CopyRichTextSegment(part, part.Text, 0);
                    unprepared.PreparedShaping = null;
                    shaped.Add(CopyRichTextSegment(unprepared, part.Text, _measure.Measure(part.Text, unprepared)));
                }
            }
            var result = new List<OfficeRichTextSegment>(start + shaped.Count);
            double width = 0;
            for (int i = 0; i < start; i++) { result.Add(_segments[i]); width += _segments[i].Width; }
            foreach (var part in shaped) { result.Add(part); width += part.Width; }
            return _pendingContext = new ContextPlan(run, text, result, width);
        }

        private static bool SimpleStyle(OfficeRichTextRun run) => run.Baseline == OfficeTextBaseline.Normal
            && !run.Underline && !run.Strikethrough && !run.BackgroundColor.HasValue;

        private static bool SameContext(OfficeRichTextSegment segment, OfficeRichTextRun run) => segment.Text.Length != 0
            && SimpleStyle(run) && segment.Baseline == OfficeTextBaseline.Normal && !segment.Underline
            && !segment.Strikethrough && !segment.BackgroundColor.HasValue && segment.FontSize == run.FontSize
            && segment.Bold == run.Bold && segment.Italic == run.Italic && segment.LanguageId == run.LanguageId
            && string.Equals(segment.FontFamily, run.FontFamily, StringComparison.Ordinal);

        private sealed record ContextPlan(OfficeRichTextRun Run, string Text, List<OfficeRichTextSegment> Segments, double Width);
    }
}
