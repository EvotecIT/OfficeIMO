namespace OfficeIMO.Pdf;

internal static partial class PdfTextEditor {
    private static IReadOnlyList<TextSearchHit> FindHits(byte[] pdf, string text, PdfTextSearchOptions? options, PdfLoadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(text, nameof(text));
        if (text.Length == 0) return Array.Empty<TextSearchHit>();
        PdfTextSearchOptions snapshot = (options ?? new PdfTextSearchOptions()).Snapshot();
        PdfReadLimits limits = readOptions?.Limits ?? new PdfReadLimits();
        PdfReadDocument document = OpenForVisualTextEditing(pdf, readOptions);
        int[] pages = snapshot.PageNumbers == null || snapshot.PageNumbers.Length == 0
            ? Enumerable.Range(1, document.Pages.Count).ToArray()
            : snapshot.PageNumbers;
        for (int index = 0; index < pages.Length; index++) ValidatePage(pages[index], document.Pages.Count, nameof(options));
        StringComparison comparison = snapshot.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var hits = new List<TextSearchHit>();
        for (int pageIndex = 0; pageIndex < pages.Length; pageIndex++) {
            int pageNumber = pages[pageIndex];
            PdfReadPage page = document.Pages[pageNumber - 1];
            (double originX, double originY) = page.GetPageBoundaryOrigin();
            IReadOnlyList<PdfTextSpan> spans = page
                .GetTextSpans()
                .Where(span => IsSearchSpan(span, snapshot.IncludeTextRenderingMode3))
                .ToArray();
            List<TextLayoutEngine.TextLine> lines = BuildSearchLines(spans);
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                TextLayoutEngine.TextLine line = lines[lineIndex];
                if (line.Spans.Count == 0) continue;
                PdfTextSpan[] lineSpans = line.Spans.ToArray();
                var unit = new TextSearchUnit(lineSpans);
                if (unit.Text.Length == 0) continue;
                int start = 0;
                while (start <= unit.Text.Length - text.Length) {
                    int found = unit.Text.IndexOf(text, start, comparison);
                    if (found < 0) break;
                    start = found + Math.Max(1, text.Length);
                    if (snapshot.WholeWords && !HasWordBoundaries(unit.Text, found, text.Length)) continue;
                    if (unit.HasUnmappedBoundary(found, text.Length)) continue;
                    IReadOnlyList<TextSourceSegment> segments = unit.GetSourceSegments(found, text.Length);
                    if (segments.Count == 0) continue;
                    if (hits.Count >= limits.MaxTextSearchMatches) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.TextSearchMatches, limits.MaxTextSearchMatches, hits.Count + 1L);
                    }
                    PdfRegionText detected = BuildRegionText(new[] { segments[0].Span });
                    SpanBounds matchBounds = GetCombinedSegmentBounds(segments);
                    bool usesTextRenderingMode3 = segments.Any(static segment => segment.Span.TextRenderingMode == 3);
                    var match = new PdfTextMatch(pageNumber, unit.Text.Substring(found, text.Length), matchBounds.X - originX, matchBounds.Y - originY, matchBounds.Width, matchBounds.Height, detected.FontSize, detected.SuggestedFont, detected.SourceFont, detected.Color, detected.RotationDegrees, usesTextRenderingMode3);
                    hits.Add(new TextSearchHit(pageNumber, segments, lineSpans, match));
                }
            }
        }
        return hits;
    }

    private static List<TextLayoutEngine.TextLine> BuildSearchLines(IReadOnlyList<PdfTextSpan> spans) {
        var lines = new List<TextLayoutEngine.TextLine>();
        foreach (IGrouping<(double Rotation, bool IsTextRenderingMode3), PdfTextSpan> rotationGroup in spans.GroupBy(
                     static span => (Math.Round(span.RotationDegrees, 1), span.TextRenderingMode == 3))) {
            PdfTextSpan[] ordered = rotationGroup
                .OrderByDescending(static span => NormalPosition(span))
                .ThenBy(static span => BaselinePosition(span))
                .ToArray();
            var current = new List<PdfTextSpan>();
            double normal = 0D;
            for (int index = 0; index < ordered.Length; index++) {
                PdfTextSpan span = ordered[index];
                double spanNormal = NormalPosition(span);
                double tolerance = Math.Max(0.5D, Math.Min(2.5D, EffectiveFontSize(span) * 0.6D));
                bool newLine = current.Count > 0 && Math.Abs(spanNormal - normal) > tolerance;
                if (newLine) {
                    AddSearchLineSegments(lines, current);
                    current.Clear();
                }
                normal = current.Count == 0 ? spanNormal : ((normal * current.Count) + spanNormal) / (current.Count + 1);
                current.Add(span);
            }
            if (current.Count > 0) AddSearchLineSegments(lines, current);
        }
        return lines.OrderByDescending(static line => line.Y).ThenBy(static line => line.XStart).ToList();
    }

    private static void AddSearchLineSegments(List<TextLayoutEngine.TextLine> lines, List<PdfTextSpan> candidates) {
        PdfTextSpan[] ordered = OrderSpansInReadingDirection(candidates);
        var segment = new List<PdfTextSpan>();
        bool rightToLeft = UsesRightToLeftReadingOrder(ordered);
        for (int index = 0; index < ordered.Length; index++) {
            PdfTextSpan span = ordered[index];
            if (segment.Count > 0) {
                PdfTextSpan previous = segment[segment.Count - 1];
                double gap = rightToLeft
                    ? BaselinePosition(previous) - (BaselinePosition(span) + Math.Abs(span.Advance))
                    : BaselinePosition(span) - (BaselinePosition(previous) + Math.Abs(previous.Advance));
                if (IsIndependentFlowGap(gap, previous, span)) {
                    lines.Add(BuildSearchLine(segment));
                    segment.Clear();
                }
            }
            segment.Add(span);
        }
        if (segment.Count > 0) lines.Add(BuildSearchLine(segment));
    }

    private static TextLayoutEngine.TextLine BuildSearchLine(List<PdfTextSpan> spans) {
        PdfTextSpan[] ordered = OrderSpansInReadingDirection(spans);
        double start = BaselinePosition(ordered[0]);
        PdfTextSpan last = ordered[ordered.Length - 1];
        return new TextLayoutEngine.TextLine(NormalPosition(ordered[0]), start, BaselinePosition(last) + Math.Abs(last.Advance), string.Empty, ordered.ToList());
    }

    private static double BaselinePosition(PdfTextSpan span) {
        double radians = span.RotationDegrees * Math.PI / 180D;
        return (Math.Cos(radians) * span.X) + (Math.Sin(radians) * span.Y);
    }

    private static double NormalPosition(PdfTextSpan span) {
        double radians = span.RotationDegrees * Math.PI / 180D;
        return (-Math.Sin(radians) * span.X) + (Math.Cos(radians) * span.Y);
    }

    private static bool HasWordBoundaries(string text, int start, int length) {
        bool left = start == 0 || !IsWordCharacter(text, start - 1);
        int end = start + length;
        bool right = end == text.Length || !IsWordCharacter(text, end);
        return left && right;
    }

    private static bool IsWordCharacter(string text, int index) {
        if (index > 0 && char.IsLowSurrogate(text[index]) && char.IsHighSurrogate(text[index - 1])) index--;
        if (text[index] == '_' || char.IsLetterOrDigit(text, index)) return true;
        System.Globalization.UnicodeCategory category = char.GetUnicodeCategory(text, index);
        return category is System.Globalization.UnicodeCategory.NonSpacingMark or
            System.Globalization.UnicodeCategory.SpacingCombiningMark or
            System.Globalization.UnicodeCategory.EnclosingMark;
    }

    private static bool IsIndependentFlowGap(double gap, PdfTextSpan previous, PdfTextSpan current) =>
        gap >= Math.Max(24D, Math.Max(EffectiveFontSize(previous), EffectiveFontSize(current)) * 1.5D);

    private sealed class TextSearchUnit {
        private readonly TextCharacterSource?[] _sources;

        internal TextSearchUnit(IReadOnlyList<PdfTextSpan> spans) {
            PdfTextSpan[] orderedSpans = OrderSpansInReadingDirection(spans);
            bool rightToLeft = UsesRightToLeftReadingOrder(orderedSpans);
            var text = new System.Text.StringBuilder();
            var sources = new List<TextCharacterSource?>();
            PdfTextSpan? previous = null;
            for (int spanIndex = 0; spanIndex < orderedSpans.Length; spanIndex++) {
                PdfTextSpan span = orderedSpans[spanIndex];
                if (previous != null && NeedsSyntheticSpace(previous, span, text, rightToLeft)) {
                    text.Append(' ');
                    sources.Add(null);
                }
                for (int characterIndex = 0; characterIndex < span.Text.Length; characterIndex++) {
                    text.Append(span.Text[characterIndex]);
                    sources.Add(new TextCharacterSource(span, characterIndex));
                }
                previous = span;
            }
            Text = text.ToString();
            _sources = sources.ToArray();
        }

        internal string Text { get; }

        internal bool HasUnmappedBoundary(int start, int length) {
            if (length <= 0 || start < 0 || start + length > _sources.Length) return true;
            return !_sources[start].HasValue || !_sources[start + length - 1].HasValue;
        }

        internal List<TextSourceSegment> GetSourceSegments(int start, int length) {
            var segments = new List<TextSourceSegment>();
            TextSourceSegment? current = null;
            int end = Math.Min(_sources.Length, start + length);
            for (int index = start; index < end; index++) {
                TextCharacterSource? source = _sources[index];
                if (!source.HasValue) continue;
                if (current.HasValue && ReferenceEquals(current.Value.Span, source.Value.Span) && current.Value.Start + current.Value.Length == source.Value.CharacterIndex) {
                    current = new TextSourceSegment(current.Value.Span, current.Value.Start, current.Value.Length + 1);
                    segments[segments.Count - 1] = current.Value;
                } else {
                    current = new TextSourceSegment(source.Value.Span, source.Value.CharacterIndex, 1);
                    segments.Add(current.Value);
                }
            }
            return segments;
        }

        private static bool NeedsSyntheticSpace(PdfTextSpan previous, PdfTextSpan current, System.Text.StringBuilder text, bool rightToLeft) {
            if (text.Length == 0 || char.IsWhiteSpace(text[text.Length - 1]) || (current.Text.Length > 0 && char.IsWhiteSpace(current.Text[0]))) return false;
            if (previous.LogicalTrailingSpace || current.LogicalLeadingSpace) return true;
            double gap = rightToLeft
                ? BaselinePosition(previous) - (BaselinePosition(current) + Math.Max(0D, Math.Abs(current.Advance)))
                : BaselinePosition(current) - (BaselinePosition(previous) + Math.Max(0D, Math.Abs(previous.Advance)));
            return gap > Math.Max(1D, Math.Min(EffectiveFontSize(previous), EffectiveFontSize(current)) * 0.18D);
        }
    }

    private static PdfTextSpan[] OrderSpansInReadingDirection(IEnumerable<PdfTextSpan> spans) {
        PdfTextSpan[] values = spans.ToArray();
        bool rightToLeft = UsesRightToLeftReadingOrder(values);
        return rightToLeft
            ? values.OrderByDescending(static span => BaselinePosition(span)).ThenByDescending(static span => NormalPosition(span)).ToArray()
            : values.OrderBy(static span => BaselinePosition(span)).ThenByDescending(static span => NormalPosition(span)).ToArray();
    }

    private static bool UsesRightToLeftReadingOrder(PdfTextSpan[] spans) {
        int rightToLeft = 0;
        int leftToRight = 0;
        for (int spanIndex = 0; spanIndex < spans.Length; spanIndex++) {
            string text = spans[spanIndex].Text;
            for (int index = 0; index < text.Length; index++) {
                int codePoint = text[index];
                if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                    codePoint = char.ConvertToUtf32(text[index], text[++index]);
                }
                if (IsRightToLeftCodePoint(codePoint)) rightToLeft++;
                else if (codePoint <= char.MaxValue && char.IsLetterOrDigit((char)codePoint)) leftToRight++;
            }
        }
        return rightToLeft > leftToRight;
    }

    private static bool IsRightToLeftCodePoint(int value) =>
        value >= 0x0590 && value <= 0x08FF ||
        value >= 0xFB1D && value <= 0xFDFF ||
        value >= 0xFE70 && value <= 0xFEFF ||
        value >= 0x10800 && value <= 0x10FFF ||
        value >= 0x1E800 && value <= 0x1EEFF;

    private sealed class TextSearchHit {
        internal TextSearchHit(int pageNumber, IReadOnlyList<TextSourceSegment> segments, PdfTextSpan[] lineSpans, PdfTextMatch match) { PageNumber = pageNumber; Segments = segments.ToArray(); LineSpans = lineSpans; Match = match; }
        internal int PageNumber { get; }
        internal TextSourceSegment[] Segments { get; }
        internal PdfTextSpan[] LineSpans { get; }
        internal PdfTextMatch Match { get; }
    }

    private readonly struct TextCharacterSource {
        internal TextCharacterSource(PdfTextSpan span, int characterIndex) { Span = span; CharacterIndex = characterIndex; }
        internal PdfTextSpan Span { get; }
        internal int CharacterIndex { get; }
    }

    private readonly struct TextSourceSegment {
        internal TextSourceSegment(PdfTextSpan span, int start, int length) { Span = span; Start = start; Length = length; }
        internal PdfTextSpan Span { get; }
        internal int Start { get; }
        internal int Length { get; }
    }
}
