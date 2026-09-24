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
        string[] queries = PdfTextSearchNormalization.NormalizeQueries(text);
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
            var pageHits = new List<(int LineOrder, int Offset, TextSearchHit Hit)>();
            foreach (int[] flow in BuildSearchFlows(lines, limits.MaxTextSearchFlowComparisons)) {
                PdfTextSpan[][] flowLines = flow.Select(lineIndex => lines[lineIndex].Spans.ToArray()).ToArray();
                var unit = new TextSearchUnit(flowLines);
                if (unit.Text.Length == 0) continue;
                foreach (TextSearchRange range in unit.FindRanges(queries, comparison, snapshot.WholeWords)) {
                    IReadOnlyList<TextSourceSegment> segments = unit.GetSourceSegments(range.Start, range.Length);
                    if (segments.Count == 0) continue;
                    if (hits.Count + pageHits.Count >= limits.MaxTextSearchMatches) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.TextSearchMatches, limits.MaxTextSearchMatches, hits.Count + pageHits.Count + 1L);
                    }
                    PdfRegionText detected = BuildRegionText(new[] { segments[0].Span });
                    SpanBounds matchBounds = GetCombinedSegmentBounds(segments);
                    bool usesTextRenderingMode3 = segments.Any(static segment => segment.Span.TextRenderingMode == 3);
                    PdfSelectionQuad visualBounds = ToVisualQuad(page, matchBounds);
                    int[] touchedLines = segments.Select(segment => unit.GetLineIndex(segment.Span)).Distinct().OrderBy(static index => index).ToArray();
                    PdfSelectionQuad[] visualLineBounds = touchedLines.Length <= 1
                        ? new[] { visualBounds }
                        : touchedLines
                            .Select(lineIndex => ToVisualQuad(page, GetCombinedSegmentBounds(segments.Where(segment => unit.GetLineIndex(segment.Span) == lineIndex).ToArray())))
                            .ToArray();
                    var match = new PdfTextMatch(pageNumber, range.MatchedText, matchBounds.X - originX, matchBounds.Y - originY, matchBounds.Width, matchBounds.Height, detected.FontSize, detected.SuggestedFont, detected.SourceFont, detected.Color, detected.RotationDegrees, visualBounds, usesTextRenderingMode3, visualLineBounds);
                    PdfTextSpan[][] hitLines = touchedLines.Select(lineIndex => flowLines[lineIndex]).ToArray();
                    pageHits.Add((flow[touchedLines[0]], range.Start, new TextSearchHit(pageNumber, segments, hitLines, match)));
                }
            }
            hits.AddRange(pageHits
                .OrderBy(static hit => hit.LineOrder)
                .ThenBy(static hit => hit.Offset)
                .Select(static hit => hit.Hit));
        }
        return hits;
    }

    private static PdfSelectionQuad ToVisualQuad(PdfReadPage page, SpanBounds bounds) {
        PdfVisualBounds visual = page.TransformBoundsToVisual(bounds.X, bounds.Y, bounds.X + bounds.Width, bounds.Y + bounds.Height);
        return new PdfSelectionQuad(
            new PdfSelectionPoint(visual.Left, visual.Top), new PdfSelectionPoint(visual.Right, visual.Top),
            new PdfSelectionPoint(visual.Right, visual.Bottom), new PdfSelectionPoint(visual.Left, visual.Bottom));
    }

    /// <summary>
    /// Chains search lines that continue the same text flow (a wrapped paragraph) so phrases can match across line breaks.
    /// A line continues into the nearest following line of the same rotation and rendering mode whose baseline extent overlaps
    /// it within a line-spacing distance; links are kept only when both lines choose each other, which keeps independent
    /// columns and side-by-side regions apart. Every line belongs to exactly one returned flow, listed in reading order.
    /// </summary>
    internal static List<int[]> BuildSearchFlows(List<TextLayoutEngine.TextLine> lines, int maxComparisons = PdfReadLimits.DefaultMaxTextSearchFlowComparisons) {
        int count = lines.Count;
        long comparisons = 0;
        var geometry = new SearchLineGeometry[count];
        for (int index = 0; index < count; index++) geometry[index] = SearchLineGeometry.Create(lines[index]);
        int[] next = Enumerable.Repeat(-1, count).ToArray();
        int[] previous = Enumerable.Repeat(-1, count).ToArray();
        double[] nextDistance = Enumerable.Repeat(double.PositiveInfinity, count).ToArray();
        double[] previousDistance = Enumerable.Repeat(double.PositiveInfinity, count).ToArray();
        double[] nextOverlap = new double[count];
        double[] previousOverlap = new double[count];
        foreach (IGrouping<(double Rotation, bool IsTextRenderingMode3), int> group in Enumerable.Range(0, count)
                     .Where(index => geometry[index].IsValid)
                     .GroupBy(index => (geometry[index].Rotation, geometry[index].IsTextRenderingMode3))) {
            int[] ordered = group.OrderByDescending(index => geometry[index].Normal).ToArray();
            double maximumFontSize = ordered.Max(index => geometry[index].FontSize);
            for (int upperOrdinal = 0; upperOrdinal < ordered.Length; upperOrdinal++) {
                int upper = ordered[upperOrdinal];
                for (int lowerOrdinal = upperOrdinal + 1; lowerOrdinal < ordered.Length; lowerOrdinal++) {
                    int lower = ordered[lowerOrdinal];
                    double distance = geometry[upper].Normal - geometry[lower].Normal;
                    if (distance > maximumFontSize * 2D) break;
                    if (++comparisons > maxComparisons) {
                        throw PdfReadLimitException.Create(PdfReadLimitKind.TextSearchFlowComparisons, maxComparisons, comparisons);
                    }
                    double fontSize = Math.Max(geometry[upper].FontSize, geometry[lower].FontSize);
                    if (distance < fontSize * 0.6D || distance > fontSize * 2D) continue;
                    double overlap = Math.Min(geometry[upper].End, geometry[lower].End) - Math.Max(geometry[upper].Start, geometry[lower].Start);
                    if (overlap <= 0D) continue;
                    if (IsBetterFlowCandidate(distance, overlap, nextDistance[upper], nextOverlap[upper])) {
                        next[upper] = lower; nextDistance[upper] = distance; nextOverlap[upper] = overlap;
                    }
                    if (IsBetterFlowCandidate(distance, overlap, previousDistance[lower], previousOverlap[lower])) {
                        previous[lower] = upper; previousDistance[lower] = distance; previousOverlap[lower] = overlap;
                    }
                }
            }
        }

        var flows = new List<int[]>();
        for (int index = 0; index < count; index++) {
            if (lines[index].Spans.Count == 0) continue;
            bool continuesPrevious = previous[index] >= 0 && next[previous[index]] == index;
            if (continuesPrevious) continue;
            var flow = new List<int> { index };
            int current = index;
            while (next[current] >= 0 && previous[next[current]] == current) {
                current = next[current];
                flow.Add(current);
            }
            flows.Add(flow.ToArray());
        }
        return flows;
    }

    private static bool IsBetterFlowCandidate(double distance, double overlap, double bestDistance, double bestOverlap) =>
        distance < bestDistance - 0.01D || (Math.Abs(distance - bestDistance) <= 0.01D && overlap > bestOverlap);

    private readonly struct SearchLineGeometry {
        private SearchLineGeometry(double rotation, bool isTextRenderingMode3, double normal, double start, double end, double fontSize) {
            Rotation = rotation; IsTextRenderingMode3 = isTextRenderingMode3; Normal = normal; Start = start; End = end; FontSize = fontSize; IsValid = true;
        }

        internal bool IsValid { get; }
        internal double Rotation { get; }
        internal bool IsTextRenderingMode3 { get; }
        internal double Normal { get; }
        internal double Start { get; }
        internal double End { get; }
        internal double FontSize { get; }

        internal static SearchLineGeometry Create(TextLayoutEngine.TextLine line) {
            if (line.Spans.Count == 0) return default;
            PdfTextSpan first = line.Spans[0];
            double start = double.PositiveInfinity, end = double.NegativeInfinity, normal = 0D, fontSize = 0D;
            for (int index = 0; index < line.Spans.Count; index++) {
                PdfTextSpan span = line.Spans[index];
                double baseline = BaselinePosition(span);
                start = Math.Min(start, baseline);
                end = Math.Max(end, baseline + Math.Abs(span.Advance));
                normal += NormalPosition(span);
                fontSize = Math.Max(fontSize, EffectiveFontSize(span));
            }
            return new SearchLineGeometry(Math.Round(first.RotationDegrees, 1), first.TextRenderingMode == 3, normal / line.Spans.Count, start, end, fontSize);
        }
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
        private readonly bool[] _lineBreaks;
        private readonly Dictionary<PdfTextSpan, int> _lineIndexes = new Dictionary<PdfTextSpan, int>();

        internal TextSearchUnit(IReadOnlyList<PdfTextSpan> spans) : this(new[] { spans }) {
        }

        /// <summary>Builds one searchable text flow from consecutive lines, separated by unmapped line-break characters.</summary>
        internal TextSearchUnit(IReadOnlyList<IReadOnlyList<PdfTextSpan>> lines) {
            var text = new System.Text.StringBuilder();
            var sources = new List<TextCharacterSource?>();
            var lineBreaks = new List<bool>();
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                PdfTextSpan[] orderedSpans = OrderSpansInReadingDirection(lines[lineIndex]);
                if (orderedSpans.Length == 0) continue;
                if (text.Length > 0) {
                    text.Append('\n');
                    sources.Add(null);
                    lineBreaks.Add(true);
                }
                bool rightToLeft = UsesRightToLeftReadingOrder(orderedSpans);
                PdfTextSpan? previous = null;
                for (int spanIndex = 0; spanIndex < orderedSpans.Length; spanIndex++) {
                    PdfTextSpan span = orderedSpans[spanIndex];
                    _lineIndexes[span] = lineIndex;
                    if (previous != null && NeedsSyntheticSpace(previous, span, text, rightToLeft)) {
                        text.Append(' ');
                        sources.Add(null);
                        lineBreaks.Add(false);
                    }
                    for (int characterIndex = 0; characterIndex < span.Text.Length; characterIndex++) {
                        text.Append(span.Text[characterIndex]);
                        sources.Add(new TextCharacterSource(span, characterIndex));
                        lineBreaks.Add(false);
                    }
                    previous = span;
                }
            }
            Text = text.ToString();
            _sources = sources.ToArray();
            _lineBreaks = lineBreaks.ToArray();
        }

        internal string Text { get; }

        internal int GetLineIndex(PdfTextSpan span) => _lineIndexes.TryGetValue(span, out int lineIndex) ? lineIndex : 0;

        internal bool HasUnmappedBoundary(int start, int length) {
            if (length <= 0 || start < 0 || start + length > _sources.Length) return true;
            return !_sources[start].HasValue || !_sources[start + length - 1].HasValue;
        }

        /// <summary>
        /// Finds non-overlapping occurrences of a whitespace-normalized query. Whitespace runs, including line breaks between
        /// chained lines, match one query space. A letter-hyphen-line-break-letter junction matches both the joined word
        /// ("hyphenation") and the hyphenated compound without the break ("well-known").
        /// </summary>
        internal List<TextSearchRange> FindRanges(IReadOnlyList<string> queries, StringComparison comparison, bool wholeWords) {
            var candidates = new List<TextSearchRange>();
            PdfNormalizedSearchText joined = PdfNormalizedSearchText.Create(Text, _lineBreaks, removeLineEndHyphens: false, out bool hasHyphenJunction);
            foreach (string query in queries) CollectRanges(joined, query, comparison, wholeWords, candidates);
            if (hasHyphenJunction) {
                PdfNormalizedSearchText dehyphenated = PdfNormalizedSearchText.Create(Text, _lineBreaks, removeLineEndHyphens: true, out _);
                foreach (string query in queries) CollectRanges(dehyphenated, query, comparison, wholeWords, candidates);
            }
            if (candidates.Count <= 1) return candidates;
            var accepted = new List<TextSearchRange>(candidates.Count);
            int acceptedEnd = 0;
            foreach (TextSearchRange candidate in candidates.OrderBy(static range => range.Start).ThenByDescending(static range => range.Length)) {
                if (accepted.Count > 0 && candidate.Start < acceptedEnd) continue;
                accepted.Add(candidate);
                acceptedEnd = candidate.Start + candidate.Length;
            }
            return accepted;
        }

        private void CollectRanges(PdfNormalizedSearchText normalized, string query, StringComparison comparison, bool wholeWords, List<TextSearchRange> ranges) {
            string value = normalized.Text;
            int start = 0;
            while (start <= value.Length - query.Length) {
                int found = value.IndexOf(query, start, comparison);
                if (found < 0) break;
                start = found + Math.Max(1, query.Length);
                if (wholeWords && !HasWordBoundaries(value, found, query.Length)) continue;
                int sourceStart = normalized.GetSourceStart(found);
                int sourceEnd = normalized.GetSourceEnd(found + query.Length - 1);
                if (HasUnmappedBoundary(sourceStart, sourceEnd - sourceStart)) continue;
                ranges.Add(new TextSearchRange(sourceStart, sourceEnd - sourceStart, value.Substring(found, query.Length)));
            }
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

    private readonly struct TextSearchRange {
        internal TextSearchRange(int start, int length, string matchedText) { Start = start; Length = length; MatchedText = matchedText; }
        internal int Start { get; }
        internal int Length { get; }
        internal string MatchedText { get; }
    }

    private sealed class TextSearchHit {
        internal TextSearchHit(int pageNumber, IReadOnlyList<TextSourceSegment> segments, PdfTextSpan[][] lines, PdfTextMatch match) { PageNumber = pageNumber; Segments = segments.ToArray(); Lines = lines; Match = match; }
        internal int PageNumber { get; }
        internal TextSourceSegment[] Segments { get; }
        /// <summary>Source lines touched by the occurrence, in reading order; more than one when the occurrence wraps.</summary>
        internal PdfTextSpan[][] Lines { get; }
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
