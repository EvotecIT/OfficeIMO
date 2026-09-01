namespace OfficeIMO.Pdf;

internal static partial class PdfTextEditor {
    private static TextSearchHit ResolveHit(byte[] pdf, PdfTextMatch match, PdfLoadOptions? readOptions) {
        if (match.Text.Length == 0) throw new ArgumentException("A located text occurrence must contain text.", nameof(match));
        IReadOnlyList<TextSearchHit> hits = FindHits(
            pdf,
            match.Text,
            new PdfTextSearchOptions { MatchCase = true, PageNumbers = new[] { match.PageNumber } },
            readOptions);
        TextSearchHit[] candidates = hits.Where(hit => SameMatchGeometry(hit.Match, match)).ToArray();
        if (candidates.Length != 1) {
            throw new InvalidOperationException("The located text occurrence no longer identifies exactly one source occurrence in this PDF revision.");
        }
        return candidates[0];
    }

    private static bool SameMatchGeometry(PdfTextMatch left, PdfTextMatch right) =>
        left.PageNumber == right.PageNumber &&
        string.Equals(left.Text, right.Text, StringComparison.Ordinal) &&
        Math.Abs(left.X - right.X) <= 0.01D &&
        Math.Abs(left.Y - right.Y) <= 0.01D &&
        Math.Abs(left.Width - right.Width) <= 0.01D &&
        Math.Abs(left.Height - right.Height) <= 0.01D;

    private static void AddExactMoveRequests(
        List<PdfStamper.TextStampRequest> requests,
        int pageNumber,
        PdfTextSpan sourceSpan,
        IReadOnlyList<TextSourceSegment> segments,
        double deltaX,
        double deltaY,
        PdfResolvedTextStyle sourceStyle,
        PdfResolvedTextStyle movedStyle) {
        TextSourceSegment[] ordered = segments.OrderBy(static segment => segment.Start).ToArray();
        int cursor = 0;
        int requestOrdinal = 0;
        for (int index = 0; index < ordered.Length; index++) {
            TextSourceSegment segment = ordered[index];
            if (segment.Start < cursor) {
                throw new InvalidOperationException("The located text occurrence contains overlapping source ranges.");
            }
            AddExactSpanSliceRequest(requests, pageNumber, sourceSpan, cursor, segment.Start - cursor, 0D, 0D, sourceStyle, requestOrdinal++);
            AddExactSpanSliceRequest(requests, pageNumber, sourceSpan, segment.Start, segment.Length, deltaX, deltaY, movedStyle, requestOrdinal++);
            cursor = segment.Start + segment.Length;
        }
        AddExactSpanSliceRequest(requests, pageNumber, sourceSpan, cursor, sourceSpan.Text.Length - cursor, 0D, 0D, sourceStyle, requestOrdinal);
    }

    private static void AddExactSpanSliceRequest(
        List<PdfStamper.TextStampRequest> requests,
        int pageNumber,
        PdfTextSpan sourceSpan,
        int start,
        int length,
        double deltaX,
        double deltaY,
        PdfResolvedTextStyle style,
        int requestOrdinal) {
        if (length <= 0) return;
        int end = start + length;
        while (start < end && char.IsWhiteSpace(sourceSpan.Text[start])) start++;
        while (end > start && char.IsWhiteSpace(sourceSpan.Text[end - 1])) end--;
        length = end - start;
        if (length <= 0) return;
        string text = GetAuthoredSlice(sourceSpan, start, length);
        if (text.Length == 0) return;
        double radians = sourceSpan.RotationDegrees * Math.PI / 180D;
        double baselineOffset = GetCharacterBoundaryAdvance(sourceSpan, start);
        double x = sourceSpan.X + Math.Cos(radians) * baselineOffset + deltaX;
        double y = sourceSpan.Y + Math.Sin(radians) * baselineOffset + deltaY;
        AddStampLines(requests, pageNumber, x, y, text, style, sourceSpan.PaintOrder + (requestOrdinal * 0.00000001D));
    }

    private static string GetAuthoredSlice(PdfTextSpan sourceSpan, int start, int length) {
        string source = sourceSpan.Text;
        string authored = TrimAuthoredEdgeWhitespace(sourceSpan.RestampText);
        int[]? boundaries = TryBuildAuthoredBoundaryMap(source, authored);
        if (boundaries == null) {
            authored = source;
            boundaries = BuildIdentityBoundaryMap(source.Length);
        }
        int authoredStart = boundaries[start];
        int authoredEnd = boundaries[start + length];
        return authored.Substring(authoredStart, authoredEnd - authoredStart);
    }

    private static double GetCharacterBoundaryAdvance(PdfTextSpan span, int characterIndex) {
        if (characterIndex <= 0) return 0D;
        if (span.CharacterAdvances != null && span.CharacterAdvances.Count == span.Text.Length) {
            return span.CharacterAdvances.Take(Math.Min(characterIndex, span.CharacterAdvances.Count)).Sum();
        }
        return Math.Abs(span.Advance) * characterIndex / Math.Max(1, span.Text.Length);
    }
}
