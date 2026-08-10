using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Canonical extraction, removal, and stamping coordinator for existing-page text edits.</summary>
internal static class PdfTextEditor {
    internal static PdfRegionText Inspect(byte[] pdf, PdfPageRegion region, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(region, nameof(region));
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        ValidatePage(region.PageNumber, document.Pages.Count, nameof(region));
        IReadOnlyList<PdfTextSpan> selected = document.Pages[region.PageNumber - 1]
            .GetTextSpans()
            .Where(span => ContainsCenter(region, GetBounds(span)))
            .ToArray();
        return BuildRegionText(selected);
    }

    internal static IReadOnlyList<PdfTextMatch> Find(byte[] pdf, string text, PdfTextSearchOptions? options, PdfReadOptions? readOptions) {
        return FindHits(pdf, text, options, readOptions).Select(static hit => hit.Match).ToArray();
    }

    internal static TextMutationResult Add(byte[] pdf, PdfPageRegion region, string text, PdfTextEditOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(text, nameof(text));
        if (text.Length == 0) throw new ArgumentException("Added text cannot be empty.", nameof(text));
        ValidateRegionPage(pdf, region, readOptions);
        PdfTextEditOptions snapshot = (options ?? new PdfTextEditOptions()).Snapshot();
        PdfResolvedTextStyle style = ResolveStyle(snapshot, detected: null);
        double baselineY = region.Top - style.FontSize;
        byte[] output = StampLines(pdf, region.PageNumber, region.X, baselineY, text, style, readOptions);
        return new TextMutationResult(output, 1, Array.Empty<string>());
    }

    internal static TextMutationResult Replace(byte[] pdf, PdfPageRegion region, string text, PdfTextEditOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(text, nameof(text));
        PdfRegionText detected = Inspect(pdf, region, readOptions);
        PdfTextEditOptions snapshot = (options ?? new PdfTextEditOptions()).Snapshot();
        PdfResolvedTextStyle style = ResolveStyle(snapshot, detected);
        TextRemovalResult removal = detected.Spans.Count == 0
            ? new TextRemovalResult(pdf.ToArray(), Array.Empty<string>())
            : RemoveTextPreservingUnmatchedSpans(pdf, new[] { region.ToRedactionArea() }, readOptions);
        byte[] output = removal.Bytes;
        if (text.Length > 0) output = StampLines(output, region.PageNumber, detected.Spans.Count == 0 ? region.X : detected.BaselineX, detected.Spans.Count == 0 ? region.Top - style.FontSize : detected.BaselineY, text, style, readOptions);
        return new TextMutationResult(output, detected.Spans.Count, removal.Warnings.Concat(BuildSubstitutionWarnings(detected, style.Font)));
    }

    internal static TextMutationResult Move(byte[] pdf, PdfPageRegion source, double deltaX, double deltaY, PdfTextEditOptions? options, PdfReadOptions? readOptions) {
        ValidateFinite(deltaX, nameof(deltaX));
        ValidateFinite(deltaY, nameof(deltaY));
        PdfRegionText detected = Inspect(pdf, source, readOptions);
        if (detected.Spans.Count == 0 || detected.Text.Length == 0) return new TextMutationResult(pdf.ToArray(), 0, Array.Empty<string>());
        PdfResolvedTextStyle style = ResolveStyle((options ?? new PdfTextEditOptions()).Snapshot(), detected);
        TextRemovalResult removal = RemoveTextPreservingUnmatchedSpans(pdf, new[] { source.ToRedactionArea() }, readOptions);
        byte[] output = StampLines(removal.Bytes, source.PageNumber, detected.BaselineX + deltaX, detected.BaselineY + deltaY, detected.Text, style, readOptions);
        return new TextMutationResult(output, detected.Spans.Count, removal.Warnings.Concat(BuildSubstitutionWarnings(detected, style.Font)));
    }

    internal static TextMutationResult ReplaceAll(byte[] pdf, string text, string replacement, PdfTextSearchOptions? searchOptions, PdfTextEditOptions? editOptions, PdfReadOptions? readOptions) {
        Guard.NotNull(text, nameof(text));
        Guard.NotNull(replacement, nameof(replacement));
        if (text.Length == 0) throw new ArgumentException("Search text cannot be empty.", nameof(text));
        IReadOnlyList<TextSearchHit> hits = FindHits(pdf, text, searchOptions, readOptions);
        if (hits.Count == 0) return new TextMutationResult(pdf.ToArray(), 0, Array.Empty<string>());

        var byUnit = hits.GroupBy(static hit => hit.Unit).ToArray();
        PdfRedactionArea[] areas = byUnit
            .Select(group => {
                TextSearchHit first = group.First();
                return new PdfRedactionArea(first.PageNumber, first.Unit.Bounds.X, first.Unit.Bounds.Y, first.Unit.Bounds.Width, first.Unit.Bounds.Height);
            })
            .ToArray();
        TextRemovalResult removal = RemoveTextPreservingUnmatchedSpans(pdf, areas, readOptions);
        byte[] current = removal.Bytes;
        PdfTextEditOptions snapshot = (editOptions ?? new PdfTextEditOptions()).Snapshot();
        var warnings = new List<string>(removal.Warnings);
        foreach (IGrouping<TextSearchUnit, TextSearchHit> group in byUnit.OrderBy(static group => group.First().PageNumber).ThenByDescending(static group => group.Key.BaselineY).ThenBy(static group => group.Key.BaselineX)) {
            TextSearchUnit unit = group.Key;
            string rewritten = ReplaceOccurrences(unit.Text, group.Select(static hit => hit.StartIndex).OrderBy(static index => index).ToArray(), text.Length, replacement);
            PdfRegionText detected = BuildRegionText(unit.Spans);
            PdfResolvedTextStyle style = ResolveStyle(snapshot, detected);
            warnings.AddRange(BuildSubstitutionWarnings(detected, style.Font));
            if (rewritten.Length > 0) current = StampLines(current, group.First().PageNumber, unit.BaselineX, unit.BaselineY, rewritten, style, readOptions);
        }

        return new TextMutationResult(current, hits.Count, warnings);
    }

    private static TextRemovalResult RemoveTextPreservingUnmatchedSpans(byte[] pdf, IReadOnlyList<PdfRedactionArea> areas, PdfReadOptions? readOptions) {
        PdfReadDocument before = PdfReadDocument.Open(pdf, readOptions);
        int[] affectedPages = areas.Select(static area => area.PageNumber).Distinct().ToArray();
        var original = new List<PageTextSpanSnapshot>();
        for (int index = 0; index < affectedPages.Length; index++) {
            int pageNumber = affectedPages[index];
            ValidatePage(pageNumber, before.Pages.Count, nameof(areas));
            IReadOnlyList<PdfTextSpan> spans = before.Pages[pageNumber - 1].GetTextSpans();
            for (int spanIndex = 0; spanIndex < spans.Count; spanIndex++) {
                PdfTextSpan span = spans[spanIndex];
                SpanBounds bounds = GetBounds(span);
                bool targeted = areas.Any(area => area.PageNumber == pageNumber && ContainsCenter(area, bounds));
                original.Add(new PageTextSpanSnapshot(pageNumber, span, targeted));
            }
        }

        byte[] removed = PdfRedactionApplier.RemoveTextInAreas(pdf, areas, readOptions: readOptions);
        PdfReadDocument after = PdfReadDocument.Open(removed, readOptions);
        var remainingByPage = affectedPages.ToDictionary(
            static page => page,
            page => after.Pages[page - 1].GetTextSpans().ToList());
        var missing = new List<PageTextSpanSnapshot>();
        for (int index = 0; index < original.Count; index++) {
            PageTextSpanSnapshot candidate = original[index];
            if (candidate.Targeted) continue;
            List<PdfTextSpan> remaining = remainingByPage[candidate.PageNumber];
            int matchIndex = remaining.FindIndex(span => SameTextPlacement(span, candidate.Span));
            if (matchIndex >= 0) remaining.RemoveAt(matchIndex);
            else missing.Add(candidate);
        }

        byte[] current = removed;
        var warnings = new List<string>();
        for (int index = 0; index < missing.Count; index++) {
            PageTextSpanSnapshot snapshot = missing[index];
            PdfRegionText detected = BuildRegionText(new[] { snapshot.Span });
            PdfResolvedTextStyle style = ResolveStyle(new PdfTextEditOptions(), detected);
            warnings.AddRange(BuildSubstitutionWarnings(detected, style.Font));
            current = StampLines(current, snapshot.PageNumber, snapshot.Span.X, snapshot.Span.Y, snapshot.Span.Text, style, readOptions);
        }
        return new TextRemovalResult(current, warnings);
    }

    private static IReadOnlyList<TextSearchHit> FindHits(byte[] pdf, string text, PdfTextSearchOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(text, nameof(text));
        if (text.Length == 0) return Array.Empty<TextSearchHit>();
        PdfTextSearchOptions snapshot = (options ?? new PdfTextSearchOptions()).Snapshot();
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        int[] pages = snapshot.PageNumbers == null || snapshot.PageNumbers.Length == 0
            ? Enumerable.Range(1, document.Pages.Count).ToArray()
            : snapshot.PageNumbers;
        for (int index = 0; index < pages.Length; index++) ValidatePage(pages[index], document.Pages.Count, nameof(options));
        StringComparison comparison = snapshot.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var hits = new List<TextSearchHit>();
        for (int pageIndex = 0; pageIndex < pages.Length; pageIndex++) {
            int pageNumber = pages[pageIndex];
            IReadOnlyList<PdfTextSpan> spans = document.Pages[pageNumber - 1].GetTextSpans();
            List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans);
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                TextLayoutEngine.TextLine line = lines[lineIndex];
                if (line.Spans.Count == 0 || line.Text.Length == 0) continue;
                var unit = new TextSearchUnit(line.Text, line.Spans, GetCombinedBounds(line.Spans));
                int start = 0;
                while (start <= unit.Text.Length - text.Length) {
                    int found = unit.Text.IndexOf(text, start, comparison);
                    if (found < 0) break;
                    start = found + Math.Max(1, text.Length);
                    if (snapshot.WholeWords && !HasWordBoundaries(unit.Text, found, text.Length)) continue;
                    PdfRegionText detected = BuildRegionText(line.Spans);
                    SpanBounds matchBounds = SliceUnitBounds(unit, found, text.Length);
                    var match = new PdfTextMatch(pageNumber, unit.Text.Substring(found, text.Length), matchBounds.X, matchBounds.Y, matchBounds.Width, matchBounds.Height, detected.FontSize, detected.SuggestedFont, detected.SourceFont, detected.Color, detected.RotationDegrees);
                    hits.Add(new TextSearchHit(pageNumber, unit, found, match));
                }
            }
        }

        return hits;
    }

    private static PdfRegionText BuildRegionText(IReadOnlyList<PdfTextSpan> spans) {
        if (spans.Count == 0) return new PdfRegionText(string.Empty, Array.Empty<PdfTextSpan>(), PdfStandardFont.Helvetica, null, 12D, PdfColor.Black, 0D, 0D, 0D);
        PdfTextSpan dominant = spans
            .GroupBy(static span => new SpanStyleKey(span.BaseFont ?? span.FontResource, Math.Round(EffectiveFontSize(span), 2), Math.Round(span.RotationDegrees, 2)))
            .OrderByDescending(static group => group.Sum(static span => Math.Max(1, span.Text.Length)))
            .First()
            .First();
        PdfTextSpan[] ordered = spans.OrderByDescending(static span => span.Y).ThenBy(static span => span.X).ToArray();
        var builder = new System.Text.StringBuilder();
        PdfTextSpan? previous = null;
        for (int index = 0; index < ordered.Length; index++) {
            PdfTextSpan span = ordered[index];
            if (previous != null) {
                bool newLine = Math.Abs(span.Y - previous.Y) > Math.Max(EffectiveFontSize(previous), EffectiveFontSize(span)) * 0.65D;
                if (newLine) builder.Append('\n');
                else if (span.X - (previous.X + Math.Max(0D, previous.Advance)) > Math.Max(1D, EffectiveFontSize(span) * 0.18D)) builder.Append(' ');
            }
            builder.Append(span.Text);
            previous = span;
        }
        return new PdfRegionText(builder.ToString(), ordered, ResolveStandardFont(dominant.BaseFont), dominant.BaseFont, EffectiveFontSize(dominant), ToPdfColor(dominant.Color), dominant.RotationDegrees, ordered[0].X, ordered[0].Y);
    }

    private static byte[] StampLines(byte[] pdf, int pageNumber, double x, double baselineY, string text, PdfResolvedTextStyle style, PdfReadOptions? readOptions) {
        string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        byte[] current = pdf;
        double radians = style.RotationDegrees * Math.PI / 180D;
        double normalX = Math.Sin(radians) * style.FontSize * 1.15D;
        double normalY = -Math.Cos(radians) * style.FontSize * 1.15D;
        for (int index = 0; index < lines.Length; index++) {
            if (lines[index].Length == 0) continue;
            current = PdfStamper.StampText(current, lines[index], new PdfTextStampOptions {
                PageNumbers = new[] { pageNumber },
                X = x + normalX * index,
                Y = baselineY + normalY * index,
                Font = style.Font,
                FontSize = style.FontSize,
                Color = style.Color,
                RotationDegrees = style.RotationDegrees
            }, readOptions);
        }
        return current;
    }

    private static PdfResolvedTextStyle ResolveStyle(PdfTextEditOptions options, PdfRegionText? detected) => new PdfResolvedTextStyle(
        options.Font ?? detected?.SuggestedFont ?? PdfStandardFont.Helvetica,
        options.FontSize ?? detected?.FontSize ?? 12D,
        options.Color ?? detected?.Color ?? PdfColor.Black,
        options.RotationDegrees ?? detected?.RotationDegrees ?? 0D);

    private static string[] BuildSubstitutionWarnings(PdfRegionText detected, PdfStandardFont targetFont) {
        string source = StripSubsetPrefix(detected.SourceFont);
        if (source.Length == 0 || string.Equals(source, targetFont.ToBaseFontName(), StringComparison.OrdinalIgnoreCase)) return Array.Empty<string>();
        return new[] { "The source font '" + source + "' is not reused by the dependency-free text editor; replacement text uses '" + targetFont.ToBaseFontName() + "'. Metrics and letterforms can differ." };
    }

    private static PdfStandardFont ResolveStandardFont(string? baseFont) {
        string name = StripSubsetPrefix(baseFont).ToLowerInvariant();
        bool bold = name.Contains("bold") || name.Contains("black") || name.Contains("heavy") || name.Contains("semibold");
        bool italic = name.Contains("italic") || name.Contains("oblique");
        if (name.Contains("times") || name.Contains("serif") || name.Contains("roman") || name.Contains("georgia")) {
            return bold && italic ? PdfStandardFont.TimesBoldItalic : bold ? PdfStandardFont.TimesBold : italic ? PdfStandardFont.TimesItalic : PdfStandardFont.TimesRoman;
        }
        if (name.Contains("courier") || name.Contains("mono") || name.Contains("consol")) {
            return bold && italic ? PdfStandardFont.CourierBoldOblique : bold ? PdfStandardFont.CourierBold : italic ? PdfStandardFont.CourierOblique : PdfStandardFont.Courier;
        }
        return bold && italic ? PdfStandardFont.HelveticaBoldOblique : bold ? PdfStandardFont.HelveticaBold : italic ? PdfStandardFont.HelveticaOblique : PdfStandardFont.Helvetica;
    }

    private static string StripSubsetPrefix(string? value) {
        string name = (value ?? string.Empty).Trim().TrimStart('/');
        return name.Length > 7 && name[6] == '+' ? name.Substring(7) : name;
    }

    private static PdfColor ToPdfColor(OfficeColor? color) => color.HasValue ? PdfColor.FromOfficeColor(color.Value) : PdfColor.Black;
    private static double EffectiveFontSize(PdfTextSpan span) => span.FontSize > 0D && !double.IsNaN(span.FontSize) && !double.IsInfinity(span.FontSize) ? span.FontSize : 12D;

    private static SpanBounds GetBounds(PdfTextSpan span) => SliceBounds(span, 0, Math.Max(1, span.Text.Length));

    private static SpanBounds GetCombinedBounds(IReadOnlyList<PdfTextSpan> spans) {
        SpanBounds first = GetBounds(spans[0]);
        double left = first.X;
        double bottom = first.Y;
        double right = first.X + first.Width;
        double top = first.Y + first.Height;
        for (int index = 1; index < spans.Count; index++) {
            SpanBounds current = GetBounds(spans[index]);
            left = Math.Min(left, current.X);
            bottom = Math.Min(bottom, current.Y);
            right = Math.Max(right, current.X + current.Width);
            top = Math.Max(top, current.Y + current.Height);
        }
        return new SpanBounds(left, bottom, Math.Max(0.1D, right - left), Math.Max(0.1D, top - bottom));
    }

    private static SpanBounds SliceUnitBounds(TextSearchUnit unit, int start, int length) {
        double textLength = Math.Max(1, unit.Text.Length);
        double x = unit.Bounds.X + unit.Bounds.Width * start / textLength;
        double width = Math.Max(0.1D, unit.Bounds.Width * length / textLength);
        return new SpanBounds(x, unit.Bounds.Y, width, unit.Bounds.Height);
    }

    private static SpanBounds SliceBounds(PdfTextSpan span, int start, int length) {
        double textLength = Math.Max(1, span.Text.Length);
        double advance = Math.Max(Math.Abs(span.Advance), EffectiveFontSize(span) * Math.Max(1, span.Text.Length) * 0.45D);
        double offset = advance * start / textLength;
        double sliceAdvance = advance * length / textLength;
        double radians = span.RotationDegrees * Math.PI / 180D;
        double ux = Math.Cos(radians);
        double uy = Math.Sin(radians);
        double nx = -uy;
        double ny = ux;
        double fontSize = EffectiveFontSize(span);
        double x0 = span.X + ux * offset - nx * fontSize * 0.25D;
        double y0 = span.Y + uy * offset - ny * fontSize * 0.25D;
        double x1 = span.X + ux * (offset + sliceAdvance) - nx * fontSize * 0.25D;
        double y1 = span.Y + uy * (offset + sliceAdvance) - ny * fontSize * 0.25D;
        double x2 = span.X + ux * offset + nx * fontSize * 0.8D;
        double y2 = span.Y + uy * offset + ny * fontSize * 0.8D;
        double x3 = span.X + ux * (offset + sliceAdvance) + nx * fontSize * 0.8D;
        double y3 = span.Y + uy * (offset + sliceAdvance) + ny * fontSize * 0.8D;
        double left = Math.Min(Math.Min(x0, x1), Math.Min(x2, x3));
        double right = Math.Max(Math.Max(x0, x1), Math.Max(x2, x3));
        double bottom = Math.Min(Math.Min(y0, y1), Math.Min(y2, y3));
        double top = Math.Max(Math.Max(y0, y1), Math.Max(y2, y3));
        return new SpanBounds(left, bottom, Math.Max(0.1D, right - left), Math.Max(0.1D, top - bottom));
    }

    private static bool ContainsCenter(PdfPageRegion region, SpanBounds bounds) {
        double centerX = bounds.X + bounds.Width / 2D;
        double centerY = bounds.Y + bounds.Height / 2D;
        return centerX >= region.X && centerX <= region.Right && centerY >= region.Y && centerY <= region.Top;
    }

    private static bool ContainsCenter(PdfRedactionArea region, SpanBounds bounds) {
        double centerX = bounds.X + bounds.Width / 2D;
        double centerY = bounds.Y + bounds.Height / 2D;
        return centerX >= region.X && centerX <= region.Right && centerY >= region.Y && centerY <= region.Top;
    }

    private static bool SameTextPlacement(PdfTextSpan left, PdfTextSpan right) =>
        string.Equals(left.Text, right.Text, StringComparison.Ordinal) &&
        Math.Abs(left.X - right.X) <= 0.2D &&
        Math.Abs(left.Y - right.Y) <= 0.2D &&
        Math.Abs(EffectiveFontSize(left) - EffectiveFontSize(right)) <= 0.2D &&
        Math.Abs(left.RotationDegrees - right.RotationDegrees) <= 0.2D;

    private static bool HasWordBoundaries(string text, int start, int length) {
        bool left = start == 0 || !IsWordCharacter(text[start - 1]);
        int end = start + length;
        bool right = end == text.Length || !IsWordCharacter(text[end]);
        return left && right;
    }

    private static bool IsWordCharacter(char value) => char.IsLetterOrDigit(value) || value == '_';

    private static string ReplaceOccurrences(string source, int[] starts, int sourceLength, string replacement) {
        var builder = new System.Text.StringBuilder(source.Length + Math.Max(0, replacement.Length - sourceLength) * starts.Length);
        int cursor = 0;
        for (int index = 0; index < starts.Length; index++) {
            int start = starts[index];
            if (start < cursor) continue;
            builder.Append(source, cursor, start - cursor);
            builder.Append(replacement);
            cursor = start + sourceLength;
        }
        builder.Append(source, cursor, source.Length - cursor);
        return builder.ToString();
    }

    private static void ValidateRegionPage(byte[] pdf, PdfPageRegion region, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(region, nameof(region));
        ValidatePage(region.PageNumber, PdfReadDocument.Open(pdf, readOptions).Pages.Count, nameof(region));
    }

    private static void ValidatePage(int pageNumber, int pageCount, string paramName) {
        if (pageNumber > pageCount) throw new ArgumentOutOfRangeException(paramName, "Page number exceeds the PDF page count.");
    }

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(paramName, "Offset must be finite.");
    }

    internal sealed class TextMutationResult {
        internal TextMutationResult(byte[] bytes, int affectedCount, IEnumerable<string> warnings) {
            Bytes = bytes;
            AffectedCount = affectedCount;
            Warnings = warnings.Distinct(StringComparer.Ordinal).ToArray();
        }
        internal byte[] Bytes { get; }
        internal int AffectedCount { get; }
        internal IReadOnlyList<string> Warnings { get; }
    }

    private readonly struct PdfResolvedTextStyle {
        internal PdfResolvedTextStyle(PdfStandardFont font, double fontSize, PdfColor color, double rotationDegrees) { Font = font; FontSize = fontSize; Color = color; RotationDegrees = rotationDegrees; }
        internal PdfStandardFont Font { get; }
        internal double FontSize { get; }
        internal PdfColor Color { get; }
        internal double RotationDegrees { get; }
    }

    private readonly struct SpanBounds {
        internal SpanBounds(double x, double y, double width, double height) { X = x; Y = y; Width = width; Height = height; }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
    }

    private readonly struct SpanStyleKey : IEquatable<SpanStyleKey> {
        internal SpanStyleKey(string font, double size, double rotation) { Font = font; Size = size; Rotation = rotation; }
        private string Font { get; }
        private double Size { get; }
        private double Rotation { get; }
        public bool Equals(SpanStyleKey other) => string.Equals(Font, other.Font, StringComparison.Ordinal) && Size.Equals(other.Size) && Rotation.Equals(other.Rotation);
        public override bool Equals(object? obj) => obj is SpanStyleKey other && Equals(other);
        public override int GetHashCode() { unchecked { int hash = StringComparer.Ordinal.GetHashCode(Font); hash = (hash * 397) ^ Size.GetHashCode(); return (hash * 397) ^ Rotation.GetHashCode(); } }
    }

    private readonly struct PageTextSpanSnapshot {
        internal PageTextSpanSnapshot(int pageNumber, PdfTextSpan span, bool targeted) { PageNumber = pageNumber; Span = span; Targeted = targeted; }
        internal int PageNumber { get; }
        internal PdfTextSpan Span { get; }
        internal bool Targeted { get; }
    }

    private readonly struct TextRemovalResult {
        internal TextRemovalResult(byte[] bytes, IEnumerable<string> warnings) { Bytes = bytes; Warnings = warnings.Distinct(StringComparer.Ordinal).ToArray(); }
        internal byte[] Bytes { get; }
        internal IReadOnlyList<string> Warnings { get; }
    }

    private sealed class TextSearchUnit {
        internal TextSearchUnit(string text, IReadOnlyList<PdfTextSpan> spans, SpanBounds bounds) {
            Text = text;
            Spans = spans.ToArray();
            Bounds = bounds;
            BaselineX = Spans.Min(static span => span.X);
            BaselineY = Spans.OrderBy(static span => span.X).First().Y;
        }
        internal string Text { get; }
        internal IReadOnlyList<PdfTextSpan> Spans { get; }
        internal SpanBounds Bounds { get; }
        internal double BaselineX { get; }
        internal double BaselineY { get; }
    }

    private sealed class TextSearchHit {
        internal TextSearchHit(int pageNumber, TextSearchUnit unit, int startIndex, PdfTextMatch match) { PageNumber = pageNumber; Unit = unit; StartIndex = startIndex; Match = match; }
        internal int PageNumber { get; }
        internal TextSearchUnit Unit { get; }
        internal int StartIndex { get; }
        internal PdfTextMatch Match { get; }
    }
}
