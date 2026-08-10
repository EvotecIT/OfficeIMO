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
            .Where(span => IsSafelyEditableSpan(span) && ContainsCenter(region, GetBounds(span)))
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
        EnsureRegionIsSafelyEditable(pdf, region, readOptions);
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
        EnsureRegionIsSafelyEditable(pdf, source, readOptions);
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

        var rewrites = new Dictionary<PageSpanKey, List<SpanTextEdit>>();
        for (int hitIndex = 0; hitIndex < hits.Count; hitIndex++) {
            TextSearchHit hit = hits[hitIndex];
            for (int segmentIndex = 0; segmentIndex < hit.Segments.Length; segmentIndex++) {
                TextSourceSegment segment = hit.Segments[segmentIndex];
                var key = new PageSpanKey(hit.PageNumber, segment.Span);
                if (!rewrites.TryGetValue(key, out List<SpanTextEdit>? edits)) {
                    edits = new List<SpanTextEdit>();
                    rewrites.Add(key, edits);
                }
                edits.Add(new SpanTextEdit(segment.Start, segment.Length, segmentIndex == 0 ? replacement : string.Empty));
            }
        }

        PdfRedactionArea[] areas = rewrites.Keys
            .Select(static key => {
                SpanBounds bounds = GetBounds(key.Span);
                return new PdfRedactionArea(key.PageNumber, bounds.X, bounds.Y, bounds.Width, bounds.Height);
            })
            .ToArray();
        TextRemovalResult removal = RemoveTextPreservingUnmatchedSpans(pdf, areas, readOptions, rewrites.Keys.ToArray());
        byte[] current = removal.Bytes;
        PdfTextEditOptions snapshot = (editOptions ?? new PdfTextEditOptions()).Snapshot();
        var warnings = new List<string>(removal.Warnings);
        foreach (KeyValuePair<PageSpanKey, List<SpanTextEdit>> rewrite in rewrites
            .OrderBy(static item => item.Key.PageNumber)
            .ThenByDescending(static item => item.Key.Span.Y)
            .ThenBy(static item => item.Key.Span.X)) {
            PdfTextSpan sourceSpan = rewrite.Key.Span;
            string rewritten = ApplySpanEdits(sourceSpan.Text, rewrite.Value);
            PdfRegionText detected = BuildRegionText(new[] { sourceSpan });
            PdfResolvedTextStyle style = ResolveStyle(snapshot, detected);
            warnings.AddRange(BuildSubstitutionWarnings(detected, style.Font));
            if (rewritten.Length > 0) current = StampLines(current, rewrite.Key.PageNumber, sourceSpan.X, sourceSpan.Y, rewritten, style, readOptions);
        }

        return new TextMutationResult(current, hits.Count, warnings);
    }

    private static TextRemovalResult RemoveTextPreservingUnmatchedSpans(byte[] pdf, IReadOnlyList<PdfRedactionArea> areas, PdfReadOptions? readOptions, IReadOnlyList<PageSpanKey>? exactTargets = null) {
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
                bool targeted = exactTargets is null
                    ? areas.Any(area => area.PageNumber == pageNumber && ContainsCenter(area, bounds))
                    : exactTargets.Any(target => target.PageNumber == pageNumber && SameExactSourceSpan(span, target.Span));
                if (targeted && !IsSafelyEditableSpan(span)) {
                    throw new NotSupportedException("The selected region contains invisible or clipped text whose rendering state cannot be recreated safely.");
                }
                original.Add(new PageTextSpanSnapshot(pageNumber, span, targeted));
            }
        }

        byte[] removed = PdfRedactionApplier.RemoveTextInAreas(pdf, areas, readOptions: readOptions);
        PdfReadOptions afterReadOptions = PdfReadOptions.WithMinimumInputBytes(readOptions, removed.LongLength);
        PdfReadDocument after = PdfReadDocument.Open(removed, afterReadOptions);
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
            if (!IsSafelyEditableSpan(snapshot.Span)) {
                throw new NotSupportedException("The text edit would require recreating invisible or clipped source text without its original rendering state.");
            }
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
            IReadOnlyList<PdfTextSpan> spans = document.Pages[pageNumber - 1]
                .GetTextSpans()
                .Where(IsSafelyEditableSpan)
                .ToArray();
            List<TextLayoutEngine.TextLine> lines = TextLayoutEngine.BuildLines(spans, new TextLayoutEngine.Options { SplitWideSameBaselineRuns = true });
            for (int lineIndex = 0; lineIndex < lines.Count; lineIndex++) {
                TextLayoutEngine.TextLine line = lines[lineIndex];
                if (line.Spans.Count == 0) continue;
                var unit = new TextSearchUnit(line.Spans);
                if (unit.Text.Length == 0) continue;
                int start = 0;
                while (start <= unit.Text.Length - text.Length) {
                    int found = unit.Text.IndexOf(text, start, comparison);
                    if (found < 0) break;
                    start = found + Math.Max(1, text.Length);
                    if (snapshot.WholeWords && !HasWordBoundaries(unit.Text, found, text.Length)) continue;
                    IReadOnlyList<TextSourceSegment> segments = unit.GetSourceSegments(found, text.Length);
                    if (segments.Count == 0) continue;
                    PdfRegionText detected = BuildRegionText(new[] { segments[0].Span });
                    SpanBounds matchBounds = GetCombinedSegmentBounds(segments);
                    var match = new PdfTextMatch(pageNumber, unit.Text.Substring(found, text.Length), matchBounds.X, matchBounds.Y, matchBounds.Width, matchBounds.Height, detected.FontSize, detected.SuggestedFont, detected.SourceFont, detected.Color, detected.RotationDegrees);
                    hits.Add(new TextSearchHit(pageNumber, segments, match));
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
            PdfReadOptions effectiveReadOptions = PdfReadOptions.WithMinimumInputBytes(readOptions, current.LongLength);
            current = PdfStamper.StampText(current, lines[index], new PdfTextStampOptions {
                PageNumbers = new[] { pageNumber },
                X = x + normalX * index,
                Y = baselineY + normalY * index,
                Font = style.Font,
                FontSize = style.FontSize,
                Color = style.Color,
                RotationDegrees = style.RotationDegrees
            }, effectiveReadOptions);
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

    private static SpanBounds GetCombinedSegmentBounds(IReadOnlyList<TextSourceSegment> segments) {
        SpanBounds combined = SliceBounds(segments[0].Span, segments[0].Start, segments[0].Length);
        double left = combined.X;
        double bottom = combined.Y;
        double right = combined.X + combined.Width;
        double top = combined.Y + combined.Height;
        for (int index = 1; index < segments.Count; index++) {
            TextSourceSegment segment = segments[index];
            SpanBounds bounds = SliceBounds(segment.Span, segment.Start, segment.Length);
            left = Math.Min(left, bounds.X);
            bottom = Math.Min(bottom, bounds.Y);
            right = Math.Max(right, bounds.X + bounds.Width);
            top = Math.Max(top, bounds.Y + bounds.Height);
        }
        return new SpanBounds(left, bottom, Math.Max(0.1D, right - left), Math.Max(0.1D, top - bottom));
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

    private static bool SameExactSourceSpan(PdfTextSpan left, PdfTextSpan right) =>
        SameTextPlacement(left, right) &&
        Math.Abs(left.Advance - right.Advance) <= 0.2D &&
        Math.Abs(left.PaintOrder - right.PaintOrder) <= 0.0001D &&
        string.Equals(left.FontResource, right.FontResource, StringComparison.Ordinal) &&
        string.Equals(left.BaseFont, right.BaseFont, StringComparison.Ordinal) &&
        Nullable.Equals(left.Color, right.Color) &&
        left.IsVisible == right.IsVisible;

    private static bool HasWordBoundaries(string text, int start, int length) {
        bool left = start == 0 || !IsWordCharacter(text[start - 1]);
        int end = start + length;
        bool right = end == text.Length || !IsWordCharacter(text[end]);
        return left && right;
    }

    private static bool IsWordCharacter(char value) => char.IsLetterOrDigit(value) || value == '_';

    private static string ApplySpanEdits(string source, IReadOnlyList<SpanTextEdit> edits) {
        SpanTextEdit[] ordered = edits.OrderBy(static edit => edit.Start).ToArray();
        var builder = new System.Text.StringBuilder(source.Length);
        int cursor = 0;
        for (int index = 0; index < ordered.Length; index++) {
            SpanTextEdit edit = ordered[index];
            if (edit.Start < cursor) continue;
            builder.Append(source, cursor, edit.Start - cursor);
            builder.Append(edit.Replacement);
            cursor = edit.Start + edit.Length;
        }
        builder.Append(source, cursor, source.Length - cursor);
        return builder.ToString();
    }

    private static bool IsSafelyEditableSpan(PdfTextSpan span) => span.IsVisible && !span.ClipPath.HasValue && !string.IsNullOrEmpty(span.Text);

    private static void ValidateRegionPage(byte[] pdf, PdfPageRegion region, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(region, nameof(region));
        ValidatePage(region.PageNumber, PdfReadDocument.Open(pdf, readOptions).Pages.Count, nameof(region));
    }

    private static void EnsureRegionIsSafelyEditable(byte[] pdf, PdfPageRegion region, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(region, nameof(region));
        PdfReadDocument document = PdfReadDocument.Open(pdf, readOptions);
        ValidatePage(region.PageNumber, document.Pages.Count, nameof(region));
        bool containsUnsafeText = document.Pages[region.PageNumber - 1]
            .GetTextSpans()
            .Any(span => !IsSafelyEditableSpan(span) && ContainsCenter(region, GetBounds(span)));
        if (containsUnsafeText) {
            throw new NotSupportedException("The selected region contains invisible or clipped text whose rendering state cannot be recreated safely.");
        }
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
        private readonly TextCharacterSource?[] _sources;

        internal TextSearchUnit(IReadOnlyList<PdfTextSpan> spans) {
            PdfTextSpan[] orderedSpans = spans.OrderBy(static span => span.X).ThenByDescending(static span => span.Y).ToArray();
            var text = new System.Text.StringBuilder();
            var sources = new List<TextCharacterSource?>();
            PdfTextSpan? previous = null;
            for (int spanIndex = 0; spanIndex < orderedSpans.Length; spanIndex++) {
                PdfTextSpan span = orderedSpans[spanIndex];
                if (previous != null && NeedsSyntheticSpace(previous, span, text)) {
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

        internal List<TextSourceSegment> GetSourceSegments(int start, int length) {
            var segments = new List<TextSourceSegment>();
            TextSourceSegment? current = null;
            int end = Math.Min(_sources.Length, start + length);
            for (int index = start; index < end; index++) {
                TextCharacterSource? source = _sources[index];
                if (!source.HasValue) continue;
                if (current.HasValue &&
                    ReferenceEquals(current.Value.Span, source.Value.Span) &&
                    current.Value.Start + current.Value.Length == source.Value.CharacterIndex) {
                    current = new TextSourceSegment(current.Value.Span, current.Value.Start, current.Value.Length + 1);
                    segments[segments.Count - 1] = current.Value;
                } else {
                    current = new TextSourceSegment(source.Value.Span, source.Value.CharacterIndex, 1);
                    segments.Add(current.Value);
                }
            }
            return segments;
        }

        private static bool NeedsSyntheticSpace(PdfTextSpan previous, PdfTextSpan current, System.Text.StringBuilder text) {
            if (text.Length == 0 || char.IsWhiteSpace(text[text.Length - 1]) || (current.Text.Length > 0 && char.IsWhiteSpace(current.Text[0]))) return false;
            if (previous.LogicalTrailingSpace || current.LogicalLeadingSpace) return true;
            double gap = current.X - (previous.X + Math.Max(0D, previous.Advance));
            return gap > Math.Max(1D, Math.Min(EffectiveFontSize(previous), EffectiveFontSize(current)) * 0.18D);
        }
    }

    private sealed class TextSearchHit {
        internal TextSearchHit(int pageNumber, IReadOnlyList<TextSourceSegment> segments, PdfTextMatch match) { PageNumber = pageNumber; Segments = segments.ToArray(); Match = match; }
        internal int PageNumber { get; }
        internal TextSourceSegment[] Segments { get; }
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

    private readonly struct SpanTextEdit {
        internal SpanTextEdit(int start, int length, string replacement) { Start = start; Length = length; Replacement = replacement; }
        internal int Start { get; }
        internal int Length { get; }
        internal string Replacement { get; }
    }

    private readonly struct PageSpanKey : IEquatable<PageSpanKey> {
        internal PageSpanKey(int pageNumber, PdfTextSpan span) { PageNumber = pageNumber; Span = span; }
        internal int PageNumber { get; }
        internal PdfTextSpan Span { get; }
        public bool Equals(PageSpanKey other) => PageNumber == other.PageNumber && ReferenceEquals(Span, other.Span);
        public override bool Equals(object? obj) => obj is PageSpanKey other && Equals(other);
        public override int GetHashCode() { unchecked { return (PageNumber * 397) ^ System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(Span); } }
    }
}
