namespace OfficeIMO.Pdf;

internal sealed class PdfRedactionTextObjectScope {
    internal PdfRedactionTextObjectScope(
        PdfContentOrderKey key,
        PdfTextSpan[] spans,
        IReadOnlyList<PdfRedactionArea>? reviewedAreas = null,
        Func<double, PdfRedactionPaintOrderContext>? paintOrderContext = null) {
        Key = key;
        Spans = spans;
        ContentStreamObjectNumber = spans.Select(static span => span.ContentStreamObjectNumber).Distinct().Count() == 1
            ? spans[0].ContentStreamObjectNumber
            : null;

        var sourceGlyphs = new List<PdfRedactionTextGlyphIdentity>();
        var expectedSurvivors = new List<PdfRedactionTextGlyphIdentity>();
        bool reviewedIntersection = false;
        bool requiresWholeObjectFallback = false;
        for (int spanIndex = 0; spanIndex < spans.Length; spanIndex++) {
            PdfTextSpan span = spans[spanIndex];
            PdfRedactionPaintOrderContext spanPaintOrderContext = paintOrderContext?.Invoke(span.PaintOrder) ?? default;
            if (!TryCreateGlyphIdentities(span, spanPaintOrderContext, out PdfRedactionTextGlyphIdentity[] glyphs)) {
                PdfRedactionTextGlyphIdentity aggregate = PdfRedactionTextGlyphIdentity.FromSpan(
                    span,
                    spanPaintOrderContext);
                sourceGlyphs.Add(aggregate);
                if (reviewedAreas != null && IntersectsAny(reviewedAreas, aggregate.Bounds)) {
                    reviewedIntersection = true;
                    requiresWholeObjectFallback = true;
                } else {
                    expectedSurvivors.Add(aggregate);
                }
                continue;
            }

            sourceGlyphs.AddRange(glyphs);
            for (int glyphIndex = 0; glyphIndex < glyphs.Length; glyphIndex++) {
                PdfRedactionTextGlyphIdentity glyph = glyphs[glyphIndex];
                if (reviewedAreas != null && IntersectsAny(reviewedAreas, glyph.Bounds)) {
                    reviewedIntersection = true;
                } else {
                    expectedSurvivors.Add(glyph);
                }
            }
        }

        SourceGlyphs = sourceGlyphs.ToArray();
        ExpectedSurvivors = requiresWholeObjectFallback
            ? Array.Empty<PdfRedactionTextGlyphIdentity>()
            : expectedSurvivors.ToArray();
        HasReviewedIntersection = reviewedIntersection;
    }

    internal PdfContentOrderKey Key { get; }
    internal PdfTextSpan[] Spans { get; }
    internal int? ContentStreamObjectNumber { get; }
    internal PdfRedactionTextGlyphIdentity[] SourceGlyphs { get; }
    internal PdfRedactionTextGlyphIdentity[] ExpectedSurvivors { get; }
    internal bool HasReviewedIntersection { get; }
    internal bool RequiresExpectedSurvivors => ExpectedSurvivors.Length > 0;

    internal bool Matches(PdfRedactionTextObjectScope candidate) {
        return (HasSameOwner(candidate) && SequenceMatches(SourceGlyphs, candidate.SourceGlyphs)) ||
            SequenceMatches(ExpectedSurvivors, candidate.SourceGlyphs);
    }

    internal bool HasSameOwner(PdfRedactionTextObjectScope candidate) =>
        Key.Equals(candidate.Key) ||
        ContentStreamObjectNumber.HasValue && ContentStreamObjectNumber == candidate.ContentStreamObjectNumber;

    private static bool TryCreateGlyphIdentities(
        PdfTextSpan span,
        PdfRedactionPaintOrderContext paintOrderContext,
        out PdfRedactionTextGlyphIdentity[] glyphs) {
        string text = span.RestampText;
        if (!PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries) ||
            boundaries.Length != text.Length + 1) {
            glyphs = Array.Empty<PdfRedactionTextGlyphIdentity>();
            return false;
        }

        IReadOnlyList<int> glyphCharacterLengths = span.GlyphCharacterLengths ??
            Enumerable.Repeat(1, text.Length).ToArray();
        if (glyphCharacterLengths.Count == 0 ||
            glyphCharacterLengths.Any(static length => length <= 0) ||
            glyphCharacterLengths.Sum() != text.Length) {
            glyphs = Array.Empty<PdfRedactionTextGlyphIdentity>();
            return false;
        }

        glyphs = new PdfRedactionTextGlyphIdentity[glyphCharacterLengths.Count];
        bool hasPaintedGlyphAdvances = span.GlyphPaintedAdvances != null &&
            span.GlyphPaintedAdvances.Count == glyphCharacterLengths.Count &&
            span.GlyphPaintedAdvances.All(static advance =>
                !double.IsNaN(advance) && !double.IsInfinity(advance) && advance >= 0D);
        int characterOffset = 0;
        for (int index = 0; index < glyphCharacterLengths.Count; index++) {
            int characterLength = glyphCharacterLengths[index];
            double start = boundaries[characterOffset];
            double end = boundaries[characterOffset + characterLength];
            double glyphOffset = hasPaintedGlyphAdvances ? start : Math.Min(start, end);
            double glyphAdvance = hasPaintedGlyphAdvances
                ? span.GlyphPaintedAdvances![index]
                : Math.Abs(end - start);
            PdfTextSpanBounds bounds = PdfTextSpanGeometry.GetAxisAlignedBounds(
                span,
                glyphOffset,
                glyphAdvance);
            glyphs[index] = PdfRedactionTextGlyphIdentity.FromGlyph(
                span,
                text.Substring(characterOffset, characterLength),
                bounds,
                span.GlyphBytes != null && span.GlyphBytes.Count == glyphCharacterLengths.Count
                    ? span.GlyphBytes[index]
                    : null,
                TranslateTextTransform(span, glyphOffset),
                paintOrderContext);
            characterOffset += characterLength;
        }
        return true;
    }

    private static bool IntersectsAny(IReadOnlyList<PdfRedactionArea> areas, PdfTextSpanBounds bounds) {
        for (int index = 0; index < areas.Count; index++) {
            PdfRedactionArea area = areas[index];
            if (area.X < bounds.Right && area.Right > bounds.Left &&
                area.Y < bounds.Top && area.Top > bounds.Bottom) return true;
        }
        return false;
    }

    private static Matrix2D? TranslateTextTransform(PdfTextSpan span, double advanceOffset) {
        if (!span.TextToPageTransform.HasValue || Math.Abs(advanceOffset) <= double.Epsilon) {
            return span.TextToPageTransform;
        }
        Matrix2D transform = span.TextToPageTransform.Value;
        double radians = span.RotationDegrees * Math.PI / 180D;
        return new Matrix2D(
            transform.A,
            transform.B,
            transform.C,
            transform.D,
            transform.E + Math.Cos(radians) * advanceOffset,
            transform.F + Math.Sin(radians) * advanceOffset);
    }

    private static bool SequenceMatches(
        PdfRedactionTextGlyphIdentity[] expected,
        PdfRedactionTextGlyphIdentity[] actual) {
        if (expected.Length == 0 || actual.Length == 0) return false;
        int expectedIndex = 0;
        for (int actualIndex = 0; actualIndex < actual.Length; actualIndex++) {
            PdfRedactionTextGlyphIdentity candidate = actual[actualIndex];
            int start = expectedIndex;
            int textLength = 0;
            while (expectedIndex < expected.Length && textLength < candidate.Text.Length) {
                textLength += expected[expectedIndex].Text.Length;
                expectedIndex++;
            }
            if (textLength != candidate.Text.Length || start == expectedIndex) return false;

            string expectedText = string.Concat(expected.Skip(start).Take(expectedIndex - start).Select(static value => value.Text));
            if (!string.Equals(expectedText, candidate.Text, StringComparison.Ordinal)) return false;
            if (!candidate.MatchesEncodedBytes(expected, start, expectedIndex)) return false;
            PdfTextSpanBounds expectedBounds = MergeBounds(expected, start, expectedIndex);
            if (!candidate.MatchesBounds(expectedBounds)) return false;
            for (int index = start; index < expectedIndex; index++) {
                if (!candidate.MatchesState(expected[index])) return false;
            }
            if (!candidate.MatchesTransform(expected[start])) return false;
        }
        return expectedIndex == expected.Length;
    }

    private static PdfTextSpanBounds MergeBounds(PdfRedactionTextGlyphIdentity[] values, int start, int end) {
        double left = values[start].Bounds.Left;
        double bottom = values[start].Bounds.Bottom;
        double right = values[start].Bounds.Right;
        double top = values[start].Bounds.Top;
        for (int index = start + 1; index < end; index++) {
            left = Math.Min(left, values[index].Bounds.Left);
            bottom = Math.Min(bottom, values[index].Bounds.Bottom);
            right = Math.Max(right, values[index].Bounds.Right);
            top = Math.Max(top, values[index].Bounds.Top);
        }
        return new PdfTextSpanBounds(left, bottom, right, top);
    }
}

internal readonly struct PdfRedactionPaintOrderContext : IEquatable<PdfRedactionPaintOrderContext> {
    internal PdfRedactionPaintOrderContext(int pathPaintsBefore, int retainedImagePaintsBefore) {
        PathPaintsBefore = pathPaintsBefore;
        RetainedImagePaintsBefore = retainedImagePaintsBefore;
    }

    internal int PathPaintsBefore { get; }
    internal int RetainedImagePaintsBefore { get; }

    public bool Equals(PdfRedactionPaintOrderContext other) =>
        PathPaintsBefore == other.PathPaintsBefore &&
        RetainedImagePaintsBefore == other.RetainedImagePaintsBefore;

    public override bool Equals(object? obj) => obj is PdfRedactionPaintOrderContext other && Equals(other);
    public override int GetHashCode() => unchecked((PathPaintsBefore * 397) ^ RetainedImagePaintsBefore);
}

internal readonly struct PdfRedactionTextGlyphIdentity {
    private const double Tolerance = 0.01D;

    private PdfRedactionTextGlyphIdentity(
        PdfTextSpan span,
        string text,
        PdfTextSpanBounds bounds,
        byte[]? encodedBytes,
        Matrix2D? textToPageTransform,
        PdfRedactionPaintOrderContext paintOrderContext) {
        Text = text;
        Bounds = bounds;
        FontResource = span.FontResource;
        BaseFont = span.BaseFont;
        FontSize = span.FontSize;
        RotationDegrees = span.RotationDegrees;
        IsVisible = span.IsVisible;
        TextRenderingMode = span.TextRenderingMode;
        Color = span.Color;
        VisualPaintIdentity = span.VisualPaintIdentity;
        MarkedContentId = span.MarkedContentId;
        ClipPath = span.ClipPath;
        TextToPageTransform = textToPageTransform;
        EncodedBytes = encodedBytes?.ToArray();
        PaintOrderContext = paintOrderContext;
    }

    internal string Text { get; }
    internal PdfTextSpanBounds Bounds { get; }
    internal string FontResource { get; }
    internal string? BaseFont { get; }
    internal double FontSize { get; }
    internal double RotationDegrees { get; }
    internal bool IsVisible { get; }
    internal int TextRenderingMode { get; }
    internal OfficeIMO.Drawing.OfficeColor? Color { get; }
    internal string? VisualPaintIdentity { get; }
    internal int? MarkedContentId { get; }
    internal PdfPageClipPath? ClipPath { get; }
    internal Matrix2D? TextToPageTransform { get; }
    internal byte[]? EncodedBytes { get; }
    internal PdfRedactionPaintOrderContext PaintOrderContext { get; }

    internal static PdfRedactionTextGlyphIdentity FromGlyph(
        PdfTextSpan span,
        string text,
        PdfTextSpanBounds bounds,
        byte[]? encodedBytes,
        Matrix2D? textToPageTransform,
        PdfRedactionPaintOrderContext paintOrderContext) =>
        new PdfRedactionTextGlyphIdentity(span, text, bounds, encodedBytes, textToPageTransform, paintOrderContext);

    internal static PdfRedactionTextGlyphIdentity FromSpan(
        PdfTextSpan span,
        PdfRedactionPaintOrderContext paintOrderContext) =>
        new PdfRedactionTextGlyphIdentity(
            span,
            span.RestampText,
            PdfTextSpanGeometry.GetAxisAlignedBounds(span),
            span.GlyphBytes?.SelectMany(static bytes => bytes).ToArray(),
            span.TextToPageTransform,
            paintOrderContext);

    internal bool MatchesBounds(PdfTextSpanBounds other) =>
        NearlyEqual(Bounds.Left, other.Left) &&
        NearlyEqual(Bounds.Bottom, other.Bottom) &&
        NearlyEqual(Bounds.Right, other.Right) &&
        NearlyEqual(Bounds.Top, other.Top);

    internal bool MatchesState(PdfRedactionTextGlyphIdentity other) =>
        string.Equals(FontResource, other.FontResource, StringComparison.Ordinal) &&
        string.Equals(BaseFont, other.BaseFont, StringComparison.Ordinal) &&
        NearlyEqual(FontSize, other.FontSize) &&
        NearlyEqual(RotationDegrees, other.RotationDegrees) &&
        IsVisible == other.IsVisible &&
        TextRenderingMode == other.TextRenderingMode &&
        Nullable.Equals(Color, other.Color) &&
        string.Equals(VisualPaintIdentity, other.VisualPaintIdentity, StringComparison.Ordinal) &&
        MarkedContentId == other.MarkedContentId &&
        ClipPathsEqual(ClipPath, other.ClipPath) &&
        PaintOrderContext.Equals(other.PaintOrderContext);

    internal bool MatchesTransform(PdfRedactionTextGlyphIdentity other) =>
        TransformsEqual(TextToPageTransform, other.TextToPageTransform);

    internal bool MatchesEncodedBytes(PdfRedactionTextGlyphIdentity[] expected, int start, int end) {
        bool hasExpectedBytes = true;
        int byteCount = 0;
        for (int index = start; index < end; index++) {
            if (expected[index].EncodedBytes == null) {
                hasExpectedBytes = false;
                break;
            }
            byteCount += expected[index].EncodedBytes!.Length;
        }
        if (!hasExpectedBytes || EncodedBytes == null) return !hasExpectedBytes && EncodedBytes == null;
        if (EncodedBytes.Length != byteCount) return false;
        int offset = 0;
        for (int index = start; index < end; index++) {
            byte[] bytes = expected[index].EncodedBytes!;
            for (int byteIndex = 0; byteIndex < bytes.Length; byteIndex++) {
                if (EncodedBytes[offset++] != bytes[byteIndex]) return false;
            }
        }
        return true;
    }

    private static bool TransformsEqual(Matrix2D? left, Matrix2D? right) {
        if (!left.HasValue || !right.HasValue) return left.HasValue == right.HasValue;
        Matrix2D l = left.Value;
        Matrix2D r = right.Value;
        return NearlyEqual(l.A, r.A) && NearlyEqual(l.B, r.B) &&
            NearlyEqual(l.C, r.C) && NearlyEqual(l.D, r.D) &&
            NearlyEqual(l.E, r.E) && NearlyEqual(l.F, r.F);
    }

    private static bool ClipPathsEqual(PdfPageClipPath? left, PdfPageClipPath? right) {
        if (!left.HasValue || !right.HasValue) return left.HasValue == right.HasValue;
        PdfPageClipPath l = left.Value;
        PdfPageClipPath r = right.Value;
        return NearlyEqual(l.X, r.X) && NearlyEqual(l.Y, r.Y) &&
            NearlyEqual(l.Width, r.Width) && NearlyEqual(l.Height, r.Height) &&
            l.IsRectangle == r.IsRectangle && l.IsExact == r.IsExact &&
            l.ContainsTextClipping == r.ContainsTextClipping && l.FillRule == r.FillRule &&
            l.Commands.SequenceEqual(r.Commands);
    }

    private static bool NearlyEqual(double left, double right) => Math.Abs(left - right) <= Tolerance;
}
