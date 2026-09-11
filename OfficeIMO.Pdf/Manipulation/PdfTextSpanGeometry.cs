using System;

namespace OfficeIMO.Pdf;

internal static class PdfTextSpanGeometry {
    internal static PdfTextSpanBounds GetAxisAlignedBounds(PdfTextSpan span) {
        if (!TryGetPaintedGlyphGeometry(
            span,
            out double[] boundaries,
            out IReadOnlyList<int> glyphCharacterLengths,
            out IReadOnlyList<double> glyphPaintedAdvances)) {
            return GetAxisAlignedBounds(span, 0D, span.Advance);
        }

        PdfTextSpanBounds result = default;
        int characterOffset = 0;
        for (int index = 0; index < glyphCharacterLengths.Count; index++) {
            PdfTextSpanBounds glyphBounds = GetAxisAlignedBounds(
                span,
                boundaries[characterOffset],
                glyphPaintedAdvances[index]);
            result = index == 0 ? glyphBounds : Union(result, glyphBounds);
            characterOffset += glyphCharacterLengths[index];
        }
        return result;
    }

    internal static PdfTextSpanBounds GetAxisAlignedBounds(PdfTextSpan span, double advanceOffset, double advance) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(span);
#else
        if (span == null) throw new ArgumentNullException(nameof(span));
#endif

        advance = Math.Max(1D, Math.Abs(advance));
        double fontSize = Math.Max(1D, Math.Abs(span.FontSize));
        double radians = span.RotationDegrees * Math.PI / 180D;
        double alongX = Math.Cos(radians);
        double alongY = Math.Sin(radians);
        double normalX = -alongY;
        double normalY = alongX;
        const double DescentFactor = 0.5D;
        double descent = fontSize * DescentFactor;
        double startX = span.X + alongX * advanceOffset;
        double startY = span.Y + alongY * advanceOffset;
        double x0 = startX - normalX * fontSize;
        double y0 = startY - normalY * fontSize;
        double x1 = startX + alongX * advance - normalX * fontSize;
        double y1 = startY + alongY * advance - normalY * fontSize;
        double x2 = startX + normalX * descent;
        double y2 = startY + normalY * descent;
        double x3 = startX + alongX * advance + normalX * descent;
        double y3 = startY + alongY * advance + normalY * descent;
        double left = Math.Min(Math.Min(x0, x1), Math.Min(x2, x3));
        double right = Math.Max(Math.Max(x0, x1), Math.Max(x2, x3));
        double bottom = Math.Min(Math.Min(y0, y1), Math.Min(y2, y3));
        double top = Math.Max(Math.Max(y0, y1), Math.Max(y2, y3));
        return new PdfTextSpanBounds(left, bottom, right, top);
    }

    internal static bool IntersectsAreaAtCharacterLevel(PdfTextSpan span, double x, double y, double width, double height) {
        if (!PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries)) {
            return Intersects(GetAxisAlignedBounds(span), x, y, width, height);
        }

        bool hasPaintedGlyphGeometry = TryGetPaintedGlyphGeometry(
            span,
            out boundaries,
            out IReadOnlyList<int> glyphCharacterLengths,
            out IReadOnlyList<double> glyphPaintedAdvances);
        int itemCount = hasPaintedGlyphGeometry ? glyphCharacterLengths!.Count : boundaries.Length - 1;
        int characterOffset = 0;
        for (int index = 0; index < itemCount; index++) {
            int characterLength = hasPaintedGlyphGeometry ? glyphCharacterLengths![index] : 1;
            double startBoundary = boundaries[characterOffset];
            double endBoundary = boundaries[characterOffset + characterLength];
            double start = hasPaintedGlyphGeometry
                ? startBoundary
                : Math.Min(startBoundary, endBoundary);
            double advance = hasPaintedGlyphGeometry
                ? glyphPaintedAdvances![index]
                : Math.Abs(endBoundary - startBoundary);
            if (Intersects(GetAxisAlignedBounds(span, start, advance), x, y, width, height)) return true;
            characterOffset += characterLength;
        }
        return false;
    }

    internal static bool IntersectsAreaAtCharacterLevel(PdfTextSpan span, PdfRedactionArea area) {
        if (!PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out double[] boundaries)) {
            PdfTextSpanBounds bounds = GetAxisAlignedBounds(span);
            return area.IntersectsRectangle(bounds.Left, bounds.Bottom, bounds.Width, bounds.Height);
        }

        bool hasPaintedGlyphGeometry = TryGetPaintedGlyphGeometry(
            span,
            out boundaries,
            out IReadOnlyList<int> glyphCharacterLengths,
            out IReadOnlyList<double> glyphPaintedAdvances);
        int itemCount = hasPaintedGlyphGeometry ? glyphCharacterLengths.Count : boundaries.Length - 1;
        int characterOffset = 0;
        for (int index = 0; index < itemCount; index++) {
            int characterLength = hasPaintedGlyphGeometry ? glyphCharacterLengths[index] : 1;
            double startBoundary = boundaries[characterOffset];
            double endBoundary = boundaries[characterOffset + characterLength];
            double start = hasPaintedGlyphGeometry ? startBoundary : Math.Min(startBoundary, endBoundary);
            double advance = hasPaintedGlyphGeometry ? glyphPaintedAdvances[index] : Math.Abs(endBoundary - startBoundary);
            PdfTextSpanBounds bounds = GetAxisAlignedBounds(span, start, advance);
            if (area.IntersectsRectangle(bounds.Left, bounds.Bottom, bounds.Width, bounds.Height)) return true;
            characterOffset += characterLength;
        }
        return false;
    }

    private static bool TryGetPaintedGlyphGeometry(
        PdfTextSpan span,
        out double[] boundaries,
        out IReadOnlyList<int> glyphCharacterLengths,
        out IReadOnlyList<double> glyphPaintedAdvances) {
        glyphCharacterLengths = span.GlyphCharacterLengths ?? Array.Empty<int>();
        glyphPaintedAdvances = span.GlyphPaintedAdvances ?? Array.Empty<double>();
        return PdfTextAdvanceProjection.TryGetResolvedBoundaries(span, out boundaries) &&
            glyphCharacterLengths.Count == glyphPaintedAdvances.Count &&
            glyphCharacterLengths.Count > 0 &&
            glyphCharacterLengths.All(static length => length > 0) &&
            glyphCharacterLengths.Sum() == span.Text.Length &&
            glyphPaintedAdvances.All(static advance =>
                !double.IsNaN(advance) && !double.IsInfinity(advance) && advance >= 0D);
    }

    private static PdfTextSpanBounds Union(PdfTextSpanBounds left, PdfTextSpanBounds right) => new(
        Math.Min(left.Left, right.Left),
        Math.Min(left.Bottom, right.Bottom),
        Math.Max(left.Right, right.Right),
        Math.Max(left.Top, right.Top));

    private static bool Intersects(PdfTextSpanBounds bounds, double x, double y, double width, double height) =>
        x < bounds.Right &&
        x + width > bounds.Left &&
        y < bounds.Top &&
        y + height > bounds.Bottom;
}

internal readonly struct PdfTextSpanBounds {
    internal PdfTextSpanBounds(double left, double bottom, double right, double top) {
        Left = left;
        Bottom = bottom;
        Right = right;
        Top = top;
    }

    internal double Left { get; }
    internal double Bottom { get; }
    internal double Right { get; }
    internal double Top { get; }
    internal double Width => Math.Max(0.1D, Right - Left);
    internal double Height => Math.Max(0.1D, Top - Bottom);
}
