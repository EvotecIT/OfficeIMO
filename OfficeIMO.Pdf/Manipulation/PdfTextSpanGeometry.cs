using System;

namespace OfficeIMO.Pdf;

internal static class PdfTextSpanGeometry {
    internal static PdfTextSpanBounds GetAxisAlignedBounds(PdfTextSpan span) =>
        GetAxisAlignedBounds(span, 0D, span.Advance);

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

        for (int index = 0; index < boundaries.Length - 1; index++) {
            double start = Math.Min(boundaries[index], boundaries[index + 1]);
            double advance = Math.Abs(boundaries[index + 1] - boundaries[index]);
            if (Intersects(GetAxisAlignedBounds(span, start, advance), x, y, width, height)) return true;
        }
        return false;
    }

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
