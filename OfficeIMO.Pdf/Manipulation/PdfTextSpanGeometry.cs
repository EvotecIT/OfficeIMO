using System;

namespace OfficeIMO.Pdf;

internal static class PdfTextSpanGeometry {
    internal static PdfTextSpanBounds GetAxisAlignedBounds(PdfTextSpan span) {
#if NET6_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(span);
#else
        if (span == null) throw new ArgumentNullException(nameof(span));
#endif

        double advance = Math.Max(1D, Math.Abs(span.Advance));
        double fontSize = Math.Max(1D, Math.Abs(span.FontSize));
        double radians = span.RotationDegrees * Math.PI / 180D;
        double alongX = Math.Cos(radians);
        double alongY = Math.Sin(radians);
        double normalX = -alongY;
        double normalY = alongX;
        double x0 = span.X - normalX * fontSize;
        double y0 = span.Y - normalY * fontSize;
        double x1 = span.X + alongX * advance - normalX * fontSize;
        double y1 = span.Y + alongY * advance - normalY * fontSize;
        double x2 = span.X;
        double y2 = span.Y;
        double x3 = span.X + alongX * advance;
        double y3 = span.Y + alongY * advance;
        double left = Math.Min(Math.Min(x0, x1), Math.Min(x2, x3));
        double right = Math.Max(Math.Max(x0, x1), Math.Max(x2, x3));
        double bottom = Math.Min(Math.Min(y0, y1), Math.Min(y2, y3));
        double top = Math.Max(Math.Max(y0, y1), Math.Max(y2, y3));
        return new PdfTextSpanBounds(left, bottom, right, top);
    }
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
