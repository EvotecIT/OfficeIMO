namespace OfficeIMO.Drawing;

// Vertical ink bounds relative to an alphabetic baseline, in the current drawing units.
internal readonly struct OfficeTextPaintBounds {
    internal OfficeTextPaintBounds(double top, double bottom) { Top = top; Bottom = bottom; }
    internal double Top { get; }
    internal double Bottom { get; }
}
