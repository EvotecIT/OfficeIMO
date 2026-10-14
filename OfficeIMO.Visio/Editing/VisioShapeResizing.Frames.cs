using System;

namespace OfficeIMO.Visio;

internal static partial class VisioShapeResizing {
    private static (double X, double Y) FrameScale(double angle, double x, double y) {
        if (!Finite(angle)) throw new NotSupportedException("Resizing requires finite frame angles.");
        if (Math.Abs(x - y) < 1e-12 || Math.Abs(Math.Sin(angle)) < 1e-12) return (x, y);
        if (Math.Abs(Math.Cos(angle)) < 1e-12) return (y, x);
        throw new NotSupportedException("Unequal scaling of this rotated child or text frame requires a shear. Use uniform resizing.");
    }

    private static void ScaleTextFrame(VisioShape source, VisioShape target, double x, double y) {
        if (target.TextStyle is not VisioTextStyle style) return;
        (double frameX, double frameY) = FrameScale(style.TextAngle ?? 0, x, y);
        style.TextPinX *= x; style.TextPinY *= y;
        if (frameX == x && frameY == y) {
            style.TextWidth *= x; style.TextHeight *= y;
        } else {
            // A rotated frame's omitted extents would otherwise default to the
            // resized shape axes rather than its own swapped axes.
            style.TextWidth = (style.TextWidth ?? source.Width) * frameX;
            style.TextHeight = (style.TextHeight ?? source.Height) * frameY;
        }
        style.TextLocPinX *= frameX; style.TextLocPinY *= frameY;
        foreach (double? value in new[] { style.TextPinX, style.TextPinY, style.TextWidth, style.TextHeight, style.TextLocPinX, style.TextLocPinY })
            if (value.HasValue && !Finite(value.Value)) throw new NotSupportedException("Resizing requires finite cached text frames.");
    }
}
