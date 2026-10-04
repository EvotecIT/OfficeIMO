using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeDrawingRasterRenderer {
    private static IDisposable PushSingleContourClip(
        OfficeRasterCanvas canvas, OfficeClipPathKind kind, IReadOnlyList<OfficePoint> contour) {
        if (kind == OfficeClipPathKind.Rectangle && contour.Count == 4) {
            OfficePoint a = contour[0], b = contour[1], c = contour[2], d = contour[3];
            bool axisAligned = (a.Y == b.Y && b.X == c.X && c.Y == d.Y && d.X == a.X)
                || (a.X == b.X && b.Y == c.Y && c.X == d.X && d.Y == a.Y);
            if (axisAligned && IsFiniteClipPoint(a) && IsFiniteClipPoint(b)
                && IsFiniteClipPoint(c) && IsFiniteClipPoint(d)) {
                return canvas.PushClipRectangleAtPixelCentres(
                    Math.Min(a.X, c.X), Math.Min(a.Y, c.Y),
                    Math.Max(a.X, c.X), Math.Max(a.Y, c.Y));
            }
        }
        return canvas.PushClipPolygon(contour);
    }

    private static bool IsFiniteClipPoint(OfficePoint point) =>
        !double.IsNaN(point.X) && !double.IsInfinity(point.X)
        && !double.IsNaN(point.Y) && !double.IsInfinity(point.Y);
}
