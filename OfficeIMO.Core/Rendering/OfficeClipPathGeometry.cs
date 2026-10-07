using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

// Shared nominal clip contours for painting and geometry inspection.
internal static class OfficeClipPathGeometry {
    internal static IReadOnlyList<IReadOnlyList<OfficePoint>> CreateContours(OfficeClipPath clipPath, Func<IReadOnlyList<OfficePoint>, IReadOnlyList<OfficePoint>> transformContour) {
        IReadOnlyList<OfficePoint> basis = transformContour(new[] { new OfficePoint(0, 0), new OfficePoint(1, 0), new OfficePoint(0, 1) });
        double dx1 = basis[1].X - basis[0].X, dy1 = basis[1].Y - basis[0].Y;
        double dx2 = basis[2].X - basis[0].X, dy2 = basis[2].Y - basis[0].Y;
        double pixelsPerUnit = Math.Sqrt(dx1 * dx1 + dy1 * dy1 + dx2 * dx2 + dy2 * dy2);

        IReadOnlyList<OfficePoint> contour;
        switch (clipPath.Kind) {
            case OfficeClipPathKind.Rectangle:
                contour = new[] {
                    new OfficePoint(0D, 0D),
                    new OfficePoint(clipPath.Width, 0D),
                    new OfficePoint(clipPath.Width, clipPath.Height),
                    new OfficePoint(0D, clipPath.Height)
                };
                return new[] { transformContour(contour) };
            case OfficeClipPathKind.RoundedRectangle:
                contour = OfficeCurveFlattening.RoundedRectangle(0D, 0D, clipPath.Width, clipPath.Height, clipPath.CornerRadius, clipPath.CornerRadius, pixelsPerUnit);
                return new[] { transformContour(contour) };
            case OfficeClipPathKind.Path:
                IReadOnlyList<OfficeFlattenedPathContour> flattened = OfficePathFlattener.Flatten(clipPath.Commands, 0D, 0D, 1D, pixelsPerUnit: pixelsPerUnit);
                List<IReadOnlyList<OfficePoint>> contours = new List<IReadOnlyList<OfficePoint>>();
                for (int i = 0; i < flattened.Count; i++) {
                    if (flattened[i].Closed && flattened[i].Points.Count >= 3) {
                        contours.Add(transformContour(flattened[i].Points));
                    }
                }

                return contours;
            default:
                return Array.Empty<IReadOnlyList<OfficePoint>>();
        }
    }

}
