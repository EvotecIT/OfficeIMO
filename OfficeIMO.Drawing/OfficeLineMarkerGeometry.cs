using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal static class OfficeLineMarkerGeometry {
    internal static IReadOnlyList<OfficePoint> CreateContour(OfficeLineMarker? marker, OfficePoint tip, OfficePoint lineDirection) {
        if (marker == null || marker.Kind == OfficeLineMarkerKind.None) {
            return Array.Empty<OfficePoint>();
        }

        double length = Math.Sqrt((lineDirection.X * lineDirection.X) + (lineDirection.Y * lineDirection.Y));
        if (length <= 0D || double.IsNaN(length) || double.IsInfinity(length)) {
            return Array.Empty<OfficePoint>();
        }

        double ux = lineDirection.X / length;
        double uy = lineDirection.Y / length;
        double px = -uy;
        double py = ux;
        double markerLength = marker.Length;
        double halfWidth = marker.Width / 2D;

        OfficePoint back = new OfficePoint(tip.X - (ux * markerLength), tip.Y - (uy * markerLength));
        OfficePoint baseLeft = new OfficePoint(back.X + (px * halfWidth), back.Y + (py * halfWidth));
        OfficePoint baseRight = new OfficePoint(back.X - (px * halfWidth), back.Y - (py * halfWidth));

        switch (marker.Kind) {
            case OfficeLineMarkerKind.Diamond:
                OfficePoint center = new OfficePoint(tip.X - (ux * markerLength / 2D), tip.Y - (uy * markerLength / 2D));
                return new[] {
                    tip,
                    new OfficePoint(center.X + (px * halfWidth), center.Y + (py * halfWidth)),
                    back,
                    new OfficePoint(center.X - (px * halfWidth), center.Y - (py * halfWidth))
                };
            case OfficeLineMarkerKind.Oval:
                return CreateOvalContour(tip, ux, uy, px, py, markerLength, halfWidth);
            case OfficeLineMarkerKind.Stealth:
                OfficePoint inset = new OfficePoint(tip.X - (ux * markerLength * 0.62D), tip.Y - (uy * markerLength * 0.62D));
                return new[] {
                    tip,
                    baseLeft,
                    inset,
                    baseRight
                };
            case OfficeLineMarkerKind.Arrow:
            case OfficeLineMarkerKind.Triangle:
            default:
                return new[] {
                    tip,
                    baseLeft,
                    baseRight
                };
        }
    }

    private static IReadOnlyList<OfficePoint> CreateOvalContour(OfficePoint tip, double ux, double uy, double px, double py, double markerLength, double halfWidth) {
        const int Segments = 24;
        var points = new List<OfficePoint>(Segments);
        double centerX = tip.X - (ux * markerLength / 2D);
        double centerY = tip.Y - (uy * markerLength / 2D);
        double radiusX = markerLength / 2D;
        for (int i = 0; i < Segments; i++) {
            double angle = (Math.PI * 2D * i) / Segments;
            double along = Math.Cos(angle) * radiusX;
            double across = Math.Sin(angle) * halfWidth;
            points.Add(new OfficePoint(
                centerX + (ux * along) + (px * across),
                centerY + (uy * along) + (py * across)));
        }

        return points;
    }
}
