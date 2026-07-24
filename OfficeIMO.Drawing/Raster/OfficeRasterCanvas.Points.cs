using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    /// <summary>
    /// Draws connected line segments through tuple points in canvas coordinates.
    /// </summary>
    public void DrawPolyline(IReadOnlyList<(double X, double Y)> points, OfficeColor color, double thickness = 1D) {
        if (color.A == 0 || points == null || points.Count < 2 || thickness <= 0D) {
            return;
        }

        DrawPolyline(ToOfficePoints(points), color, thickness);
    }

    /// <summary>
    /// Draws connected line segments through tuple points using a shared Office stroke dash style.
    /// </summary>
    public void DrawStyledPolyline(
        IReadOnlyList<(double X, double Y)> points,
        OfficeColor color,
        double thickness = 1D,
        OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid,
        bool resetDashPatternForEachSegment = false) {
        if (color.A == 0 || points == null || points.Count < 2 || thickness <= 0D) {
            return;
        }

        DrawStyledPolyline(ToOfficePoints(points), color, thickness, dashStyle, resetDashPatternForEachSegment);
    }

    /// <summary>
    /// Draws connected line segments through tuple points using an alternating dash and gap pattern.
    /// </summary>
    public void DrawPatternedPolyline(
        IReadOnlyList<(double X, double Y)> points,
        OfficeColor color,
        double thickness,
        IReadOnlyList<double>? dashPattern,
        bool resetDashPatternForEachSegment = false) {
        if (color.A == 0 || points == null || points.Count < 2 || thickness <= 0D) {
            return;
        }

        DrawPatternedPolyline(ToOfficePoints(points), color, thickness, dashPattern, resetDashPatternForEachSegment);
    }

    /// <summary>
    /// Draws a dashed polyline through tuple points.
    /// </summary>
    public void DrawDashedPolyline(
        IReadOnlyList<(double X, double Y)> points,
        OfficeColor color,
        double thickness = 1D,
        double dashLength = 6D,
        double gapLength = 4D,
        bool resetDashPatternForEachSegment = false) {
        if (color.A == 0 || points == null || points.Count < 2 || thickness <= 0D) {
            return;
        }

        DrawDashedPolyline(ToOfficePoints(points), color, thickness, dashLength, gapLength, resetDashPatternForEachSegment);
    }

    /// <summary>
    /// Strokes a polygon outline through tuple points using a shared Office stroke dash style.
    /// </summary>
    public void DrawStyledPolygon(
        IReadOnlyList<(double X, double Y)> points,
        OfficeColor color,
        double thickness = 1D,
        OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid,
        bool resetDashPatternForEachSegment = false) {
        if (color.A == 0 || points == null || points.Count < 2 || thickness <= 0D) {
            return;
        }

        DrawStyledPolygon(ToOfficePoints(points), color, thickness, dashStyle, resetDashPatternForEachSegment);
    }

    /// <summary>
    /// Strokes a polygon outline through points using a shared Office stroke dash style.
    /// </summary>
    public void DrawStyledPolygon(
        IReadOnlyList<OfficePoint> points,
        OfficeColor color,
        double thickness = 1D,
        OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid,
        bool resetDashPatternForEachSegment = false) {
        if (color.A == 0 || points == null || points.Count < 2 || thickness <= 0D) {
            return;
        }

        List<OfficePoint> closed = new(points.Count + 1);
        for (int i = 0; i < points.Count; i++) {
            closed.Add(points[i]);
        }

        closed.Add(points[0]);
        DrawStyledPolyline(closed, color, thickness, dashStyle, resetDashPatternForEachSegment);
    }

    /// <summary>
    /// Fills a polygon described by tuple points.
    /// </summary>
    public void FillPolygon(IReadOnlyList<(double X, double Y)> points, OfficeColor color) {
        if (color.A == 0 || points == null || points.Count < 3) {
            return;
        }

        FillPolygon(ToOfficePoints(points), color);
    }

    /// <summary>
    /// Fills multiple tuple-point polygon contours using the even-odd fill rule.
    /// </summary>
    public void FillPolygonsEvenOdd(IReadOnlyList<IReadOnlyList<(double X, double Y)>> contours, OfficeColor color) {
        if (color.A == 0 || contours == null || contours.Count == 0) {
            return;
        }

        List<IReadOnlyList<OfficePoint>> converted = new(contours.Count);
        for (int i = 0; i < contours.Count; i++) {
            converted.Add(ToOfficePoints(contours[i]));
        }

        FillPolygonsEvenOdd(converted, color);
    }

    private static List<OfficePoint> ToOfficePoints(IReadOnlyList<(double X, double Y)> points) {
        List<OfficePoint> converted = new(points.Count);
        for (int i = 0; i < points.Count; i++) {
            converted.Add(new OfficePoint(points[i].X, points[i].Y));
        }

        return converted;
    }
}
