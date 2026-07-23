using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    /// <summary>
    /// Draws connected line segments through the supplied points.
    /// </summary>
    /// <param name="points">Polyline points in canvas coordinates.</param>
    /// <param name="color">Stroke color.</param>
    /// <param name="thickness">Stroke thickness in canvas pixels.</param>
    public void DrawPolyline(IReadOnlyList<OfficePoint> points, OfficeColor color, double thickness = 1D) {
        if (color.A == 0 || points == null || points.Count < 2 || thickness <= 0D) {
            return;
        }

        for (int i = 1; i < points.Count; i++) {
            DrawLine(points[i - 1].X, points[i - 1].Y, points[i].X, points[i].Y, color, thickness);
        }
    }

    /// <summary>
    /// Draws connected line segments through the supplied points using a shared Office stroke dash style.
    /// </summary>
    /// <param name="points">Polyline points in canvas coordinates.</param>
    /// <param name="color">Stroke color.</param>
    /// <param name="thickness">Stroke thickness in canvas pixels.</param>
    /// <param name="dashStyle">Shared Office stroke dash style.</param>
    /// <param name="resetDashPatternForEachSegment">Whether the dash pattern should restart for every segment.</param>
    public void DrawStyledPolyline(
        IReadOnlyList<OfficePoint> points,
        OfficeColor color,
        double thickness = 1D,
        OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid,
        bool resetDashPatternForEachSegment = false) {
        if (dashStyle == OfficeStrokeDashStyle.Solid) {
            DrawPolyline(points, color, thickness);
            return;
        }

        DrawPatternedPolyline(points, color, thickness, dashStyle.GetDashPattern(thickness), resetDashPatternForEachSegment);
    }

    /// <summary>
    /// Draws connected line segments through the supplied points using an alternating dash and gap pattern.
    /// </summary>
    /// <param name="points">Polyline points in canvas coordinates.</param>
    /// <param name="color">Stroke color.</param>
    /// <param name="thickness">Stroke thickness in canvas pixels.</param>
    /// <param name="dashPattern">Alternating dash and gap lengths in canvas pixels.</param>
    /// <param name="resetDashPatternForEachSegment">Whether the dash pattern should restart for every segment.</param>
    public void DrawPatternedPolyline(
        IReadOnlyList<OfficePoint> points,
        OfficeColor color,
        double thickness,
        IReadOnlyList<double>? dashPattern,
        bool resetDashPatternForEachSegment = false) {
        if (color.A == 0 || points == null || points.Count < 2 || thickness <= 0D) {
            return;
        }

        List<double> pattern = NormalizeDashPattern(dashPattern);
        if (pattern.Count == 0) {
            DrawPolyline(points, color, thickness);
            return;
        }

        double patternPosition = 0D;
        for (int i = 1; i < points.Count; i++) {
            OfficePoint previous = points[i - 1];
            OfficePoint current = points[i];
            if (resetDashPatternForEachSegment) {
                DrawPatternedLine(previous.X, previous.Y, current.X, current.Y, color, thickness, pattern);
            } else {
                DrawPatternedPathSegment(previous, current, color, thickness, pattern, ref patternPosition);
            }
        }
    }

    /// <summary>
    /// Draws a dashed polyline through the supplied points.
    /// </summary>
    /// <param name="points">Polyline points in canvas coordinates.</param>
    /// <param name="color">Stroke color.</param>
    /// <param name="thickness">Stroke thickness in canvas pixels.</param>
    /// <param name="dashLength">Visible dash length in canvas pixels.</param>
    /// <param name="gapLength">Transparent gap length in canvas pixels.</param>
    /// <param name="resetDashPatternForEachSegment">Whether the dash pattern should restart for every segment.</param>
    public void DrawDashedPolyline(
        IReadOnlyList<OfficePoint> points,
        OfficeColor color,
        double thickness = 1D,
        double dashLength = 6D,
        double gapLength = 4D,
        bool resetDashPatternForEachSegment = false) {
        if (color.A == 0 || points == null || points.Count < 2 || thickness <= 0D || dashLength <= 0D || gapLength < 0D) {
            return;
        }

        double patternPosition = 0D;
        for (int i = 1; i < points.Count; i++) {
            OfficePoint previous = points[i - 1];
            OfficePoint current = points[i];
            if (resetDashPatternForEachSegment) {
                DrawDashedLine(previous.X, previous.Y, current.X, current.Y, color, thickness, dashLength, gapLength);
            } else {
                DrawDashedPathSegment(previous, current, color, thickness, dashLength, gapLength, ref patternPosition);
            }
        }
    }

    private static List<double> NormalizeDashPattern(IReadOnlyList<double>? dashPattern) {
        List<double> pattern = new();
        if (dashPattern == null) {
            return pattern;
        }

        for (int i = 0; i < dashPattern.Count; i++) {
            double value = dashPattern[i];
            if (IsFinite(value) && value > 0D) {
                pattern.Add(value);
            }
        }

        if ((pattern.Count & 1) == 1) {
            int originalCount = pattern.Count;
            for (int index = 0; index < originalCount; index++) pattern.Add(pattern[index]);
        }

        return pattern;
    }
}
