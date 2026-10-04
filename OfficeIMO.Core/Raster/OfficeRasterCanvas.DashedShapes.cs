using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    /// <summary>
    /// Draws a filled and/or stroked ellipse using a shared Office stroke dash style.
    /// </summary>
    /// <param name="centerX">Ellipse center X coordinate.</param>
    /// <param name="centerY">Ellipse center Y coordinate.</param>
    /// <param name="radiusX">Horizontal ellipse radius.</param>
    /// <param name="radiusY">Vertical ellipse radius.</param>
    /// <param name="fill">Fill color.</param>
    /// <param name="stroke">Stroke color.</param>
    /// <param name="thickness">Stroke thickness in canvas pixels.</param>
    /// <param name="dashStyle">Shared Office stroke dash style.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="segments">Number of line segments used to approximate dashed outlines.</param>
    public void DrawStyledEllipse(
        double centerX,
        double centerY,
        double radiusX,
        double radiusY,
        OfficeColor fill,
        OfficeColor stroke,
        double thickness = 1D,
        OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        int segments = 72) {
        if (dashStyle == OfficeStrokeDashStyle.Solid || stroke.A == 0 || thickness <= 0D) {
            DrawEllipse(centerX, centerY, radiusX, radiusY, fill, stroke, thickness, rotationDegrees, rotationCenterX, rotationCenterY);
            return;
        }

        if (fill.A > 0) {
            DrawEllipse(centerX, centerY, radiusX, radiusY, fill, OfficeColor.Transparent, 0D, rotationDegrees, rotationCenterX, rotationCenterY);
        }

        DrawPatternedEllipse(
            centerX,
            centerY,
            radiusX,
            radiusY,
            stroke,
            thickness,
            dashStyle.GetDashPattern(thickness),
            rotationDegrees,
            rotationCenterX,
            rotationCenterY,
            segments);
    }

    /// <summary>
    /// Draws a dashed elliptical outline using center/radius coordinates and optional rotation.
    /// </summary>
    /// <param name="centerX">Ellipse center X coordinate.</param>
    /// <param name="centerY">Ellipse center Y coordinate.</param>
    /// <param name="radiusX">Horizontal ellipse radius.</param>
    /// <param name="radiusY">Vertical ellipse radius.</param>
    /// <param name="color">Stroke color.</param>
    /// <param name="thickness">Stroke thickness in canvas pixels.</param>
    /// <param name="dashLength">Visible dash length in canvas pixels.</param>
    /// <param name="gapLength">Transparent gap length in canvas pixels.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="segments">Number of line segments used to approximate the ellipse.</param>
    /// <param name="resetDashPatternForEachSegment">Whether the dash pattern should restart for every approximation segment.</param>
    public void DrawDashedEllipse(
        double centerX,
        double centerY,
        double radiusX,
        double radiusY,
        OfficeColor color,
        double thickness = 1D,
        double dashLength = 6D,
        double gapLength = 4D,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        int segments = 72,
        bool resetDashPatternForEachSegment = false) {
        if (color.A == 0 || thickness <= 0D || radiusX <= 0D || radiusY <= 0D || dashLength <= 0D || gapLength < 0D || segments < 4) {
            return;
        }

        var points = CreateEllipseStrokePoints(centerX, centerY, radiusX, radiusY, rotationDegrees, rotationCenterX, rotationCenterY, segments);
        DrawDashedPolyline(points, color, thickness, dashLength, gapLength, resetDashPatternForEachSegment);
    }
    /// <summary>
    /// Draws an elliptical outline using an alternating dash and gap pattern.
    /// </summary>
    /// <param name="centerX">Ellipse center X coordinate.</param>
    /// <param name="centerY">Ellipse center Y coordinate.</param>
    /// <param name="radiusX">Horizontal ellipse radius.</param>
    /// <param name="radiusY">Vertical ellipse radius.</param>
    /// <param name="color">Stroke color.</param>
    /// <param name="thickness">Stroke thickness in canvas pixels.</param>
    /// <param name="dashPattern">Alternating dash and gap lengths in canvas pixels.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="segments">Number of line segments used to approximate the ellipse.</param>
    public void DrawPatternedEllipse(
        double centerX,
        double centerY,
        double radiusX,
        double radiusY,
        OfficeColor color,
        double thickness,
        IReadOnlyList<double>? dashPattern,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        int segments = 72) {
        if (color.A == 0 || thickness <= 0D || radiusX <= 0D || radiusY <= 0D || segments < 4) {
            return;
        }

        var points = CreateEllipseStrokePoints(centerX, centerY, radiusX, radiusY, rotationDegrees, rotationCenterX, rotationCenterY, segments);
        DrawPatternedPolyline(points, color, thickness, dashPattern);
    }

    private static List<OfficePoint> CreateEllipseStrokePoints(double centerX, double centerY, double radiusX, double radiusY,
        double rotationDegrees, double rotationCenterX, double rotationCenterY, int segments) {
        segments = Math.Max(segments, OfficeCurveFlattening.ArcSegments(Math.Max(radiusX, radiusY), Math.PI * 2D));
        double rotation = OfficeGeometry.DegreesToRadians(rotationDegrees);
        var points = new List<OfficePoint> { CreateArcStartPoint(centerX, centerY, radiusX, radiusY, 0D, rotation, rotationCenterX, rotationCenterY) };
        points.AddRange(OfficeGeometry.CreateEllipticalArcPoints(centerX, centerY, radiusX, radiusY, 0D, Math.PI * 2D, segments, rotation, rotationCenterX, rotationCenterY));
        return points;
    }
    private static void NormalizeRasterDashLengths(ref double dashLength, ref double gapLength) {
        double smallest = gapLength > 0D ? Math.Min(dashLength, gapLength) : dashLength;
        if (!IsFinite(smallest) || smallest <= 0D || smallest >= MinimumRasterDashLength) return;
        double scale = MinimumRasterDashLength / smallest;
        double normalizedDash = dashLength * scale;
        double normalizedGap = gapLength * scale;
        if (!IsFinite(scale) || !IsFinite(normalizedDash) || !IsFinite(normalizedGap)) {
            dashLength = Math.Max(dashLength, MinimumRasterDashLength);
            if (gapLength > 0D) gapLength = Math.Max(gapLength, MinimumRasterDashLength);
            return;
        }
        dashLength = normalizedDash;
        gapLength = normalizedGap;
    }

}
