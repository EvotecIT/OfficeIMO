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

        StrokePolyline(points, color, thickness);
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
    /// <param name="dashOffset">Distance into the dash pattern at which stroking begins.</param>
    public void DrawPatternedPolyline(
        IReadOnlyList<OfficePoint> points,
        OfficeColor color,
        double thickness,
        IReadOnlyList<double>? dashPattern,
        bool resetDashPatternForEachSegment = false,
        double dashOffset = 0D) {
        if (color.A == 0 || points == null || points.Count < 2 || thickness <= 0D) {
            return;
        }

        StrokePolyline(points, color, thickness, pattern: dashPattern, offset: dashOffset, reset: resetDashPatternForEachSegment);
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
        if (color.A == 0 || points == null || points.Count < 2 || thickness <= 0D || dashLength <= 0D || gapLength < 0D
            || !IsFinite(dashLength) || !IsFinite(gapLength)) {
            return;
        }

        NormalizeRasterDashLengths(ref dashLength, ref gapLength);
        StrokePolyline(points, color, thickness, pattern: new[] { dashLength, gapLength }, reset: resetDashPatternForEachSegment);
    }
}
