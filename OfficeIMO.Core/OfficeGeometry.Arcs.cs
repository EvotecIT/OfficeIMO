using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeGeometry {
    private const double MaxCubicArcSegmentRadians = Math.PI / 2D;

    /// <summary>
    /// Converts an elliptical arc into cubic Bezier path commands, excluding the current move point.
    /// </summary>
    /// <param name="start">Current path point. It must represent the point at <paramref name="startRadians"/> on the ellipse.</param>
    /// <param name="radiusX">Horizontal ellipse radius.</param>
    /// <param name="radiusY">Vertical ellipse radius.</param>
    /// <param name="startRadians">Start angle in radians, using the renderer's y-down coordinate system.</param>
    /// <param name="sweepRadians">Sweep angle in radians. Positive values sweep clockwise in y-down coordinates.</param>
    /// <returns>Cubic Bezier commands that approximate the requested arc.</returns>
    public static List<OfficePathCommand> CreateEllipticalArcCubicBezierCommands(
        OfficePoint start,
        double radiusX,
        double radiusY,
        double startRadians,
        double sweepRadians) {
        if (!IsFinite(radiusX) || radiusX <= 0D || !IsFinite(radiusY) || radiusY <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(radiusX), "Arc radii must be positive finite values.");
        }

        if (!IsFinite(startRadians) || !IsFinite(sweepRadians)) {
            throw new ArgumentOutOfRangeException(nameof(startRadians), "Arc angles must be finite values.");
        }

        var commands = new List<OfficePathCommand>();
        if (Math.Abs(sweepRadians) < GeometryTolerance) {
            return commands;
        }

        double centerX = start.X - (Math.Cos(startRadians) * radiusX);
        double centerY = start.Y - (Math.Sin(startRadians) * radiusY);
        int segments = Math.Max(1, (int)Math.Ceiling(Math.Abs(sweepRadians) / MaxCubicArcSegmentRadians));
        double segmentSweep = sweepRadians / segments;
        double segmentStart = startRadians;

        for (int i = 0; i < segments; i++) {
            double segmentEnd = segmentStart + segmentSweep;
            OfficePoint startPoint = CreateEllipsePoint(centerX, centerY, radiusX, radiusY, segmentStart);
            OfficePoint endPoint = CreateEllipsePoint(centerX, centerY, radiusX, radiusY, segmentEnd);
            double controlScale = 4D / 3D * Math.Tan(segmentSweep / 4D);
            OfficePoint startTangent = CreateEllipseTangent(radiusX, radiusY, segmentStart);
            OfficePoint endTangent = CreateEllipseTangent(radiusX, radiusY, segmentEnd);
            commands.Add(OfficePathCommand.CubicBezierTo(
                new OfficePoint(startPoint.X + (controlScale * startTangent.X), startPoint.Y + (controlScale * startTangent.Y)),
                new OfficePoint(endPoint.X - (controlScale * endTangent.X), endPoint.Y - (controlScale * endTangent.Y)),
                endPoint));
            segmentStart = segmentEnd;
        }

        return commands;
    }

    private static OfficePoint CreateEllipsePoint(double centerX, double centerY, double radiusX, double radiusY, double angleRadians) =>
        new OfficePoint(
            centerX + (Math.Cos(angleRadians) * radiusX),
            centerY + (Math.Sin(angleRadians) * radiusY));

    private static OfficePoint CreateEllipseTangent(double radiusX, double radiusY, double angleRadians) =>
        new OfficePoint(
            -Math.Sin(angleRadians) * radiusX,
            Math.Cos(angleRadians) * radiusY);
}
