using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Drawing;

/// <summary>
/// Creates reusable <see cref="OfficeShape"/> descriptors for common DrawingML preset geometries.
/// </summary>
public static class OfficeShapePresets {
    /// <summary>
    /// Attempts to create a shared OfficeIMO shape for a DrawingML-style preset name.
    /// </summary>
    /// <param name="presetName">Preset name such as <c>rect</c>, <c>triangle</c>, or <c>rightArrow</c>.</param>
    /// <param name="width">Requested shape width in the caller's layout unit.</param>
    /// <param name="height">Requested shape height in the caller's layout unit.</param>
    /// <param name="horizontalFlip">Whether the shape geometry should be mirrored horizontally.</param>
    /// <param name="verticalFlip">Whether the shape geometry should be mirrored vertically.</param>
    /// <param name="shape">Created shape when the preset is supported.</param>
    /// <returns><c>true</c> when the preset could be mapped; otherwise <c>false</c>.</returns>
    public static bool TryCreate(string? presetName, double width, double height, bool horizontalFlip, bool verticalFlip, out OfficeShape? shape) {
        shape = null;
        string normalized = NormalizePresetName(presetName);
        if (normalized.Length == 0 || !IsFiniteNonNegative(width) || !IsFiniteNonNegative(height)) {
            return false;
        }

        switch (normalized) {
            case "rect":
            case "rectangle":
                if (!HasArea(width, height)) return false;
                shape = OfficeShape.Rectangle(width, height);
                return true;
            case "roundrect":
            case "roundrectangle":
                if (!HasArea(width, height)) return false;
                shape = OfficeShape.RoundedRectangle(width, height, Math.Min(width, height) * 0.18D);
                return true;
            case "ellipse":
            case "oval":
                if (!HasArea(width, height)) return false;
                shape = OfficeShape.Ellipse(width, height);
                return true;
            case "line":
                if (width <= 0D) return false;
                shape = OfficeShape.Line(
                    horizontalFlip ? width : 0D,
                    height / 2D,
                    horizontalFlip ? 0D : width,
                    height / 2D);
                return true;
            case "straightconnector1":
                if (width == 0D && height == 0D) return false;
                shape = OfficeShape.Line(
                    horizontalFlip ? width : 0D,
                    verticalFlip ? height : 0D,
                    horizontalFlip ? 0D : width,
                    verticalFlip ? 0D : height);
                return true;
            case "triangle":
                shape = Polygon(width, height, horizontalFlip, verticalFlip, (0.5D, 0D), (1D, 1D), (0D, 1D));
                return shape != null;
            case "rttriangle":
            case "righttriangle":
                shape = Polygon(width, height, horizontalFlip, verticalFlip, (0D, 0D), (1D, 1D), (0D, 1D));
                return shape != null;
            case "diamond":
                shape = Polygon(width, height, horizontalFlip, verticalFlip, (0.5D, 0D), (1D, 0.5D), (0.5D, 1D), (0D, 0.5D));
                return shape != null;
            case "parallelogram":
                shape = Polygon(width, height, horizontalFlip, verticalFlip, (0.22D, 0D), (1D, 0D), (0.78D, 1D), (0D, 1D));
                return shape != null;
            case "trapezoid":
                shape = Polygon(width, height, horizontalFlip, verticalFlip, (0.22D, 0D), (0.78D, 0D), (1D, 1D), (0D, 1D));
                return shape != null;
            case "pentagon":
                shape = RegularPolygon(width, height, horizontalFlip, verticalFlip, 5, -90D);
                return shape != null;
            case "hexagon":
                shape = RegularPolygon(width, height, horizontalFlip, verticalFlip, 6, 30D);
                return shape != null;
            case "octagon":
                shape = RegularPolygon(width, height, horizontalFlip, verticalFlip, 8, 22.5D);
                return shape != null;
            case "plus":
                shape = Polygon(width, height, horizontalFlip, verticalFlip,
                    (0.38D, 0D), (0.62D, 0D), (0.62D, 0.38D), (1D, 0.38D),
                    (1D, 0.62D), (0.62D, 0.62D), (0.62D, 1D), (0.38D, 1D),
                    (0.38D, 0.62D), (0D, 0.62D), (0D, 0.38D), (0.38D, 0.38D));
                return shape != null;
            case "chevron":
                shape = Polygon(width, height, horizontalFlip, verticalFlip, (0D, 0D), (0.72D, 0D), (1D, 0.5D), (0.72D, 1D), (0D, 1D), (0.28D, 0.5D));
                return shape != null;
            case "rightarrow":
                shape = Polygon(width, height, horizontalFlip, verticalFlip, (0D, 0.25D), (0.62D, 0.25D), (0.62D, 0D), (1D, 0.5D), (0.62D, 1D), (0.62D, 0.75D), (0D, 0.75D));
                return shape != null;
            case "leftarrow":
                shape = Polygon(width, height, horizontalFlip, verticalFlip, (1D, 0.25D), (0.38D, 0.25D), (0.38D, 0D), (0D, 0.5D), (0.38D, 1D), (0.38D, 0.75D), (1D, 0.75D));
                return shape != null;
            case "uparrow":
                shape = Polygon(width, height, horizontalFlip, verticalFlip, (0.25D, 1D), (0.25D, 0.38D), (0D, 0.38D), (0.5D, 0D), (1D, 0.38D), (0.75D, 0.38D), (0.75D, 1D));
                return shape != null;
            case "downarrow":
                shape = Polygon(width, height, horizontalFlip, verticalFlip, (0.25D, 0D), (0.25D, 0.62D), (0D, 0.62D), (0.5D, 1D), (1D, 0.62D), (0.75D, 0.62D), (0.75D, 0D));
                return shape != null;
            case "leftrightarrow":
                shape = Polygon(width, height, horizontalFlip, verticalFlip, (0D, 0.5D), (0.25D, 0D), (0.25D, 0.25D), (0.75D, 0.25D), (0.75D, 0D), (1D, 0.5D), (0.75D, 1D), (0.75D, 0.75D), (0.25D, 0.75D), (0.25D, 1D));
                return shape != null;
            case "star5":
                shape = Star(width, height, horizontalFlip, verticalFlip, 5);
                return shape != null;
            case "heart":
                shape = Path(width, height, horizontalFlip, verticalFlip,
                    OfficePathCommand.MoveTo(0.5D, 1D),
                    OfficePathCommand.CubicBezierTo(0.18D, 0.72D, 0D, 0.52D, 0D, 0.28D),
                    OfficePathCommand.CubicBezierTo(0D, 0.08D, 0.16D, 0D, 0.31D, 0D),
                    OfficePathCommand.CubicBezierTo(0.42D, 0D, 0.49D, 0.07D, 0.5D, 0.18D),
                    OfficePathCommand.CubicBezierTo(0.51D, 0.07D, 0.58D, 0D, 0.69D, 0D),
                    OfficePathCommand.CubicBezierTo(0.84D, 0D, 1D, 0.08D, 1D, 0.28D),
                    OfficePathCommand.CubicBezierTo(1D, 0.52D, 0.82D, 0.72D, 0.5D, 1D),
                    OfficePathCommand.Close());
                return shape != null;
            case "cloud":
                shape = Path(width, height, horizontalFlip, verticalFlip,
                    OfficePathCommand.MoveTo(0.18D, 0.7D),
                    OfficePathCommand.CubicBezierTo(0.05D, 0.7D, 0D, 0.58D, 0.09D, 0.48D),
                    OfficePathCommand.CubicBezierTo(0.03D, 0.32D, 0.19D, 0.18D, 0.34D, 0.26D),
                    OfficePathCommand.CubicBezierTo(0.42D, 0.04D, 0.72D, 0.08D, 0.75D, 0.32D),
                    OfficePathCommand.CubicBezierTo(0.94D, 0.27D, 1D, 0.46D, 0.91D, 0.61D),
                    OfficePathCommand.CubicBezierTo(0.84D, 0.75D, 0.63D, 0.76D, 0.54D, 0.68D),
                    OfficePathCommand.CubicBezierTo(0.46D, 0.82D, 0.25D, 0.82D, 0.18D, 0.7D),
                    OfficePathCommand.Close());
                return shape != null;
            case "donut":
                shape = Donut(width, height, horizontalFlip, verticalFlip);
                return shape != null;
            case "can":
                shape = Path(width, height, horizontalFlip, verticalFlip,
                    OfficePathCommand.MoveTo(0D, 0.18D),
                    OfficePathCommand.CubicBezierTo(0D, 0.1026D, 0.2239D, 0.04D, 0.5D, 0.04D),
                    OfficePathCommand.CubicBezierTo(0.7761D, 0.04D, 1D, 0.1026D, 1D, 0.18D),
                    OfficePathCommand.LineTo(1D, 0.82D),
                    OfficePathCommand.CubicBezierTo(1D, 0.8974D, 0.7761D, 0.96D, 0.5D, 0.96D),
                    OfficePathCommand.CubicBezierTo(0.2239D, 0.96D, 0D, 0.8974D, 0D, 0.82D),
                    OfficePathCommand.Close());
                return shape != null;
            case "cube":
                shape = Polygon(width, height, horizontalFlip, verticalFlip, (0.32D, 0D), (1D, 0.18D), (1D, 0.72D), (0.62D, 1D), (0D, 0.82D), (0D, 0.28D));
                return shape != null;
            default:
                return false;
        }
    }

    /// <summary>
    /// Attempts to create a shared OfficeIMO shape for a DrawingML-style preset name.
    /// </summary>
    public static bool TryCreate(string? presetName, double width, double height, out OfficeShape? shape) =>
        TryCreate(presetName, width, height, horizontalFlip: false, verticalFlip: false, out shape);

    private static OfficeShape? Polygon(double width, double height, bool horizontalFlip, bool verticalFlip, params (double X, double Y)[] points) {
        if (!HasArea(width, height)) {
            return null;
        }

        return OfficeShape.Polygon(ToPoints(width, height, horizontalFlip, verticalFlip, points));
    }

    private static OfficeShape? RegularPolygon(double width, double height, bool horizontalFlip, bool verticalFlip, int sides, double rotationDegrees) {
        if (!HasArea(width, height) || sides < 3) {
            return null;
        }

        var points = new List<(double X, double Y)>(sides);
        double rotation = Math.PI * rotationDegrees / 180D;
        for (int i = 0; i < sides; i++) {
            double angle = rotation + Math.PI * 2D * i / sides;
            points.Add((0.5D + Math.Cos(angle) * 0.5D, 0.5D + Math.Sin(angle) * 0.5D));
        }

        return Polygon(width, height, horizontalFlip, verticalFlip, NormalizePoints(points).ToArray());
    }

    private static OfficeShape? Star(double width, double height, bool horizontalFlip, bool verticalFlip, int points) {
        if (!HasArea(width, height) || points < 3) {
            return null;
        }

        var coordinates = new List<(double X, double Y)>(points * 2);
        for (int i = 0; i < points * 2; i++) {
            double radius = i % 2 == 0 ? 0.5D : 0.22D;
            double angle = -Math.PI / 2D + Math.PI * i / points;
            coordinates.Add((0.5D + Math.Cos(angle) * radius, 0.5D + Math.Sin(angle) * radius));
        }

        return Polygon(width, height, horizontalFlip, verticalFlip, NormalizePoints(coordinates).ToArray());
    }

    private static OfficeShape? Donut(double width, double height, bool horizontalFlip, bool verticalFlip) {
        if (!HasArea(width, height)) {
            return null;
        }

        List<OfficePathCommand> commands = CreateEllipsePath(0.5D, 0.5D, 0.5D, 0.5D, clockwise: true);
        commands.AddRange(CreateEllipsePath(0.5D, 0.5D, 0.22D, 0.22D, clockwise: false));
        return Path(width, height, horizontalFlip, verticalFlip, commands);
    }

    private static OfficeShape? Path(double width, double height, bool horizontalFlip, bool verticalFlip, params OfficePathCommand[] commands) =>
        Path(width, height, horizontalFlip, verticalFlip, (IReadOnlyList<OfficePathCommand>)commands);

    private static OfficeShape? Path(double width, double height, bool horizontalFlip, bool verticalFlip, IReadOnlyList<OfficePathCommand> commands) {
        if (!HasArea(width, height)) {
            return null;
        }

        var transformed = new List<OfficePathCommand>(commands.Count);
        for (int i = 0; i < commands.Count; i++) {
            OfficePathCommand command = commands[i];
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    transformed.Add(OfficePathCommand.MoveTo(TransformPoint(command.Point, width, height, horizontalFlip, verticalFlip)));
                    break;
                case OfficePathCommandKind.LineTo:
                    transformed.Add(OfficePathCommand.LineTo(TransformPoint(command.Point, width, height, horizontalFlip, verticalFlip)));
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    transformed.Add(OfficePathCommand.QuadraticBezierTo(
                        TransformPoint(command.ControlPoint1, width, height, horizontalFlip, verticalFlip),
                        TransformPoint(command.Point, width, height, horizontalFlip, verticalFlip)));
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    transformed.Add(OfficePathCommand.CubicBezierTo(
                        TransformPoint(command.ControlPoint1, width, height, horizontalFlip, verticalFlip),
                        TransformPoint(command.ControlPoint2, width, height, horizontalFlip, verticalFlip),
                        TransformPoint(command.Point, width, height, horizontalFlip, verticalFlip)));
                    break;
                case OfficePathCommandKind.Close:
                    transformed.Add(OfficePathCommand.Close());
                    break;
            }
        }

        return OfficeShape.Path(transformed);
    }

    private static List<OfficePathCommand> CreateEllipsePath(double centerX, double centerY, double radiusX, double radiusY, bool clockwise) {
        const double k = 0.5522847498307936D;
        if (clockwise) {
            return new List<OfficePathCommand> {
                OfficePathCommand.MoveTo(centerX + radiusX, centerY),
                OfficePathCommand.CubicBezierTo(centerX + radiusX, centerY + radiusY * k, centerX + radiusX * k, centerY + radiusY, centerX, centerY + radiusY),
                OfficePathCommand.CubicBezierTo(centerX - radiusX * k, centerY + radiusY, centerX - radiusX, centerY + radiusY * k, centerX - radiusX, centerY),
                OfficePathCommand.CubicBezierTo(centerX - radiusX, centerY - radiusY * k, centerX - radiusX * k, centerY - radiusY, centerX, centerY - radiusY),
                OfficePathCommand.CubicBezierTo(centerX + radiusX * k, centerY - radiusY, centerX + radiusX, centerY - radiusY * k, centerX + radiusX, centerY),
                OfficePathCommand.Close()
            };
        }

        return new List<OfficePathCommand> {
            OfficePathCommand.MoveTo(centerX + radiusX, centerY),
            OfficePathCommand.CubicBezierTo(centerX + radiusX, centerY - radiusY * k, centerX + radiusX * k, centerY - radiusY, centerX, centerY - radiusY),
            OfficePathCommand.CubicBezierTo(centerX - radiusX * k, centerY - radiusY, centerX - radiusX, centerY - radiusY * k, centerX - radiusX, centerY),
            OfficePathCommand.CubicBezierTo(centerX - radiusX, centerY + radiusY * k, centerX - radiusX * k, centerY + radiusY, centerX, centerY + radiusY),
            OfficePathCommand.CubicBezierTo(centerX + radiusX * k, centerY + radiusY, centerX + radiusX, centerY + radiusY * k, centerX + radiusX, centerY),
            OfficePathCommand.Close()
        };
    }

    private static IReadOnlyList<(double X, double Y)> NormalizePoints(IReadOnlyList<(double X, double Y)> points) {
        double minX = double.MaxValue;
        double minY = double.MaxValue;
        double maxX = double.MinValue;
        double maxY = double.MinValue;
        for (int i = 0; i < points.Count; i++) {
            minX = Math.Min(minX, points[i].X);
            minY = Math.Min(minY, points[i].Y);
            maxX = Math.Max(maxX, points[i].X);
            maxY = Math.Max(maxY, points[i].Y);
        }

        double spanX = maxX - minX;
        double spanY = maxY - minY;
        if (spanX <= 0D || spanY <= 0D) {
            return points;
        }

        var normalized = new List<(double X, double Y)>(points.Count);
        for (int i = 0; i < points.Count; i++) {
            normalized.Add(((points[i].X - minX) / spanX, (points[i].Y - minY) / spanY));
        }

        return normalized;
    }

    private static IReadOnlyList<OfficePoint> ToPoints(double width, double height, bool horizontalFlip, bool verticalFlip, IReadOnlyList<(double X, double Y)> points) {
        var result = new List<OfficePoint>(points.Count);
        for (int i = 0; i < points.Count; i++) {
            double x = horizontalFlip ? 1D - points[i].X : points[i].X;
            double y = verticalFlip ? 1D - points[i].Y : points[i].Y;
            result.Add(new OfficePoint(x * width, y * height));
        }

        return result;
    }

    private static OfficePoint TransformPoint(OfficePoint point, double width, double height, bool horizontalFlip, bool verticalFlip) {
        double x = horizontalFlip ? 1D - point.X : point.X;
        double y = verticalFlip ? 1D - point.Y : point.Y;
        return new OfficePoint(x * width, y * height);
    }

    private static string NormalizePresetName(string? presetName) {
        if (string.IsNullOrWhiteSpace(presetName)) {
            return string.Empty;
        }

        var chars = new char[presetName!.Length];
        int count = 0;
        for (int i = 0; i < presetName.Length; i++) {
            char value = presetName[i];
            if (char.IsLetterOrDigit(value)) {
                chars[count++] = char.ToLowerInvariant(value);
            }
        }

        if (count == 0) {
            return string.Empty;
        }

        string normalized = new string(chars, 0, count);
        const string openXmlEnumPrefix = "shapetypevaluesinnertext";
        if (normalized.StartsWith(openXmlEnumPrefix, StringComparison.Ordinal)) {
            return normalized.Substring(openXmlEnumPrefix.Length);
        }

        return normalized;
    }

    private static bool HasArea(double width, double height) => width > 0D && height > 0D;

    private static bool IsFiniteNonNegative(double value) => !double.IsNaN(value) && !double.IsInfinity(value) && value >= 0D;
}
