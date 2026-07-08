using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free vector shape descriptor shared by OfficeIMO document packages.
/// Dimensions are expressed in the caller's layout unit, such as points for PDF.
/// </summary>
public sealed class OfficeShape {
    /// <summary>Shape type.</summary>
    public OfficeShapeKind Kind { get; set; }

    /// <summary>Shape width in the caller's layout unit.</summary>
    public double Width { get; set; }

    /// <summary>Shape height in the caller's layout unit.</summary>
    public double Height { get; set; }

    /// <summary>Optional stroke color. Null means no outline.</summary>
    public OfficeColor? StrokeColor { get; set; }

    /// <summary>Stroke width in the caller's layout unit.</summary>
    public double StrokeWidth { get; set; } = 1;

    /// <summary>Stroke dash style.</summary>
    public OfficeStrokeDashStyle StrokeDashStyle { get; set; }

    /// <summary>Optional stroke ending style for open paths. Null lets each renderer choose its default.</summary>
    public OfficeStrokeLineCap? StrokeLineCap { get; set; }

    /// <summary>Optional stroke corner style for connected path segments. Null lets each renderer choose its default.</summary>
    public OfficeStrokeLineJoin? StrokeLineJoin { get; set; }

    /// <summary>Optional marker drawn at the first point of open line shapes.</summary>
    public OfficeLineMarker? StrokeStartMarker { get; set; }

    /// <summary>Optional marker drawn at the last point of open line shapes.</summary>
    public OfficeLineMarker? StrokeEndMarker { get; set; }

    /// <summary>Optional fill color. Null means transparent fill.</summary>
    public OfficeColor? FillColor { get; set; }

    /// <summary>Optional gradient fill. When set, renderers should prefer it over <see cref="FillColor"/> for fill-capable shapes.</summary>
    public OfficeLinearGradient? FillGradient { get; set; }

    /// <summary>Optional radial gradient fill. When set, renderers should prefer it over <see cref="FillGradient"/> and <see cref="FillColor"/> for fill-capable shapes.</summary>
    public OfficeRadialGradient? FillRadialGradient { get; set; }

    /// <summary>Optional linear gradient stroke. When set, renderers should prefer it over <see cref="StrokeColor"/> for stroked shapes.</summary>
    public OfficeLinearGradient? StrokeGradient { get; set; }

    /// <summary>Optional radial gradient stroke. When set, renderers should prefer it over <see cref="StrokeGradient"/> and <see cref="StrokeColor"/> for stroked shapes.</summary>
    public OfficeRadialGradient? StrokeRadialGradient { get; set; }

    /// <summary>Optional shadow intent rendered behind the shape when supported by a format-specific renderer.</summary>
    public OfficeShadow? Shadow { get; set; }

    /// <summary>Optional glow intent rendered around the shape when supported by a format-specific renderer.</summary>
    public OfficeGlow? Glow { get; set; }

    /// <summary>Optional fill opacity from 0.0 (transparent) to 1.0 (opaque). Null lets each renderer use opaque fill.</summary>
    public double? FillOpacity { get; set; }

    /// <summary>Optional stroke opacity from 0.0 (transparent) to 1.0 (opaque). Null lets each renderer use opaque stroke.</summary>
    public double? StrokeOpacity { get; set; }

    /// <summary>Optional local transform applied before format-specific rendering.</summary>
    public OfficeTransform? Transform { get; set; }

    /// <summary>Optional local clipping path applied before format-specific rendering.</summary>
    public OfficeClipPath? ClipPath { get; set; }

    /// <summary>Fill rule used for multi-contour path shapes.</summary>
    public OfficeFillRule FillRule { get; set; } = OfficeFillRule.EvenOdd;

    /// <summary>Corner radius for rounded rectangle shapes.</summary>
    public double CornerRadius { get; set; }

    /// <summary>Local points for point-based shapes such as polygons.</summary>
    public IReadOnlyList<OfficePoint> Points { get; private set; } = Array.Empty<OfficePoint>();

    /// <summary>Local commands for freeform path shapes.</summary>
    public IReadOnlyList<OfficePathCommand> PathCommands { get; private set; } = Array.Empty<OfficePathCommand>();

    /// <summary>Creates a rectangle descriptor.</summary>
    public static OfficeShape Rectangle(double width, double height) => new OfficeShape {
        Kind = OfficeShapeKind.Rectangle,
        Width = width,
        Height = height
    };

    /// <summary>Creates a rounded rectangle descriptor.</summary>
    public static OfficeShape RoundedRectangle(double width, double height, double cornerRadius) {
        ValidatePositiveFinite(width, nameof(width), "Shape dimensions must be finite positive numbers.");
        ValidatePositiveFinite(height, nameof(height), "Shape dimensions must be finite positive numbers.");
        ValidateFiniteNonNegative(cornerRadius, nameof(cornerRadius), "Corner radius must be a finite non-negative number.");

        double maxRadius = Math.Min(width, height) / 2D;
        if (cornerRadius > maxRadius) {
            throw new ArgumentOutOfRangeException(nameof(cornerRadius), "Corner radius cannot exceed half of the rounded rectangle width or height.");
        }

        return new OfficeShape {
            Kind = OfficeShapeKind.RoundedRectangle,
            Width = width,
            Height = height,
            CornerRadius = cornerRadius
        };
    }

    /// <summary>Creates a straight line descriptor from two points in a local top-left coordinate space.</summary>
    public static OfficeShape Line(double x1, double y1, double x2, double y2) => Line(new OfficePoint(x1, y1), new OfficePoint(x2, y2));

    /// <summary>Creates a straight line descriptor from two points in a local top-left coordinate space.</summary>
    public static OfficeShape Line(OfficePoint start, OfficePoint end) {
        ValidateFinitePoint(start, nameof(start), "Line points must be finite numbers.");
        ValidateFinitePoint(end, nameof(end), "Line points must be finite numbers.");

        if (start == end) {
            throw new ArgumentException("Line shapes require two different points.", nameof(end));
        }

        double minX = Math.Min(start.X, end.X);
        double minY = Math.Min(start.Y, end.Y);
        double width = Math.Abs(end.X - start.X);
        double height = Math.Abs(end.Y - start.Y);

        return new OfficeShape {
            Kind = OfficeShapeKind.Line,
            Width = width,
            Height = height,
            Points = new ReadOnlyCollection<OfficePoint>(new List<OfficePoint> {
                new OfficePoint(start.X - minX, start.Y - minY),
                new OfficePoint(end.X - minX, end.Y - minY)
            })
        };
    }

    /// <summary>Creates an ellipse descriptor bounded by the supplied width and height.</summary>
    public static OfficeShape Ellipse(double width, double height) => new OfficeShape {
        Kind = OfficeShapeKind.Ellipse,
        Width = width,
        Height = height
    };

    /// <summary>Creates a polygon descriptor from points in a local top-left coordinate space.</summary>
    public static OfficeShape Polygon(params OfficePoint[] points) => Polygon((IEnumerable<OfficePoint>)points);

    /// <summary>Creates a polygon descriptor from points in a local top-left coordinate space.</summary>
    public static OfficeShape Polygon(IEnumerable<OfficePoint> points) {
        if (points is null) {
            throw new ArgumentNullException(nameof(points));
        }

        var source = new List<OfficePoint>();
        double minX = 0;
        double minY = 0;
        double maxX = 0;
        double maxY = 0;
        bool hasPoint = false;

        foreach (var point in points) {
            ValidateFinitePoint(point, nameof(points), "Polygon points must be finite numbers.");

            if (!hasPoint) {
                minX = maxX = point.X;
                minY = maxY = point.Y;
                hasPoint = true;
            } else {
                if (point.X < minX) minX = point.X;
                if (point.Y < minY) minY = point.Y;
                if (point.X > maxX) maxX = point.X;
                if (point.Y > maxY) maxY = point.Y;
            }

            source.Add(point);
        }

        if (source.Count < 3) {
            throw new ArgumentException("Polygon shapes require at least three points.", nameof(points));
        }

        double width = maxX - minX;
        double height = maxY - minY;
        if (width <= 0 || height <= 0) {
            throw new ArgumentException("Polygon points must describe a non-empty two-dimensional area.", nameof(points));
        }

        var normalized = new List<OfficePoint>(source.Count);
        for (int i = 0; i < source.Count; i++) {
            normalized.Add(new OfficePoint(source[i].X - minX, source[i].Y - minY));
        }

        return new OfficeShape {
            Kind = OfficeShapeKind.Polygon,
            Width = width,
            Height = height,
            Points = new ReadOnlyCollection<OfficePoint>(normalized)
        };
    }

    /// <summary>Creates a freeform path descriptor from commands in a local top-left coordinate space.</summary>
    public static OfficeShape Path(params OfficePathCommand[] commands) => Path((IEnumerable<OfficePathCommand>)commands);

    /// <summary>Creates a freeform path descriptor from commands in a local top-left coordinate space.</summary>
    public static OfficeShape Path(IEnumerable<OfficePathCommand> commands) {
        if (commands is null) {
            throw new ArgumentNullException(nameof(commands));
        }

        var source = new List<OfficePathCommand>();
        double minX = 0;
        double minY = 0;
        double maxX = 0;
        double maxY = 0;
        bool hasPoint = false;
        bool hasDraw = false;

        foreach (var command in commands) {
            if (source.Count == 0 && command.Kind != OfficePathCommandKind.MoveTo) {
                throw new ArgumentException("Path shapes must start with a MoveTo command.", nameof(commands));
            }

            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    IncludePoint(command.Point, ref minX, ref minY, ref maxX, ref maxY, ref hasPoint, nameof(commands));
                    break;
                case OfficePathCommandKind.LineTo:
                    IncludePoint(command.Point, ref minX, ref minY, ref maxX, ref maxY, ref hasPoint, nameof(commands));
                    hasDraw = true;
                    break;
                case OfficePathCommandKind.QuadraticBezierTo:
                    IncludePoint(command.ControlPoint1, ref minX, ref minY, ref maxX, ref maxY, ref hasPoint, nameof(commands));
                    IncludePoint(command.Point, ref minX, ref minY, ref maxX, ref maxY, ref hasPoint, nameof(commands));
                    hasDraw = true;
                    break;
                case OfficePathCommandKind.CubicBezierTo:
                    IncludePoint(command.ControlPoint1, ref minX, ref minY, ref maxX, ref maxY, ref hasPoint, nameof(commands));
                    IncludePoint(command.ControlPoint2, ref minX, ref minY, ref maxX, ref maxY, ref hasPoint, nameof(commands));
                    IncludePoint(command.Point, ref minX, ref minY, ref maxX, ref maxY, ref hasPoint, nameof(commands));
                    hasDraw = true;
                    break;
                case OfficePathCommandKind.Close:
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(commands), "Unsupported path command kind.");
            }

            source.Add(command);
        }

        if (source.Count == 0 || !hasDraw) {
            throw new ArgumentException("Path shapes require at least one drawing command.", nameof(commands));
        }

        double width = maxX - minX;
        double height = maxY - minY;
        if (width <= 0 || height <= 0) {
            throw new ArgumentException("Path commands must describe a non-empty two-dimensional area.", nameof(commands));
        }

        var normalized = new List<OfficePathCommand>(source.Count);
        for (int i = 0; i < source.Count; i++) {
            normalized.Add(source[i].Translate(minX, minY));
        }

        return new OfficeShape {
            Kind = OfficeShapeKind.Path,
            Width = width,
            Height = height,
            PathCommands = new ReadOnlyCollection<OfficePathCommand>(normalized)
        };
    }

    /// <summary>Creates a detached copy of this shape.</summary>
    public OfficeShape Clone() => new OfficeShape {
        Kind = Kind,
        Width = Width,
        Height = Height,
        StrokeColor = StrokeColor,
        StrokeWidth = StrokeWidth,
        StrokeDashStyle = StrokeDashStyle,
        StrokeLineCap = StrokeLineCap,
        StrokeLineJoin = StrokeLineJoin,
        StrokeStartMarker = StrokeStartMarker?.Clone(),
        StrokeEndMarker = StrokeEndMarker?.Clone(),
        FillColor = FillColor,
        FillGradient = FillGradient?.Clone(),
        FillRadialGradient = FillRadialGradient?.Clone(),
        StrokeGradient = StrokeGradient?.Clone(),
        StrokeRadialGradient = StrokeRadialGradient?.Clone(),
        Shadow = Shadow?.Clone(),
        Glow = Glow?.Clone(),
        FillOpacity = FillOpacity,
        StrokeOpacity = StrokeOpacity,
        Transform = Transform,
        ClipPath = ClipPath?.Clone(),
        FillRule = FillRule,
        CornerRadius = CornerRadius,
        Points = new ReadOnlyCollection<OfficePoint>(new List<OfficePoint>(Points)),
        PathCommands = new ReadOnlyCollection<OfficePathCommand>(new List<OfficePathCommand>(PathCommands))
    };

    private static void IncludePoint(OfficePoint point, ref double minX, ref double minY, ref double maxX, ref double maxY, ref bool hasPoint, string paramName) {
        ValidateFinitePoint(point, paramName, "Path command points must be finite numbers.");

        if (!hasPoint) {
            minX = maxX = point.X;
            minY = maxY = point.Y;
            hasPoint = true;
        } else {
            if (point.X < minX) minX = point.X;
            if (point.Y < minY) minY = point.Y;
            if (point.X > maxX) maxX = point.X;
            if (point.Y > maxY) maxY = point.Y;
        }
    }

    private static void ValidateFinitePoint(OfficePoint point, string paramName, string message) {
        if (double.IsNaN(point.X) || double.IsInfinity(point.X) || double.IsNaN(point.Y) || double.IsInfinity(point.Y)) {
            throw new ArgumentOutOfRangeException(paramName, message);
        }
    }

    private static void ValidatePositiveFinite(double value, string paramName, string message) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
            throw new ArgumentOutOfRangeException(paramName, message);
        }
    }

    private static void ValidateFiniteNonNegative(double value, string paramName, string message) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) {
            throw new ArgumentOutOfRangeException(paramName, message);
        }
    }
}
