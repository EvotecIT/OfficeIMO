using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free path command used by shared drawing descriptors.
/// Coordinates are expressed in the path's local top-left coordinate space.
/// </summary>
public struct OfficePathCommand : IEquatable<OfficePathCommand> {
    /// <summary>Command type.</summary>
    public OfficePathCommandKind Kind { get; }

    /// <summary>Primary point. For cubic Bezier commands this is the end point.</summary>
    public OfficePoint Point { get; }

    /// <summary>First cubic Bezier control point.</summary>
    public OfficePoint ControlPoint1 { get; }

    /// <summary>Second cubic Bezier control point.</summary>
    public OfficePoint ControlPoint2 { get; }

    private OfficePathCommand(OfficePathCommandKind kind, OfficePoint point, OfficePoint controlPoint1, OfficePoint controlPoint2) {
        Kind = kind;
        Point = point;
        ControlPoint1 = controlPoint1;
        ControlPoint2 = controlPoint2;
    }

    /// <summary>Creates a move command.</summary>
    public static OfficePathCommand MoveTo(double x, double y) => MoveTo(new OfficePoint(x, y));

    /// <summary>Creates a move command.</summary>
    public static OfficePathCommand MoveTo(OfficePoint point) => new OfficePathCommand(OfficePathCommandKind.MoveTo, point, default(OfficePoint), default(OfficePoint));

    /// <summary>Creates a line command.</summary>
    public static OfficePathCommand LineTo(double x, double y) => LineTo(new OfficePoint(x, y));

    /// <summary>Creates a line command.</summary>
    public static OfficePathCommand LineTo(OfficePoint point) => new OfficePathCommand(OfficePathCommandKind.LineTo, point, default(OfficePoint), default(OfficePoint));

    /// <summary>Creates a cubic Bezier curve command.</summary>
    public static OfficePathCommand CubicBezierTo(double control1X, double control1Y, double control2X, double control2Y, double endX, double endY)
        => CubicBezierTo(new OfficePoint(control1X, control1Y), new OfficePoint(control2X, control2Y), new OfficePoint(endX, endY));

    /// <summary>Creates a cubic Bezier curve command.</summary>
    public static OfficePathCommand CubicBezierTo(OfficePoint controlPoint1, OfficePoint controlPoint2, OfficePoint endPoint)
        => new OfficePathCommand(OfficePathCommandKind.CubicBezierTo, endPoint, controlPoint1, controlPoint2);

    /// <summary>Creates a close command.</summary>
    public static OfficePathCommand Close() => new OfficePathCommand(OfficePathCommandKind.Close, default(OfficePoint), default(OfficePoint), default(OfficePoint));

    internal OfficePathCommand Translate(double offsetX, double offsetY) {
        switch (Kind) {
            case OfficePathCommandKind.MoveTo:
            case OfficePathCommandKind.LineTo:
                return new OfficePathCommand(Kind, new OfficePoint(Point.X - offsetX, Point.Y - offsetY), default(OfficePoint), default(OfficePoint));
            case OfficePathCommandKind.CubicBezierTo:
                return new OfficePathCommand(
                    Kind,
                    new OfficePoint(Point.X - offsetX, Point.Y - offsetY),
                    new OfficePoint(ControlPoint1.X - offsetX, ControlPoint1.Y - offsetY),
                    new OfficePoint(ControlPoint2.X - offsetX, ControlPoint2.Y - offsetY));
            default:
                return this;
        }
    }

    /// <inheritdoc />
    public bool Equals(OfficePathCommand other) =>
        Kind == other.Kind &&
        Point.Equals(other.Point) &&
        ControlPoint1.Equals(other.ControlPoint1) &&
        ControlPoint2.Equals(other.ControlPoint2);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficePathCommand other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = (int)Kind;
            hash = (hash * 397) ^ Point.GetHashCode();
            hash = (hash * 397) ^ ControlPoint1.GetHashCode();
            hash = (hash * 397) ^ ControlPoint2.GetHashCode();
            return hash;
        }
    }

    /// <summary>Compares two path commands for equality.</summary>
    public static bool operator ==(OfficePathCommand left, OfficePathCommand right) => left.Equals(right);

    /// <summary>Compares two path commands for inequality.</summary>
    public static bool operator !=(OfficePathCommand left, OfficePathCommand right) => !left.Equals(right);
}
