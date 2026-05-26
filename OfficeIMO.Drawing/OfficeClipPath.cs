using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free clipping path descriptor in local top-left coordinates.
/// </summary>
public sealed class OfficeClipPath {
    /// <summary>Clip path type.</summary>
    public OfficeClipPathKind Kind { get; private set; }

    /// <summary>Clip path width in the caller's layout unit.</summary>
    public double Width { get; private set; }

    /// <summary>Clip path height in the caller's layout unit.</summary>
    public double Height { get; private set; }

    /// <summary>Corner radius for rounded rectangle clipping paths.</summary>
    public double CornerRadius { get; private set; }

    /// <summary>Local commands for freeform clipping paths.</summary>
    public IReadOnlyList<OfficePathCommand> Commands { get; private set; } = Array.Empty<OfficePathCommand>();

    private OfficeClipPath() {
    }

    /// <summary>Creates a rectangular clipping path from the shape's local top-left corner.</summary>
    public static OfficeClipPath Rectangle(double width, double height) {
        ValidatePositiveFinite(width, nameof(width));
        ValidatePositiveFinite(height, nameof(height));

        return new OfficeClipPath {
            Kind = OfficeClipPathKind.Rectangle,
            Width = width,
            Height = height
        };
    }

    /// <summary>Creates a rounded rectangular clipping path from the shape's local top-left corner.</summary>
    public static OfficeClipPath RoundedRectangle(double width, double height, double cornerRadius) {
        ValidatePositiveFinite(width, nameof(width));
        ValidatePositiveFinite(height, nameof(height));
        ValidateFiniteNonNegative(cornerRadius, nameof(cornerRadius));

        double maxRadius = Math.Min(width, height) / 2D;
        if (cornerRadius > maxRadius) {
            throw new ArgumentOutOfRangeException(nameof(cornerRadius), "Clip path corner radius cannot exceed half of the width or height.");
        }

        return new OfficeClipPath {
            Kind = OfficeClipPathKind.RoundedRectangle,
            Width = width,
            Height = height,
            CornerRadius = cornerRadius
        };
    }

    /// <summary>Creates a freeform clipping path from commands in local top-left coordinates.</summary>
    public static OfficeClipPath Path(params OfficePathCommand[] commands) => Path((IEnumerable<OfficePathCommand>)commands);

    /// <summary>Creates a freeform clipping path from commands in local top-left coordinates.</summary>
    public static OfficeClipPath Path(IEnumerable<OfficePathCommand> commands) {
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
                throw new ArgumentException("Clip paths must start with a MoveTo command.", nameof(commands));
            }

            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    IncludePoint(command.Point, ref minX, ref minY, ref maxX, ref maxY, ref hasPoint, nameof(commands));
                    break;
                case OfficePathCommandKind.LineTo:
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
                    throw new ArgumentOutOfRangeException(nameof(commands), "Unsupported clip path command kind.");
            }

            source.Add(command);
        }

        if (source.Count == 0 || !hasDraw) {
            throw new ArgumentException("Clip paths require at least one drawing command.", nameof(commands));
        }

        double width = maxX - minX;
        double height = maxY - minY;
        if (width <= 0 || height <= 0) {
            throw new ArgumentException("Clip path commands must describe a non-empty two-dimensional area.", nameof(commands));
        }

        var normalized = new List<OfficePathCommand>(source.Count);
        for (int i = 0; i < source.Count; i++) {
            normalized.Add(source[i].Translate(minX, minY));
        }

        return new OfficeClipPath {
            Kind = OfficeClipPathKind.Path,
            Width = width,
            Height = height,
            Commands = new ReadOnlyCollection<OfficePathCommand>(normalized)
        };
    }

    /// <summary>Creates a detached copy of this clipping path.</summary>
    public OfficeClipPath Clone() => new OfficeClipPath {
        Kind = Kind,
        Width = Width,
        Height = Height,
        CornerRadius = CornerRadius,
        Commands = new ReadOnlyCollection<OfficePathCommand>(new List<OfficePathCommand>(Commands))
    };

    private static void IncludePoint(OfficePoint point, ref double minX, ref double minY, ref double maxX, ref double maxY, ref bool hasPoint, string paramName) {
        ValidateFinitePoint(point, paramName);

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

    private static void ValidateFinitePoint(OfficePoint point, string paramName) {
        if (double.IsNaN(point.X) || double.IsInfinity(point.X) || double.IsNaN(point.Y) || double.IsInfinity(point.Y)) {
            throw new ArgumentOutOfRangeException(paramName, "Clip path points must be finite numbers.");
        }
    }

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
            throw new ArgumentOutOfRangeException(paramName, "Clip path dimensions must be finite positive numbers.");
        }
    }

    private static void ValidateFiniteNonNegative(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) {
            throw new ArgumentOutOfRangeException(paramName, "Clip path corner radius must be a finite non-negative number.");
        }
    }
}
