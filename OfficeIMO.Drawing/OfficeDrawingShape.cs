using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Positioned shape inside an <see cref="OfficeDrawing"/> canvas.
/// Coordinates use the drawing's local top-left coordinate space.
/// </summary>
public sealed class OfficeDrawingShape {
    /// <summary>Shape horizontal position inside the drawing.</summary>
    public double X { get; }

    /// <summary>Shape vertical position inside the drawing.</summary>
    public double Y { get; }

    /// <summary>Detached shape descriptor.</summary>
    public OfficeShape Shape { get; }

    /// <summary>Creates a positioned shape.</summary>
    public OfficeDrawingShape(OfficeShape shape, double x, double y) {
        if (shape is null) {
            throw new ArgumentNullException(nameof(shape));
        }

        ValidateFiniteNonNegative(x, nameof(x));
        ValidateFiniteNonNegative(y, nameof(y));
        ValidatePositiveFinite(shape.Width, nameof(shape.Width));
        ValidatePositiveFinite(shape.Height, nameof(shape.Height));

        Shape = shape.Clone();
        X = x;
        Y = y;
    }

    /// <summary>Creates a detached copy of this positioned shape.</summary>
    public OfficeDrawingShape Clone() => new OfficeDrawingShape(Shape, X, Y);

    private static void ValidateFiniteNonNegative(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing coordinates must be finite non-negative numbers.");
        }
    }

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing shape dimensions must be finite positive numbers.");
        }
    }
}
