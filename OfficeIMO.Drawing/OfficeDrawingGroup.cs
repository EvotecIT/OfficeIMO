using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Clipped nested drawing element inside an <see cref="OfficeDrawing"/> canvas.
/// </summary>
public sealed class OfficeDrawingGroup : OfficeDrawingElement {
    private readonly OfficeDrawing _drawing;

    /// <summary>Creates a clipped nested drawing element.</summary>
    public OfficeDrawingGroup(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath)
        : this(drawing, x, y, clipPath, null) {
    }

    /// <summary>Creates a clipped nested drawing element with an optional destination-space frame transform.</summary>
    public OfficeDrawingGroup(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath, OfficeImageFrameTransform? frameTransform) {
        if (drawing == null) {
            throw new ArgumentNullException(nameof(drawing));
        }

        if (clipPath == null) {
            throw new ArgumentNullException(nameof(clipPath));
        }

        ValidateFiniteNonNegative(x, nameof(x));
        ValidateFiniteNonNegative(y, nameof(y));
        _drawing = drawing.Clone();
        X = x;
        Y = y;
        ClipPath = clipPath.Clone();
        FrameTransform = frameTransform;
    }

    /// <summary>Detached nested drawing content.</summary>
    public OfficeDrawing Drawing => _drawing.Clone();

    /// <summary>Nested drawing content used by renderers without additional cloning.</summary>
    internal OfficeDrawing InnerDrawing => _drawing;

    /// <summary>Group horizontal position inside the parent drawing.</summary>
    public double X { get; }

    /// <summary>Group vertical position inside the parent drawing.</summary>
    public double Y { get; }

    /// <summary>Clipping path in group-local coordinates.</summary>
    public OfficeClipPath ClipPath { get; }

    /// <summary>Optional destination-space frame transform applied to the group and its clipping path.</summary>
    public OfficeImageFrameTransform? FrameTransform { get; }

    internal override OfficeDrawingElement CloneElement() =>
        new OfficeDrawingGroup(_drawing, X, Y, ClipPath, FrameTransform);

    private static void ValidateFiniteNonNegative(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing group coordinates must be finite non-negative numbers.");
        }
    }
}
