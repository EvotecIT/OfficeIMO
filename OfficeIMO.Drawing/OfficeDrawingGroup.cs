using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Clipped nested drawing element inside an <see cref="OfficeDrawing"/> canvas.
/// Finite negative coordinates are retained for clipped intermediate scenes; drawing insertion APIs still validate the visible clip against their canvas.
/// </summary>
public sealed class OfficeDrawingGroup : OfficeDrawingElement {
    private readonly OfficeDrawing _drawing;

    /// <summary>Creates a clipped nested drawing element.</summary>
    public OfficeDrawingGroup(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath)
        : this(drawing, x, y, clipPath, 0D, 0D, null) {
    }

    /// <summary>Creates a clipped nested drawing element with an optional destination-space frame transform.</summary>
    public OfficeDrawingGroup(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath, OfficeImageFrameTransform? frameTransform)
        : this(drawing, x, y, clipPath, 0D, 0D, frameTransform) {
    }

    /// <summary>Creates a clipped nested drawing with an independent content offset inside the clip rectangle.</summary>
    public OfficeDrawingGroup(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath, double contentOffsetX, double contentOffsetY)
        : this(drawing, x, y, clipPath, contentOffsetX, contentOffsetY, null) {
    }

    /// <summary>Creates a clipped nested drawing with independent content offset and optional destination-space frame transform.</summary>
    public OfficeDrawingGroup(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath, double contentOffsetX, double contentOffsetY, OfficeImageFrameTransform? frameTransform) {
        if (drawing == null) {
            throw new ArgumentNullException(nameof(drawing));
        }

        if (clipPath == null) {
            throw new ArgumentNullException(nameof(clipPath));
        }

        ValidateFinite(x, nameof(x));
        ValidateFinite(y, nameof(y));
        ValidateFinite(contentOffsetX, nameof(contentOffsetX));
        ValidateFinite(contentOffsetY, nameof(contentOffsetY));
        _drawing = drawing.Clone();
        X = x;
        Y = y;
        ClipPath = clipPath.Clone();
        ContentOffsetX = contentOffsetX;
        ContentOffsetY = contentOffsetY;
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

    /// <summary>Horizontal translation applied to nested content after the clip is positioned.</summary>
    public double ContentOffsetX { get; }

    /// <summary>Vertical translation applied to nested content after the clip is positioned.</summary>
    public double ContentOffsetY { get; }

    /// <summary>Optional destination-space frame transform applied to the group and its clipping path.</summary>
    public OfficeImageFrameTransform? FrameTransform { get; }

    internal override OfficeDrawingElement CloneElement() =>
        new OfficeDrawingGroup(_drawing, X, Y, ClipPath, ContentOffsetX, ContentOffsetY, FrameTransform);

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing group content offsets must be finite.");
        }
    }

}
