using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Ordered nested drawing content painted through one affine transform and isolated opacity group.
/// </summary>
public sealed class OfficeDrawingEffectGroup : OfficeDrawingElement {
    private readonly OfficeDrawing _drawing;

    /// <summary>Creates a transformed, isolated nested drawing group.</summary>
    public OfficeDrawingEffectGroup(OfficeDrawing drawing, OfficeTransform transform, double opacity = 1D) {
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        if (double.IsNaN(opacity) || double.IsInfinity(opacity) || opacity < 0D || opacity > 1D) {
            throw new ArgumentOutOfRangeException(nameof(opacity), "Group opacity must be between zero and one.");
        }
        _drawing = drawing.Clone();
        Transform = transform;
        Opacity = opacity;
    }

    /// <summary>Detached nested drawing content.</summary>
    public OfficeDrawing Drawing => _drawing.Clone();

    /// <summary>Nested drawing content used by renderers without another clone.</summary>
    internal OfficeDrawing InnerDrawing => _drawing;

    /// <summary>Destination-space affine transform applied to the whole group.</summary>
    public OfficeTransform Transform { get; }

    /// <summary>Isolated group opacity from zero through one.</summary>
    public double Opacity { get; }

    internal override OfficeDrawingElement CloneElement() => new OfficeDrawingEffectGroup(_drawing, Transform, Opacity);
}
