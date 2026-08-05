using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Determines how an isolated mask drawing controls group opacity.
/// </summary>
public enum OfficeSoftMaskMode {
    /// <summary>Uses the alpha channel of the rendered mask.</summary>
    Alpha,
    /// <summary>Uses the alpha-weighted luminosity of the rendered mask.</summary>
    Luminosity
}

/// <summary>
/// Reusable vector soft mask applied while an isolated drawing group is composited.
/// </summary>
public sealed class OfficeDrawingSoftMask {
    private readonly OfficeDrawing _drawing;

    /// <summary>Creates a mask from drawing content in the source group's local coordinate system.</summary>
    public OfficeDrawingSoftMask(
        OfficeDrawing drawing,
        OfficeSoftMaskMode mode = OfficeSoftMaskMode.Alpha,
        OfficeTransform? transform = null,
        OfficeColor? backdropColor = null) {
        _drawing = drawing?.Clone() ?? throw new ArgumentNullException(nameof(drawing));
        Mode = mode;
        Transform = transform ?? OfficeTransform.Identity;
        BackdropColor = backdropColor ?? OfficeColor.Transparent;
    }

    /// <summary>Detached mask drawing.</summary>
    public OfficeDrawing Drawing => _drawing.Clone();

    /// <summary>Mask interpretation.</summary>
    public OfficeSoftMaskMode Mode { get; }

    /// <summary>Local transform applied to the mask before it is sampled.</summary>
    public OfficeTransform Transform { get; }

    /// <summary>Color used where the mask drawing has no coverage.</summary>
    public OfficeColor BackdropColor { get; }

    internal OfficeDrawing InnerDrawing => _drawing;

    internal OfficeDrawingSoftMask Clone() => new OfficeDrawingSoftMask(_drawing, Mode, Transform, BackdropColor);
}
