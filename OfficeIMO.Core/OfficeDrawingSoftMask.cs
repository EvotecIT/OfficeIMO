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

/// <summary>Defines the luminance coefficients used by a luminosity soft mask.</summary>
public enum OfficeSoftMaskLuminosityStandard {
    /// <summary>Uses the sRGB luminance coefficients defined by SVG and CSS.</summary>
    Srgb,
    /// <summary>Uses the DeviceRGB luminosity coefficients defined by PDF transparency groups.</summary>
    PdfDeviceRgb
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
        OfficeColor? backdropColor = null)
        : this(drawing, mode, transform, backdropColor, OfficeSoftMaskLuminosityStandard.Srgb) {
    }

    /// <summary>Creates a mask with an explicit luminosity standard.</summary>
    public static OfficeDrawingSoftMask CreateWithLuminosityStandard(
        OfficeDrawing drawing,
        OfficeSoftMaskLuminosityStandard luminosityStandard,
        OfficeSoftMaskMode mode = OfficeSoftMaskMode.Alpha,
        OfficeTransform? transform = null,
        OfficeColor? backdropColor = null) =>
        new OfficeDrawingSoftMask(drawing, mode, transform, backdropColor, luminosityStandard);

    /// <summary>Creates a mask with all interpretation settings supplied positionally.</summary>
    public OfficeDrawingSoftMask(
        OfficeDrawing drawing,
        OfficeSoftMaskMode mode,
        OfficeTransform? transform,
        OfficeColor? backdropColor,
        OfficeSoftMaskLuminosityStandard luminosityStandard) {
        _drawing = drawing?.Clone() ?? throw new ArgumentNullException(nameof(drawing));
        if (!Enum.IsDefined(typeof(OfficeSoftMaskMode), mode)) {
            throw new ArgumentOutOfRangeException(nameof(mode));
        }
        if (!Enum.IsDefined(typeof(OfficeSoftMaskLuminosityStandard), luminosityStandard)) {
            throw new ArgumentOutOfRangeException(nameof(luminosityStandard));
        }
        Mode = mode;
        Transform = transform ?? OfficeTransform.Identity;
        BackdropColor = backdropColor ?? OfficeColor.Transparent;
        LuminosityStandard = luminosityStandard;
    }

    /// <summary>Detached mask drawing.</summary>
    public OfficeDrawing Drawing => _drawing.Clone();

    /// <summary>Mask interpretation.</summary>
    public OfficeSoftMaskMode Mode { get; }

    /// <summary>Local transform applied to the mask before it is sampled.</summary>
    public OfficeTransform Transform { get; }

    /// <summary>Color used where the mask drawing has no coverage.</summary>
    public OfficeColor BackdropColor { get; }

    /// <summary>Luminance coefficients used when <see cref="Mode"/> is <see cref="OfficeSoftMaskMode.Luminosity"/>.</summary>
    public OfficeSoftMaskLuminosityStandard LuminosityStandard { get; }

    internal OfficeDrawing InnerDrawing => _drawing;

    internal OfficeDrawingSoftMask Clone() => new OfficeDrawingSoftMask(_drawing, Mode, Transform, BackdropColor, LuminosityStandard);
}
