namespace OfficeIMO.Drawing;

/// <summary>Identifies the standard ICC rendering intent used for a managed color conversion.</summary>
public enum OfficeIccRenderingIntent {
    /// <summary>Preserves the visual relationship between source colors.</summary>
    Perceptual = 0,
    /// <summary>Preserves in-gamut colors relative to the destination white point.</summary>
    RelativeColorimetric = 1,
    /// <summary>Preserves source saturation for vivid business graphics.</summary>
    Saturation = 2,
    /// <summary>Preserves source colorimetry without mapping source white to destination white.</summary>
    AbsoluteColorimetric = 3
}
