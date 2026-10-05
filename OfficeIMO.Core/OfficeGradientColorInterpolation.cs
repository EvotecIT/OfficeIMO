namespace OfficeIMO.Drawing;

/// <summary>Color space used between gradient stops. Alpha remains a linear scalar.</summary>
public enum OfficeGradientColorInterpolation {
    /// <summary>Interpolate encoded sRGB components.</summary>
    Srgb,
    /// <summary>Interpolate linear-light sRGB components, then encode the result as sRGB.</summary>
    LinearRgb
}
