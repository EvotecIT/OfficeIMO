namespace OfficeIMO.Drawing;

/// <summary>
/// Blend operation used when an isolated drawing group is composited over existing content.
/// </summary>
public enum OfficeBlendMode {
    /// <summary>Standard source-over compositing.</summary>
    Normal,
    /// <summary>Multiplies source and backdrop colors.</summary>
    Multiply,
    /// <summary>Screens source and backdrop colors.</summary>
    Screen,
    /// <summary>Combines multiply and screen according to the backdrop.</summary>
    Overlay,
    /// <summary>Selects the darker component.</summary>
    Darken,
    /// <summary>Selects the lighter component.</summary>
    Lighten,
    /// <summary>Brightens the backdrop to reflect the source.</summary>
    ColorDodge,
    /// <summary>Darkens the backdrop to reflect the source.</summary>
    ColorBurn,
    /// <summary>Combines multiply and screen according to the source.</summary>
    HardLight,
    /// <summary>Applies a softer lightening or darkening effect.</summary>
    SoftLight,
    /// <summary>Subtracts the darker component from the lighter component.</summary>
    Difference,
    /// <summary>Produces a lower-contrast difference effect.</summary>
    Exclusion,
    /// <summary>Uses source hue with backdrop saturation and luminosity.</summary>
    Hue,
    /// <summary>Uses source saturation with backdrop hue and luminosity.</summary>
    Saturation,
    /// <summary>Uses source hue and saturation with backdrop luminosity.</summary>
    Color,
    /// <summary>Uses source luminosity with backdrop hue and saturation.</summary>
    Luminosity
}
