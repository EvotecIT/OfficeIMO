namespace OfficeIMO.Drawing;

/// <summary>
/// Visual styling for dependency-free callout rendering.
/// </summary>
public sealed class OfficeCalloutStyle {
    /// <summary>Callout body fill.</summary>
    public OfficeColor FillColor { get; set; } = OfficeColor.FromRgb(255, 251, 230);

    /// <summary>Header band fill.</summary>
    public OfficeColor HeaderFillColor { get; set; } = OfficeColor.FromRgb(255, 242, 204);

    /// <summary>Border and pointer stroke color.</summary>
    public OfficeColor StrokeColor { get; set; } = OfficeColor.FromRgb(214, 168, 67);

    /// <summary>Drop-shadow color.</summary>
    public OfficeColor ShadowColor { get; set; } = OfficeColor.FromRgba(15, 23, 42, 46);

    /// <summary>Horizontal shadow offset in CSS pixels before renderer scale is applied.</summary>
    public double ShadowOffsetX { get; set; } = 2D;

    /// <summary>Vertical shadow offset in CSS pixels before renderer scale is applied.</summary>
    public double ShadowOffsetY { get; set; } = 2D;

    /// <summary>Soft shadow spread in CSS pixels before renderer scale is applied.</summary>
    public double ShadowSpread { get; set; } = 1D;

    /// <summary>Vertical accent line color.</summary>
    public OfficeColor AccentColor { get; set; } = OfficeColor.FromRgb(192, 0, 0);

    /// <summary>Title text color.</summary>
    public OfficeColor TitleColor { get; set; } = OfficeColor.FromRgb(92, 64, 14);

    /// <summary>Body text color.</summary>
    public OfficeColor TextColor { get; set; } = OfficeColor.FromRgb(31, 41, 55);

    /// <summary>Inner padding in CSS pixels before renderer scale is applied.</summary>
    public double Padding { get; set; } = 7D;

    /// <summary>Header band height in CSS pixels before renderer scale is applied.</summary>
    public double HeaderHeight { get; set; } = 20D;

    /// <summary>Title font size in CSS pixels before renderer scale is applied.</summary>
    public double TitleFontSize { get; set; } = 9.5D;

    /// <summary>Body font size in CSS pixels before renderer scale is applied.</summary>
    public double TextFontSize { get; set; } = 9D;

    /// <summary>Text line-height multiplier.</summary>
    public double LineHeightFactor { get; set; } = 1.18D;

    /// <summary>Font family used for SVG text output and raster font lookup.</summary>
    public string FontFamily { get; set; } = "Calibri, Arial, sans-serif";
}
