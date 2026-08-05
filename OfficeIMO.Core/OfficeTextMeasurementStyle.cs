namespace OfficeIMO.Drawing;

/// <summary>
/// Normalized font and DPI data used by <see cref="OfficeTextMeasurer"/>.
/// </summary>
public readonly struct OfficeTextMeasurementStyle {
    internal OfficeTextMeasurementStyle(OfficeFontInfo fontInfo, double dpi) {
        FontInfo = OfficeTextMeasurer.NormalizeFontInfo(fontInfo);
        Dpi = OfficeTextMeasurer.NormalizeDpi(dpi);
        FontSizePixels = FontInfo.Size * Dpi / OfficeTextMeasurer.PointsPerInch;
        SpaceWidthPixels = FontSizePixels * 0.34D * OfficeTextMeasurer.GetFontFamilyWidthFactor(FontInfo) * OfficeTextMeasurer.GetStyleWidthFactor(FontInfo);
        MaximumDigitWidthPixels = FontSizePixels * OfficeTextMeasurer.DefaultDigitEmWidth * OfficeTextMeasurer.GetFontFamilyWidthFactor(FontInfo) * OfficeTextMeasurer.GetStyleWidthFactor(FontInfo);
    }

    /// <summary>Normalized font descriptor used for this measurement style.</summary>
    public OfficeFontInfo FontInfo { get; }

    /// <summary>Measurement DPI.</summary>
    public double Dpi { get; }

    /// <summary>Font size converted to pixels for the selected DPI.</summary>
    public double FontSizePixels { get; }

    /// <summary>Estimated width of a space character in pixels.</summary>
    public double SpaceWidthPixels { get; }

    /// <summary>Estimated maximum digit width in pixels.</summary>
    public double MaximumDigitWidthPixels { get; }
}
