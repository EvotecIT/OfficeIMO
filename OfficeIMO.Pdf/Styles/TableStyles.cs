namespace OfficeIMO.Pdf;

/// <summary>
/// Friendly presets for common table appearances.
/// </summary>
public static class TableStyles {
    /// <summary>
    /// Light preset: subtle header fill, soft grid, and gentle row striping.
    /// </summary>
    public static PdfTableStyle Light() => new PdfTableStyle {
        HeaderFill = PdfColor.LightGray,
        RowStripeFill = PdfColor.FromRgb(245, 245, 245),
        BorderColor = PdfColor.FromRgb(210, 210, 210),
        BorderWidth = 0.5,
        // Padding and baseline offset use sane defaults from PdfTableStyle
    };

    /// <summary>
    /// Minimal preset: grid only with no header or row fills.
    /// </summary>
    public static PdfTableStyle Minimal() => new PdfTableStyle {
        BorderColor = PdfColor.FromRgb(210, 210, 210),
        BorderWidth = 0.5
    };

    /// <summary>
    /// Light preset with automatic right alignment for numeric-looking values.
    /// </summary>
    public static PdfTableStyle RightAlignedNumbers() {
        var t = Light();
        t.RightAlignNumeric = true;
        return t;
    }
}
