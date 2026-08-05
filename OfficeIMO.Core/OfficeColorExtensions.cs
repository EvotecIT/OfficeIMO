namespace OfficeIMO.Drawing;

/// <summary>
/// Convenience color conversion helpers shared by OfficeIMO packages.
/// </summary>
public static class OfficeColorExtensions {
    /// <summary>
    /// Converts a color to a hexadecimal RRGGBB value.
    /// </summary>
    public static string ToHexColor(this OfficeColor color) => color.ToRgbHex();
}
