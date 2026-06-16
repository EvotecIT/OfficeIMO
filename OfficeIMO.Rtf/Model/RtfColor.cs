namespace OfficeIMO.Rtf;

/// <summary>
/// RGB color entry used by an RTF document.
/// </summary>
public sealed class RtfColor {
    /// <summary>
    /// Initializes a color.
    /// </summary>
    public RtfColor(byte red, byte green, byte blue) {
        Red = red;
        Green = green;
        Blue = blue;
    }

    /// <summary>Red component.</summary>
    public byte Red { get; }

    /// <summary>Green component.</summary>
    public byte Green { get; }

    /// <summary>Blue component.</summary>
    public byte Blue { get; }

    /// <summary>Optional theme color token associated with this RGB color.</summary>
    public RtfThemeColor? ThemeColor { get; set; }

    /// <summary>Optional tint value from <c>\ctint</c>.</summary>
    public int? Tint { get; set; }

    /// <summary>Optional shade value from <c>\cshade</c>.</summary>
    public int? Shade { get; set; }

    /// <inheritdoc />
    public override string ToString() => $"#{Red:X2}{Green:X2}{Blue:X2}";
}
