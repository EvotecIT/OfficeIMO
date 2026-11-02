using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint;

/// <summary>Text and measured native geometry from a slide's Drawing scene.</summary>
public sealed class PowerPointTextSpanInfo {
    internal PowerPointTextSpanInfo(string text, string? family, double fontSize, double x, double y,
        double advance, bool bold, bool italic, OfficeColor? color, double rotation) {
        Text = text; BaseFont = family; FontSize = fontSize; X = x; Y = y;
        Advance = advance; IsBold = bold; IsItalic = italic; Color = color; RotationDegrees = rotation;
    }
    /// <summary>Text painted by the native layout.</summary>
    public string Text { get; }
    /// <summary>Resolved font family.</summary>
    public string? BaseFont { get; }
    /// <summary>Painted font size in points.</summary>
    public double FontSize { get; }
    /// <summary>Left text origin in slide points.</summary>
    public double X { get; }
    /// <summary>Baseline in slide points, measured upward from the bottom edge.</summary>
    public double Y { get; }
    /// <summary>Measured advance in points.</summary>
    public double Advance { get; }
    /// <summary>Whether the native font is bold.</summary>
    public bool IsBold { get; }
    /// <summary>Whether the native font is italic.</summary>
    public bool IsItalic { get; }
    /// <summary>Native text color.</summary>
    public OfficeColor? Color { get; }
    /// <summary>Counter-clockwise baseline rotation in degrees.</summary>
    public double RotationDegrees { get; }
}
