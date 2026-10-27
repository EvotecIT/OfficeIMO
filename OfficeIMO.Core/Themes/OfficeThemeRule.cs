using System;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>A horizontal rule (colour and stroke width) drawn under a title, a table header or above a footer.</summary>
public sealed class OfficeThemeRule {
    /// <summary>Gets a rule that is not drawn.</summary>
    public static OfficeThemeRule None { get; } = new OfficeThemeRule(OfficeColor.Text, OfficeLength.Zero);

    /// <summary>Gets the rule colour, which may refer to a theme role.</summary>
    public OfficeColor Color { get; }
    /// <summary>Gets the stroke width; zero means no rule.</summary>
    public OfficeLength Width { get; }
    /// <summary>Gets whether the rule is drawn.</summary>
    public bool IsVisible => Width > OfficeLength.Zero;

    /// <summary>Creates an immutable rule.</summary>
    public OfficeThemeRule(OfficeColor color, OfficeLength width) {
        Color = color;
        Width = ThemeValues.NonNegative(width, nameof(width));
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeRule With(OfficeColor? color = null, OfficeLength? width = null) =>
        new OfficeThemeRule(color ?? Color, width ?? Width);
}
