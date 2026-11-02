using System;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>Page footer text and the optional rule above it.</summary>
public sealed class OfficeThemeFooter {
    /// <summary>Gets the footer text size.</summary>
    public OfficeLength FontSize { get; }
    /// <summary>Gets the footer text colour.</summary>
    public OfficeColor Color { get; }
    /// <summary>Gets the rule above the footer, or <see cref="OfficeThemeRule.None"/>.</summary>
    public OfficeThemeRule Rule { get; }

    /// <summary>Creates immutable footer styling.</summary>
    public OfficeThemeFooter(OfficeLength fontSize, OfficeColor color, OfficeThemeRule rule) {
        FontSize = ThemeValues.NonNegative(fontSize, nameof(fontSize));
        Color = color;
        Rule = rule ?? throw new ArgumentNullException(nameof(rule));
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeFooter With(OfficeLength? fontSize = null, OfficeColor? color = null, OfficeThemeRule? rule = null) =>
        new OfficeThemeFooter(fontSize ?? FontSize, color ?? Color, rule ?? Rule);
}
