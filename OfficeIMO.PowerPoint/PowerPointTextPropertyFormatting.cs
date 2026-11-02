using System;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static class PowerPointTextPropertyFormatting {
    internal static void SetFontName(A.TextCharacterPropertiesType properties, string? value) {
        properties.RemoveAllChildren<A.LatinFont>();
        if (value != null) properties.AddChild(new A.LatinFont { Typeface = value }, true);
    }

    internal static void SetColor(A.TextCharacterPropertiesType properties, string? value) {
        A.RgbColorModelHex? color = value == null ? null : LiteralColor(value);
        properties.RemoveAllChildren<A.SolidFill>();
        if (color == null) return;

        // Explicit RGB replaces any imported fill choice. Schema insertion preserves
        // script fonts and actions without removing and rebuilding their elements.
        properties.RemoveAllChildren<A.NoFill>();
        properties.RemoveAllChildren<A.GradientFill>();
        properties.RemoveAllChildren<A.BlipFill>();
        properties.RemoveAllChildren<A.PatternFill>();
        properties.RemoveAllChildren<A.GroupFill>();
        properties.AddChild(new A.SolidFill(color), true);
        PowerPointTextHyperlinkColor.UseExplicitTextColor(properties);
    }

    /// <summary>Writes highlight RGB and opacity with the same literal convention as text fill.</summary>
    internal static void SetHighlightColor(A.TextCharacterPropertiesType properties, string? value) {
        A.RgbColorModelHex? color = value == null ? null : LiteralColor(value);
        properties.RemoveAllChildren<A.Highlight>();
        if (color != null) properties.AddChild(new A.Highlight(color), true);
    }

    private static A.RgbColorModelHex LiteralColor(string value) {
        var literal = new A.RgbColorModelHex { Val = value };
        // Keep raw native values for validation/import workflows; only expand an actual RGBA literal.
        if (value.Length != 8 || !OfficeColor.TryParseHex(value, out OfficeColor color)) return literal;
        literal.Val = color.ToRgbHex();
        if (color.A < 255) literal.Append(new A.Alpha { Val = (int)Math.Round(color.A * 100000D / 255D) });
        return literal;
    }
}
