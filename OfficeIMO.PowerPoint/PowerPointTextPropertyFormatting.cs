using System;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static class PowerPointTextPropertyFormatting {
    internal static void SetFontFamily(A.TextCharacterPropertiesType properties, string? value) {
        properties.RemoveAllChildren<A.LatinFont>();
        if (value != null) properties.AddChild(new A.LatinFont { Typeface = value }, true);
    }

    /// <summary>Writes an RGB text fill; alpha below 255 becomes native DrawingML opacity.</summary>
    internal static void SetColor(A.TextCharacterPropertiesType properties, OfficeColor? value) =>
        SetColorLiteral(properties, value.HasValue ? ToLiteral(value.Value) : null);

    /// <summary>Writes a raw RGB or RGBA literal text fill; used by chart text styles.</summary>
    internal static void SetColorLiteral(A.TextCharacterPropertiesType properties, string? value) {
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

    /// <summary>Writes the run background (highlight) RGB and opacity with the same literal convention as text fill.</summary>
    internal static void SetBackgroundColor(A.TextCharacterPropertiesType properties, OfficeColor? value) {
        A.RgbColorModelHex? color = value.HasValue ? LiteralColor(ToLiteral(value.Value)) : null;
        properties.RemoveAllChildren<A.Highlight>();
        if (color != null) properties.AddChild(new A.Highlight(color), true);
    }

    /// <summary>Reads a DrawingML RGB literal, including native alpha, as an <see cref="OfficeColor"/>.</summary>
    internal static OfficeColor? ReadColor(A.RgbColorModelHex? rgb) {
        string? value = rgb?.Val?.Value;
        if (value == null || !OfficeColor.TryParseHex(value, out OfficeColor color)) return null;
        int? alpha = rgb!.GetFirstChild<A.Alpha>()?.Val?.Value;
        if (!alpha.HasValue) return color;
        int a = (int)Math.Round(Math.Max(0, Math.Min(100000, alpha.Value)) * 255D / 100000D);
        return OfficeColor.FromRgba(color.R, color.G, color.B, (byte)a);
    }

    private static string ToLiteral(OfficeColor color) {
        if (color.IsThemeRole) color = color.Resolve(OfficeTheme.Default);
        return color.A < 255 ? color.ToHex() : color.ToRgbHex();
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
