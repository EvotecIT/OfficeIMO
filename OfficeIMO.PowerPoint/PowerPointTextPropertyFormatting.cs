using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static class PowerPointTextPropertyFormatting {
    internal static void SetFontFamily(A.TextCharacterPropertiesType properties, string? value) {
        properties.RemoveAllChildren<A.LatinFont>();
        if (value != null) properties.AddChild(new A.LatinFont { Typeface = value }, true);
    }

    /// <summary>Writes a text fill, retaining native theme slots and opacity.</summary>
    internal static void SetColor(A.TextCharacterPropertiesType properties, OfficeColor? value, OfficeTheme theme) {
        properties.RemoveAllChildren<A.SolidFill>();
        if (!value.HasValue) return;
        properties.RemoveAllChildren<A.NoFill>();
        properties.RemoveAllChildren<A.GradientFill>();
        properties.RemoveAllChildren<A.BlipFill>();
        properties.RemoveAllChildren<A.PatternFill>();
        properties.RemoveAllChildren<A.GroupFill>();
        properties.AddChild(PowerPointDrawingColor.Fill(value, theme), true);
        PowerPointTextHyperlinkColor.UseExplicitTextColor(properties);
    }

    /// <summary>Writes the run highlight color, retaining native theme slots and opacity.</summary>
    internal static void SetBackgroundColor(A.TextCharacterPropertiesType properties, OfficeColor? value, OfficeTheme theme) {
        properties.RemoveAllChildren<A.Highlight>();
        if (value.HasValue) properties.AddChild(new A.Highlight(PowerPointDrawingColor.Create(value.Value, theme)), true);
    }
}
