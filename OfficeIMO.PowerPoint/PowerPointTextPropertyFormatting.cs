using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint;

internal static class PowerPointTextPropertyFormatting {
    internal static void SetFontName(A.TextCharacterPropertiesType properties, string? value) {
        properties.RemoveAllChildren<A.LatinFont>();
        if (value != null) properties.AddChild(new A.LatinFont { Typeface = value }, true);
    }

    internal static void SetColor(A.TextCharacterPropertiesType properties, string? value) {
        properties.RemoveAllChildren<A.SolidFill>();
        if (value == null) return;

        // Explicit RGB replaces any imported fill choice. Schema insertion preserves
        // script fonts and actions without removing and rebuilding their elements.
        properties.RemoveAllChildren<A.NoFill>();
        properties.RemoveAllChildren<A.GradientFill>();
        properties.RemoveAllChildren<A.BlipFill>();
        properties.RemoveAllChildren<A.PatternFill>();
        properties.RemoveAllChildren<A.GroupFill>();
        properties.AddChild(new A.SolidFill(new A.RgbColorModelHex { Val = value }), true);
    }
}
