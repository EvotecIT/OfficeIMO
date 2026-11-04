using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word;

/// <summary>Maps authoring colors to native Word theme references and their literal fallback.</summary>
internal static class WordRunColor {
    internal static OfficeColor Resolve(OfficeColor color, WordDocument? document) {
        if (!color.IsThemeRole) return color;
        var scheme = document?._wordprocessingDocument.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        return Resolve(color, scheme, document?.EffectiveCreationTheme ?? OfficeTheme.Default);
    }

    /// <summary>Resolves retained native color context without requiring an open package.</summary>
    internal static OfficeColor Resolve(OfficeColor color, A.ColorScheme? scheme, OfficeTheme theme) =>
        !color.IsThemeRole ? color : OfficeOpenXmlThemeColorResolver.ResolveSchemeColor(scheme, color.ThemeSlot) ?? color.Resolve(theme);

    internal static Color Create(OfficeColor color, WordDocument? document) {
        var native = new Color { Val = Resolve(color, document).ToRgbHex() };
        if (color.ThemeSlot != null) native.ThemeColor = color.Role switch {
            OfficeColorRole.Accent1 => ThemeColorValues.Accent1,
            OfficeColorRole.Accent2 => ThemeColorValues.Accent2,
            OfficeColorRole.Accent3 => ThemeColorValues.Accent3,
            OfficeColorRole.Accent4 => ThemeColorValues.Accent4,
            OfficeColorRole.Accent5 => ThemeColorValues.Accent5,
            OfficeColorRole.Accent6 => ThemeColorValues.Accent6,
            OfficeColorRole.Text => ThemeColorValues.Dark1,
            OfficeColorRole.TextMuted => ThemeColorValues.Dark2,
            OfficeColorRole.Background => ThemeColorValues.Light1,
            _ => ThemeColorValues.None
        };
        return native;
    }

    /// <summary>Writes native DrawingML slots, resolving status colors against the destination document.</summary>
    internal static A.SolidFill CreateDrawingFill(OfficeColor color, WordDocument? document) {
        if (color.ThemeSlot is string slot) {
            var scheme = new A.SchemeColor();
            scheme.SetAttribute(new OpenXmlAttribute("val", string.Empty, slot));
            return new A.SolidFill(scheme);
        }
        var literal = Resolve(color, document);
        var rgb = new A.RgbColorModelHex { Val = literal.ToRgbHex() };
        if (literal.A < 255) rgb.Append(new A.Alpha { Val = (int)Math.Round(literal.A * 100000D / 255D) });
        return new A.SolidFill(rgb);
    }

    /// <summary>Replaces a DrawingML fill choice without leaving competing fill nodes.</summary>
    internal static void SetDrawingFill(OpenXmlCompositeElement properties, OfficeColor color, WordDocument? document) {
        properties.RemoveAllChildren<A.NoFill>();
        properties.RemoveAllChildren<A.SolidFill>();
        properties.RemoveAllChildren<A.GradientFill>();
        properties.RemoveAllChildren<A.PatternFill>();
        properties.RemoveAllChildren<A.BlipFill>();
        properties.RemoveAllChildren<A.GroupFill>();
        properties.AddChild(CreateDrawingFill(color, document), true);
    }

    internal static OfficeColor? ReadDrawingFill(A.SolidFill? fill, A.ColorScheme? scheme) {
        var schemeColor = fill?.GetFirstChild<A.SchemeColor>();
        if (schemeColor != null && !schemeColor.HasChildren) {
            var role = ReadRole(schemeColor.GetAttribute("val", string.Empty).Value);
            if (role.HasValue) return role;
        }
        return OfficeOpenXmlThemeColorResolver.ResolveColor(fill, scheme);
    }

    internal static OfficeColor? Read(Color? native) {
        if (native?.ThemeColor != null && native.ThemeTint == null && native.ThemeShade == null) {
            OfficeColor? role = ReadRole(native.ThemeColor.InnerText);
            if (role.HasValue) return role;
        }
        return OfficeColor.TryParseHex(native?.Val?.Value, out var color) ? color : null;
    }

    private static OfficeColor? ReadRole(string? slot) => slot switch {
        "accent1" => OfficeColor.Accent1, "accent2" => OfficeColor.Accent2,
        "accent3" => OfficeColor.Accent3, "accent4" => OfficeColor.Accent4,
        "accent5" => OfficeColor.Accent5, "accent6" => OfficeColor.Accent6,
        "dark1" or "dk1" => OfficeColor.Text, "dark2" or "dk2" => OfficeColor.TextMuted,
        "light1" or "lt1" => OfficeColor.Background, _ => null
    };
}
