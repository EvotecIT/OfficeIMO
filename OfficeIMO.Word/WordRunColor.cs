using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

/// <summary>Maps authoring colors to native Word theme references and their literal fallback.</summary>
internal static class WordRunColor {
    internal static Color Create(OfficeColor color, WordDocument? document) {
        var scheme = document?._wordprocessingDocument.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        OfficeColor literal = OfficeOpenXmlThemeColorResolver.ResolveSchemeColor(scheme, color.ThemeSlot)
            ?? color.Resolve(document?.EffectiveCreationTheme ?? OfficeTheme.Default);
        var native = new Color { Val = literal.ToRgbHex() };
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

    internal static OfficeColor? Read(Color? native) {
        if (native?.ThemeColor != null && native.ThemeTint == null && native.ThemeShade == null) {
            OfficeColor? role = native.ThemeColor.InnerText switch {
                "accent1" => OfficeColor.Accent1, "accent2" => OfficeColor.Accent2,
                "accent3" => OfficeColor.Accent3, "accent4" => OfficeColor.Accent4,
                "accent5" => OfficeColor.Accent5, "accent6" => OfficeColor.Accent6,
                "dark1" => OfficeColor.Text, "dark2" => OfficeColor.TextMuted,
                "light1" => OfficeColor.Background, _ => null
            };
            if (role.HasValue) return role;
        }
        return OfficeColor.TryParseHex(native?.Val?.Value, out var color) ? color : null;
    }
}
