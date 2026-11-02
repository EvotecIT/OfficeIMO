using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Writes a literal fallback and retains native theme-slot identity for semantic colors.</summary>
    private void SetRunColor(OfficeColor color) {
        string? slot = color.ThemeSlot;
        var scheme = _document?._wordprocessingDocument.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        OfficeColor literal = OfficeOpenXmlThemeColorResolver.ResolveSchemeColor(scheme, slot)
            ?? color.Resolve(_document?.EffectiveCreationTheme ?? OfficeTheme.Default);
        ColorHex = literal.ToRgbHex();
        if (slot == null) return;
        // ColorHex owns run/SDT/hyperlink selection and clears any previous theme transforms.
        var native = ScopedRunProperties!.Color!;
        native.ThemeColor = color.Role switch {
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
        UpdatePageNumberFormatting(native);
    }
}
