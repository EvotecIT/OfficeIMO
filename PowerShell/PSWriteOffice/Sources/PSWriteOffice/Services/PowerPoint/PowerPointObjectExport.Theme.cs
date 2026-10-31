using System;
using OfficeIMO;
using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.PowerPoint;

internal static partial class PowerPointObjectExport {
    /// <summary>Creates an explicitly persisted native deck, preserving template masters and deriving brand colours with the native helper.</summary>
    internal static PowerPointDocument CreateDocument(OfficeThemeInput? input, OfficeColor? accent) {
        OfficeTheme theme = input?.Theme ?? OfficeTheme.Modern;
        if (input?.TemplatePath == null) {
            if (accent.HasValue) {
                var brand = PowerPointDesignTheme.FromBrand(accent.Value.Resolve(theme).ToRgbHex());
                var palette = theme.Palette.With(accent1: OfficeColor.Parse(brand.AccentColor),
                    accent2: OfficeColor.Parse(brand.Accent2Color), accent3: OfficeColor.Parse(brand.Accent3Color));
                var surfaces = OfficeThemeSurfaces.FromPalette(palette, theme.Surfaces.Surface, theme.Surfaces.Rule);
                theme = theme.WithPalette(palette).WithSurfaces(theme.Surfaces.With(accent1Tint: surfaces.Accent1Tint));
            }
            return PowerPointDocument.Create(new PowerPointCreateOptions {
                PersistenceMode = DocumentPersistenceMode.Explicit, Theme = theme
            });
        }

        string extension = System.IO.Path.GetExtension(input.TemplatePath);
        if (!extension.Equals(".pptx", StringComparison.OrdinalIgnoreCase) && !extension.Equals(".potx", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("Choose a .pptx or .potx PowerPoint theme template.", nameof(input));
        var document = PowerPointDocument.Create(new PowerPointCreateOptions {
            PersistenceMode = DocumentPersistenceMode.Explicit, Template = input.TemplatePath
        });
        try {
            if (accent.HasValue) {
                OfficeColor color = ResolveTemplateAccent(document, accent.Value, theme);
                var brand = PowerPointDesignTheme.FromBrand(color.ToRgbHex());
                document.SetThemeColorForAllMasters(PowerPointThemeColor.Accent1, brand.AccentColor);
                document.SetThemeColorForAllMasters(PowerPointThemeColor.Accent2, brand.Accent2Color);
                document.SetThemeColorForAllMasters(PowerPointThemeColor.Accent3, brand.Accent3Color);
            }
            // Create(template) copies content as well as its native masters. An object export supplies its own slides.
            for (int index = document.Slides.Count - 1; index >= 0; index--) document.RemoveSlide(index);
            return document;
        } catch { document.Dispose(); throw; }
    }

    private static OfficeColor ResolveTemplateAccent(PowerPointDocument document, OfficeColor color, OfficeTheme fallback) {
        PowerPointThemeColor? slot = color.Role switch {
            OfficeColorRole.Accent1 => PowerPointThemeColor.Accent1,
            OfficeColorRole.Accent2 => PowerPointThemeColor.Accent2,
            OfficeColorRole.Accent3 => PowerPointThemeColor.Accent3,
            OfficeColorRole.Accent4 => PowerPointThemeColor.Accent4,
            OfficeColorRole.Accent5 => PowerPointThemeColor.Accent5,
            OfficeColorRole.Accent6 => PowerPointThemeColor.Accent6,
            OfficeColorRole.Text => PowerPointThemeColor.Dark1,
            OfficeColorRole.TextMuted => PowerPointThemeColor.Dark2,
            OfficeColorRole.Background => PowerPointThemeColor.Light1,
            _ => null
        };
        string? hex = slot.HasValue ? document.GetThemeColor(slot.Value) : null;
        return hex == null ? color.Resolve(fallback) : OfficeColor.Parse(hex);
    }
}
