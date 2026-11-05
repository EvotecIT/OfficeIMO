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
                var palette = theme.Palette.With(accent1: brand.AccentColor,
                    accent2: brand.Accent2Color, accent3: brand.Accent3Color);
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
                OfficeColor color = accent.Value.Resolve(document.GetTheme());
                var brand = PowerPointDesignTheme.FromBrand(color.ToRgbHex());
                document.SetThemeColorForAllMasters(PowerPointThemeColor.Accent1, brand.AccentColor.ToRgbHex());
                document.SetThemeColorForAllMasters(PowerPointThemeColor.Accent2, brand.Accent2Color.ToRgbHex());
                document.SetThemeColorForAllMasters(PowerPointThemeColor.Accent3, brand.Accent3Color.ToRgbHex());
            }
            // Create(template) copies content as well as its native masters. An object export supplies its own slides.
            for (int index = document.Slides.Count - 1; index >= 0; index--) document.RemoveSlide(index);
            return document;
        } catch { document.Dispose(); throw; }
    }

}
