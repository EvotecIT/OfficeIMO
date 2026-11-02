using System;
using OfficeIMO.PowerPoint;

namespace PSWriteOffice.Services.PowerPoint;

public static partial class PowerPointDocumentService {
    /// <summary>Translates PowerShell design choices to the native brief while retaining the document theme.</summary>
    internal static PowerPointDeckDesign CreateDesign(PowerPointDocument document, PowerPointDeckStyle style,
        string? purpose, PowerPointAutoLayoutStrategy? strategy, string? seed) {
        var theme = document.GetTheme();
        var brief = PowerPointDesignBrief.FromBrand(theme.Palette.Accent1.ToRgbHex(), seed ?? "document", purpose);
        if (style == PowerPointDeckStyle.Auto && purpose == null) brief.AddDirection(PowerPointDesignDirection.Structured);
        if (style != PowerPointDeckStyle.Auto) {
            brief.WithCreativeDirectionPack(style switch {
                PowerPointDeckStyle.FieldProof => PowerPointCreativeDirectionPack.FieldProof,
                PowerPointDeckStyle.EditorialCaseStudy => PowerPointCreativeDirectionPack.EditorialCaseStudy,
                PowerPointDeckStyle.TechnicalMap => PowerPointCreativeDirectionPack.TechnicalMap,
                PowerPointDeckStyle.QuietAppendix => PowerPointCreativeDirectionPack.QuietAppendix,
                _ => throw new ArgumentOutOfRangeException(nameof(style))
            });
        }
        if (strategy.HasValue) brief.WithLayoutStrategy(strategy.Value);
        brief.WithFonts(theme.Fonts.Heading.Family, theme.Fonts.Body.Family);
        var design = brief.CreateDesign();
        // The brief selects composition intent. Theme remains the user's palette and typeface choice.
        design.Theme.AccentColor = theme.Palette.Accent1.ToRgbHex();
        design.Theme.Accent2Color = theme.Palette.Accent2.ToRgbHex();
        design.Theme.Accent3Color = theme.Palette.Accent3.ToRgbHex();
        return design;
    }
}
