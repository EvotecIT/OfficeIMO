using System;
using System.Globalization;

namespace OfficeIMO.Drawing;

internal enum OfficeStandardFontFace {
    Helvetica,
    HelveticaBold,
    TimesRoman,
    TimesBold,
    TimesItalic,
    TimesBoldItalic
}

internal static partial class OfficeStandardFontMetrics {
    /// <summary>Measures a conservative advance across standard sans, serif and monospace fallbacks.</summary>
    internal static double MeasureFallbackWidth(string text, double points, OfficeFontStyle style) {
        bool bold = (style & OfficeFontStyle.Bold) != 0;
        bool italic = (style & OfficeFontStyle.Italic) != 0;
        OfficeStandardFontFace sans = bold ? OfficeStandardFontFace.HelveticaBold : OfficeStandardFontFace.Helvetica;
        OfficeStandardFontFace serif = bold
            ? italic ? OfficeStandardFontFace.TimesBoldItalic : OfficeStandardFontFace.TimesBold
            : italic ? OfficeStandardFontFace.TimesItalic : OfficeStandardFontFace.TimesRoman;
        double sansWidth = 0D;
        double serifWidth = 0D;
        double monoWidth = 0D;
        var fallback = new OfficeFontInfo("Arial", points, style);
        foreach (char character in text) {
            // WinAnsi paints a soft hyphen as the ordinary hyphen glyph.
            char measured = character == '\u00AD' ? '-' : character;
            UnicodeCategory category = char.GetUnicodeCategory(measured);
            if (category is UnicodeCategory.Control or UnicodeCategory.Format or UnicodeCategory.NonSpacingMark
                    or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark) {
                continue;
            }
            double estimate = OfficeTextMeasurer.GetCharacterWidthFactor(measured, fallback) * 1000D;
            sansWidth += TryGetWidth1000(sans, measured, out int sansAdvance) ? sansAdvance : estimate;
            serifWidth += TryGetWidth1000(serif, measured, out int serifAdvance) ? serifAdvance : estimate;
            monoWidth += 600D;
        }
        return Math.Max(monoWidth, Math.Max(sansWidth, serifWidth)) * points / 1000D;
    }
}
