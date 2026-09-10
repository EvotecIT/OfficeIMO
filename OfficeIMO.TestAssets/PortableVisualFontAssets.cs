using System;
using System.IO;
using OfficeIMO.Drawing;

namespace OfficeIMO.TestAssets;

internal static class PortableVisualFontAssets {
    internal static OfficeFontFaceCollection CreateSpreadsheetFonts() {
        var fonts = new OfficeFontFaceCollection();
        foreach (var face in new[] {
            ("Regular", OfficeFontStyle.Regular), ("Bold", OfficeFontStyle.Bold),
            ("Italic", OfficeFontStyle.Italic), ("BoldItalic", OfficeFontStyle.Bold | OfficeFontStyle.Italic)
        }) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts",
                "OfficeIMOBaselineSans-" + face.Item1 + ".ttf"));
            fonts.Add("Calibri", bytes, face.Item2);
        }
        return fonts.AddAlias("Arial", "Calibri").AddAlias("Aptos", "Calibri");
    }
}
