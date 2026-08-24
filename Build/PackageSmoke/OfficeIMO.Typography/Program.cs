using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Drawing.HarfBuzz;

internal static class Program {
    private static string AssetPath(string fileName) =>
        Path.Combine(AppContext.BaseDirectory, "TestAssets", fileName);

    public static void Main() {
        byte[] arabicFont = File.ReadAllBytes(AssetPath("NotoSansArabic-Regular.ttf"));
        OfficeTextShapingResult? arabic = OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(
            new OfficeTextShapingRequest(
                "مرحبا بالعالم",
                "Noto Sans Arabic",
                arabicFont,
                isOpenTypeCff: false,
                unitsPerEm: 1000,
                OfficeTextDirection.RightToLeft,
                "ar"));
        if (arabic == null || arabic.Glyphs.Count < 3 || arabic.Glyphs.Any(glyph => glyph.GlyphId <= 0)) {
            throw new InvalidOperationException("The packed HarfBuzz provider did not shape the Arabic smoke run.");
        }

        OfficeRenderingProfile harfBuzzProfile = OfficeHarfBuzzRenderingProfile.Create(language: "ar");
        if (!string.Equals(harfBuzzProfile.Name, "officeimo-harfbuzz", StringComparison.Ordinal)) {
            throw new InvalidOperationException("The packed HarfBuzz rendering profile is inconsistent.");
        }

#if NET8_0_OR_GREATER
        byte[] woff2 = File.ReadAllBytes(AssetPath("OpenSans-Regular.woff2"));
        var fonts = new OfficeFontFaceCollection()
            .Add("Open Sans", woff2)
            .AddFallbackFamily("Open Sans");
        OfficeFontFace face = fonts.Faces.Single();
        if (face.ContainerFormat != OfficeFontContainerFormat.Woff2 ||
            !face.CanEmbedAsStaticPdfFont ||
            !face.Program.HasGlyphs("OfficeIMO") ||
            face.Program.Measure("OfficeIMO", 12) <= 0 ||
            face.Program.GetTextContours("OfficeIMO", 0, 20, 12).Count == 0) {
            throw new InvalidOperationException("The packed first-party font engine did not decode and outline the WOFF2 smoke face.");
        }

        OfficeTextShapingResult? combined = OfficeHarfBuzzTextShapingProvider.Instance.ShapeText(
            new OfficeTextShapingRequest(
                "OfficeIMO",
                "Open Sans",
                face.Program.GetFontDataForShaping(),
                face.Program.IsOpenTypeCff,
                face.Program.UnitsPerEm,
                OfficeTextDirection.LeftToRight,
                "en"));
        if (combined == null || combined.Glyphs.Count == 0) {
            throw new InvalidOperationException("The packed HarfBuzz adapter and first-party font engine did not interoperate.");
        }

        byte[] variableData = File.ReadAllBytes(AssetPath("RobotoFlex.ttf"));
        var variableFonts = new OfficeFontFaceCollection {
            FontVariationResolver = _ => new System.Collections.Generic.Dictionary<string, float> { ["wght"] = 800F }
        }.Add("Roboto Flex", variableData);
        OfficeFontFace variableFace = variableFonts.Faces.Single();
        if (variableFace.CanEmbedAsStaticPdfFont || variableFace.Program.GetTextContours("Variable", 0, 20, 12).Count == 0) {
            throw new InvalidOperationException("The packed first-party font engine did not evaluate the TrueType variable face.");
        }

        var pack = new OfficeFontFallbackPack("package-smoke", "Open Sans", fonts);
        if (pack.Fingerprint.Length != 64) {
            throw new InvalidOperationException("The packed fallback-pack fingerprint is invalid.");
        }
#endif

        Console.WriteLine("OfficeIMO typography packed API smoke passed on " +
            System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription + ".");
    }
}
