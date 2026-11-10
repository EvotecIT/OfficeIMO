namespace OfficeIMO.Pdf.Testing.Tests;

internal static class SampleReport {
    /// <summary>Pixel area of the optional badge at the default snapshot scale of 2, with a margin for antialiasing.</summary>
    public static PdfPixelRegion BadgeRegion { get; } = new(0, 0, 1224, 346);

    /// <summary>
    /// Builds the sample with the repository's baseline fonts, so every machine writes the same PDF. Without them
    /// (<paramref name="embedFonts"/> false) it uses the non-embedded standard fonts; <paramref name="font"/> replaces them.
    /// </summary>
    public static byte[] Create(string total, OfficeColor? badge = null, int pages = 1, bool embedFonts = true, PdfEmbeddedFontFamily? font = null) {
        var options = new PdfCreateOptions { EmbedDefaultFonts = false };
        if (embedFonts) options.RegisterFontFamily(PdfStandardFont.Helvetica, font ?? BaselineSans);
        return PdfDocument.Create(pdf => pdf.Content(content => {
            if (badge is OfficeColor fill) content.Rectangle(OfficeLength.Points(80), OfficeLength.Points(24), fillColor: fill);
            content.H1("Quarterly report")
                .AddTable(new[] {
                    new[] { "Item", "Amount" },
                    new[] { "Licences", "900.00" },
                    new[] { "Total", total }
                });
            for (int page = 2; page <= pages; page++) {
                content.PageBreak().AddParagraph(paragraph => paragraph.AddText("Appendix " + page));
            }
        }), options).ToBytes();
    }

    private static PdfEmbeddedFontFamily BaselineSans { get; } = PdfEmbeddedFontFamily.FromFiles("OfficeIMO Baseline Sans",
        Font("Regular"), Font("Bold"), Font("Italic"), Font("BoldItalic"));

    public static PdfEmbeddedFontFamily OpenSans { get; } = PdfEmbeddedFontFamily.FromFiles("Open Sans",
        Path.Combine(AppContext.BaseDirectory, "Fonts", "OpenSans-Bold.ttf"));

    private static string Font(string style) => Path.Combine(AppContext.BaseDirectory, "Fonts", "OfficeIMOBaselineSans-" + style + ".ttf");
}
