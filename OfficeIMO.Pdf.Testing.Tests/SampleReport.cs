namespace OfficeIMO.Pdf.Testing.Tests;

internal static class SampleReport {
    /// <summary>Pixel area of the optional badge at the default 1.5 snapshot scale, with a margin for antialiasing.</summary>
    public static PdfPixelRegion BadgeRegion { get; } = new(0, 0, 918, 260);

    public static byte[] Create(string total, OfficeColor? badge = null, int pages = 1, bool embedFonts = true) =>
        PdfDocument.Create(pdf => pdf.Content(content => {
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
        }), new PdfCreateOptions { EmbedDefaultFonts = embedFonts }).ToBytes();
}
