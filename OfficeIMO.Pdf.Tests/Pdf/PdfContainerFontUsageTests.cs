using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfContainerFontUsageTests {
    [Fact]
    public void MultiParagraphPanelRetainsGlyphsForEmbeddedCffFont() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        if (fontPath == null) return;

        byte[] pdf = PdfDocument.Create(new PdfOptions())
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath), "Panel CFF")
            .Panel(content => content
                .Paragraph(paragraph => paragraph.Text("PlainBox"))
                .Paragraph(paragraph => paragraph.Text("BulletBox"))
                .Paragraph(paragraph => paragraph.Text("UnmarkedBox")))
            .ToBytes();

        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("PlainBox", text, StringComparison.Ordinal);
        Assert.Contains("BulletBox", text, StringComparison.Ordinal);
        Assert.Contains("UnmarkedBox", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MultiParagraphPanelRetainsGlyphsIntroducedAfterFirstChild(bool namedFont) {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) return;

        byte[] fontData = File.ReadAllBytes(fontPath);
        var family = new PdfEmbeddedFontFamily("Panel Font", fontData);
        PdfOptions options = namedFont
            ? new PdfOptions().RegisterNamedFontFamily(family)
            : new PdfOptions().UseFontFamily(family);
        byte[] pdf = PdfDocument.Create(options)
            .Panel(content => content
                .Paragraph(paragraph => {
                    if (namedFont) paragraph.FontFamily("Panel Font");
                    paragraph.Text("PlainBox");
                })
                .Panel(nested => nested
                    .Paragraph(paragraph => {
                        if (namedFont) paragraph.FontFamily("Panel Font");
                        paragraph.Text("BulletBox");
                    })
                    .Paragraph(paragraph => {
                        if (namedFont) paragraph.FontFamily("Panel Font");
                        paragraph.Text("UnmarkedBox");
                    })))
            .ToBytes();

        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("PlainBox", text, StringComparison.Ordinal);
        Assert.Contains("BulletBox", text, StringComparison.Ordinal);
        Assert.Contains("UnmarkedBox", text, StringComparison.Ordinal);
    }
}
