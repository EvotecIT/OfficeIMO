using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfContainerFontUsageTests {
    [Fact]
    public void MultiParagraphPanelRetainsGlyphsForEmbeddedCffFont() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        if (fontPath == null) return;

        byte[] pdf = PdfDocument.Create(new PdfCreateOptions())
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath), "Panel CFF")
            .Panel(content => content
                .AddParagraph(paragraph => paragraph.AddText("PlainBox"))
                .AddParagraph(paragraph => paragraph.AddText("BulletBox"))
                .AddParagraph(paragraph => paragraph.AddText("UnmarkedBox")))
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
        PdfCreateOptions options = namedFont
            ? new PdfCreateOptions().RegisterNamedFontFamily(family)
            : new PdfCreateOptions().UseFontFamily(family);
        byte[] pdf = PdfDocument.Create(options)
            .Panel(content => content
                .AddParagraph(paragraph => {
                    if (namedFont) { paragraph.RunDefaults.FontFamily = "Panel Font"; paragraph.RunDefaults.StandardFont = null; }
                    paragraph.AddText("PlainBox");
                })
                .Panel(nested => nested
                    .AddParagraph(paragraph => {
                        if (namedFont) { paragraph.RunDefaults.FontFamily = "Panel Font"; paragraph.RunDefaults.StandardFont = null; }
                        paragraph.AddText("BulletBox");
                    })
                    .AddParagraph(paragraph => {
                        if (namedFont) { paragraph.RunDefaults.FontFamily = "Panel Font"; paragraph.RunDefaults.StandardFont = null; }
                        paragraph.AddText("UnmarkedBox");
                    })))
            .ToBytes();

        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("PlainBox", text, StringComparison.Ordinal);
        Assert.Contains("BulletBox", text, StringComparison.Ordinal);
        Assert.Contains("UnmarkedBox", text, StringComparison.Ordinal);
    }
}
