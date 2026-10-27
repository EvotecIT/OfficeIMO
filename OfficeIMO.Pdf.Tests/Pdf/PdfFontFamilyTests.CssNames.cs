using System;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfFontFamilyTests {
    [Theory]
    [InlineData("Scoped, Literal", ",")]
    [InlineData("Scoped\\Literal\"Face", ",")]
    [InlineData("Scoped; Literal", ";")]
    public void NamedFontRequestsPreserveQuotedNamesAndLegacyFallbackSeparators(string family, string separator) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('B');
        var options = new PdfCreateOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family, font));
        string request = "Missing Family" + separator + OfficeFontFamilyParser.Quote(family);
        Assert.True(options.HasNamedFontFamily(request));

        byte[] bytes = PdfDocument.Create(options).Paragraph(paragraph => paragraph.FontFamily(request).Text("BB")).ToBytes();
        Assert.Equal(1, PdfFontInspector.Inspect(PdfReadDocument.Open(bytes)).EmbeddedFontCount);
        Assert.Contains("BB", PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        var plainOptions = new PdfCreateOptions().RegisterNamedFontFamily(new PdfEmbeddedFontFamily("PlainFixture", font));
        byte[] control = PdfDocument.Create(plainOptions).Paragraph(paragraph => paragraph.FontFamily("PlainFixture").Text("BB")).ToBytes();
        Assert.Equal(OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(control)).GetPixels(),
            OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(bytes)).GetPixels());
    }
}
