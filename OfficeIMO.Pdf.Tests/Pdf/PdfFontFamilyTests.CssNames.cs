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

    [Theory]
    [InlineData("comma-registered")]
    [InlineData("comma-unregistered")]
    [InlineData("semicolon-registered")]
    [InlineData("semicolon-unregistered")]
    public void NamedFontLiteralRequestsDoNotBorrowComponentSubstitutions(string scenario) {
        (string literal, bool registered) = scenario switch {
            "comma-registered" => ("Scoped, Literal", true),
            "comma-unregistered" => ("Scoped, Literal", false),
            "semicolon-registered" => ("Scoped; Literal", true),
            "semicolon-unregistered" => ("Scoped; Literal", false),
            _ => throw new ArgumentOutOfRangeException(nameof(scenario))
        };
        const string decoy = "Component Decoy";
        var options = new PdfCreateOptions { DefaultFont = PdfStandardFont.Courier, DefaultFontSize = OfficeLength.Points(12) }
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(decoy,
                ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 900, 'B')))
            .RegisterFontFamilySubstitution("Literal", decoy);
        if (registered) options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(literal,
            ManagedTextShapingTestAssets.CreateFontWithUniformAdvance(1000, 300, 'B')));

        byte[] bytes = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.FontFamily(OfficeFontFamilyParser.Quote(literal)).Text("BB"))
            .ToBytes();

        var page = Assert.Single(PdfReadDocument.Open(bytes).Pages);
        var spans = page.GetTextSpans().Where(span => !string.IsNullOrEmpty(span.Text)).ToArray();
        Assert.Equal("BB", string.Concat(spans.Select(span => span.Text)));
        var glyphs = ReadCallerCollisionGlyphs(bytes);
        Assert.Equal(2, glyphs.Count);
        Assert.All(glyphs, glyph => Assert.Equal(registered ? 3.6D : 7.2D, glyph.PaintedAdvance, 3));
        Assert.All(spans, span => {
            Assert.Equal(12D, span.FontSize, 3);
            if (registered) Assert.EndsWith("ScopedLiteral-Regular", span.BaseFont, StringComparison.OrdinalIgnoreCase);
            else Assert.Equal("Courier", span.BaseFont);
        });
    }
}
