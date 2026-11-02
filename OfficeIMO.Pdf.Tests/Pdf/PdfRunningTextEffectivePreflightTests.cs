using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfRunningTextEffectivePreflightTests {
    [Theory]
    [InlineData("Header")]
    [InlineData("Footer")]
    public void CustomRunningTextSegmentsPreflightAndRenderWithoutSupersededFormatGlyphs(string owner) {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'x');
        var options = new PdfCreateOptions { HeaderFormat = "Page {page}/{pages}", CompressContentStreams = false }
            .EmbedStandardFont(PdfStandardFont.Helvetica, font, "Selected running text");
        PdfDocument document = PdfDocument.Create(options).Compose(composition => composition.Page(page => {
            page.Content(content => content.Text("x"));
            if (owner == "Header") page.Header(header => header.Text(text => text.Run(new PdfTextRun("xxx"))));
            else page.Footer(footer => footer.Text(text => text.Run(new PdfTextRun("xxx"))));
        }));

        Assert.Empty(document.AnalyzeTextEncoding());
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal("xxxx", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)
            .Where(value => !string.IsNullOrWhiteSpace(value))));
    }
}
