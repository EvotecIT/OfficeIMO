using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlPdfLinkConformanceTests {
    [Theory]
    [InlineData("<a href='https://example.test/details'>Details</a>")]
    [InlineData("<svg width='80' height='40' role='img' aria-label='Green chart'><a href='https://example.test/details'><rect width='80' height='40' fill='green'/></a></svg>")]
    [InlineData("<a href='https://example.test/details'><math alttext='one half'><mfrac><mn>1</mn><mn>2</mn></mfrac></math></a>")]
    public void InvisibleLinkAreasRetainAccessibleAnnotationsWithoutUnnamedDrawingFailures(string body) {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fonts", "SourceSerif4-Regular.otf"));
        var options = new HtmlToPdfOptions {
            MathMlSourceModificationDate = new DateTimeOffset(2026, 10, 2, 0, 0, 0, TimeSpan.Zero),
            PdfOptions = new PdfOptions()
                .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3B)
                .ConfigurePdfUaGroundwork()
                .EmbedStandardFont(PdfStandardFont.Helvetica, font, "Source Serif")
                .RequireCompliance(PdfComplianceProfile.PdfUa1)
        };
        byte[] bytes = HtmlConversionDocument.Parse("<html lang='en'><title>Accessible links</title><body>" + body + "</body></html>")
            .ToPdfBytes(options);
        Assert.NotEmpty(PdfDocumentReadResult.Load(bytes).GetLinksByUri("https://example.test/details"));
        var tagged = PdfInspector.Inspect(bytes).TaggedContent!;
        Assert.Contains(tagged.StructureElements, element => element.StructureType == "Link" && element.ObjectReferenceCount > 0);
    }
}
