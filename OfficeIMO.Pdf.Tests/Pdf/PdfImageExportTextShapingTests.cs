using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfImageExportTextShapingTests {
    [Fact]
    public void LoadedPdfRasterExport_ReachesTheSharedTextShapingProvider() {
        byte[] fontData = ManagedTextShapingTestAssets.CreateFont(' ', 'A');
        var family = new PdfEmbeddedFontFamily(
            ManagedTextShapingTestAssets.FamilyName,
            fontData);
        byte[] pdf = PdfDocument.Create()
            .UseFontFamily(family)
            .Paragraph(paragraph => paragraph.Text("A"))
            .ToBytes();
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new PdfImageExportOptions {
            TextShapingProvider = provider,
            TextShapingLanguage = "ar-SA"
        };
        options.Fonts.Add(ManagedTextShapingTestAssets.FamilyName, fontData);

        OfficeImageExportResult result =
            document.Pages[0].ExportImage(OfficeImageExportFormat.Png, options);

        Assert.Equal(OfficeImageExportFormat.Png, result.Format);
        Assert.NotEmpty(provider.Requests);
        Assert.All(provider.Requests, request =>
            Assert.Equal("ar-SA", request.Language));
    }
}
