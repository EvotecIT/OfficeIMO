using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfImageExportTextShapingTests {
    [Fact]
    public void UnembeddedPdfRasterExport_UsesSharedFallbackShapingProvider() {
        byte[] fontData = ManagedTextShapingTestAssets.CreateFont(' ', 'A');
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica
            })
            .Paragraph(paragraph => paragraph.Text("A"))
            .ToBytes();
        PdfFontInfo font = Assert.Single(PdfDocument.Load(pdf).Resources.Fonts().Fonts);
        Assert.Equal("Helvetica", font.BaseFontName);
        Assert.False(font.IsEmbedded);
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new PdfImageExportOptions {
            TextShapingProvider = provider,
            TextShapingLanguage = "ar-SA"
        };
        options.Fonts.Add("Helvetica", fontData).AddFallbackFamily("Helvetica");

        OfficeImageExportResult result =
            document.Pages[0].ExportImage(OfficeImageExportFormat.Png, options);

        Assert.Equal(OfficeImageExportFormat.Png, result.Format);
        Assert.NotEmpty(result.Bytes);
        Assert.True(result.Width > 0);
        Assert.True(result.Height > 0);
        Assert.NotEmpty(provider.Requests);
        Assert.All(provider.Requests, request => {
            Assert.Equal("A", request.Text);
            Assert.Equal("ar-SA", request.Language);
            Assert.Equal(fontData, request.FontData);
        });
    }
}
