using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRasterExport_ReachesTheSharedTextShapingProvider() {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(
            "<p style=\"font-family:'OfficeIMO Shaping Test'\">A</p>");
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new HtmlRenderOptions {
            ViewportWidth = 180D,
            TextShapingProvider = provider,
            TextShapingLanguage = "ar-SA"
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));

        OfficeImageExportResult result = document.ExportImage(OfficeImageExportFormat.Png, options);

        Assert.Equal(OfficeImageExportFormat.Png, result.Format);
        Assert.Contains(provider.Requests, request =>
            request.Text == "A" &&
            request.Language == "ar-SA");
    }
}
