using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class VisioImageExportTextShapingTests {
    [Fact]
    public void VisioRasterExport_ReachesTheSharedTextShapingProvider() {
        using var package = new MemoryStream();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Shaping").Size(2D, 1D);
        VisioShape shape = page.AddRectangle(1D, 0.5D, 1.2D, 0.5D, "A");
        shape.TextStyle = new VisioTextStyle {
            FontFamily = ManagedTextShapingTestAssets.FamilyName
        };
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new VisioImageExportOptions {
            Scale = 0.5D,
            Supersampling = 1,
            TextShapingProvider = provider,
            TextShapingLanguage = "ar-SA"
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));

        OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Png, options);

        Assert.Equal(OfficeImageExportFormat.Png, result.Format);
        Assert.Contains(provider.Requests, request =>
            request.Text == "A" &&
            request.Language == "ar-SA");
    }

    [Fact]
    public void RetainedVisioPngExport_ReachesTheSharedTextShapingProvider() {
        using var package = new MemoryStream();
        VisioDocument document = VisioDocument.Create(package);
        VisioPage page = document.AddPage("Retained shaping").Size(2D, 1D);
        VisioShape shape = page.AddRectangle(1D, 0.5D, 1.2D, 0.5D, "A");
        shape.TextStyle = new VisioTextStyle {
            FontFamily = ManagedTextShapingTestAssets.FamilyName
        };
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new VisioPngSaveOptions {
            PixelsPerInch = 48D,
            Supersampling = 1,
            TextShapingProvider = provider,
            TextShapingLanguage = "ar-SA"
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));

        byte[] png = page.ToPng(options);

        Assert.NotEmpty(png);
        Assert.Contains(provider.Requests, request =>
            request.Text == "A" &&
            request.Language == "ar-SA");
    }
}
