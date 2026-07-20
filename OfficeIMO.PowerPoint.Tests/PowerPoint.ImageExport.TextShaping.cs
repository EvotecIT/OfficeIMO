using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public partial class PowerPointImageExportTests {
    [Fact]
    public void PowerPointRasterExport_ReachesTheSharedTextShapingProvider() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240D, 160D);
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox textBox = slide.AddTextBoxPoints("A", 20D, 20D, 100D, 30D);
        textBox.Paragraphs[0].SetFontName(ManagedTextShapingTestAssets.FamilyName);
        var provider = new ManagedTextShapingTestAssets.RecordingProvider();
        var options = new PowerPointImageExportOptions {
            TextShapingProvider = provider,
            TextShapingLanguage = "ar-SA"
        };
        options.Fonts.Add(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFont('A'));

        OfficeImageExportResult result = slide.ExportImage(OfficeImageExportFormat.Png, options);

        Assert.Equal(OfficeImageExportFormat.Png, result.Format);
        Assert.Contains(provider.Requests, request =>
            request.Text == "A" &&
            request.Language == "ar-SA");
    }
}
