using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OdgImageConversionTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public async Task DirectImagesUsePointDimensionsAndOrderedPageSelection(OfficeImageExportFormat format) {
        using var document = OdgDocument.Create();
        document.AddPage("Landscape", OdfLength.Points(72), OdfLength.Points(36));
        document.AddPage("Portrait", OdfLength.Points(36), OdfLength.Points(72));
        var options = new OfficeImageExportOptions { Format = format, TargetDpi = 144, PageSelector = OfficePageSelector.Parse("2,1,2") };
        var images = await document.ToImagesAsync(options);
        Assert.Equal(new[] { (72, 144), (144, 72), (72, 144) }, images.Select(image => (image.Width, image.Height)));
        Assert.All(images, image => Assert.Equal(format, image.Format));
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-odg-images-" + Guid.NewGuid().ToString("N"));
        try {
            OfficeConversionReport report = document.SaveAsImages(folder, options);
            Assert.Equal("ODG", report.Source); Assert.Equal(format.ToString(), report.Target); Assert.Equal(3, report.ItemsWritten);
            Assert.Equal(3, Directory.GetFiles(folder, "*" + format.GetFileExtension()).Length);
        } finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }
}
