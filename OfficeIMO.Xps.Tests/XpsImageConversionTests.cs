using System;
using System.IO;
using System.Linq;
using Xunit;
using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsImageConversionTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public async Task OrderedPageImagesAndAsyncSavesUseSharedOptionsAndReportMetadata(OfficeImageExportFormat format) {
        XpsDocument document = XpsDocument.Create();
        document.AddPage(96, 48).AddPath("M0,0 H96 V48 H0 Z", "#FFFF0000");
        document.AddPage(48, 96).AddPath("M0,0 H48 V96 H0 Z", "#FF0000FF");
        var options = new OfficeImageExportOptions { Format = format, PageSelector = OfficePageSelector.Parse("2,1,2"), TargetDpi = 192 };
        var images = await document.ToImagesAsync(options);
        Assert.Equal(new[] { (96, 192), (192, 96), (96, 192) }, images.Select(image => (image.Width, image.Height)));
        Assert.All(images, image => Assert.Equal("XPS", image.Source));
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-xps-images-" + Guid.NewGuid().ToString("N"));
        try {
            OfficeConversionReport report = await document.SaveAsImagesAsync(folder, options);
            Assert.Equal("XPS", report.Source); Assert.Equal(format.ToString(), report.Target); Assert.Equal(3, report.ItemsWritten);
            Assert.Equal(3, Directory.GetFiles(folder, "*" + format.GetFileExtension()).Length);
            Assert.Equal(1D, options.Scale);
        } finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }
}
