using System.Threading.Tasks;
using OfficeIMO.Drawing;

namespace OfficeIMO.Publisher.Tests;

public sealed class PublisherImageConversionTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public async Task SharedImageOptionsUsePointCoordinatesAndPreserveSourceDiagnostics(OfficeImageExportFormat format) {
        PublisherDocument document = PublisherDocument.Create(OfficeLength.Points(72), OfficeLength.Points(36));
        var options = new OfficeImageExportOptions { TargetDpi = 144, Format = format };
        OfficeImageExportResult image = Assert.Single(await document.ToImagesAsync(options));
        Assert.Equal((144, 72), (image.Width, image.Height));
        Assert.Equal("Publisher", image.Source);
        Assert.Equal(document.ReadReport.Diagnostics.Select(diagnostic => diagnostic.Code), image.CreateReport().Diagnostics.Select(diagnostic => diagnostic.Code));
        Assert.Equal(1D, options.Scale);
    }
}
