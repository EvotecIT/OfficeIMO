using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfImageConversionTests {
    [Fact]
    public async Task ImageBridgeSelectsPagesAndSavesTheSameReportedEncoding() {
        RtfDocument document = RtfDocument.Parse(@"{\rtf1\ansi\paperw2880\paperh2880\margl144\margr144\margt144\margb144 First\page Second}");
        var options = new PdfImageExportOptions { Format = OfficeImageExportFormat.Png, PageSelector = OfficePageSelector.Parse("2,1,2"), Scale = .5 };
        var results = await document.ToImagesAsync(options);
        Assert.Equal(3, results.Count);
        Assert.All(results, image => {
            Assert.Equal(OfficeImageExportFormat.Png, image.Format);
            Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, image.Bytes.Take(8));
            Assert.InRange(image.Width, 70, 73);
        });
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-rtf-images-" + Guid.NewGuid().ToString("N"));
        try {
            OfficeConversionReport report = await document.SaveAsImagesAsync(folder, options);
            Assert.Equal("RTF", report.Source);
            Assert.Equal("Png", report.Target);
            Assert.Equal(3, report.ItemsWritten);
            Assert.Equal(3, Directory.GetFiles(folder, "*.png").Length);
            Assert.Equal("2,1,2", options.PageSelector.Expression);
            Assert.Equal(.5, options.Scale);
        } finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [Fact]
    public async Task ImageBridgeRetainsUpstreamLossAndRejectsBeforePublishingAnImage() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Visible text");
        document.AddImage(RtfImageFormat.Emf, new byte[] { 1, 2, 3 });
        OfficeImageExportResult result = Assert.Single(document.ToImages(new PdfImageExportOptions { Scale = .1 }));
        Assert.Contains(result.CreateReport().Diagnostics, finding => finding.Code == "unsupported-image" && finding.LossKind != OfficeConversionLossKind.None);
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-rtf-loss-" + Guid.NewGuid().ToString("N"));
        try {
            await Assert.ThrowsAnyAsync<OfficeConversionException>(() => document.SaveAsImagesAsync(folder, new PdfImageExportOptions { FailOnLoss = true, Scale = .1 }));
            Assert.Empty(Directory.Exists(folder) ? Directory.GetFiles(folder) : Array.Empty<string>());
        } finally { if (Directory.Exists(folder)) Directory.Delete(folder, true); }
    }

    [Fact]
    public async Task CancelledImageSaveDoesNotCreateItsDestination() {
        RtfDocument document = RtfDocument.Create();
        document.AddParagraph("Cancelled");
        string folder = Path.Combine(Path.GetTempPath(), "officeimo-rtf-cancelled-" + Guid.NewGuid().ToString("N"));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => document.SaveAsImagesAsync(folder, cancellationToken: cancellation.Token));
        Assert.False(Directory.Exists(folder));
    }
}
