using System.Threading.Tasks;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Visio.Pdf;
using Xunit;

namespace OfficeIMO.Visio.Tests;

public sealed class VisioDrawingPdfFailureEvidenceTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task StrictDiagramFailureRetainsLocatedSourceEvidenceAndLeavesDestinationUntouched(
        bool asynchronous, bool streamDestination) {
        VisioDocument document = VisioDocument.Create();
        document.AddPage("Page").AddRectangle(1, 1, 1, 1, "Label").AddHyperlink("https://example.com");
        byte[] source = document.ToLegacyXmlResult().Value;
        var options = new VisioToPdfOptions {
            Mode = VisioPdfProjectionMode.DiagramPages,
            SourceName = "source.vdx",
            DrawingOptions = new VisioDrawingOptions { RequireNoLoss = true }
        };
        byte[] sentinel = { 1, 2, 3 };
        string path = Path.Combine(Path.GetTempPath(), "visio-pdf-failure-" + Guid.NewGuid().ToString("N") + ".pdf");
        using var stream = new MemoryStream();
        stream.Write(sentinel, 0, sentinel.Length);
        long position = stream.Position;
        try {
            File.WriteAllBytes(path, sentinel);
            PdfSaveResult result = streamDestination
                ? asynchronous
                    ? await document.SaveAsPdfResultAsync(stream, options)
                    : document.SaveAsPdfResult(stream, options)
                : asynchronous
                    ? await document.SaveAsPdfResultAsync(path, options)
                    : document.SaveAsPdfResult(path, options);

            Assert.False(result.Succeeded);
            Assert.True(result.HasLoss);
            Assert.Equal(0, result.BytesWritten);
            var exception = Assert.IsType<OfficeConversionException>(result.Exception);
            var report = Assert.IsType<VisioDrawingConversionReport>(result.ConversionReports[0]);
            Assert.Same(exception.Report, report);
            Assert.Equal("source.vdx", report.SourceName);
            Assert.Contains(result.FidelityDiagnostics, diagnostic =>
                diagnostic.Code == "VISIO_DRAWING_METADATA" &&
                diagnostic.LossKind == OfficeConversionLossKind.Omission &&
                diagnostic.Location!.StartsWith("page:1:Page:shape:"));
            Assert.Equal(sentinel, stream.ToArray());
            Assert.Equal(position, stream.Position);
            Assert.Equal(sentinel, File.ReadAllBytes(path));
            Assert.Equal(source, document.ToLegacyXmlResult().Value);
        } finally {
            File.Delete(path);
        }
    }
}
