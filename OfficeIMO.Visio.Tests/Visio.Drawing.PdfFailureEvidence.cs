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
        document.AddPage("Page").AddRectangle(OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, VisioMeasurementUnit.Inches)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, VisioMeasurementUnit.Inches)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, VisioMeasurementUnit.Inches)), OfficeLength.Inches(VisioMeasurementUnitExtensions.ToInches(1, VisioMeasurementUnit.Inches)), "Label").AddHyperlink("https://example.com");
        byte[] source = document.ToLegacyXmlResult().Value;
        var options = new VisioToPdfOptions {
            Mode = VisioPdfProjectionMode.DiagramPages,
            SourceName = "source.vdx",
            DrawingOptions = new VisioDrawingOptions { FailOnLoss = true }
        };
        byte[] sentinel = { 1, 2, 3 };
        string path = Path.Combine(Path.GetTempPath(), "visio-pdf-failure-" + Guid.NewGuid().ToString("N") + ".pdf");
        using var stream = new MemoryStream();
        stream.Write(sentinel, 0, sentinel.Length);
        long position = stream.Position;
        try {
            File.WriteAllBytes(path, sentinel);
            OfficeConversionException exception = streamDestination
                ? asynchronous
                    ? await Assert.ThrowsAsync<OfficeConversionException>(() => document.SaveAsPdfAsync(stream, options))
                    : Assert.Throws<OfficeConversionException>(() => document.SaveAsPdf(stream, options))
                : asynchronous
                    ? await Assert.ThrowsAsync<OfficeConversionException>(() => document.SaveAsPdfAsync(path, options))
                    : Assert.Throws<OfficeConversionException>(() => document.SaveAsPdf(path, options));

            var report = Assert.IsType<VisioDrawingConversionReport>(exception.Report);
            Assert.Equal("source.vdx", report.SourceName);
            Assert.True(report.HasLoss);
            Assert.Contains(report.Diagnostics, diagnostic =>
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
