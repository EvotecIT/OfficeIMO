using global::ChartForgeX.VisualArtifacts;
using OfficeIMO.ChartForgeX;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.ChartForgeX.Tests;

// A visual placed in an Office document stays vector when the document converts to PDF: its text remains
// text, and no image XObject stands in for the drawing.
public sealed class OfficeVisualPdfConversionTests {
    [Theory]
    [InlineData("Word")]
    [InlineData("PowerPoint")]
    public void VisualPlacedInOfficeDocumentConvertsToVectorPdf(string format) {
        VisualArtifact artifact = TableArtifact.Create("service-health")
            .WithTitle("Service Health")
            .AddColumn("service", "Service")
            .AddColumn("state", "State", TableArtifactColumnType.Status)
            .AddRow("api", "API", "Healthy")
            .AddRow("worker", "Worker", "Warning")
            .ToVisualArtifact();
        OfficeConversionReport report;
        byte[] pdf;
        if (format == "Word") {
            using WordDocument document = WordDocument.Create();
            document.AddParagraph().AddVisualArtifact(artifact);
            pdf = document.ToPdf(null, out report).ToBytes();
        } else {
            using PowerPointDocument presentation = PowerPointDocument.Create();
            presentation.AddSlide().AddVisualArtifact(artifact, 36D, 36D);
            pdf = presentation.ToPdf(null, out report).ToBytes(new PdfSaveOptions { AllowLoss = true });
        }

        using PdfDocument read = PdfDocument.Load(new MemoryStream(pdf));
        Assert.Empty(read.Images.Extract());
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.SvgImageRasterFallback);
        string text = read.ExtractText();
        Assert.Contains("Service Health", text);
        Assert.Contains("Worker", text);
    }
}
