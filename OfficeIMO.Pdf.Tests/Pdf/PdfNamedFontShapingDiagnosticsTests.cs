using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfNamedFontShapingDiagnosticsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NamedFontPreflightDefersShapingWarningsUntilTheProviderAttemptsTheRun(bool declines) {
        const string text = "ܫܠܡ";
        const string family = "Named Syriac";
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(text.Select(character => (int)character).ToArray());
        IOfficeTextShapingProvider provider = declines
            ? new DecliningProvider()
            : new ManagedTextShapingTestAssets.RecordingProvider();
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions()
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily(family, font))
            .SetTextShapingProvider(provider)
            .ReportDiagnosticsTo(report);
        PdfDocument document = PdfDocument.Create(options).Page(page => page.Canvas(canvas =>
            canvas.Text(new[] { new PdfTextRun(text, fontFamily: family) }, 8D, 8D, 100D, 40D)));

        Assert.Empty(document.AnalyzeTextEncoding());
        Assert.Empty(report.Warnings);
        byte[] bytes = document.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });

        Assert.Contains(text, PdfReadDocument.Open(bytes).ExtractText(), StringComparison.Ordinal);
        if (declines) {
            Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-complex-script-shaping");
            Assert.Contains(report.Warnings, warning => warning.Code == "unsupported-bidirectional-text-layout");
        } else {
            report.RequireNoLoss();
        }
    }

    private sealed class DecliningProvider : IOfficeTextShapingProvider {
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) => null;
    }
}
