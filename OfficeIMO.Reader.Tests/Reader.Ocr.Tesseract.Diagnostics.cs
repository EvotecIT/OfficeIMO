using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Process;
using OfficeIMO.Ocr.Tesseract;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class TesseractDiagnosticsTests {
    [Theory]
    [InlineData("Estimating resolution as 447\n", false, OcrDiagnosticSeverity.Info)]
    [InlineData("Estimating resolution as 447\r\nEstimating resolution as 300\r\n", false, OcrDiagnosticSeverity.Info)]
    [InlineData("Warning: Invalid resolution 0 dpi. Using 70 instead.\n", false, OcrDiagnosticSeverity.Warning)]
    [InlineData("Estimating resolution as 447\nunknown provider warning\n", false, OcrDiagnosticSeverity.Warning)]
    [InlineData("Estimating resolution as 447", true, OcrDiagnosticSeverity.Warning)]
    public void OnlyCompleteKnownInformationalStderrCanPassReviewPolicy(string message, bool truncated, OcrDiagnosticSeverity severity) {
        OcrDiagnostic diagnostic = new TesseractOcrEngine().CreateStandardErrorDiagnostic(new OcrProcessResult {
            StandardError = message, StandardErrorTruncated = truncated
        }, "tesseract-stderr")!;
        Assert.Equal(severity, diagnostic.Severity);
        Assert.Equal(message, diagnostic.Message);
        Assert.Equal(truncated ? "true" : "false", diagnostic.Attributes["truncated"]);
        var assessment = new OcrReviewPolicy().Assess(new OcrResult {
            Text = "Invoice", Spans = new[] { new OcrTextSpan { Level = OcrTextSpanLevel.Word, Text = "Invoice", Confidence = 0.99 } },
            Diagnostics = new[] { diagnostic }
        });
        Assert.Equal(severity == OcrDiagnosticSeverity.Info, assessment.MeetsThresholds);
    }
}
