using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class OcrReliabilityWorkflowTests {
    [Theory]
    [InlineData("terminal")]
    [InlineData("exception")]
    [InlineData("empty")]
    public async Task FailedRecognitionPreservesPdfAndImageDestinations(string failure) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-ocr-failure-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "output.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).Save(source);
            byte[] original = File.ReadAllBytes(source), previous = [1, 2, 3];
            File.WriteAllBytes(output, previous);
            var engine = new DelegateOcrEngine("failure", (_, _) => failure == "exception"
                ? throw new InvalidOperationException("Authorization=FAKE_AUDIT_SENTINEL")
                : Task.FromResult(failure == "empty" ? new OcrResult() : new OcrResult {
                    Text = "Do not publish", Spans = [Word("Do not publish")], Diagnostics = [
                        new OcrDiagnostic { Severity = OcrDiagnosticSeverity.Error, Code = "terminal", IsRecoverable = false }
                    ] }));
            var result = await new OfficeWorkflowRunner().MakePdfSearchableAsync(new() {
                InputPath = source, OutputPath = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            }, engine);
            Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
            Assert.Null(result.OutputPath);
            Assert.Equal(previous, File.ReadAllBytes(output));
            Assert.Equal(original, File.ReadAllBytes(source));
            Assert.DoesNotContain("FAKE_AUDIT_SENTINEL", System.Text.Json.JsonSerializer.Serialize(result));
            string image = Path.Combine(root, "source.png"), text = Path.Combine(root, "output.txt");
            File.WriteAllBytes(image, OcrSessionWorkflowTests.Png()); File.WriteAllBytes(text, previous);
            var imageResult = await new OfficeWorkflowRunner().RecognizeImageAsync(new() {
                InputPath = image, OutputPath = text, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            }, engine);
            Assert.Equal(OfficeWorkflowStatus.Failed, imageResult.Status);
            Assert.Equal(previous, File.ReadAllBytes(text));
            Assert.DoesNotContain("FAKE_AUDIT_SENTINEL", System.Text.Json.JsonSerializer.Serialize(imageResult));
            Assert.Empty(Directory.GetFiles(root, ".ocr-*.tmp"));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task CorrectionsAndRecoverableDiagnosticsSurviveSessionAndPdfReopen() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-ocr-correct-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf"), output = Path.Combine(root, "output.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).Save(source);
            byte[] original = File.ReadAllBytes(source);
            var engine = new DelegateOcrEngine("fixture", (_, _) => Task.FromResult(new OcrResult {
                Spans = [Word("Mispelled")], Diagnostics = [new OcrDiagnostic {
                    Code = "uncertain", Message = "Review uncertain words.", Severity = OcrDiagnosticSeverity.Warning, IsRecoverable = true
                }]
            }));
            var request = new PdfSearchableWorkflowRequest {
                InputPath = source, OutputPath = output,
                ReviewCorrectionsAsync = (review, _) => {
                    var word = Assert.Single(review.Ocr.Pages[0].Words);
                    var replacements = new Dictionary<PdfRecognizedWord, string> { [word] = "Corrected" };
                    Assert.Equal("Corrected", review.ExtractText(replacements));
                    Assert.Equal("Mispelled", word.Text);
                    Assert.Equal(OcrDiagnosticSeverity.Warning, Assert.Single(review.Ocr.Pages[0].ProviderDiagnostics).Severity);
                    return Task.FromResult<IReadOnlyDictionary<PdfRecognizedWord, string>>(replacements);
                }
            };
            var results = await new OfficeWorkflowRunner().RunOcrSessionAsync([new("correct", request)], engine);
            var result = Assert.Single(results);
            Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
            Assert.Equal("Corrected", PdfReadDocument.Open(File.ReadAllBytes(output)).ExtractText().Trim());
            Assert.Equal(original, File.ReadAllBytes(source));
            var diagnostic = Assert.Single(result.Diagnostics, item => item.Code == "uncertain");
            Assert.Equal(OfficeWorkflowDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("1", diagnostic.Details["page"]);
        } finally { Directory.Delete(root, true); }
    }

    private static OcrTextSpan Word(string text) => new() {
        Text = text, Level = OcrTextSpanLevel.Word, Confidence = 0.95, CoordinateUnit = OcrCoordinateUnit.Points,
        Region = new() { X = 20, Y = 30, Width = 80, Height = 12 }
    };
}
