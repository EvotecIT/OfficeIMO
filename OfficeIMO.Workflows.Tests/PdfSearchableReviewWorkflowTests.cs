using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfSearchableReviewWorkflowTests {
    [Theory]
    [InlineData("commit")]
    [InlineData("source-changed")]
    [InlineData("cancel")]
    public async Task ReviewPausesPublicationAndRechecksTheSource(string action) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-ocr-review-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            string output = Path.Combine(root, "output.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).Save(source);
            byte[] previous = [1, 2, 3];
            File.WriteAllBytes(output, previous);
            using var cancellation = new CancellationTokenSource();
            var entered = new TaskCompletionSource<PdfSearchableOcrReview>(TaskCreationOptions.RunContinuationsAsynchronously);
            var decision = new TaskCompletionSource<IReadOnlyList<PdfRecognizedWord>>(TaskCreationOptions.RunContinuationsAsynchronously);
            var request = new PdfSearchableWorkflowRequest {
                InputPath = source, OutputPath = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                ReviewAsync = (review, _) => { entered.SetResult(review); return decision.Task; }
            };
            var engine = new DelegateOcrEngine("fixture", (_, _) => Task.FromResult(new OcrResult { Spans = [
                new OcrTextSpan { Text = "Chosen", Level = OcrTextSpanLevel.Word, Confidence = 1,
                    CoordinateUnit = OcrCoordinateUnit.Points, Region = new OcrRegion { X = 20, Y = 30, Width = 80, Height = 12 } }
            ] }));
            var running = new OfficeWorkflowRunner().MakePdfSearchableAsync(request, engine, cancellation.Token);
            var review = await entered.Task.WaitAsync(TimeSpan.FromSeconds(30));
            Assert.False(running.IsCompleted);
            Assert.Equal(previous, File.ReadAllBytes(output));
            if (action == "source-changed") File.WriteAllBytes(source, [9, 8, 7]);
            if (action == "cancel") cancellation.Cancel();
            else decision.SetResult(review.Ocr.Pages[0].Words);
            var result = await running.WaitAsync(TimeSpan.FromSeconds(30));
            if (action == "commit") {
                Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
                Assert.Equal(1, result.AddedWordCount);
                Assert.Equal("Chosen", PdfReadDocument.Open(File.ReadAllBytes(output)).Pages.Single().ExtractText().Trim());
            } else {
                Assert.Equal(action == "cancel" ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed, result.Status);
                Assert.Null(result.OutputPath);
                Assert.Equal(previous, File.ReadAllBytes(output));
            }
            decision.TrySetResult([]);
            Assert.Empty(Directory.GetFiles(root, ".ocr-*.tmp"));
        } finally { Directory.Delete(root, recursive: true); }
    }
}
