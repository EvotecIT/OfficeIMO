using OfficeIMO.Ocr;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class ImageOcrWorkflowTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProviderTextIsVerifiedOrRetainedForRecovery(bool failWrite) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-image-ocr-provider-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            byte[] image = OcrSessionWorkflowTests.Png();
            byte[] stored = [];
            var recovery = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root, "recovery"));
            var request = new ImageOcrWorkflowRequest {
                InputPath = "content://scans/input", InputStream = new("scan.png", _ => Task.FromResult<Stream>(new MemoryStream(image))),
                OutputPath = "content://scans/output", ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                OutputStream = new("reviewed.txt", _ => Task.FromResult<Stream>(new MemoryStream(stored)), _ => {
                    if (failWrite) throw new IOException("Provider write is unavailable.");
                    return Task.FromResult<Stream>(new TextDestination(bytes => stored = bytes));
                }, recovery),
                ReviewAsync = (_, _) => Task.FromResult("Reviewed zażółć 🚀")
            };
            var result = await new OfficeWorkflowRunner().RecognizeImageAsync(request,
                new DelegateOcrEngine("fixture", (_, _) => Task.FromResult(new OcrResult { Text = "Recognized" })));
            Assert.Equal(failWrite ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Completed, result.Status);
            if (failWrite) {
                Assert.Null(result.OutputPath);
                var retained = Assert.Single(recovery.GetRecoveries());
                await recovery.VerifyAsync(retained);
                Assert.Equal("Reviewed zażółć 🚀", File.ReadAllText(retained.FilePath));
                recovery.Discard(retained);
            } else {
                Assert.Equal(request.OutputPath, result.OutputPath);
                Assert.Equal("Reviewed zażółć 🚀", System.Text.Encoding.UTF8.GetString(stored));
                Assert.Empty(recovery.GetRecoveries());
            }
        } finally { Directory.Delete(root, recursive: true); }
    }

    private sealed class TextDestination(Action<byte[]> commit) : MemoryStream {
        private bool _closed;
        protected override void Dispose(bool disposing) {
            if (disposing && !_closed) { _closed = true; commit(ToArray()); }
            base.Dispose(disposing);
        }
    }

    [Theory]
    [InlineData("commit")]
    [InlineData("source-changed")]
    [InlineData("cancel")]
    [InlineData("too-large")]
    public async Task ReviewedTextIsPublishedOnlyAfterSourceAndOutputChecks(string action) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-image-ocr-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "scan.png");
            string output = Path.Combine(root, "recognized.txt");
            byte[] image = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
            File.WriteAllBytes(source, image);
            File.WriteAllText(output, "Previous output");
            using var cancellation = new CancellationTokenSource();
            var entered = new TaskCompletionSource<ImageOcrWorkflowReview>(TaskCreationOptions.RunContinuationsAsynchronously);
            var decision = new TaskCompletionSource<string>(TaskCreationOptions.RunContinuationsAsynchronously);
            var request = new ImageOcrWorkflowRequest {
                InputPath = source, OutputPath = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                Limits = new() { MaximumOutputBytes = action == "too-large" ? 1 : 1024 },
                ReviewAsync = (review, _) => { entered.SetResult(review); return decision.Task; }
            };
            var engine = new DelegateOcrEngine("image-fixture", (_, _) => Task.FromResult(new OcrResult {
                Text = "Recognized text", Provider = "fixture", Confidence = 0.95
            }));
            var running = new OfficeWorkflowRunner().RecognizeImageAsync(request, engine, cancellation.Token);
            var completed = await Task.WhenAny(entered.Task, running).WaitAsync(TimeSpan.FromSeconds(30));
            Assert.True(completed == entered.Task, running.IsCompleted ? (await running).Summary : "Review was not reached.");
            var review = await entered.Task;
            Assert.Equal(image, review.GetImageBytes());
            byte[] preview = review.GetImageBytes();
            preview[0] = 0;
            Assert.Equal(image, review.GetImageBytes());
            Assert.Equal("Recognized text", review.Text);
            Assert.Equal(1, review.Recognition.Report.RecognizedCandidateCount);
            Assert.Equal("Previous output", File.ReadAllText(output));
            if (action == "source-changed") File.WriteAllBytes(source, [1, 2, 3]);
            if (action == "cancel") cancellation.Cancel();
            else decision.SetResult("Reviewed zażółć 🚀");
            var result = await running.WaitAsync(TimeSpan.FromSeconds(30));
            if (action == "commit") {
                Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
                Assert.Equal("Reviewed zażółć 🚀", File.ReadAllText(output));
                Assert.Equal("Reviewed zażółć 🚀".Length, result.CharacterCount);
                Assert.Equal(image, File.ReadAllBytes(source));
            } else {
                Assert.Equal(action == "cancel" ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed, result.Status);
                Assert.Null(result.OutputPath);
                Assert.Equal("Previous output", File.ReadAllText(output));
            }
            decision.TrySetResult(string.Empty);
            Assert.Empty(Directory.GetFiles(root, ".image-ocr-*.tmp"));
        } finally { Directory.Delete(root, recursive: true); }
    }
}
