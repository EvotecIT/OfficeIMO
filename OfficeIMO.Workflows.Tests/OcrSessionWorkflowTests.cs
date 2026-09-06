using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class OcrSessionWorkflowTests {
    [Fact]
    public async Task DeferredProviderDestinationsCannotOverwriteAnEarlierSessionOutput() {
        string root = NewDirectory();
        try {
            string source = Path.Combine(root, "scan.png");
            File.WriteAllBytes(source, Png());
            byte[] stored = [];
            int writes = 0;
            var recovery = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root, "recovery"));
            OfficeOcrSessionRequest Request(int index) => new(index.ToString(), new ImageOcrWorkflowRequest {
                InputPath = source, OutputPath = "content://folder/planned-" + index, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                OutputStream = new(index + ".txt", _ => Task.FromResult<Stream>(new MemoryStream(stored)), _ => {
                    writes++;
                    return Task.FromResult<Stream>(new Destination(bytes => stored = bytes));
                }, recovery, _ => Task.FromResult("content://folder/same-actual-file")),
                ReviewAsync = (_, _) => Task.FromResult("Reviewed " + index)
            });
            var results = await new OfficeWorkflowRunner().RunOcrSessionAsync([Request(0), Request(1), Request(2)], Engine());
            Assert.Equal(OfficeWorkflowStatus.Completed, results[0].Status);
            Assert.Equal(OfficeWorkflowStatus.Unconfirmed, results[1].Status);
            Assert.Equal(OfficeWorkflowStatus.Cancelled, results[2].Status);
            Assert.Equal(1, writes);
            Assert.Equal("Reviewed 0", System.Text.Encoding.UTF8.GetString(stored));
            var retained = Assert.Single(recovery.GetRecoveries());
            Assert.Equal("Reviewed 1", File.ReadAllText(retained.FilePath));
            var retry = await new OfficeWorkflowRunner().RunOcrSessionAsync([Request(2)], Engine(),
                protectedOutputPaths: [results[0].OutputPath!]);
            Assert.Equal(OfficeWorkflowStatus.Unconfirmed, Assert.Single(retry).Status);
            Assert.Equal(1, writes);
            Assert.Equal("Reviewed 0", System.Text.Encoding.UTF8.GetString(stored));
        } finally { Directory.Delete(root, recursive: true); }
    }
    private sealed class Destination(Action<byte[]> commit) : MemoryStream {
        private bool _closed;
        protected override void Dispose(bool disposing) {
            if (disposing && !_closed) { _closed = true; commit(ToArray()); }
            base.Dispose(disposing);
        }
    }

    [Fact]
    public async Task CancellationRetainsCompletedTextAndMarksUnstartedFiles() {
        string root = NewDirectory();
        try {
            string image = Path.Combine(root, "scan.png");
            File.WriteAllBytes(image, Png());
            using var cancellation = new CancellationTokenSource();
            var requests = Enumerable.Range(0, 3).Select(index => new OfficeOcrSessionRequest(index.ToString(), new ImageOcrWorkflowRequest {
                InputPath = image, OutputPath = Path.Combine(root, index + ".txt"),
                ReviewAsync = (_, _) => {
                    if (index == 1) cancellation.Cancel();
                    return Task.FromResult("Reviewed " + index);
                }
            })).ToArray();
            var results = await new OfficeWorkflowRunner().RunOcrSessionAsync(requests, Engine(), cancellationToken: cancellation.Token);
            Assert.Equal(new[] { OfficeWorkflowStatus.Completed, OfficeWorkflowStatus.Cancelled, OfficeWorkflowStatus.Cancelled }, results.Select(item => item.Status));
            Assert.Equal("Reviewed 0", File.ReadAllText(Path.Combine(root, "0.txt")));
            Assert.False(File.Exists(Path.Combine(root, "1.txt")));
            Assert.False(File.Exists(Path.Combine(root, "2.txt")));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task MixedSessionSnapshotsFutureRequestsBeforeReviewAndContinuesAfterFailure() {
        string root = NewDirectory();
        try {
            string image = Path.Combine(root, "scan.png");
            File.WriteAllBytes(image, Png());
            string pdf = Path.Combine(root, "scan.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).Save(pdf);
            string output = Path.Combine(root, "searchable.pdf");
            var future = new PdfSearchableWorkflowRequest { InputPath = pdf, OutputPath = output };
            var first = new ImageOcrWorkflowRequest {
                InputPath = image, OutputPath = Path.Combine(root, "text.txt"),
                ReviewAsync = (_, _) => { future.OutputPath = pdf; future.Ocr.Language = "changed"; return Task.FromResult("Chosen text"); }
            };
            var results = await new OfficeWorkflowRunner().RunOcrSessionAsync([
                new("image", first), new("missing", new ImageOcrWorkflowRequest {
                    InputPath = Path.Combine(root, "missing.png"), OutputPath = Path.Combine(root, "missing.txt"),
                    InputStream = new("missing.png", _ => Task.FromResult<Stream>(File.OpenRead(Path.Combine(root, "missing.png"))))
                }),
                new("pdf", future)
            ], Engine());
            Assert.Equal(new[] { OfficeWorkflowStatus.Completed, OfficeWorkflowStatus.Failed, OfficeWorkflowStatus.Completed }, results.Select(item => item.Status));
            Assert.Equal(output, results[2].OutputPath);
            Assert.Contains("Recognized", PdfReadDocument.Open(File.ReadAllBytes(output)).ExtractText());
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task OnePdfCannotOverwriteAnotherSelectedSource(bool providerAccess) {
        string root = NewDirectory();
        try {
            string first = Path.Combine(root, "first.pdf");
            string second = Path.Combine(root, "second.pdf");
            byte[] source = PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).ToBytes();
            File.WriteAllBytes(first, source); File.WriteAllBytes(second, source);
            var results = await new OfficeWorkflowRunner().RunOcrSessionAsync([
                new("first", new PdfSearchableWorkflowRequest { InputPath = first, OutputPath = second, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace }),
                new("second", new PdfSearchableWorkflowRequest { InputPath = second, OutputPath = Path.Combine(root, "result.pdf") })
            ], Engine());
            Assert.Equal(OfficeWorkflowStatus.Failed, results[0].Status);
            Assert.Equal(OfficeWorkflowStatus.Completed, results[1].Status);
            Assert.Equal(source, File.ReadAllBytes(second));
            int scopeOpens = 0;
            var access = providerAccess ? new OfficeWorkflowStreamInput("second.pdf", _ => {
                scopeOpens++;
                return Task.FromResult<Stream>(File.OpenRead(second));
            }) : null;
            var retry = await new OfficeWorkflowRunner().RunOcrSessionAsync([
                new("first", new PdfSearchableWorkflowRequest { InputPath = first, OutputPath = second, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace })
            ], Engine(), protectedOutputPaths: [results[1].OutputPath!], protectedInputs: [new(second, access)]);
            Assert.Equal(OfficeWorkflowStatus.Failed, Assert.Single(retry).Status);
            Assert.Equal(source, File.ReadAllBytes(second));
            if (providerAccess) Assert.True(scopeOpens > 0);
        } finally { Directory.Delete(root, recursive: true); }
    }
    private static string NewDirectory() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-ocr-session-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root); return root;
    }
    internal static byte[] Png() => Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
    private static IOcrEngine Engine() => new DelegateOcrEngine("session-fixture", (_, _) => Task.FromResult(new OcrResult {
        Text = "Recognized", Spans = [new OcrTextSpan { Text = "Recognized", Confidence = 1, Level = OcrTextSpanLevel.Word,
            CoordinateUnit = OcrCoordinateUnit.Points, Region = new OcrRegion { X = 20, Y = 60, Width = 90, Height = 12 } }]
    }));
}
