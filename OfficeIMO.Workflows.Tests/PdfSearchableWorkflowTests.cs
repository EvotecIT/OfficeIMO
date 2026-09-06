using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class PdfSearchableWorkflowTests {
    [Theory]
    [InlineData("success")]
    [InlineData("source-changed")]
    [InlineData("source-replaced")]
    [InlineData("host-denied")]
    [InlineData("cancelled")]
    public async Task PublicationChecksTheSourceAndHostAfterRecognition(string mode) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-ocr-workflow-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            string output = Path.Combine(root, "output.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).Save(source);
            byte[] original = File.ReadAllBytes(source);
            byte[] previous = [1, 2, 3, 4];
            File.WriteAllBytes(output, previous);
            using var cancellation = new CancellationTokenSource();
            bool recognized = false;
            bool guardCalled = false;
            var request = new PdfSearchableWorkflowRequest {
                InputPath = source, OutputPath = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                PublicationGuard = new Guard(() => {
                    Assert.True(recognized);
                    guardCalled = true;
                    return mode != "host-denied";
                })
            };
            var engine = new DelegateOcrEngine("fixture", (_, _) => {
                recognized = true;
                if (mode == "source-changed") File.WriteAllBytes(source, [9, 8, 7]);
                if (mode == "source-replaced") {
                    File.Move(source, Path.Combine(root, "old.pdf"));
                    File.WriteAllBytes(source, original);
                }
                if (mode == "cancelled") cancellation.Cancel();
                // Mutating the caller's options must not redirect the captured request.
                request.OutputPath = source;
                return Task.FromResult(RecognizedWord());
            });

            var result = await new OfficeWorkflowRunner().MakePdfSearchableAsync(request, engine, cancellation.Token);

            Assert.True(recognized);
            if (mode == "success") {
                Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
                Assert.Equal(output, result.OutputPath);
                Assert.Equal(1, result.AddedWordCount);
                Assert.Equal("Recognized", PdfReadDocument.Open(File.ReadAllBytes(output)).Pages.Single().ExtractText().Trim());
                Assert.True(guardCalled);
            } else {
                Assert.Equal(mode == "cancelled" ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed, result.Status);
                Assert.Null(result.OutputPath);
                Assert.Equal(previous, File.ReadAllBytes(output));
            }
            if (mode != "source-changed") Assert.Equal(original, File.ReadAllBytes(source));
            Assert.Empty(Directory.GetFiles(root, ".ocr-*.tmp"));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProviderInputIsReopenedBeforePublication(bool revoked) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-ocr-provider-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            byte[] bytes = PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).ToBytes();
            bool recognized = false;
            int reads = 0;
            string output = Path.Combine(root, "searchable.pdf");
            var request = new PdfSearchableWorkflowRequest {
                InputPath = "content://documents/scan", OutputPath = output,
                InputStream = new("scan.pdf", _ => {
                    reads++;
                    if (recognized && revoked) throw new UnauthorizedAccessException("Provider access revoked.");
                    return Task.FromResult<Stream>(new MemoryStream(bytes, writable: false));
                })
            };
            var engine = new DelegateOcrEngine("fixture", (_, _) => {
                recognized = true;
                return Task.FromResult(RecognizedWord());
            });
            var result = await new OfficeWorkflowRunner().MakePdfSearchableAsync(request, engine);
            Assert.Equal(2, reads);
            Assert.Equal(revoked ? OfficeWorkflowStatus.Failed : OfficeWorkflowStatus.Completed, result.Status);
            Assert.Equal(!revoked, File.Exists(output));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProviderOutputIsVerifiedOrRetainedForRecovery(bool failWrite) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-ocr-output-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).Save(source);
            byte[] bytes = [];
            var recoveryStore = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root, "recovery"));
            var request = new PdfSearchableWorkflowRequest {
                InputPath = source, OutputPath = "content://documents/output", ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                OutputStream = new("searchable.pdf", _ => Task.FromResult<Stream>(new MemoryStream(bytes)), _ => {
                    if (failWrite) throw new IOException("Provider unavailable during publication.");
                    return Task.FromResult<Stream>(new DestinationStream(result => bytes = result));
                }, recoveryStore)
            };
            var result = await new OfficeWorkflowRunner().MakePdfSearchableAsync(request,
                new DelegateOcrEngine("fixture", (_, _) => Task.FromResult(RecognizedWord())));
            Assert.Equal(failWrite ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Completed, result.Status);
            if (failWrite) {
                Assert.Null(result.OutputPath);
                Assert.NotNull(result.Recovery);
                var retained = Assert.Single(recoveryStore.GetRecoveries());
                await recoveryStore.VerifyAsync(retained);
                Assert.Equal("Recognized", PdfReadDocument.Open(File.ReadAllBytes(retained.FilePath)).Pages.Single().ExtractText().Trim());
                recoveryStore.Discard(retained);
            } else {
                Assert.Equal(request.OutputPath, result.OutputPath);
                Assert.Equal("Recognized", PdfReadDocument.Open(bytes).Pages.Single().ExtractText().Trim());
                Assert.Empty(recoveryStore.GetRecoveries());
            }
        } finally { Directory.Delete(root, recursive: true); }
    }

    private sealed class DestinationStream(Action<byte[]> commit) : MemoryStream {
        private bool _closed;
        protected override void Dispose(bool disposing) {
            if (!_closed) { _closed = true; commit(ToArray()); }
            base.Dispose(disposing);
        }
    }

    private static OcrResult RecognizedWord() => new() {
        Provider = "fixture", Spans = [new OcrTextSpan {
            Text = "Recognized", Level = OcrTextSpanLevel.Word, Confidence = 1,
            CoordinateUnit = OcrCoordinateUnit.Points,
            Region = new OcrRegion { X = 20, Y = 30, Width = 80, Height = 12 }
        }]
    };

    private sealed class Guard(Func<bool> check) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) => new(check());
    }
}
