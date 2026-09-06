using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class ProviderOutputPreparationTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public async Task BatchCannotOverwriteAnotherRequestsSource(bool deferred, bool invalidSecond) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-batch-output-protection-" + Guid.NewGuid().ToString("N"));
        var store = new OfficeWorkflowOutputRecoveryStore(root);
        byte[] html = System.Text.Encoding.UTF8.GetBytes("<html><body><p>Batch conversion</p></body></html>");
        byte[] pdf = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).ToBytes();
        int writes = 0;
        try {
            var output = new OfficeWorkflowStreamOutput("result.pdf", _ => Task.FromResult<Stream>(new MemoryStream(pdf)),
                _ => { writes++; return Task.FromResult<Stream>(new CommitStream(value => pdf = value)); }, store,
                deferred ? _ => Task.FromResult("content://provider/second") : null);
            var results = await new OfficeWorkflowRunner().RunBatchAsync([
                new() { InputPath = "content://provider/first", InputStream = new("first.html", _ => Task.FromResult<Stream>(new MemoryStream(html))),
                    Operation = OfficeWorkflowOperation.Convert, ConversionRouteId = "html-pdf", OutputPath = deferred ? "content://provider/folder" : "content://provider/second",
                    OutputStream = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace },
                new() { Id = invalidSecond ? "" : "second", InputPath = "content://provider/second", InputStream = new("second.pdf", _ => Task.FromResult<Stream>(new MemoryStream(pdf))), Operation = OfficeWorkflowOperation.Inspect }
            ]);
            Assert.Equal(0, writes);
            Assert.Equal(deferred ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Failed, results[0].Status);
            Assert.Equal(invalidSecond ? OfficeWorkflowStatus.Failed : OfficeWorkflowStatus.Completed, results[1].Status);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("success")]
    [InlineData("create-failure")]
    [InlineData("cancel")]
    [InlineData("source-alias")]
    [InlineData("duplicate")]
    public async Task DeferredCreationIsRecoverableAndUsesActualAuthorizedLocations(string mode) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-output-preparation-" + Guid.NewGuid().ToString("N"));
        var recovery = new OfficeWorkflowOutputRecoveryStore(root);
        using var cancellation = new CancellationTokenSource();
        byte[] source = PdfDocument.Create(compose => { compose.Page(page => page.Size(300, 400)); compose.Page(page => page.Size(300, 400)); }).ToBytes();
        int creations = 0, writes = 0;
        var bytes = new Dictionary<string, byte[]>();
        try {
            var result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(new() {
                InputPath = "content://provider/source", InputStream = new("source.pdf", _ => Task.FromResult<Stream>(new MemoryStream(source))),
                OutputDirectory = "content://provider/selected-folder", MaximumDimension = 32, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
                DirectoryOutput = new((name, _) => {
                    string? actual = null;
                    var output = new OfficeWorkflowStreamOutput(name,
                        _ => Task.FromResult<Stream>(new MemoryStream(bytes[actual!])),
                        _ => { writes++; return Task.FromResult<Stream>(new CommitStream(value => bytes[actual!] = value)); }, recovery,
                        _ => {
                            Assert.NotEmpty(Directory.GetFiles(root, "record.json", SearchOption.AllDirectories));
                            creations++;
                            actual = mode == "source-alias" ? "content://provider/source"
                                : "content://provider/assigned/" + (mode == "duplicate" ? 1 : creations);
                            if (!bytes.ContainsKey(actual)) bytes[actual] = [];
                            if (mode == "create-failure") throw new IOException("The provider failed after creating a child.");
                            if (mode == "cancel") cancellation.Cancel();
                            return Task.FromResult(actual);
                        });
                    return Task.FromResult(new OfficeWorkflowDirectoryOutputFile("content://provider/selected-folder", output));
                })
            }, cancellationToken: cancellation.Token);
            Assert.Equal(mode == "success" ? OfficeWorkflowStatus.Completed : OfficeWorkflowStatus.Unconfirmed, result.Status);
            Assert.Equal(mode == "success" ? 2 : mode == "duplicate" ? 1 : 0, writes);
            Assert.Equal(writes, result.Files.Count);
            Assert.All(result.Files, file => Assert.StartsWith("content://provider/assigned/", file.Path));
            if (mode == "success") Assert.Empty(recovery.GetRecoveries());
            else {
                var retained = Assert.Single(result.OutputRecoveries);
                Assert.Equal("content://provider/selected-folder", retained.Destination);
                await recovery.VerifyAsync(retained);
            }
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    private sealed class CommitStream(Action<byte[]> commit) : MemoryStream {
        private bool _closed;
        protected override void Dispose(bool disposing) {
            if (!_closed) { _closed = true; commit(ToArray()); }
            base.Dispose(disposing);
        }
    }
}
