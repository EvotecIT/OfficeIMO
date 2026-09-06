using System.Text;
using System.IO.Compression;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeWorkflowStreamInputTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProviderZipExpansionIsPrivateAndCleanupPrecedesPublication(bool failCleanup) {
        using var root = new Scope();
        string marker = Guid.NewGuid().ToString("N") + ".pdf";
        using var buffer = new MemoryStream();
        using (var archive = new ZipArchive(buffer, ZipArchiveMode.Create, leaveOpen: true)) {
            Write(marker, CreatePdf(1));
            Write("page/source.html", Encoding.UTF8.GetBytes("<html><head><link rel=\"stylesheet\" href=\"style.css\"></head><body><p>Private HTML resource</p><img src=\"pixel.png\"></body></html>"));
            Write("page/style.css", Encoding.UTF8.GetBytes("p { color: #123456; }"));
            Write("page/pixel.png", Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
            void Write(string name, byte[] bytes) {
                using Stream entry = archive.CreateEntry(name).Open();
                entry.Write(bytes);
            }
        }
        var source = new Source(buffer.ToArray());
        string output = Path.Combine(root.Path, "assembled.pdf");
        string? extractionRoot = null;
        bool inspected = false;
        try {
            var result = await new OfficeWorkflowRunner().AssemblePdfAsync(new() {
                Sources = ["content://provider/archive"],
                SourceStreams = new Dictionary<string, OfficeWorkflowStreamInput> { ["content://provider/archive"] = source.Input("documents.zip") },
                OutputPath = output
            }, new DirectProgress(update => {
                if (update.Stage == "normalize" && !inspected) {
                    extractionRoot = Directory.GetDirectories(Path.GetTempPath(), "officeimo-assembly-*")
                        .Single(path => File.Exists(Path.Combine(path, "archive-0001", marker)));
                    if (!OperatingSystem.IsWindows()) {
                        Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute,
                            File.GetUnixFileMode(extractionRoot));
                    }
                    Assert.True(File.Exists(Path.Combine(extractionRoot, "archive-0001", "page", "pixel.png")));
                    inspected = true;
                }
                if (update.Stage == "publish" && failCleanup) {
                    // Replace this test's exact extraction directory with a file to force a portable cleanup failure.
                    Directory.Move(extractionRoot!, Path.Combine(root.Path, "retained-extraction"));
                    File.WriteAllText(extractionRoot!, "cleanup obstruction");
                }
            }));
            Assert.True(inspected);
            if (failCleanup) {
                Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
                Assert.False(File.Exists(output));
                Assert.Contains(result.Diagnostics, item => item.Code == "InputStagingCleanupFailed");
            } else {
                Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
                Assert.Equal(2, PdfDocument.Load(output).Inspect().PageCount);
                Assert.False(Directory.Exists(extractionRoot));
                Assert.DoesNotContain(result.Diagnostics, item => item.Code == "HtmlRenderResourceUnavailable");
            }
            Assert.Equal(source.Reads, source.Closed);
        } finally {
            if (extractionRoot is not null && File.Exists(extractionRoot)) File.Delete(extractionRoot);
        }
    }

    private sealed class DirectProgress(Action<OfficeWorkflowProgress> report) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress value) => report(value);
    }

    [Fact]
    public async Task OpaqueHtmlInputConvertsWithClosedScopesAndVerifiedOutput() {
        using var root = new Scope();
        var source = new Source(Encoding.UTF8.GetBytes("<html><body><p>Provider conversion</p></body></html>"));
        var result = await new OfficeWorkflowRunner().RunAsync(new() {
            InputPath = "content://provider/opaque", InputStream = source.Input("report.html"),
            Operation = OfficeWorkflowOperation.Convert, ConversionRouteId = "html-pdf",
            OutputPath = Path.Combine(root.Path, "report.pdf")
        });
        Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
        Assert.NotEmpty(PdfDocument.Load(result.OutputPath!).Inspect().Pages);
        Assert.Equal(2, source.Reads);
        Assert.Equal(source.Reads, source.Closed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ChangedOrRevokedSourceBeforePublicationLeavesExistingOutputUnchanged(bool revoked) {
        using var root = new Scope();
        string output = Path.Combine(root.Path, "report.pdf");
        byte[] original = [4, 5, 6];
        File.WriteAllBytes(output, original);
        var source = new Source(CreatePdf(1)) { ChangeOnVerification = !revoked, DenyVerification = revoked };
        var result = await new OfficeWorkflowRunner().RunAsync(new() {
            InputPath = "content://provider/opaque", InputStream = source.Input("report.pdf"),
            Operation = OfficeWorkflowOperation.Optimize, OutputPath = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(original, File.ReadAllBytes(output));
        Assert.Null(result.OutputPath);
        Assert.Equal(revoked ? 1 : 2, source.Closed);
    }

    [Fact]
    public async Task ProviderInputCannotOverwriteItsOriginalLocalLocation() {
        using var root = new Scope();
        string path = Path.Combine(root.Path, "input.pdf");
        byte[] bytes = CreatePdf(1);
        File.WriteAllBytes(path, bytes);
        var source = new Source(bytes);
        var result = await new OfficeWorkflowRunner().RunAsync(new() {
            InputPath = path, InputStream = source.Input("input.pdf"),
            Operation = OfficeWorkflowOperation.Optimize, OutputPath = path, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(bytes, File.ReadAllBytes(path));
    }

    [Fact]
    public async Task ProviderAssemblyPreservesSelectedPageOrderAndVerifiesBothSources() {
        using var root = new Scope();
        var first = new Source(CreatePdf(1, 200));
        var second = new Source(CreatePdf(2, 400));
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(new() {
            Sources = ["content://provider/first", "content://provider/second"],
            SourceStreams = new Dictionary<string, OfficeWorkflowStreamInput> {
                ["content://provider/first"] = first.Input("first.pdf"),
                ["content://provider/second"] = second.Input("second.pdf")
            },
            OutputPath = Path.Combine(root.Path, "assembled.pdf")
        });
        Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
        var pages = PdfDocument.Load(result.OutputPath!).Inspect().Pages;
        Assert.Equal(3, pages.Count);
        Assert.Equal(200, pages[0].Width);
        Assert.Equal(400, pages[1].Width);
        Assert.Equal(400, pages[2].Width);
        Assert.Equal(new[] { "first.pdf", "second.pdf" }, result.Diagnostics
            .Where(item => item.Code == "AssemblySourceNormalized").Select(item => item.Details["name"]));
        Assert.Equal(2, first.Reads);
        Assert.Equal(2, second.Reads);
        Assert.Equal(first.Reads, first.Closed);
        Assert.Equal(second.Reads, second.Closed);
    }

    [Fact]
    public async Task ChangedAssemblyProviderPreventsPublication() {
        using var root = new Scope();
        var source = new Source(CreatePdf(1)) { ChangeOnVerification = true };
        string output = Path.Combine(root.Path, "assembled.pdf");
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(new() {
            Sources = ["content://provider/source"],
            SourceStreams = new Dictionary<string, OfficeWorkflowStreamInput> { ["content://provider/source"] = source.Input("source.pdf") },
            OutputPath = output
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.False(File.Exists(output));
        Assert.Equal(source.Reads, source.Closed);
    }

    [Fact]
    public async Task NonSeekableInputLimitAndExpectedFingerprintFailBeforeExecution() {
        using var root = new Scope();
        var source = new Source(CreatePdf(1));
        string output = Path.Combine(root.Path, "output.pdf");
        var request = new OfficeWorkflowRequest {
            InputPath = "content://provider/source", InputStream = source.Input("source.pdf"),
            Operation = OfficeWorkflowOperation.Optimize, OutputPath = output,
            Limits = new() { MaximumInputBytes = 8 }
        };
        var runner = new OfficeWorkflowRunner();
        Assert.Equal(OfficeWorkflowStatus.Failed, (await runner.RunAsync(request)).Status);
        request.Limits.MaximumInputBytes = 1024 * 1024;
        request.InputStream = source.Input("source.pdf", new string('0', 64));
        Assert.Equal(OfficeWorkflowStatus.Failed, (await runner.RunAsync(request)).Status);
        Assert.False(File.Exists(output));
        Assert.Equal(source.Reads, source.Closed);
    }

    [Fact]
    public async Task ProviderContentsAreCheckedAfterHostPublicationAuthorization() {
        using var root = new Scope();
        var source = new Source(CreatePdf(1));
        string output = Path.Combine(root.Path, "output.pdf");
        var result = await new OfficeWorkflowRunner().RunAsync(new() {
            InputPath = "content://provider/source", InputStream = source.Input("source.pdf"),
            Operation = OfficeWorkflowOperation.Optimize, OutputPath = output,
            PublicationGuard = new CallbackGuard(() => source.ChangeOnVerification = true)
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.False(File.Exists(output));
        Assert.Equal(source.Reads, source.Closed);
    }

    [Fact]
    public async Task ProviderCancellationAtPublicationRetainsNoOutput() {
        using var root = new Scope();
        using var cancellation = new CancellationTokenSource();
        var source = new Source(CreatePdf(1));
        string output = Path.Combine(root.Path, "output.pdf");
        var result = await new OfficeWorkflowRunner().RunAsync(new() {
            InputPath = "content://provider/source", InputStream = source.Input("source.pdf"),
            Operation = OfficeWorkflowOperation.Optimize, OutputPath = output,
            PublicationGuard = new CallbackGuard(cancellation.Cancel)
        }, cancellationToken: cancellation.Token);
        Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status);
        Assert.False(File.Exists(output));
        Assert.Equal(source.Reads, source.Closed);
    }

    private sealed class CallbackGuard(Action callback) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            callback();
            return ValueTask.FromResult(true);
        }
    }

    private static byte[] CreatePdf(int count, int width = 300) => PdfDocument.Create(compose => {
        for (int index = 0; index < count; index++) compose.Page(page => page.Size(width, 500));
    }).ToBytes();

    private sealed class Source(byte[] bytes) {
        internal int Reads;
        internal int Closed;
        internal bool ChangeOnVerification;
        internal bool DenyVerification;
        internal OfficeWorkflowStreamInput Input(string name, string? expected = null) => new(name, token => {
            token.ThrowIfCancellationRequested();
            Reads++;
            if (Reads > 1 && DenyVerification) throw new UnauthorizedAccessException("Provider access was revoked.");
            return Task.FromResult<Stream>(new ProviderStream(Reads > 1 && ChangeOnVerification ? [1, 2, 3] : bytes, () => Closed++));
        }, expected);
    }

    private sealed class ProviderStream(byte[] bytes, Action closed) : MemoryStream(bytes, writable: false) {
        private bool _closed;
        public override bool CanSeek => false;
        protected override void Dispose(bool disposing) {
            if (!_closed) { _closed = true; closed(); }
            base.Dispose(disposing);
        }
    }

    private sealed class Scope : IDisposable {
        internal string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeimo-stream-tests-" + Guid.NewGuid().ToString("N"));
        internal Scope() => Directory.CreateDirectory(Path);
        public void Dispose() => Directory.Delete(Path, true);
    }
}
