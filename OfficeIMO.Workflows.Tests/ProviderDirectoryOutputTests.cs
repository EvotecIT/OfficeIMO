using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Workflows.Tests;

public sealed class ProviderDirectoryOutputTests {
    [Fact]
    public async Task RequiresReplaceBeforeResolvingAnyDestination() {
        using var provider = new Provider();
        var request = provider.Request();
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Rename;
        var result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(request);
        Assert.Equal(OfficeWorkflowFailureKind.ValidationFailed, result.FailureKind);
        Assert.Empty(result.Files);
        Assert.Equal(0, provider.WriteCount);
        Assert.False(Directory.Exists(provider.Recovery.DirectoryPath));
    }

    [Fact]
    public async Task CannotReplaceTheSelectedSourceThroughAProviderChild() {
        using var provider = new Provider();
        var request = provider.Request();
        byte[] original = provider.Source.ToArray();
        request.Pages = "1";
        request.DirectoryOutput = new((name, _) => Task.FromResult(new OfficeWorkflowDirectoryOutputFile(request.InputPath,
            new(name, _ => Task.FromResult<Stream>(new MemoryStream(provider.Source)),
                _ => { provider.WriteCount++; return Task.FromResult<Stream>(new CommitStream(bytes => provider.Source = bytes)); }, provider.Recovery))));
        var result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(request);
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(0, provider.WriteCount);
        Assert.Equal(original, provider.Source);
        Assert.Empty(result.OutputRecoveries);
    }

    [Theory]
    [InlineData(OfficeImageExportFormat.Png)]
    [InlineData(OfficeImageExportFormat.Svg)]
    [InlineData(OfficeImageExportFormat.Jpeg)]
    [InlineData(OfficeImageExportFormat.Tiff)]
    [InlineData(OfficeImageExportFormat.Webp)]
    public async Task EveryFormatPublishesReopenedImages(OfficeImageExportFormat format) {
        using var provider = new Provider();
        var request = provider.Request();
        request.Format = format;
        var result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(request);
        Assert.True(result.Succeeded, result.Summary);
        Assert.Equal(3, result.Files.Count);
        Assert.Empty(result.OutputRecoveries);
        Assert.Empty(provider.Recovery.GetRecoveries());
        Assert.Equal(result.OutputBytes, provider.Bytes.Values.Sum(bytes => (long)bytes.Length));
        foreach (var file in result.Files) {
            Assert.True(OfficeImageReader.TryValidateContent(provider.Bytes[file.Path], file.Path, default, out var info));
            Assert.Equal(file.Width, info.Width);
        }
    }

    [Theory]
    [InlineData("open")]
    [InlineData("verify")]
    [InlineData("cancel-write")]
    public async Task UnconfirmedSecondWriteRetainsFirstAndRecovery(string failure) {
        using var provider = new Provider { Failure = failure };
        var result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(provider.Request(), cancellationToken: provider.Cancellation.Token);
        Assert.Equal(OfficeWorkflowStatus.Unconfirmed, result.Status);
        Assert.Single(result.Files);
        Assert.Equal(2, provider.WriteCount);
        var recovery = Assert.Single(result.OutputRecoveries);
        Assert.Single(provider.Recovery.GetRecoveries());
        await provider.Recovery.VerifyAsync(recovery);
        Assert.True(OfficeImageReader.TryValidateContent(File.ReadAllBytes(recovery.FilePath), recovery.Name, default, out _));
        Assert.Equal(result.Files[0].SizeBytes, result.OutputBytes);
    }

    [Theory]
    [InlineData("deny")]
    [InlineData("cancel-before")]
    [InlineData("source-change")]
    public async Task StopsBeforeSecondWriteWithoutLosingFirst(string failure) {
        using var provider = new Provider { Failure = failure };
        var request = provider.Request();
        request.PublicationGuard = new Guard(provider);
        var result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(request, cancellationToken: provider.Cancellation.Token);
        Assert.Equal(failure == "cancel-before" ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed, result.Status);
        Assert.Single(result.Files);
        Assert.Equal(1, provider.WriteCount);
        Assert.Empty(result.OutputRecoveries);
        Assert.Empty(provider.Recovery.GetRecoveries());
    }

    [Theory]
    [InlineData("resolve")]
    [InlineData("duplicate")]
    [InlineData("name")]
    public async Task InvalidResolutionCannotPartiallyPublish(string failure) {
        using var provider = new Provider { Failure = failure };
        var result = await new OfficeWorkflowRunner().ExportPdfPagesAsync(provider.Request());
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Empty(result.Files);
        Assert.Equal(0, provider.WriteCount);
        Assert.Empty(provider.Recovery.GetRecoveries());
    }

    private sealed class Guard(Provider provider) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            if (provider.WriteCount == 1) {
                if (provider.Failure == "cancel-before") provider.Cancellation.Cancel();
                if (provider.Failure == "source-change") provider.Source = CreatePdf(4);
                if (provider.Failure == "deny") return ValueTask.FromResult(false);
            }
            return ValueTask.FromResult(true);
        }
    }

    private sealed class Provider : IDisposable {
        internal string? Failure;
        internal byte[] Source = CreatePdf(3);
        internal readonly CancellationTokenSource Cancellation = new();
        internal readonly Dictionary<string, byte[]> Bytes = new();
        internal readonly OfficeWorkflowOutputRecoveryStore Recovery = new(Path.Combine(Path.GetTempPath(), "officeimo-provider-output-" + Guid.NewGuid().ToString("N")));
        internal int WriteCount;
        private int _resolved;

        internal PdfPageImageExportRequest Request() => new() {
            InputPath = "content://provider/input", InputStream = new("input.pdf", _ => Task.FromResult<Stream>(new MemoryStream(Source))),
            OutputDirectory = "content://provider/output", ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
            MaximumDimension = 32, DirectoryOutput = new((name, token) => {
                _resolved++;
                if (Failure == "resolve" && _resolved == 2) throw new IOException("Provider resolution failed.");
                string location = "content://provider/output/" + (Failure == "duplicate" ? "same" : Uri.EscapeDataString(name));
                var output = new OfficeWorkflowStreamOutput(Failure == "name" ? "changed.png" : name,
                    _ => Task.FromResult<Stream>(new MemoryStream(Bytes.TryGetValue(location, out var bytes) ? bytes : throw new FileNotFoundException())),
                    _ => {
                        WriteCount++;
                        // Recovery must already be durable even when creation throws before returning a stream.
                        Assert.NotEmpty(Directory.GetFiles(Recovery.DirectoryPath, "record.json", SearchOption.AllDirectories));
                        if (WriteCount == 2 && Failure == "open") { Bytes[location] = []; throw new IOException("Creation failed after mutation."); }
                        if (WriteCount == 2 && Failure == "cancel-write") Cancellation.Cancel();
                        return Task.FromResult<Stream>(new CommitStream(bytes => Bytes[location] = WriteCount == 2 && Failure == "verify" ? [] : bytes));
                    }, Recovery);
                return Task.FromResult(new OfficeWorkflowDirectoryOutputFile(location, output));
            })
        };

        public void Dispose() {
            Cancellation.Dispose();
            if (Directory.Exists(Recovery.DirectoryPath)) Directory.Delete(Recovery.DirectoryPath, true);
        }
    }

    private sealed class CommitStream(Action<byte[]> commit) : MemoryStream {
        private bool _closed;
        protected override void Dispose(bool disposing) {
            if (!_closed) { _closed = true; commit(ToArray()); }
            base.Dispose(disposing);
        }
    }

    private static byte[] CreatePdf(int count) => PdfDocument.Create(compose => {
        for (int index = 0; index < count; index++) compose.Page(page => page.Size(300, 400));
    }).ToBytes();
}
