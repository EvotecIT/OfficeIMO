using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeWorkflowStreamOutputTests {
    [Fact]
    public async Task ActiveRecoveryIsHiddenAndCannotBeDiscardedByAnotherStore() {
        using var root = new Scope();
        var store = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root.Path, "recovery"));
        var other = new OfficeWorkflowOutputRecoveryStore(store.DirectoryPath);
        var lease = await store.CreateAsync([1, 2, 3], "content://provider/selected", "selected.pdf", default);
        try {
            Assert.Empty(other.GetRecoveries());
            Assert.Throws<IOException>(() => other.Discard(lease.Recovery));
            Assert.True(File.Exists(lease.Recovery.FilePath));
        } finally { lease.Dispose(); }
        Assert.Single(other.GetRecoveries());
        await other.VerifyAsync(lease.Recovery);
        other.Discard(lease.Recovery);
    }

    [Fact]
    public async Task ConcurrentStoresShareTheRecoveryAdmissionBudget() {
        using var root = new Scope();
        string directory = Path.Combine(root.Path, "recovery");
        async Task<bool> RetainAsync() {
            try {
                var store = new OfficeWorkflowOutputRecoveryStore(directory, 12000);
                using var lease = await store.CreateAsync(new byte[8000], "content://provider/selected", "selected.pdf", default);
                return true;
            } catch (IOException) { return false; }
        }
        bool[] admitted = await Task.WhenAll(Task.Run(RetainAsync), Task.Run(RetainAsync));
        Assert.Equal(1, admitted.Count(value => value));
        var reader = new OfficeWorkflowOutputRecoveryStore(directory, 12000);
        var record = Assert.Single(reader.GetRecoveries());
        await reader.VerifyAsync(record);
        reader.Discard(record);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AssemblyPublishesVerifiedBytesOrPreservesItsCompleteCandidate(bool failWrite) {
        using var root = new Scope();
        string source = Path.Combine(root.Path, "source.pdf");
        File.WriteAllBytes(source, PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).ToBytes());
        byte[] outputBytes = [];
        var store = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root.Path, "recovery"));
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(new() {
            Sources = [source], OutputPath = "content://provider/selected", ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
            OutputStream = new("selected.pdf", _ => Task.FromResult<Stream>(new MemoryStream(outputBytes)), _ => {
                if (failWrite) throw new IOException("Provider write failed.");
                return Task.FromResult<Stream>(new DestinationStream(bytes => outputBytes = bytes));
            }, store)
        });
        Assert.Equal(failWrite ? OfficeWorkflowStatus.Unconfirmed : OfficeWorkflowStatus.Completed, result.Status);
        if (failWrite) {
            var recovery = Assert.Single(store.GetRecoveries());
            await store.VerifyAsync(recovery);
            Assert.Equal(1, PdfDocument.Load(recovery.FilePath).Inspect().PageCount);
            store.Discard(recovery);
        } else {
            Assert.Equal(1, PdfDocument.Load(outputBytes).Inspect().PageCount);
            Assert.Empty(store.GetRecoveries());
        }
    }

    [Fact]
    public async Task FolderAssemblyCanPublishToAProvider() {
        using var root = new Scope();
        string sources = Path.Combine(root.Path, "sources");
        Directory.CreateDirectory(sources);
        File.WriteAllBytes(Path.Combine(sources, "source.pdf"), PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).ToBytes());
        byte[] outputBytes = [];
        var store = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root.Path, "recovery"));
        var result = await new OfficeWorkflowRunner().AssemblePdfAsync(new() {
            Sources = [sources], OutputPath = "content://provider/selected", ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
            OutputStream = new("selected.pdf", _ => Task.FromResult<Stream>(new MemoryStream(outputBytes)),
                _ => Task.FromResult<Stream>(new DestinationStream(bytes => outputBytes = bytes)), store)
        });
        Assert.True(result.Status == OfficeWorkflowStatus.Completed, result.Summary);
        Assert.Equal(1, PdfDocument.Load(outputBytes).Inspect().PageCount);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task InterruptedPrePublicationArtifactDoesNotPermanentlyConsumeCapacity(bool longPath) {
        using var root = new Scope();
        string directory = Path.Combine(root.Path, longPath ? new string('a', 160) : "recovery");
        string abandoned = Path.Combine(directory, "output-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(abandoned);
        File.WriteAllBytes(Path.Combine(abandoned, "output.pdf"), new byte[8000]);
        File.WriteAllBytes(Path.Combine(abandoned, ".officeimo-" + Guid.NewGuid().ToString("N") + ".tmp"), new byte[1000]);
        File.WriteAllBytes(Path.Combine(abandoned, ".lease"), []);
        foreach (string target in new[] { "output.pdf", "record.json" }) {
            // Use the actual writer's claim path, including its long-path fallback.
            string claim = (string)System.Reflection.Assembly.Load("OfficeIMO.Core")
                .GetType("OfficeIMO.Core.Internal.OfficeFileCommit")!
                .GetMethod("CreateClaimPath", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!
                .Invoke(null, [Path.Combine(abandoned, target)])!;
            File.WriteAllBytes(claim, []);
        }
        var store = new OfficeWorkflowOutputRecoveryStore(directory, 12000);
        Assert.Empty(store.GetRecoveries());
        using var lease = await store.CreateAsync(new byte[8000], "content://provider/selected", "selected.pdf", default);
        Assert.False(Directory.Exists(abandoned));
        Assert.True(File.Exists(lease.Recovery.FilePath));
    }

    [Theory]
    [InlineData("record.json")]
    [InlineData("unknown.bin")]
    [InlineData(".officeimo-not-a-guid.tmp")]
    [InlineData(".unrelated.pdf.officeimo-commit")]
    [InlineData(".officeimo-0123456789abcdef01234567.commit")]
    public async Task AdmissionPreservesUnrecognizedAndFutureRecords(string filename) {
        using var root = new Scope();
        string directory = Path.Combine(root.Path, "recovery");
        string preserved = Path.Combine(directory, "output-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(preserved);
        byte[] bytes = System.Text.Encoding.UTF8.GetBytes("{\"Version\":2}");
        string path = Path.Combine(preserved, filename);
        File.WriteAllBytes(path, bytes);
        var store = new OfficeWorkflowOutputRecoveryStore(directory);
        using var lease = await store.CreateAsync([1], "content://provider/selected", "selected.pdf", default);
        Assert.Equal(bytes, File.ReadAllBytes(path));
    }

    [Fact]
    public async Task AdmissionPreservesAnIncompleteRecordWithAnActiveLease() {
        using var root = new Scope();
        string directory = Path.Combine(root.Path, "recovery");
        string active = Path.Combine(directory, "output-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(active);
        string artifact = Path.Combine(active, "output.pdf");
        File.WriteAllBytes(artifact, [1, 2, 3]);
        using var activeLease = new FileStream(Path.Combine(active, ".lease"), FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
        var store = new OfficeWorkflowOutputRecoveryStore(directory);
        using var lease = await store.CreateAsync([1], "content://provider/selected", "selected.pdf", default);
        Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(artifact));
    }

    [Fact]
    public async Task AdmissionPreservesAnIncompleteRecordWithAnActiveCommitClaim() {
        using var root = new Scope();
        string directory = Path.Combine(root.Path, "recovery");
        string active = Path.Combine(directory, "output-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(active);
        string artifact = Path.Combine(active, "output.pdf");
        File.WriteAllBytes(artifact, [1, 2, 3]);
        string claim = Path.Combine(active, ".output.pdf.officeimo-commit");
        using var activeClaim = new FileStream(claim, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
        var store = new OfficeWorkflowOutputRecoveryStore(directory);
        using var lease = await store.CreateAsync([1], "content://provider/selected", "selected.pdf", default);
        Assert.Equal(new byte[] { 1, 2, 3 }, File.ReadAllBytes(artifact));
        Assert.True(File.Exists(claim));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public async Task ProviderOutputVerifiesOrRetainsRestartRecovery(int failureMode) {
        using var root = new Scope();
        string input = Path.Combine(root.Path, "source.pdf");
        byte[] original = PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).ToBytes();
        File.WriteAllBytes(input, original);
        using var cancellation = new CancellationTokenSource();
        byte[] destinationBytes = [9, 9, 9];
        int writes = 0;
        var store = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root.Path, "recovery"));
        var output = new OfficeWorkflowStreamOutput("selected.pdf", _ => Task.FromResult<Stream>(new MemoryStream(destinationBytes)), _ => {
            writes++;
            destinationBytes = [];
            if (failureMode == 1) throw new IOException("Provider failed after truncating.");
            if (failureMode == 3) cancellation.Cancel();
            return Task.FromResult<Stream>(new DestinationStream(bytes => destinationBytes = failureMode == 2 ? [.. bytes, 1] : bytes));
        }, store);
        var result = await new OfficeWorkflowRunner().RunAsync(new() {
            InputPath = input, Operation = OfficeWorkflowOperation.Optimize,
            OutputPath = "content://provider/selected", OutputStream = output, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
        }, cancellationToken: cancellation.Token);
        Assert.Equal(original, File.ReadAllBytes(input));
        Assert.Equal(1, writes);
        if (failureMode == 0) {
            Assert.Equal(OfficeWorkflowStatus.Completed, result.Status);
            Assert.Equal("content://provider/selected", result.OutputPath);
            Assert.NotEmpty(PdfDocument.Load(destinationBytes).Inspect().Pages);
            Assert.Empty(store.GetRecoveries());
            Assert.Null(result.Recovery);
        } else {
            Assert.Equal(OfficeWorkflowStatus.Unconfirmed, result.Status);
            Assert.Null(result.OutputPath);
            Assert.NotNull(result.Recovery);
            var restarted = new OfficeWorkflowOutputRecoveryStore(store.DirectoryPath);
            var recovery = Assert.Single(restarted.GetRecoveries());
            Assert.Equal(result.Recovery.Id, recovery.Id);
            await restarted.VerifyAsync(recovery);
            Assert.NotEmpty(PdfDocument.Load(recovery.FilePath).Inspect().Pages);
            if (!OperatingSystem.IsWindows()) {
                Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute,
                    File.GetUnixFileMode(Path.GetDirectoryName(recovery.FilePath)!));
            }
            byte[] damaged = File.ReadAllBytes(recovery.FilePath);
            damaged[0] ^= 0xFF;
            File.WriteAllBytes(recovery.FilePath, damaged);
            await Assert.ThrowsAsync<IOException>(() => restarted.VerifyAsync(recovery));
            File.WriteAllBytes(recovery.FilePath, [1]);
            Assert.Single(restarted.GetRecoveries());
            await Assert.ThrowsAsync<IOException>(() => restarted.VerifyAsync(recovery));
            restarted.Discard(recovery);
            Assert.Empty(restarted.GetRecoveries());
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DeniedPublicationOrRecoveryCapacityPreventsOpeningDestination(bool fullStore) {
        using var root = new Scope();
        string input = Path.Combine(root.Path, "source.pdf");
        File.WriteAllBytes(input, PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).ToBytes());
        var store = new OfficeWorkflowOutputRecoveryStore(Path.Combine(root.Path, "recovery"), fullStore ? 1 : 1024 * 1024);
        int writes = 0;
        var result = await new OfficeWorkflowRunner().RunAsync(new() {
            InputPath = input, Operation = OfficeWorkflowOperation.Optimize,
            OutputPath = "content://provider/selected", ConflictPolicy = OfficeWorkflowConflictPolicy.Replace,
            OutputStream = new("selected.pdf", _ => Task.FromResult<Stream>(new MemoryStream()), _ => {
                writes++;
                return Task.FromResult<Stream>(new MemoryStream());
            }, store),
            PublicationGuard = new DenyGuard()
        });
        Assert.Equal(OfficeWorkflowStatus.Failed, result.Status);
        Assert.Equal(0, writes);
        Assert.Empty(store.GetRecoveries());
    }

    private sealed class DenyGuard : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) => ValueTask.FromResult(false);
    }

    private sealed class DestinationStream(Action<byte[]> commit) : MemoryStream {
        private bool _closed;
        protected override void Dispose(bool disposing) {
            if (!_closed) { _closed = true; commit(ToArray()); }
            base.Dispose(disposing);
        }
    }

    private sealed class Scope : IDisposable {
        internal string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeimo-stream-output-tests-" + Guid.NewGuid().ToString("N"));
        internal Scope() => Directory.CreateDirectory(Path);
        public void Dispose() => Directory.Delete(Path, true);
    }
}
