using System.Text;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficePdfArchiveWorkflowTests {
    [Fact]
    public void MinimalJsonRequestUsesDeclaredResourceDefaults() {
        var request = OfficePdfArchiveSerializer.ParseRequest(Encoding.UTF8.GetBytes(
            "{\"InputDirectory\":\"input\",\"OutputDirectory\":\"output\",\"CheckpointDirectory\":\"state\"}"));
        var defaults = new OfficePdfArchiveRequest { InputDirectory = "input", OutputDirectory = "output", CheckpointDirectory = "state" };
        Assert.Equal(defaults, request);
    }
    [Fact]
    public async Task ResumeRetainsImportAndRenderingFidelityStages() {
        using var scope = new ArchiveDirectory();
        File.WriteAllBytes(Path.Combine(scope.Input, "loss.doc"), LegacyDocPdfStageTests.CreateLegacySource());
        var request = scope.Request with { AllowLegacyImportLoss = true };
        var first = new CaptureItem();
        Assert.Equal(1, (await OfficePdfArchiveWorkflow.RunAsync(request, first)).Completed);
        Assert.Contains(first.Item!.Diagnostics, item => item.Stage == "import" && item.Details.ContainsKey("lossKind"));
        Assert.Contains(first.Item.Diagnostics, item => item.Stage == "convert" && item.Code == "NativeFontFamilySubstituted");
        var resumed = new CaptureItem();
        Assert.Equal(1, (await OfficePdfArchiveWorkflow.RunAsync(request, resumed)).Reused);
        foreach (var diagnostic in first.Item.Diagnostics.Where(item => item.Details.ContainsKey("lossKind"))) {
            var restored = Assert.Single(resumed.Item!.Diagnostics, item => item.Code == diagnostic.Code && item.Stage == diagnostic.Stage);
            Assert.Equal(diagnostic.Details["lossKind"], restored.Details["lossKind"]);
            Assert.Equal(diagnostic.Details["source"], restored.Details["source"]);
        }
    }

    private sealed class CaptureItem : IProgress<OfficePdfArchiveItemResult> {
        public OfficePdfArchiveItemResult? Item { get; private set; }
        public void Report(OfficePdfArchiveItemResult item) => Item = item;
    }
    [Fact]
    public async Task CancellationKeepsCompletedReceiptsAndResumeFinishesTheRest() {
        using var scope = new ArchiveDirectory();
        for (int index = 0; index < 4; index++) File.WriteAllText(Path.Combine(scope.Input, $"{index}.txt"), "cancellation evidence");
        using var cancellation = new CancellationTokenSource();
        var interrupted = await OfficePdfArchiveWorkflow.RunAsync(scope.Request with { MaximumConcurrency = 1 },
            new CancelAfterCompletion(cancellation), cancellation.Token);
        Assert.True(interrupted.Cancelled);
        Assert.Equal(1, interrupted.Completed);
        var resumed = await OfficePdfArchiveWorkflow.RunAsync(scope.Request);
        Assert.Equal(4, resumed.Completed);
        Assert.Equal(1, resumed.Reused);
        Assert.Equal(0, resumed.Failed);
    }

    [Fact]
    public async Task HostPublicationPolicyIsCheckedBeforeCreatingArchiveState() {
        using var scope = new ArchiveDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "protected publication");
        await Assert.ThrowsAsync<UnauthorizedAccessException>(() => OfficePdfArchiveWorkflow.RunAsync(scope.Request,
            publicationGuard: new RejectPublication()));
        Assert.False(Directory.Exists(scope.Output));
        Assert.False(Directory.Exists(scope.State));
    }

    private sealed class RejectPublication : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool directory, CancellationToken token) => ValueTask.FromResult(false);
    }
    private sealed class CancelAfterCompletion(CancellationTokenSource cancellation) : IProgress<OfficePdfArchiveItemResult> {
        public void Report(OfficePdfArchiveItemResult item) { if (item.Status == OfficeWorkflowStatus.Completed) cancellation.Cancel(); }
    }

    [Fact]
    public async Task ArchivePreservesRelativeNamesAndResumesOnlyVerifiedArtifacts() {
        using var scope = new ArchiveDirectory();
        Directory.CreateDirectory(Path.Combine(scope.Input, "nested"));
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "first archive item");
        File.WriteAllText(Path.Combine(scope.Input, "nested", "one.txt"), "second archive item");
        var request = scope.Request;
        var first = await OfficePdfArchiveWorkflow.RunAsync(request);
        Assert.Equal(2, first.Completed);
        Assert.Equal(0, first.Failed);
        Assert.True(File.Exists(Path.Combine(scope.Output, "nested", "one.txt.pdf")));
        string output = Path.Combine(scope.Output, "one.txt.pdf");
        byte[] original = File.ReadAllBytes(output);
        var resumed = await OfficePdfArchiveWorkflow.RunAsync(request with { MaximumConcurrency = 1 });
        Assert.Equal(2, resumed.Reused);
        Assert.Equal(original, File.ReadAllBytes(output));
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "changed completed source");
        var changed = await OfficePdfArchiveWorkflow.RunAsync(request with { RetryFailed = true });
        Assert.Equal(1, changed.Failed);
        Assert.Equal(original, File.ReadAllBytes(output));
    }

    [Fact]
    public async Task RecordedFailureRequiresExplicitRetryAndPermitsRepairOfFailedSource() {
        using var scope = new ArchiveDirectory();
        string input = Path.Combine(scope.Input, "invalid.txt");
        File.WriteAllBytes(input, [0xff]);
        Assert.Equal(1, (await OfficePdfArchiveWorkflow.RunAsync(scope.Request)).Failed);
        File.WriteAllText(input, "repaired input");
        Assert.Equal(1, (await OfficePdfArchiveWorkflow.RunAsync(scope.Request)).Failed);
        var repaired = await OfficePdfArchiveWorkflow.RunAsync(scope.Request with { RetryFailed = true });
        Assert.Equal(1, repaired.Completed);
        Assert.Equal(0, repaired.Reused);
    }

    [Fact]
    public async Task MissingReceiptAndChangedOutputNeverAuthorizeReplacement() {
        using var scope = new ArchiveDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "original");
        string output = Path.Combine(scope.Output, "one.txt.pdf");
        Directory.CreateDirectory(scope.Output);
        File.WriteAllText(output, "user-owned existing output");
        Assert.Equal(1, (await OfficePdfArchiveWorkflow.RunAsync(scope.Request)).Failed);
        Assert.Equal("user-owned existing output", File.ReadAllText(output));
        File.Delete(output);
        Assert.Equal(1, (await OfficePdfArchiveWorkflow.RunAsync(scope.Request)).Completed);
        File.WriteAllText(output, "changed after completion");
        Assert.Equal(1, (await OfficePdfArchiveWorkflow.RunAsync(scope.Request with { RetryFailed = true })).Failed);
        Assert.Equal("changed after completion", File.ReadAllText(output));
    }

    [Fact]
    public async Task ConfigurationChangesAndConcurrentOwnersFailClosed() {
        using var scope = new ArchiveDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "original");
        await OfficePdfArchiveWorkflow.RunAsync(scope.Request);
        await Assert.ThrowsAsync<InvalidDataException>(() => OfficePdfArchiveWorkflow.RunAsync(scope.Request with { TabSize = 4 }));
        using var lease = new FileStream(Path.Combine(scope.State, "archive.lock"), FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        await Assert.ThrowsAsync<IOException>(() => OfficePdfArchiveWorkflow.RunAsync(scope.Request));
    }

    [Fact]
    public async Task NestedOutputsAndUnknownRequestFieldsAreRejected() {
        using var scope = new ArchiveDirectory();
        await Assert.ThrowsAsync<ArgumentException>(() => OfficePdfArchiveWorkflow.RunAsync(scope.Request with {
            OutputDirectory = Path.Combine(scope.Input, "output")
        }));
        Assert.Throws<System.Text.Json.JsonException>(() => OfficePdfArchiveSerializer.ParseRequest(
            Encoding.UTF8.GetBytes("{\"InputDirectory\":\"input\",\"OutputDirectory\":\"out\",\"CheckpointDirectory\":\"state\",\"AllowLoss\":true}")));
    }

    private sealed class ArchiveDirectory : IDisposable {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-archive-" + Guid.NewGuid().ToString("N"));
        public ArchiveDirectory() { Directory.CreateDirectory(Input); }
        public string Input => Path.Combine(_root, "input");
        public string Output => Path.Combine(_root, "output");
        public string State => Path.Combine(_root, "state");
        public OfficePdfArchiveRequest Request => new() { InputDirectory = Input, OutputDirectory = Output, CheckpointDirectory = State };
        public void Dispose() => Directory.Delete(_root, true);
    }
}
