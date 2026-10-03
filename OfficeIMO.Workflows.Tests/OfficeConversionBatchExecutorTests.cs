using System.Text;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeConversionBatchExecutorTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task PublicationRejectsStageChangesDuringTheHostGuard(bool resume, bool replace) {
        using var scope = new BatchDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "Validated source");
        if (resume) {
            using var cancellation = new CancellationTokenSource();
            await new OfficeWorkflowRunner().RunBatchAsync(scope.Request, cancellationToken: cancellation.Token,
                publicationGuard: new CancelBeforeFinalMove(cancellation));
        }
        var result = await new OfficeWorkflowRunner().RunBatchAsync(scope.Request,
            publicationGuard: new ChangePendingStage(scope.Output, resume ? 1 : 2, replace));
        Assert.Equal(1, result.Failed);
        Assert.False(File.Exists(Path.Combine(scope.Output, "one.txt.pdf")));
    }

    private sealed class ChangePendingStage(string directory, int check, bool replace) : IOfficeWorkflowPublicationGuard {
        private int _finalChecks;
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) {
            if (!isDirectory && Path.GetFileName(path) == "one.txt.pdf" && Interlocked.Increment(ref _finalChecks) == check) {
                string staged = Assert.Single(Directory.GetFiles(directory, "*.conversion.pdf"));
                if (replace) File.Delete(staged);
                File.WriteAllText(staged, "Modified during publication policy check");
            }
            return ValueTask.FromResult(true);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PendingPublicationRecoversBeforeAndAfterTheFinalMove(bool alreadyMoved) {
        using var scope = new BatchDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "Recoverable publication");
        using var cancellation = new CancellationTokenSource();
        var interrupted = await new OfficeWorkflowRunner().RunBatchAsync(scope.Request, cancellationToken: cancellation.Token,
            publicationGuard: new CancelBeforeFinalMove(cancellation));
        Assert.True(interrupted.Cancelled);
        string output = Path.Combine(scope.Output, "one.txt.pdf");
        Assert.False(File.Exists(output));
        string staged = Assert.Single(Directory.GetFiles(scope.Output, "*.conversion.pdf"));
        byte[] original = File.ReadAllBytes(staged);
        if (alreadyMoved) File.Move(staged, output); // Process interruption after move, before completion checkpoint.
        var resumed = await new OfficeWorkflowRunner().RunBatchAsync(scope.Request);
        Assert.Equal(1, resumed.Completed); Assert.Equal(1, resumed.Reused); Assert.Equal(0, resumed.Failed);
        Assert.Equal(original, File.ReadAllBytes(output));
        Assert.Empty(Directory.GetFiles(scope.Output, "*.conversion.pdf"));
        Assert.Single(OfficeIMO.Pdf.PdfDocument.Load(output).Read().Pages);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task PendingPublicationRejectsChangedSourceOrArtifact(bool changeSource) {
        using var scope = new BatchDirectory();
        string input = Path.Combine(scope.Input, "one.txt");
        File.WriteAllText(input, "Recoverable publication");
        using var cancellation = new CancellationTokenSource();
        await new OfficeWorkflowRunner().RunBatchAsync(scope.Request, cancellationToken: cancellation.Token,
            publicationGuard: new CancelBeforeFinalMove(cancellation));
        string staged = Assert.Single(Directory.GetFiles(scope.Output, "*.conversion.pdf"));
        File.WriteAllText(changeSource ? input : staged, "Changed after durable intent");
        var resumed = await new OfficeWorkflowRunner().RunBatchAsync(scope.Request with { RetryFailed = true });
        Assert.Equal(1, resumed.Failed); Assert.False(File.Exists(Path.Combine(scope.Output, "one.txt.pdf")));
    }

    private sealed class CancelBeforeFinalMove(CancellationTokenSource cancellation) : IOfficeWorkflowPublicationGuard {
        private int _finalChecks;
        public ValueTask<bool> CanPublishAsync(string path, bool directory, CancellationToken token) {
            if (!directory && Path.GetFileName(path) == "one.txt.pdf" && Interlocked.Increment(ref _finalChecks) == 2)
                cancellation.Cancel();
            return ValueTask.FromResult(true);
        }
    }

    [Fact]
    public async Task ResumeRetainsImportAndRenderingFidelityStages() {
        using var scope = new BatchDirectory();
        File.WriteAllBytes(Path.Combine(scope.Input, "loss.doc"), LegacyDocPdfStageTests.CreateLegacySource());
        var request = scope.Request with { ConversionOptions = new() { LegacyDocLossPolicy = OfficeConversionLossPolicy.Allow } };
        var first = new CaptureItem();
        Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(request, first)).Completed);
        Assert.Contains(first.Item!.Diagnostics, item => item.Stage == "import" && item.Details.ContainsKey("lossKind"));
        Assert.Contains(first.Item.Diagnostics, item => item.Stage == "convert" && item.Code == "NativeFontFamilySubstituted");
        var resumed = new CaptureItem();
        Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(request, resumed)).Reused);
        foreach (var diagnostic in first.Item.Diagnostics.Where(item => item.Details.ContainsKey("lossKind"))) {
            var restored = Assert.Single(resumed.Item!.Diagnostics, item => item.Code == diagnostic.Code && item.Stage == diagnostic.Stage);
            Assert.Equal(diagnostic.Details["lossKind"], restored.Details["lossKind"]);
            Assert.Equal(diagnostic.Details["source"], restored.Details["source"]);
        }
    }

    private sealed class CaptureItem : IProgress<OfficeConversionBatchItemResult> {
        public OfficeConversionBatchItemResult? Item { get; private set; }
        public void Report(OfficeConversionBatchItemResult item) => Item = item;
    }
    [Fact]
    public async Task CancellationKeepsCompletedReceiptsAndResumeFinishesTheRest() {
        using var scope = new BatchDirectory();
        for (int index = 0; index < 4; index++) File.WriteAllText(Path.Combine(scope.Input, $"{index}.txt"), "cancellation evidence");
        using var cancellation = new CancellationTokenSource();
        var interrupted = await new OfficeWorkflowRunner().RunBatchAsync(scope.Request with { MaximumConcurrency = 1 },
            new CancelAfterCompletion(cancellation), cancellation.Token);
        Assert.True(interrupted.Cancelled);
        Assert.Equal(1, interrupted.Completed);
        var resumed = await new OfficeWorkflowRunner().RunBatchAsync(scope.Request);
        Assert.Equal(4, resumed.Completed);
        Assert.Equal(1, resumed.Reused);
        Assert.Equal(0, resumed.Failed);
    }

    [Fact]
    public async Task HostPublicationPolicyIsCheckedBeforeCreatingBatchState() {
        using var scope = new BatchDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "protected publication");
        await Assert.ThrowsAsync<UnauthorizedAccessException>(() => new OfficeWorkflowRunner().RunBatchAsync(scope.Request,
            publicationGuard: new RejectPublication()));
        Assert.False(Directory.Exists(scope.Output));
        Assert.False(Directory.Exists(scope.State));
    }

    private sealed class RejectPublication : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool directory, CancellationToken token) => ValueTask.FromResult(false);
    }
    private sealed class CancelAfterCompletion(CancellationTokenSource cancellation) : IProgress<OfficeConversionBatchItemResult> {
        public void Report(OfficeConversionBatchItemResult item) { if (item.Status == OfficeWorkflowStatus.Completed) cancellation.Cancel(); }
    }

    [Fact]
    public async Task BatchPreservesRelativeNamesAndResumesOnlyVerifiedArtifacts() {
        using var scope = new BatchDirectory();
        Directory.CreateDirectory(Path.Combine(scope.Input, "nested"));
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "first batch item");
        File.WriteAllText(Path.Combine(scope.Input, "nested", "one.txt"), "second batch item");
        var request = scope.Request;
        var first = await new OfficeWorkflowRunner().RunBatchAsync(request);
        Assert.Equal(2, first.Completed);
        Assert.Equal(0, first.Failed);
        Assert.True(File.Exists(Path.Combine(scope.Output, "nested", "one.txt.pdf")));
        string output = Path.Combine(scope.Output, "one.txt.pdf");
        byte[] original = File.ReadAllBytes(output);
        var resumed = await new OfficeWorkflowRunner().RunBatchAsync(request with { MaximumConcurrency = 1 });
        Assert.Equal(2, resumed.Reused);
        Assert.Equal(original, File.ReadAllBytes(output));
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "changed completed source");
        var changed = await new OfficeWorkflowRunner().RunBatchAsync(request with { RetryFailed = true });
        Assert.Equal(1, changed.Failed);
        Assert.Equal(original, File.ReadAllBytes(output));
    }

    [Fact]
    public async Task RecordedFailureRequiresExplicitRetryAndPermitsRepairOfFailedSource() {
        using var scope = new BatchDirectory();
        string input = Path.Combine(scope.Input, "invalid.txt");
        File.WriteAllBytes(input, [0xff]);
        Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(scope.Request)).Failed);
        File.WriteAllText(input, "repaired input");
        Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(scope.Request)).Failed);
        var repaired = await new OfficeWorkflowRunner().RunBatchAsync(scope.Request with { RetryFailed = true });
        Assert.Equal(1, repaired.Completed);
        Assert.Equal(0, repaired.Reused);
    }

    [Fact]
    public async Task MissingReceiptAndChangedOutputNeverAuthorizeReplacement() {
        using var scope = new BatchDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "original");
        string output = Path.Combine(scope.Output, "one.txt.pdf");
        Directory.CreateDirectory(scope.Output);
        File.WriteAllText(output, "user-owned existing output");
        Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(scope.Request)).Failed);
        Assert.Equal("user-owned existing output", File.ReadAllText(output));
        File.Delete(output);
        Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(scope.Request)).Completed);
        File.WriteAllText(output, "changed after completion");
        Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(scope.Request with { RetryFailed = true })).Failed);
        Assert.Equal("changed after completion", File.ReadAllText(output));
    }

    [Fact]
    public async Task ConfigurationChangesAndConcurrentOwnersFailClosed() {
        using var scope = new BatchDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "one.txt"), "original");
        await new OfficeWorkflowRunner().RunBatchAsync(scope.Request);
        Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(scope.Request with { ConversionOptions = new() { PlainText = new() { TabSize = 4 } } })).Failed);
        using var lease = new FileStream(Path.Combine(scope.State, "batch.lock"), FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        await Assert.ThrowsAsync<IOException>(() => new OfficeWorkflowRunner().RunBatchAsync(scope.Request));
    }

    [Fact]
    public async Task NestedOutputAndCheckpointTreesAreRejected() {
        using var scope = new BatchDirectory();
        await Assert.ThrowsAsync<ArgumentException>(() => new OfficeWorkflowRunner().RunBatchAsync(scope.Request with {
            OutputDirectory = Path.Combine(scope.Input, "output")
        }));
    }

    [Fact]
    public async Task OrdinaryBatchRejectsLinkedOutputSubdirectoryBeforeConversion() {
        using var scope = new BatchDirectory();
        string nestedInput = Path.Combine(scope.Input, "nested");
        Directory.CreateDirectory(nestedInput);
        File.WriteAllText(Path.Combine(nestedInput, "one.txt"), "batch source");
        string outside = Path.Combine(Path.GetDirectoryName(scope.Output)!, "outside");
        Directory.CreateDirectory(outside);
        string displacedOutput = Path.Combine(outside, "one.txt.pdf");
        File.WriteAllText(displacedOutput, "keep existing file");
        Directory.CreateDirectory(scope.Output);
        Directory.CreateSymbolicLink(Path.Combine(scope.Output, "nested"), outside);

        OfficeConversionBatchResult result = await new OfficeWorkflowRunner().RunBatchAsync(scope.Request with {
            CheckpointDirectory = null, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
        });

        Assert.Equal(1, result.Failed);
        Assert.Equal("keep existing file", File.ReadAllText(displacedOutput));
    }

    [Fact]
    public async Task OrdinaryBatchRejectsOutputLinkSwappedDuringConversionBeforeStaging() {
        using var scope = new BatchDirectory();
        string nestedInput = Path.Combine(scope.Input, "nested");
        Directory.CreateDirectory(nestedInput);
        File.WriteAllText(Path.Combine(nestedInput, "one.txt"), "batch source");
        string outside = Path.Combine(Path.GetDirectoryName(scope.Output)!, "outside");
        Directory.CreateDirectory(outside);
        string displacedOutput = Path.Combine(outside, "one.txt.pdf");
        File.WriteAllText(displacedOutput, "keep existing file");

        var runner = new SwapOutputBeforeConversion(scope.Output, outside);
        OfficeConversionBatchResult result = await ((IOfficeWorkflowRunner)runner)
            .RunBatchAsync(scope.Request with { CheckpointDirectory = null, ConflictPolicy = OfficeWorkflowConflictPolicy.Replace });

        Assert.Equal(1, result.Failed);
        Assert.False(runner.SawOutsideStage);
        Assert.Equal("keep existing file", File.ReadAllText(displacedOutput));
        Assert.Empty(Directory.GetFiles(outside, "*.tmp"));
    }

    private sealed class SwapOutputBeforeConversion(string outputRoot, string outside) : IOfficeWorkflowRunner {
        private readonly OfficeWorkflowRunner _inner = new();
        private readonly string _outside = outside;
        public bool SawOutsideStage { get; private set; }

        public Task<OfficeWorkflowResult> RunAsync(OfficeWorkflowRequest request,
            IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) {
            string nested = Path.Combine(outputRoot, "nested");
            Directory.CreateSymbolicLink(nested, _outside);
            return _inner.RunAsync(request, new CaptureStaging(this, progress), cancellationToken);
        }

        private sealed class CaptureStaging(SwapOutputBeforeConversion owner, IProgress<OfficeWorkflowProgress>? next)
            : IProgress<OfficeWorkflowProgress> {
            public void Report(OfficeWorkflowProgress value) {
                if (value.Stage == "validate-output" && Directory.GetFiles(owner._outside, "*.tmp").Length != 0)
                    owner.SawOutsideStage = true;
                next?.Report(value);
            }
        }

        public Task<IReadOnlyList<OfficeWorkflowResult>> RunBatchAsync(IEnumerable<OfficeWorkflowRequest> requests,
            IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) =>
            _inner.RunBatchAsync(requests, progress, cancellationToken);
    }

    [Theory]
    [InlineData(".html", false, false)]
    [InlineData(".html", true, false)]
    [InlineData(".md", false, false)]
    [InlineData(".md", true, false)]
    [InlineData(".md", false, true)]
    [InlineData(".md", true, true)]
    public async Task SelectedFileCheckpointsRejectGeneratedTreesInsideResourceRootsBeforeCreatingState(string extension, bool nestedCheckpoint, bool explicitBase) {
        using var scope = new BatchDirectory();
        string input = Path.Combine(scope.Input, "page" + extension);
        File.WriteAllText(input, extension == ".html" ? "<p>Selected HTML</p>" : "# Selected Markdown");
        string resourceRoot = explicitBase ? Path.Combine(scope.Input, "assets") : scope.Input;
        Directory.CreateDirectory(resourceRoot);
        var request = scope.Request with {
            InputDirectory = null, InputPaths = [input],
            OutputDirectory = nestedCheckpoint ? scope.Output : Path.Combine(resourceRoot, "output"),
            CheckpointDirectory = nestedCheckpoint ? Path.Combine(resourceRoot, "state") : scope.State,
            ConversionOptions = new() { Markdown = extension == ".md" ? new() {
                BaseDirectory = explicitBase ? resourceRoot : null,
                ResourcePolicy = new OfficeIMO.Pdf.PdfResourcePolicy { AllowLocalFileAccess = true }
            } : null }
        };
        await Assert.ThrowsAsync<ArgumentException>(() => new OfficeWorkflowRunner().RunBatchAsync(request));
        Assert.False(Directory.Exists(request.OutputDirectory));
        Assert.False(Directory.Exists(request.CheckpointDirectory));
    }

    [Theory]
    [InlineData(".html")]
    [InlineData(".md")]
    public async Task SelectedResourceFilesConvertAndReuseWithSeparateOutputAndCheckpointTrees(string extension) {
        using var scope = new BatchDirectory();
        string input = Path.Combine(scope.Input, "page" + extension);
        File.WriteAllText(input, extension == ".html" ? "<p>Selected HTML</p>" : "# Selected Markdown");
        var request = scope.Request with { InputDirectory = null, InputPaths = [input],
            ConversionOptions = new() { Markdown = extension == ".md" ? new() {
                ResourcePolicy = new OfficeIMO.Pdf.PdfResourcePolicy { AllowLocalFileAccess = true }
            } : null }
        };
        var first = await new OfficeWorkflowRunner().RunBatchAsync(request);
        Assert.Equal(1, first.Completed); Assert.Equal(0, first.Failed);
        Assert.NotEmpty(OfficeIMO.Pdf.PdfDocument.Load(Path.Combine(scope.Output, "page" + extension + ".pdf")).Read().Pages);
        Assert.Equal(1, (await new OfficeWorkflowRunner().RunBatchAsync(request)).Reused);
    }

    private sealed class BatchDirectory : IDisposable {
        private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-pdf-batch-" + Guid.NewGuid().ToString("N"));
        public BatchDirectory() { Directory.CreateDirectory(Input); }
        public string Input => Path.Combine(_root, "input");
        public string Output => Path.Combine(_root, "output");
        public string State => Path.Combine(_root, "state");
        public OfficeConversionBatchRequest Request => new() { InputDirectory = Input, OutputDirectory = Output, CheckpointDirectory = State };
        public void Dispose() => Directory.Delete(_root, true);
    }
}
