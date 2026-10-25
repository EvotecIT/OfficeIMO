using System.Text.Json;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeConversionJobTests {
    [Fact]
    public void DefinitionRequiresSchemaAndRejectsOversizedOrEmptyPathCapture() {
        Assert.Throws<JsonException>(() => OfficeConversionJobSerializer.Deserialize("{\"inputDirectory\":\"input\",\"outputDirectory\":\"output\"}"));
        Assert.Throws<NotSupportedException>(() => OfficeConversionJobSerializer.Deserialize("{\"schema\":\"officeimo.conversion.job.v99\",\"inputDirectory\":\"input\",\"outputDirectory\":\"output\"}"));
        Assert.Throws<InvalidDataException>(() => OfficeConversionJobSerializer.Deserialize(new string(' ', OfficeConversionJobSerializer.MaximumJobBytes + 1)));
        using var scope = new JobDirectory();
        Assert.Throws<ArgumentException>(() => OfficeConversionJob.FromRequest(new OfficeConversionBatchRequest {
            InputDirectory = scope.Input, OutputDirectory = string.Empty
        }, scope.Root));
    }

    [Fact]
    public void SavedJobRetainsRelativePathsAndPortableSettingsWithoutDependingOnProcessDirectory() {
        using var scope = new JobDirectory();
        var request = new OfficeConversionBatchRequest {
            InputDirectory = scope.Input, OutputDirectory = scope.Output, CheckpointDirectory = scope.State,
            InputPaths = [Path.Combine(scope.Input, "notes.txt")], MaximumConcurrency = 3,
            ConversionOptions = new OfficeWorkflowConversionOptions {
                PlainText = new PdfPlainTextOptions { EncodingName = "utf-16", TabSize = 4 }
            }
        };
        OfficeConversionJob job = OfficeConversionJob.FromRequest(request, scope.Root);
        string json = OfficeConversionJobSerializer.Serialize(job);
        Assert.DoesNotContain(scope.Root.Replace("\\", "\\\\"), json);
        Assert.Contains("\"outputProfile\": \"Faithful\"", json);
        OfficeConversionBatchRequest restored = OfficeConversionJobSerializer.Deserialize(json).CreateRequest(scope.Root);
        Assert.Equal(scope.Input, restored.InputDirectory);
        Assert.Equal(scope.Output, restored.OutputDirectory);
        Assert.Equal(scope.State, restored.CheckpointDirectory);
        Assert.Equal(request.InputPaths, restored.InputPaths);
        Assert.Equal(3, restored.MaximumConcurrency);
        Assert.Equal("utf-16", restored.ConversionOptions.PlainText!.EncodingName);
        Assert.Equal(4, restored.ConversionOptions.PlainText.TabSize);
    }

    [Fact]
    public void CapturedJobUsesPortableSeparatorsWhenRebasedToAnotherDirectory() {
        using var scope = new JobDirectory();
        string source = Path.Combine(scope.Input, "nested");
        var request = new OfficeConversionBatchRequest {
            InputDirectory = source, InputPaths = [Path.Combine(source, "notes.txt")],
            OutputDirectory = Path.Combine(scope.Output, "pdf"), CheckpointDirectory = Path.Combine(scope.State, "checkpoints")
        };
        string jobDirectory = Path.Combine(scope.Root, "jobs", "saved");
        string json = OfficeConversionJobSerializer.Serialize(OfficeConversionJob.FromRequest(request, jobDirectory));
        using (var document = JsonDocument.Parse(json)) {
            Assert.Equal("../../input/nested", document.RootElement.GetProperty("inputDirectory").GetString());
            Assert.Equal("../../input/nested/notes.txt", document.RootElement.GetProperty("inputPaths")[0].GetString());
            Assert.Equal("../../output/pdf", document.RootElement.GetProperty("outputDirectory").GetString());
            Assert.Equal("../../state/checkpoints", document.RootElement.GetProperty("checkpointDirectory").GetString());
        }
        string relocatedRoot = Path.Combine(scope.Root, "relocated");
        OfficeConversionBatchRequest restored = OfficeConversionJobSerializer.Deserialize(json)
            .CreateRequest(Path.Combine(relocatedRoot, "jobs", "saved"));
        Assert.Equal(Path.Combine(relocatedRoot, "input", "nested"), restored.InputDirectory);
        Assert.Equal(Path.Combine(restored.InputDirectory!, "notes.txt"), Assert.Single(restored.InputPaths!));
        Assert.Equal(Path.Combine(relocatedRoot, "output", "pdf"), restored.OutputDirectory);
        Assert.Equal(Path.Combine(relocatedRoot, "state", "checkpoints"), restored.CheckpointDirectory);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void DefinitionRejectsAmbiguousUnknownAndSecretBearingFields(int invalidCase) {
        string suffix = invalidCase switch {
            0 => ",\"unknownSetting\":true",
            1 => ",\"outputDirectory\":\"second-output\"",
            2 => ",\"pdfPassword\":\"must-not-be-stored\"",
            3 => ",\"outputProfile\":0",
            _ => ",\"options\":{\"nativeOptions\":{}}"
        };
        Assert.Throws<JsonException>(() => OfficeConversionJobSerializer.Deserialize("{\"schema\":\"officeimo.conversion.job.v1\",\"inputDirectory\":\"input\",\"outputDirectory\":\"output\"" + suffix + "}"));
    }

    [Fact]
    public void RuntimePasswordsRequireReferencesAndAdvancedSettingsAreNeverSilentlyDropped() {
        using var scope = new JobDirectory();
        var request = new OfficeConversionBatchRequest { InputDirectory = scope.Input, OutputDirectory = scope.Output, PdfPassword = "runtime-value" };
        Assert.Throws<NotSupportedException>(() => OfficeConversionJob.FromRequest(request, scope.Root));
        request.PdfPassword = null;
        request.ConversionOptions.Word = new OfficeIMO.Word.Pdf.WordToPdfOptions { Title = "A custom title" };
        Assert.Throws<NotSupportedException>(() => OfficeConversionJob.FromRequest(request, scope.Root));
        request.ConversionOptions.Word = null;
        request.ConversionOptions.PlainText = new PdfPlainTextOptions { PdfOptions = new PdfOptions { DefaultFontSize = 16 } };
        Assert.Throws<NotSupportedException>(() => OfficeConversionJob.FromRequest(request, scope.Root));
        var job = scope.Job();
        job.Secrets.PdfPasswordEnvironmentVariable = "OFFICEIMO_TEST_PASSWORD";
        string json = OfficeConversionJobSerializer.Serialize(job);
        Assert.Contains("OFFICEIMO_TEST_PASSWORD", json);
        Assert.DoesNotContain("runtime-value", json);
        Assert.Throws<InvalidOperationException>(() => job.CreateRequest(scope.Root, new OfficeConversionJobRuntimeOptions()));
        OfficeConversionBatchRequest resolved = job.CreateRequest(scope.Root, new OfficeConversionJobRuntimeOptions {
            ResolveEnvironmentVariable = name => name == "OFFICEIMO_TEST_PASSWORD" ? "runtime-value" : null
        });
        Assert.Equal("runtime-value", resolved.PdfPassword);
        Assert.Null(job.CreateRequest(scope.Root).PdfPassword);
    }

    [Fact]
    public async Task PlanningFindsCollisionsUnsupportedInputsAndLimitsWithoutCreatingOutputOrCheckpointState() {
        using var scope = new JobDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "notes.txt"), "A real source");
        File.WriteAllText(Path.Combine(scope.Input, "unsupported.bin"), "Unselected source");
        var job = scope.Job();
        job.CheckpointDirectory = "state";
        job.Secrets.SourcePasswordEnvironmentVariable = "NOT_RESOLVED_WHILE_PLANNING";
        OfficeConversionJobPlan plan = await OfficeConversionJobRunner.PlanAsync(job, scope.Root, new OfficeWorkflowRunner());
        Assert.True(plan.CanRun);
        Assert.Equal(1, plan.Selected);
        Assert.Equal(1, plan.Skipped);
        Assert.Equal(2, plan.Discovered);
        Assert.False(Directory.Exists(scope.Output));
        Assert.False(Directory.Exists(scope.State));
        Assert.Equal("txt-pdf", Assert.Single(plan.Items, item => item.Disposition == "Selected").RouteId);
        job.CheckpointDirectory = null;
        Directory.CreateDirectory(scope.Output);
        string existing = Path.Combine(scope.Output, "notes.txt.pdf");
        File.WriteAllText(existing, "Keep this destination");
        plan = await OfficeConversionJobRunner.PlanAsync(job, scope.Root, new OfficeWorkflowRunner());
        Assert.False(plan.CanRun);
        Assert.Equal(1, plan.Collisions);
        Assert.Equal(1, plan.Failed);
        Assert.Equal("Keep this destination", File.ReadAllText(existing));
        job.MaximumInputBytes = 1;
        plan = await OfficeConversionJobRunner.PlanAsync(job, scope.Root, new OfficeWorkflowRunner());
        Assert.False(plan.CanRun);
        Assert.Contains(plan.Items, item => item.Disposition == "Blocked" && item.Message.Contains("byte limit"));
    }

    [Fact]
    public async Task PlanRetainsBoundedObservationsAndReportsAnEmptyEligibleSelection() {
        using var scope = new JobDirectory();
        for (int index = 0; index < 260; index++) File.WriteAllText(Path.Combine(scope.Input, index + ".bin"), "ignored");
        OfficeConversionJobPlan plan = await OfficeConversionJobRunner.PlanAsync(scope.Job(), scope.Root, new OfficeWorkflowRunner());
        Assert.False(plan.CanRun);
        Assert.Equal(0, plan.Selected);
        Assert.Equal(0, plan.Failed);
        Assert.Equal(260, plan.Skipped);
        Assert.Equal(250, plan.Items.Count);
        Assert.True(plan.ItemsTruncated);
        Assert.Contains(plan.Diagnostics, note => note.Contains("No input files"));
        var invalid = scope.Job(); invalid.InputDirectory = "missing-folder";
        plan = await OfficeConversionJobRunner.PlanAsync(invalid, scope.Root, new OfficeWorkflowRunner());
        Assert.False(plan.CanRun); Assert.Equal(1, plan.Failed); Assert.Empty(plan.Items);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfficeConversionJobRunner.PlanAsync(scope.Job(), scope.Root,
            new OfficeWorkflowRunner(), cancellationToken: cancelled.Token));
    }

    [Fact]
    public async Task EmptyInputPlanReportsNoFailuresWithoutCreatingOutputOrCheckpointState() {
        using var scope = new JobDirectory();
        var job = scope.Job(); job.CheckpointDirectory = "state";
        OfficeConversionJobPlan plan = await OfficeConversionJobRunner.PlanAsync(job, scope.Root, new OfficeWorkflowRunner());
        Assert.False(plan.CanRun);
        Assert.Equal(0, plan.Discovered);
        Assert.Equal(0, plan.Selected);
        Assert.Equal(0, plan.Failed);
        Assert.Empty(plan.Items);
        Assert.Contains(plan.Diagnostics, note => note.Contains("No input files"));
        Assert.False(Directory.Exists(scope.Output));
        Assert.False(Directory.Exists(scope.State));
    }

    [Fact]
    public async Task SavedJobExecutesRealConversionAndReusesVerifiedCheckpointArtifacts() {
        using var scope = new JobDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "notes.txt"), "Saved conversion job proof");
        var job = scope.Job(); job.CheckpointDirectory = "state";
        string path = Path.Combine(scope.Root, "conversion.officeimo.json");
        OfficeConversionJobSerializer.SaveFile(job, path);
        var runner = new OfficeWorkflowRunner();
        OfficeConversionBatchResult first = await OfficeConversionJobRunner.RunAsync(OfficeConversionJobSerializer.ReadFile(path), scope.Root, runner);
        Assert.Equal(1, first.Completed); Assert.Equal(0, first.Failed);
        string output = Path.Combine(scope.Output, "notes.txt.pdf");
        byte[] original = File.ReadAllBytes(output);
        using (var pdf = UglyToad.PdfPig.PdfDocument.Open(output)) Assert.Contains("Saved conversion job proof", pdf.GetPage(1).Text);
        OfficeConversionBatchResult resumed = await OfficeConversionJobRunner.RunAsync(OfficeConversionJobSerializer.ReadFile(path), scope.Root, runner);
        Assert.Equal(1, resumed.Reused); Assert.Equal(original, File.ReadAllBytes(output));
    }

    [Fact]
    public async Task PlanAppliesRouteProfilesAndRejectsDuplicateCheckpointDestinations() {
        using var scope = new JobDirectory();
        string first = Path.Combine(scope.Input, "notes.txt");
        File.WriteAllText(first, "First source");
        var job = scope.Job(); job.OutputProfile = OfficeWorkflowOutputProfile.PrintReady;
        var runner = new OfficeWorkflowRunner();
        OfficeConversionJobPlan plan = await OfficeConversionJobRunner.PlanAsync(job, scope.Root, runner);
        Assert.False(plan.CanRun); Assert.Equal(1, plan.Failed);
        Assert.Contains(plan.Items, item => item.Disposition == "Blocked" && item.Message.Contains("output profile"));
        string other = Path.Combine(scope.Root, "other"); Directory.CreateDirectory(other);
        string second = Path.Combine(other, "notes.txt"); File.WriteAllText(second, "Different source");
        job = new OfficeConversionJob { InputPaths = [first, second], OutputDirectory = scope.Output, CheckpointDirectory = scope.State };
        plan = await OfficeConversionJobRunner.PlanAsync(job, scope.Root, runner);
        Assert.False(plan.CanRun); Assert.Equal(1, plan.Collisions); Assert.Equal(1, plan.Failed);
        Assert.Equal(2, plan.Selected);
        Assert.False(Directory.Exists(scope.Output)); Assert.False(Directory.Exists(scope.State));
    }

    [Fact]
    public async Task SavedJobCancellationRetainsCompletedOutputAndResumesRemainingFiles() {
        using var scope = new JobDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "first.txt"), "First retained output");
        File.WriteAllText(Path.Combine(scope.Input, "second.txt"), "Second retained output");
        var job = scope.Job(); job.CheckpointDirectory = "state"; job.MaximumConcurrency = 1;
        var runner = new OfficeWorkflowRunner();
        using var cancellation = new CancellationTokenSource();
        var progress = new CancelAfterFirstCompletion(cancellation);
        OfficeConversionBatchResult interrupted = await OfficeConversionJobRunner.RunAsync(job, scope.Root, runner,
            progress: progress, cancellationToken: cancellation.Token);
        Assert.True(interrupted.Cancelled); Assert.Equal(1, interrupted.Completed);
        string output = Assert.IsType<string>(progress.OutputPath);
        byte[] firstBytes = File.ReadAllBytes(output);
        Assert.NotEmpty(firstBytes);
        OfficeConversionBatchResult resumed = await OfficeConversionJobRunner.RunAsync(job, scope.Root, runner);
        Assert.False(resumed.Cancelled); Assert.Equal(2, resumed.Completed); Assert.Equal(1, resumed.Reused); Assert.Equal(0, resumed.Failed);
        Assert.Equal(firstBytes, File.ReadAllBytes(output));
        Assert.Equal(2, Directory.GetFiles(scope.Output, "*.pdf").Length);
    }

    [Fact]
    public async Task DiscoveryLimitReturnsTerminalSummaryAndKeepsVerifiedOutputsForResume() {
        using var scope = new JobDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "first.txt"), "First real source");
        File.WriteAllText(Path.Combine(scope.Input, "second.txt"), "Second real source");
        var job = scope.Job(); job.CheckpointDirectory = "state"; job.MaximumConcurrency = 1; job.MaximumFiles = 1;
        var runner = new OfficeWorkflowRunner();
        OfficeConversionBatchResult aborted = await OfficeConversionJobRunner.RunAsync(job, scope.Root, runner);
        Assert.NotNull(aborted.FailureSummary); Assert.Contains("file limit", aborted.FailureSummary);
        Assert.Equal(1, aborted.Selected); Assert.Equal(1, aborted.Completed); Assert.Equal(0, aborted.Failed); Assert.False(aborted.Cancelled);
        string output = Assert.Single(Directory.GetFiles(scope.Output, "*.pdf"));
        byte[] original = File.ReadAllBytes(output);
        using (var pdf = UglyToad.PdfPig.PdfDocument.Open(output)) Assert.Single(pdf.GetPages());
        using (var report = JsonDocument.Parse(OfficeConversionBatchSerializer.SerializeResult(aborted)))
            Assert.Contains("file limit", report.RootElement.GetProperty("FailureSummary").GetString());
        job.MaximumFiles = 2;
        OfficeConversionBatchResult resumed = await OfficeConversionJobRunner.RunAsync(job, scope.Root, runner);
        Assert.Null(resumed.FailureSummary); Assert.Equal(2, resumed.Completed); Assert.Equal(1, resumed.Reused); Assert.Equal(0, resumed.Failed);
        Assert.Equal(original, File.ReadAllBytes(output));
        using var completedReport = JsonDocument.Parse(OfficeConversionBatchSerializer.SerializeResult(resumed));
        Assert.False(completedReport.RootElement.TryGetProperty("FailureSummary", out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DiscoveryFailureDoesNotReportSiblingCancellationAsCallerCancellation(bool checkpointed) {
        using var scope = new JobDirectory();
        File.WriteAllText(Path.Combine(scope.Input, "first.txt"), "First source");
        File.WriteAllText(Path.Combine(scope.Input, "second.txt"), "Second source");
        var job = scope.Job();
        job.CheckpointDirectory = checkpointed ? "state" : null;
        job.MaximumConcurrency = 2;
        job.MaximumFiles = 1;
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30));
        var observed = new CapturedBatchItems();
        OfficeConversionBatchResult result = await OfficeConversionJobRunner.RunAsync(job, scope.Root, new WaitForCancellationRunner(),
            progress: observed, cancellationToken: timeout.Token);
        Assert.False(timeout.IsCancellationRequested);
        Assert.NotNull(result.FailureSummary);
        Assert.Contains("file limit", result.FailureSummary);
        Assert.False(result.Cancelled);
        Assert.Equal(1, result.Selected);
        Assert.Equal(0, result.Completed);
        Assert.Equal(0, result.Failed);
        Assert.Equal(OfficeWorkflowStatus.Cancelled, Assert.Single(observed.Items).Status);
        Assert.Empty(Directory.GetFiles(scope.Output, "*.pdf"));
    }

    [Fact]
    public void SavingProtectsInputsAndJobDefinitions() {
        using var scope = new JobDirectory();
        string source = Path.Combine(scope.Input, "notes.txt"); File.WriteAllText(source, "Keep original");
        var job = scope.Job();
        Assert.Throws<InvalidDataException>(() => OfficeConversionJobSerializer.SaveFile(job, source, overwrite: true));
        Assert.Equal("Keep original", File.ReadAllText(source));
        string path = Path.Combine(scope.Root, "O'Brien $report.officeimo.json");
        OfficeConversionJobSerializer.SaveFile(job, path);
        string before = File.ReadAllText(path);
        job.Name = "Changed";
        Assert.Throws<IOException>(() => OfficeConversionJobSerializer.SaveFile(job, path));
        Assert.Equal(before, File.ReadAllText(path));
        OfficeConversionJobSerializer.SaveFile(job, path, overwrite: true);
        Assert.Equal("Changed", OfficeConversionJobSerializer.ReadFile(path).Name);
    }

    [Theory]
    [InlineData(0x27)]
    [InlineData(0x2018)]
    [InlineData(0x2019)]
    [InlineData(0x201A)]
    [InlineData(0x201B)]
    public void CopiedCommandsKeepEveryPowerShellApostropheInsideTheLiteralPath(int apostropheCode) {
        char apostrophe = (char)apostropheCode;
        string path = "O" + apostrophe + "Brien $report; & value.json";
        string powerShellLiteral = "'O" + apostrophe + apostrophe + "Brien $report; & value.json'";
        Assert.Equal("Invoke-OfficeConversionJob -Path " + powerShellLiteral, OfficeConversionJobCommands.GetPowerShellCommand(path));
        string cliLiteral = OperatingSystem.IsWindows() ? powerShellLiteral : apostrophe == '\''
            ? "'O'\\''Brien $report; & value.json'" : "'" + path + "'";
        Assert.Equal("officeimo workflow job run " + cliLiteral, OfficeConversionJobCommands.GetCliCommand(path));
    }

    private sealed class JobDirectory : IDisposable {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-conversion-job-" + Guid.NewGuid().ToString("N"));
        internal string Input => Path.Combine(Root, "input");
        internal string Output => Path.Combine(Root, "output");
        internal string State => Path.Combine(Root, "state");
        internal JobDirectory() => Directory.CreateDirectory(Input);
        internal OfficeConversionJob Job() => new() { InputDirectory = "input", OutputDirectory = "output" };
        public void Dispose() { if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true); }
    }

    private sealed class CancelAfterFirstCompletion(CancellationTokenSource cancellation) : IProgress<OfficeConversionBatchItemResult> {
        internal string? OutputPath { get; private set; }
        public void Report(OfficeConversionBatchItemResult item) {
            if (item.Status != OfficeWorkflowStatus.Completed) return;
            OutputPath = item.OutputPath;
            cancellation.Cancel();
        }
    }

    private sealed class CapturedBatchItems : IProgress<OfficeConversionBatchItemResult> {
        internal System.Collections.Concurrent.ConcurrentQueue<OfficeConversionBatchItemResult> Items { get; } = new();
        public void Report(OfficeConversionBatchItemResult item) => Items.Enqueue(item);
    }

    private sealed class WaitForCancellationRunner : IOfficeWorkflowRunner {
        private readonly OfficeWorkflowRunner _inner = new();
        public async Task<OfficeWorkflowResult> RunAsync(OfficeWorkflowRequest request,
            IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) {
            try {
                await Task.Delay(Timeout.Infinite, cancellationToken);
            } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                return await _inner.RunAsync(request, progress, cancellationToken);
            }
            throw new InvalidOperationException("The batch must cancel the waiting sibling after discovery fails.");
        }

        public Task<IReadOnlyList<OfficeWorkflowResult>> RunBatchAsync(IEnumerable<OfficeWorkflowRequest> requests,
            IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) =>
            _inner.RunBatchAsync(requests, progress, cancellationToken);
    }
}
