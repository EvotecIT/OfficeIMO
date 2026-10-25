using System.Security.Cryptography;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeConversionJobRunReportTests {
    [Fact]
    public async Task RunningReportIsVisibleBeforeExecutionAndBecomesCompleted() {
        using var scope = new RunDirectory();
        bool entered = false;
        var runner = new BatchBoundary((_, progress, _) => {
            entered = true;
            OfficeConversionJobRunReport running = OfficeConversionJobRunReportSerializer.ReadFile(scope.ReportPath);
            Assert.Equal(OfficeConversionJobRunState.Running, running.State);
            Assert.Null(running.Batch); Assert.Null(running.FinishedUtc);
            Assert.True(OfficeConversionJobRunReportSerializer.IsActive(scope.ReportPath));
            Assert.False(OfficeConversionJobRunReportSerializer.TryDeleteFinished(scope.ReportPath));
            progress?.Report(scope.Item());
            return Task.FromResult(new OfficeConversionBatchResult(1, 1, 0, 0, false));
        });
        OfficeConversionJobRunReport report = await OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, scope.ReportPath, runner);
        Assert.True(entered);
        Assert.Equal(OfficeConversionJobRunState.Completed, report.State);
        Assert.Null(report.RecordingError); Assert.NotNull(report.FinishedUtc);
        Assert.True(report.Batch!.Succeeded);
        Assert.Single(report.Batch.Items);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(scope.JobPath))), report.JobSha256);
        OfficeConversionJobRunReport reopened = OfficeConversionJobRunReportSerializer.ReadFile(scope.ReportPath);
        Assert.Equal(report.RunId, reopened.RunId);
        Assert.Equal(OfficeConversionJobRunState.Completed, reopened.State);
        Assert.Equal(1, reopened.Batch!.Summary.Completed);
        Assert.False(OfficeConversionJobRunReportSerializer.IsActive(scope.ReportPath));
    }

    [Fact]
    public void AbandonedRunningReportIsInterruptedRegardlessOfTimestamp() {
        using var scope = new RunDirectory();
        var running = scope.Report() with { StartedUtc = DateTimeOffset.UtcNow.AddHours(1) };
        File.WriteAllText(scope.ReportPath, OfficeConversionJobRunReportSerializer.Serialize(running));
        Assert.Equal(OfficeConversionJobRunState.Running, OfficeConversionJobRunReportSerializer.Deserialize(File.ReadAllText(scope.ReportPath)).State);
        Assert.Equal(OfficeConversionJobRunState.Interrupted, OfficeConversionJobRunReportSerializer.ReadFile(scope.ReportPath).State);
        File.WriteAllText(scope.ReportPath + ".lease", running.RunId);
        Assert.Equal(OfficeConversionJobRunState.Interrupted, OfficeConversionJobRunReportSerializer.ReadFile(scope.ReportPath).State);
        Assert.True(OfficeConversionJobRunReportSerializer.TryDeleteFinished(scope.ReportPath));
        Assert.False(File.Exists(scope.ReportPath));
        Assert.False(File.Exists(scope.ReportPath + ".lease"));
        Assert.True(OfficeConversionJobRunReportSerializer.TryDeleteFinished(scope.ReportPath));
    }

    [Fact]
    public async Task RecordedRetryPreservesFirstAttemptAndVerifiedCompletedOutput() {
        using var scope = new RunDirectory(checkpointed: true);
        File.WriteAllText(Path.Combine(scope.Input, "first.txt"), "First retained source");
        string second = Path.Combine(scope.Input, "second.txt");
        File.WriteAllText(second, new string('x', 256));
        scope.Job.MaximumInputBytes = 64;
        scope.SaveJob();
        var runner = new OfficeWorkflowRunner();
        OfficeConversionJobRunReport first = await OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, scope.ReportPath, runner);
        Assert.Equal(OfficeConversionJobRunState.Failed, first.State);
        Assert.Equal(1, first.Batch!.Summary.Completed); Assert.Equal(1, first.Batch.Summary.Failed);
        byte[] firstReport = File.ReadAllBytes(scope.ReportPath);
        string firstOutput = Path.Combine(scope.Output, "first.txt.pdf");
        byte[] firstBytes = File.ReadAllBytes(firstOutput);
        File.WriteAllText(second, "Repaired second source");
        scope.Job.RetryFailed = true;
        string retryPath = Path.Combine(scope.Root, "retry.json");
        OfficeConversionJobRunReport retried = await OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, retryPath, runner);
        Assert.Equal(OfficeConversionJobRunState.Completed, retried.State);
        Assert.True(retried.RetryFailed);
        Assert.Equal(2, retried.Batch!.Summary.Completed); Assert.Equal(1, retried.Batch.Summary.Reused); Assert.Equal(0, retried.Batch.Summary.Failed);
        Assert.NotEqual(first.RunId, retried.RunId); Assert.Equal(first.JobSha256, retried.JobSha256);
        Assert.Equal(firstReport, File.ReadAllBytes(scope.ReportPath)); Assert.Equal(firstBytes, File.ReadAllBytes(firstOutput));
        Assert.Equal(2, Directory.GetFiles(scope.Output, "*.pdf").Length);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(Path.Combine(scope.Output, "second.txt.pdf"));
        Assert.Contains("Repaired second source", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public async Task InitialRecordingFailuresPreventExecution(int invalidCase) {
        using var scope = new RunDirectory(checkpointed: true);
        string destination = invalidCase switch {
            0 => scope.ReportPath,
            1 => scope.JobPath,
            2 => Path.Combine(scope.Input, "report.json"),
            3 => Path.Combine(scope.Output, "report.json"),
            _ => Path.Combine(scope.State, "report.json")
        };
        if (invalidCase == 0) File.WriteAllText(destination, "Unrelated existing file");
        byte[] jobBytes = File.ReadAllBytes(scope.JobPath);
        bool entered = false;
        var runner = new BatchBoundary((_, _, _) => { entered = true; return Task.FromResult(new OfficeConversionBatchResult(0, 0, 0, 0, false)); });
        if (invalidCase == 0)
            await Assert.ThrowsAsync<IOException>(() => OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, destination, runner));
        else
            await Assert.ThrowsAsync<InvalidDataException>(() => OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, destination, runner));
        Assert.False(entered); Assert.Equal(jobBytes, File.ReadAllBytes(scope.JobPath));
        Assert.False(Directory.Exists(scope.Output)); Assert.False(Directory.Exists(scope.State));
        if (invalidCase == 0) Assert.Equal("Unrelated existing file", File.ReadAllText(destination));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FinalRecordingFailurePreservesReturnedConversionOutcomeAndChangedFile(bool replaceReport) {
        using var scope = new RunDirectory();
        var runner = new BatchBoundary((_, progress, _) => {
            progress?.Report(scope.Item());
            if (replaceReport) File.WriteAllText(scope.ReportPath, "External replacement must survive");
            return Task.FromResult(new OfficeConversionBatchResult(1, 1, 0, 0, false));
        });
        var guard = replaceReport ? null : new InitialReportOnlyGuard();
        OfficeConversionJobRunReport report = await OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, scope.ReportPath, runner, publicationGuard: guard);
        Assert.Equal(OfficeConversionJobRunState.Completed, report.State);
        Assert.True(report.Batch!.Succeeded); Assert.Equal(1, report.Batch.Summary.Completed);
        Assert.NotEmpty(report.RecordingError!);
        if (replaceReport) Assert.Equal("External replacement must survive", File.ReadAllText(scope.ReportPath));
        else Assert.Equal(OfficeConversionJobRunState.Interrupted, OfficeConversionJobRunReportSerializer.ReadFile(scope.ReportPath).State);
    }

    [Fact]
    public async Task UnexpectedExecutionFailureWritesFailedHistoryAndRethrowsOriginal() {
        using var scope = new RunDirectory();
        var original = new InvalidOperationException("Original runner failure");
        var runner = new BatchBoundary((_, _, _) => Task.FromException<OfficeConversionBatchResult>(original));
        Exception caught = await Assert.ThrowsAsync<InvalidOperationException>(() => OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, scope.ReportPath, runner));
        Assert.Same(original, caught);
        OfficeConversionJobRunReport report = OfficeConversionJobRunReportSerializer.ReadFile(scope.ReportPath);
        Assert.Equal(OfficeConversionJobRunState.Failed, report.State);
        Assert.Contains("InvalidOperationException", report.FailureSummary);
        Assert.False(report.Batch!.Succeeded);
        Assert.Equal(0, report.Batch.Summary.Failed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CallerCancellationAndDiscoveryFailureRetainTheirOwnStates(bool callerCancelled) {
        using var scope = new RunDirectory(checkpointed: true);
        File.WriteAllText(Path.Combine(scope.Input, "first.txt"), "First source");
        File.WriteAllText(Path.Combine(scope.Input, "second.txt"), "Second source");
        scope.Job.MaximumConcurrency = 1; scope.Job.MaximumFiles = callerCancelled ? 2 : 1; scope.SaveJob();
        using var cancellation = new CancellationTokenSource();
        var progress = new BatchProgress(item => { if (callerCancelled && item.Status == OfficeWorkflowStatus.Completed) cancellation.Cancel(); });
        OfficeConversionJobRunReport report = await OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, scope.ReportPath,
            new OfficeWorkflowRunner(), progress: progress, cancellationToken: cancellation.Token);
        Assert.Equal(callerCancelled ? OfficeConversionJobRunState.Cancelled : OfficeConversionJobRunState.Failed, report.State);
        Assert.Equal(callerCancelled, report.Batch!.Summary.Cancelled);
        Assert.Equal(1, report.Batch.Summary.Completed); Assert.Equal(0, report.Batch.Summary.Failed);
        if (!callerCancelled) Assert.Contains("file limit", report.Batch.Summary.FailureSummary);
        Assert.Equal(report.State, OfficeConversionJobRunReportSerializer.ReadFile(scope.ReportPath).State);
        Assert.Single(Directory.GetFiles(scope.Output, "*.pdf"));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public async Task RetentionOptOutRemovesOwnedHistoryWithoutCancellingConversion(int disabledPhase) {
        using var scope = new RunDirectory();
        bool retain = disabledPhase != 0;
        int guardCalls = 0;
        var guard = new CallbackGuard(() => { if (++guardCalls == 2 && disabledPhase == 2) retain = false; return true; });
        var runner = new BatchBoundary((_, progress, _) => {
            Assert.Equal(disabledPhase != 0, File.Exists(scope.ReportPath));
            if (disabledPhase == 1) retain = false;
            progress?.Report(scope.Item());
            return Task.FromResult(new OfficeConversionBatchResult(1, 1, 0, 0, false));
        });
        OfficeConversionJobRunReport report = await OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, scope.ReportPath, runner,
            publicationGuard: guard, retainReport: () => retain);
        Assert.Equal(1, report.Batch!.Summary.Completed); Assert.Null(report.RecordingError);
        Assert.False(File.Exists(scope.ReportPath)); Assert.False(File.Exists(scope.ReportPath + ".lease"));
        Assert.True(File.Exists(scope.JobPath));
    }

    [Fact]
    public async Task PrivacyOptOutPreservesAReportReplacedByAnotherWriter() {
        using var scope = new RunDirectory();
        bool retain = true;
        const string replacement = "Another writer's replacement must survive privacy cleanup";
        var runner = new BatchBoundary((_, progress, _) => {
            progress?.Report(scope.Item());
            File.WriteAllText(scope.ReportPath, replacement);
            retain = false;
            return Task.FromResult(new OfficeConversionBatchResult(1, 1, 0, 0, false));
        });
        OfficeConversionJobRunReport report = await OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, scope.ReportPath,
            runner, retainReport: () => retain);
        Assert.Equal(OfficeConversionJobRunState.Completed, report.State);
        Assert.Equal(1, report.Batch!.Summary.Completed);
        Assert.NotEmpty(report.RecordingError!);
        Assert.Equal(replacement, File.ReadAllText(scope.ReportPath));
        Assert.False(File.Exists(scope.ReportPath + ".lease"));
        Assert.Equal(2, Directory.GetFiles(scope.Root).Length);
    }

    [Fact]
    public void LeaseCleanupPreservesDifferentOwnerAndRemovesOnlyMatchingSnapshot() {
        using var scope = new RunDirectory();
        string previousOwner = Guid.NewGuid().ToString("N");
        string replacementOwner = Guid.NewGuid().ToString("N");
        string leasePath = scope.ReportPath + ".lease";
        File.WriteAllText(leasePath, replacementOwner);
        OfficeConversionJobRunReportSerializer.DeleteLease(leasePath, previousOwner);
        Assert.Equal(replacementOwner, File.ReadAllText(leasePath));
        OfficeConversionJobRunReportSerializer.DeleteLease(leasePath, replacementOwner);
        Assert.False(File.Exists(leasePath));
        Assert.Equal(new[] { scope.JobPath }, Directory.GetFiles(scope.Root));
    }

    [Fact]
    public async Task RecordedExecutionAdmitsOnlyExplicitRetryOverride() {
        using var scope = new RunDirectory();
        scope.Job.MaximumConcurrency = 3;
        bool entered = false;
        var runner = new BatchBoundary((_, _, _) => { entered = true; return Task.FromResult(new OfficeConversionBatchResult(0, 0, 0, 0, false)); });
        await Assert.ThrowsAsync<InvalidDataException>(() => OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, scope.ReportPath, runner));
        Assert.False(entered); Assert.False(File.Exists(scope.ReportPath));
        scope.Job.MaximumConcurrency = 2; scope.Job.RetryFailed = true;
        OfficeConversionJobRunReport report = await OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, scope.ReportPath, runner, allowEmpty: true);
        Assert.True(entered); Assert.True(report.RetryFailed); Assert.True(report.AllowEmpty); Assert.True(report.Batch!.Succeeded);
        Assert.False(OfficeConversionJobSerializer.ReadFile(scope.JobPath).RetryFailed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EmptySelectionPolicyIsRecordedWithOutcome(bool allowEmpty) {
        using var scope = new RunDirectory();
        OfficeConversionJobRunReport report = await OfficeConversionJobRunner.RunRecordedAsync(scope.Job, scope.JobPath, scope.ReportPath,
            new OfficeWorkflowRunner(), allowEmpty: allowEmpty);
        Assert.Equal(allowEmpty, report.AllowEmpty); Assert.Equal(allowEmpty, report.Batch!.Succeeded);
        Assert.Equal(allowEmpty ? OfficeConversionJobRunState.Completed : OfficeConversionJobRunState.Failed, report.State);
        Assert.Equal(0, report.Batch.Summary.Selected); Assert.Equal(0, report.Batch.Summary.Failed);
    }

    [Fact]
    public void StrictRunReaderAndFinishedDeletionPreserveMalformedAndForeignContent() {
        using var scope = new RunDirectory();
        var report = scope.Report();
        string json = OfficeConversionJobRunReportSerializer.Serialize(report);
        Assert.Throws<JsonException>(() => OfficeConversionJobRunReportSerializer.Deserialize(json.Insert(1, "\"extra\":true,")));
        Assert.Throws<JsonException>(() => OfficeConversionJobRunReportSerializer.Deserialize(json.Insert(1, "\"state\":\"Running\",")));
        Assert.Throws<JsonException>(() => OfficeConversionJobRunReportSerializer.Deserialize(json.Replace("\"Running\"", "0")));
        File.WriteAllText(scope.ReportPath, "Foreign content");
        Assert.ThrowsAny<JsonException>(() => OfficeConversionJobRunReportSerializer.TryDeleteFinished(scope.ReportPath));
        Assert.Equal("Foreign content", File.ReadAllText(scope.ReportPath));
        File.WriteAllText(scope.ReportPath, json);
        File.WriteAllText(scope.ReportPath + ".lease", new string('a', 32));
        Assert.Throws<InvalidDataException>(() => OfficeConversionJobRunReportSerializer.TryDeleteFinished(scope.ReportPath));
        Assert.Equal(json, File.ReadAllText(scope.ReportPath));
        using (var file = new FileStream(scope.ReportPath, FileMode.Create, FileAccess.Write)) file.SetLength(OfficeConversionJobRunReportSerializer.MaximumReportBytes + 1L);
        Assert.Throws<InvalidDataException>(() => OfficeConversionJobRunReportSerializer.ReadFile(scope.ReportPath));
    }

    private sealed class RunDirectory : IDisposable {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-run-report-" + Guid.NewGuid().ToString("N"));
        internal string Input => Path.Combine(Root, "input");
        internal string Output => Path.Combine(Root, "output");
        internal string State => Path.Combine(Root, "state");
        internal string JobPath => Path.Combine(Root, "job.json");
        internal string ReportPath => Path.Combine(Root, "run.json");
        internal OfficeConversionJob Job { get; }
        internal RunDirectory(bool checkpointed = false) {
            Directory.CreateDirectory(Input);
            Job = new() { InputDirectory = "input", OutputDirectory = "output", CheckpointDirectory = checkpointed ? "state" : null };
            SaveJob();
        }
        internal void SaveJob() => OfficeConversionJobSerializer.SaveFile(Job, JobPath, overwrite: File.Exists(JobPath));
        internal OfficeConversionBatchItemResult Item() => new(Path.Combine(Input, "source.txt"), Path.Combine(Output, "source.txt.pdf"), OfficeWorkflowStatus.Completed, false, "Completed", []);
        internal OfficeConversionJobRunReport Report() => new() {
            RunId = Guid.NewGuid().ToString("N"), StartedUtc = DateTimeOffset.UtcNow, JobPath = JobPath,
            JobSha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(JobPath))), OutputDirectory = Output,
            CheckpointDirectory = Job.CheckpointDirectory == null ? null : State, State = OfficeConversionJobRunState.Running
        };
        public void Dispose() { if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true); }
    }

    private sealed class BatchBoundary(Func<OfficeConversionBatchRequest, IProgress<OfficeConversionBatchItemResult>?, CancellationToken, Task<OfficeConversionBatchResult>> execute)
        : IOfficeWorkflowRunner {
        public Task<OfficeConversionBatchResult> RunBatchAsync(OfficeConversionBatchRequest request, IProgress<OfficeConversionBatchItemResult>? progress = null,
            CancellationToken cancellationToken = default, IOfficeWorkflowPublicationGuard? publicationGuard = null) => execute(request, progress, cancellationToken);
        public Task<OfficeWorkflowResult> RunAsync(OfficeWorkflowRequest request, IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default)
            => throw new NotSupportedException();
        public Task<IReadOnlyList<OfficeWorkflowResult>> RunBatchAsync(IEnumerable<OfficeWorkflowRequest> requests,
            IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) => throw new NotSupportedException();
    }

    private sealed class BatchProgress(Action<OfficeConversionBatchItemResult> action) : IProgress<OfficeConversionBatchItemResult> {
        public void Report(OfficeConversionBatchItemResult value) => action(value);
    }

    private sealed class InitialReportOnlyGuard : IOfficeWorkflowPublicationGuard {
        private int _calls;
        public ValueTask<bool> CanPublishAsync(string absoluteDestination, bool isDirectory, CancellationToken cancellationToken) => new(++_calls == 1);
    }

    private sealed class CallbackGuard(Func<bool> callback) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string absoluteDestination, bool isDirectory, CancellationToken cancellationToken) => new(callback());
    }
}
