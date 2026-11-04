using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Nodes;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OfficeBatchRunReportTests {
    [Fact]
    public async Task RunningReportIsVisibleBeforePublicationAndBecomesCompleted() {
        using var scope = new RunDirectory();
        scope.Source("source.txt", "Recorded conversion");
        bool entered = false;
        scope.Job.PublicationGuard = new CallbackGuard((path, directory) => {
            if (!directory && path.EndsWith(".pdf", StringComparison.Ordinal)) {
                entered = true;
                OfficeBatchReport running = OfficeBatchReport.Load(scope.ReportPath);
                Assert.Equal(OfficeBatchRunState.Running, running.State);
                Assert.Null(running.FinishedUtc);
                Assert.True(OfficeBatchReport.IsActive(scope.ReportPath));
                Assert.False(OfficeBatchReport.TryDeleteFinished(scope.ReportPath));
            }
            return true;
        });
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(scope.Job);
        Assert.True(entered);
        Assert.Equal(OfficeBatchRunState.Completed, report.State);
        Assert.Null(report.RecordingError); Assert.NotNull(report.FinishedUtc);
        Assert.True(report.Succeeded); Assert.Single(report.Items);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(scope.JobPath))), report.JobSha256);
        OfficeBatchReport reopened = OfficeBatchReport.Load(scope.ReportPath);
        Assert.Equal(report.RunId, reopened.RunId);
        Assert.Equal(1, reopened.Summary.Completed);
        Assert.False(OfficeBatchReport.IsActive(scope.ReportPath));
    }

    [Fact]
    public void AbandonedRunningReportIsInterruptedRegardlessOfTimestamp() {
        using var scope = new RunDirectory();
        OfficeBatchReport running = scope.Report() with { StartedUtc = DateTimeOffset.UtcNow.AddHours(1) };
        File.WriteAllText(scope.ReportPath, running.ToJson());
        Assert.Equal(OfficeBatchRunState.Running, OfficeBatchReport.Parse(File.ReadAllText(scope.ReportPath)).State);
        Assert.Equal(OfficeBatchRunState.Interrupted, OfficeBatchReport.Load(scope.ReportPath).State);
        File.WriteAllText(scope.ReportPath + ".lease", running.RunId);
        Assert.Equal(OfficeBatchRunState.Interrupted, OfficeBatchReport.Load(scope.ReportPath).State);
        Assert.True(OfficeBatchReport.TryDeleteFinished(scope.ReportPath));
        Assert.False(File.Exists(scope.ReportPath)); Assert.False(File.Exists(scope.ReportPath + ".lease"));
        Assert.True(OfficeBatchReport.TryDeleteFinished(scope.ReportPath));
    }

    [Fact]
    public void ReaderPreservesTerminalReportPublishedBeforeLeaseReleaseObservation() {
        using var scope = new RunDirectory();
        var running = scope.Report();
        var terminal = running with { State = OfficeBatchRunState.Completed, FinishedUtc = DateTimeOffset.UtcNow,
            Summary = new(1, 1, 0, 0, false), Succeeded = true };
        File.WriteAllText(scope.ReportPath, running.ToJson());
        using var lease = new FileStream(scope.ReportPath + ".lease", FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
        bool probed = false;
        OfficeBatchReport observed = OfficeBatchReportStore.ReadFile(scope.ReportPath, path => {
            Assert.True(OfficeBatchReport.IsActive(path)); probed = true;
            string staged = Path.Combine(scope.Root, "terminal.json");
            File.WriteAllText(staged, terminal.ToJson());
            File.Move(staged, scope.ReportPath, overwrite: true); lease.Dispose();
            return OfficeBatchReport.IsActive(path);
        });
        Assert.True(probed); Assert.Equal(OfficeBatchRunState.Completed, observed.State);
        Assert.Equal(running.RunId, observed.RunId); Assert.Equal(1, observed.Summary.Completed);
        Assert.Equal(terminal.FinishedUtc, observed.FinishedUtc);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void MalformedDiagnosticsAreRejectedWithoutBreakingSnapshotProtection(int invalidCase) {
        using var scope = new RunDirectory();
        var report = scope.Report() with { State = OfficeBatchRunState.Completed, FinishedUtc = DateTimeOffset.UtcNow,
            Summary = new(1, 1, 0, 0, false), Items = [new("source.txt", "output.pdf", OfficeBatchItemStatus.Completed,
                false, "Completed", [new(OfficeDiagnosticCodes.TableCellContentClipped, OfficeDiagnosticSeverity.Warning, "A layout warning")])] };
        string json = report.ToJson();
        string expectedHash = Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(json)));
        JsonNode data = JsonNode.Parse(json)!;
        JsonObject diagnostic = data["items"]![0]!["diagnostics"]![0]!.AsObject();
        diagnostic[invalidCase % 2 == 0 ? "code" : "message"] = invalidCase < 2 ? null : " ";
        string malformed = data.ToJsonString();
        var error = Assert.Throws<InvalidDataException>(() => OfficeBatchReport.Parse(malformed));
        Assert.IsAssignableFrom<ArgumentException>(error.InnerException);
        File.WriteAllText(scope.ReportPath, malformed);
        Assert.Throws<InvalidDataException>(() => OfficeBatchReport.Load(scope.ReportPath));
        Assert.False(OfficeBatchReportStore.MatchesSnapshot(scope.ReportPath, expectedHash));
        Assert.Throws<IOException>(() => OfficeBatchReportStore.DeleteOwned(scope.ReportPath, expectedHash));
        Assert.Equal(malformed, File.ReadAllText(scope.ReportPath));
    }

    [Fact]
    public void InvalidUtf8ReportIsRejectedWithoutBreakingSnapshotProtection() {
        using var scope = new RunDirectory();
        byte[] malformed = [0xff]; File.WriteAllBytes(scope.ReportPath, malformed);
        Assert.Throws<InvalidDataException>(() => OfficeBatchReport.Load(scope.ReportPath));
        Assert.False(OfficeBatchReportStore.MatchesSnapshot(scope.ReportPath, new string('0', 64)));
        Assert.Throws<IOException>(() => OfficeBatchReportStore.DeleteOwned(scope.ReportPath, new string('0', 64)));
        Assert.Equal(malformed, File.ReadAllBytes(scope.ReportPath));
    }

    [Fact]
    public async Task RecordedRetryPreservesFirstAttemptAndVerifiedCompletedOutput() {
        using var scope = new RunDirectory(checkpointed: true);
        scope.Source("first.txt", "First retained source");
        string second = scope.Source("second.txt", new string('x', 256));
        scope.Job.MaximumInputBytes = 64; scope.SaveJob();
        OfficeBatchReport first = await OfficeConverter.ConvertAsync(scope.Job);
        Assert.Equal(OfficeBatchRunState.Failed, first.State);
        Assert.Equal(1, first.Summary.Completed); Assert.Equal(1, first.Summary.Failed);
        byte[] firstReport = File.ReadAllBytes(scope.ReportPath);
        string firstOutput = Path.Combine(scope.Output, "first.pdf"); byte[] firstBytes = File.ReadAllBytes(firstOutput);
        File.WriteAllText(second, "Repaired second source"); scope.Job.RetryFailed = true;
        scope.Job.ReportPath = Path.Combine(scope.Root, "retry.json");
        OfficeBatchReport retried = await OfficeConverter.ConvertAsync(scope.Job);
        Assert.Equal(OfficeBatchRunState.Completed, retried.State); Assert.True(retried.RetryFailed);
        Assert.Equal(2, retried.Summary.Completed); Assert.Equal(1, retried.Summary.Reused); Assert.Equal(0, retried.Summary.Failed);
        Assert.NotEqual(first.RunId, retried.RunId); Assert.Equal(first.JobSha256, retried.JobSha256);
        Assert.Equal(firstReport, File.ReadAllBytes(scope.ReportPath)); Assert.Equal(firstBytes, File.ReadAllBytes(firstOutput));
        using var pdf = OfficeIMO.Pdf.PdfDocument.Load(Path.Combine(scope.Output, "second.pdf"));
        Assert.Contains("Repaired second source", pdf.ExtractText());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public async Task InitialRecordingFailuresPreventPublication(int invalidCase) {
        using var scope = new RunDirectory(checkpointed: true); scope.Source("source.txt", "Protected source");
        string destination = invalidCase switch { 0 => scope.ReportPath, 1 => scope.JobPath,
            2 => Path.Combine(scope.Input, "report.json"), 3 => Path.Combine(scope.Output, "report.json"),
            _ => Path.Combine(scope.Job.CheckpointDirectory!, "report.json") };
        if (invalidCase == 0) File.WriteAllText(destination, "Unrelated existing file");
        byte[] jobBytes = File.ReadAllBytes(scope.JobPath); bool entered = false;
        scope.Job.ReportPath = destination;
        scope.Job.PublicationGuard = new CallbackGuard((path, directory) => { if (!directory && path.EndsWith(".pdf", StringComparison.Ordinal)) entered = true; return true; });
        if (invalidCase == 0) await Assert.ThrowsAsync<IOException>(() => OfficeConverter.ConvertAsync(scope.Job));
        else await Assert.ThrowsAsync<InvalidDataException>(() => OfficeConverter.ConvertAsync(scope.Job));
        Assert.False(entered); Assert.Equal(jobBytes, File.ReadAllBytes(scope.JobPath));
        Assert.False(Directory.Exists(scope.Output)); Assert.False(Directory.Exists(scope.Job.CheckpointDirectory));
        if (invalidCase == 0) Assert.Equal("Unrelated existing file", File.ReadAllText(destination));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task FinalRecordingFailurePreservesConversionOutcomeAndChangedFile(bool replaceReport) {
        using var scope = new RunDirectory(); scope.Source("source.txt", "Completed source");
        int reportChecks = 0;
        scope.Job.PublicationGuard = new CallbackGuard((path, _) => path != scope.ReportPath || replaceReport || ++reportChecks == 1);
        var progress = new BatchProgress(_ => { if (replaceReport) File.WriteAllText(scope.ReportPath, "External replacement must survive"); });
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(scope.Job, progress);
        Assert.Equal(OfficeBatchRunState.Completed, report.State); Assert.True(report.Succeeded);
        Assert.Equal(1, report.Summary.Completed); Assert.NotEmpty(report.RecordingError!);
        if (replaceReport) Assert.Equal("External replacement must survive", File.ReadAllText(scope.ReportPath));
        else Assert.Equal(OfficeBatchRunState.Interrupted, OfficeBatchReport.Load(scope.ReportPath).State);
    }

    [Fact]
    public async Task UnexpectedExecutionFailureWritesFailedHistoryAndRethrowsOriginal() {
        using var scope = new RunDirectory(); var original = new InvalidOperationException("Original host failure");
        scope.Job.PublicationGuard = new CallbackGuard((path, directory) => directory && path == scope.Output ? throw original : true);
        Exception caught = await Assert.ThrowsAsync<InvalidOperationException>(() => OfficeConverter.ConvertAsync(scope.Job));
        Assert.Same(original, caught);
        OfficeBatchReport report = OfficeBatchReport.Load(scope.ReportPath);
        Assert.Equal(OfficeBatchRunState.Failed, report.State); Assert.Contains("InvalidOperationException", report.FailureSummary);
        Assert.False(report.Succeeded); Assert.Equal(0, report.Summary.Failed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CallerCancellationAndDiscoveryFailureRetainTheirOwnStates(bool callerCancelled) {
        using var scope = new RunDirectory(checkpointed: true);
        scope.Source("first.txt", "First source"); scope.Source("second.txt", "Second source");
        scope.Job.MaxParallelism = 1; scope.Job.MaximumFiles = callerCancelled ? 2 : 1; scope.SaveJob();
        using var cancellation = new CancellationTokenSource();
        var progress = new BatchProgress(item => { if (callerCancelled && item.Status == OfficeBatchItemStatus.Completed) cancellation.Cancel(); });
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(scope.Job, progress, cancellation.Token);
        Assert.Equal(callerCancelled ? OfficeBatchRunState.Cancelled : OfficeBatchRunState.Failed, report.State);
        Assert.Equal(callerCancelled, report.Summary.Cancelled); Assert.Equal(1, report.Summary.Completed); Assert.Equal(0, report.Summary.Failed);
        if (!callerCancelled) Assert.Contains("file limit", report.Summary.FailureSummary);
        Assert.Equal(report.State, OfficeBatchReport.Load(scope.ReportPath).State); Assert.Single(Directory.GetFiles(scope.Output, "*.pdf"));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public async Task RetentionOptOutRemovesOwnedHistoryWithoutCancellingConversion(int disabledPhase) {
        using var scope = new RunDirectory(); scope.Source("source.txt", "Retained conversion");
        bool retain = disabledPhase != 0; int reportChecks = 0;
        scope.Job.RetainReport = () => retain;
        scope.Job.PublicationGuard = new CallbackGuard((path, _) => { if (path == scope.ReportPath && ++reportChecks == 2 && disabledPhase == 2) retain = false; return true; });
        var progress = new BatchProgress(_ => { Assert.Equal(disabledPhase != 0, File.Exists(scope.ReportPath)); if (disabledPhase == 1) retain = false; });
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(scope.Job, progress);
        Assert.Equal(1, report.Summary.Completed); Assert.Null(report.RecordingError);
        Assert.False(File.Exists(scope.ReportPath)); Assert.False(File.Exists(scope.ReportPath + ".lease")); Assert.True(File.Exists(scope.JobPath));
    }

    [Fact]
    public async Task PrivacyOptOutPreservesAReportReplacedByAnotherWriter() {
        using var scope = new RunDirectory(); scope.Source("source.txt", "Converted source"); bool retain = true;
        const string replacement = "Another writer's replacement must survive privacy cleanup";
        scope.Job.RetainReport = () => retain;
        var progress = new BatchProgress(_ => { File.WriteAllText(scope.ReportPath, replacement); retain = false; });
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(scope.Job, progress);
        Assert.Equal(OfficeBatchRunState.Completed, report.State); Assert.Equal(1, report.Summary.Completed);
        Assert.NotEmpty(report.RecordingError!); Assert.Equal(replacement, File.ReadAllText(scope.ReportPath));
        Assert.False(File.Exists(scope.ReportPath + ".lease"));
    }

    [Fact]
    public void LeaseCleanupPreservesDifferentOwnerAndRemovesOnlyMatchingSnapshot() {
        using var scope = new RunDirectory(); string previous = Guid.NewGuid().ToString("N"), replacement = Guid.NewGuid().ToString("N");
        string lease = scope.ReportPath + ".lease"; File.WriteAllText(lease, replacement);
        OfficeBatchReportStore.DeleteLease(lease, previous); Assert.Equal(replacement, File.ReadAllText(lease));
        OfficeBatchReportStore.DeleteLease(lease, replacement); Assert.False(File.Exists(lease));
    }

    [Fact]
    public async Task RecordedExecutionAdmitsOnlyObservationalAndRetryOverrides() {
        using var scope = new RunDirectory(); scope.Job.MaxParallelism = 3;
        await Assert.ThrowsAsync<InvalidDataException>(() => OfficeConverter.ConvertAsync(scope.Job)); Assert.False(File.Exists(scope.ReportPath));
        scope.Job.MaxParallelism = 2; scope.Job.RetryFailed = true; scope.Job.AllowEmpty = true;
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(scope.Job);
        Assert.True(report.RetryFailed); Assert.True(report.AllowEmpty); Assert.True(report.Succeeded);
        Assert.False(OfficeBatchOptions.Load(scope.JobPath).RetryFailed);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task EmptySelectionPolicyIsRecordedWithOutcome(bool allowEmpty) {
        using var scope = new RunDirectory(); scope.Job.AllowEmpty = allowEmpty;
        OfficeBatchReport report = await OfficeConverter.ConvertAsync(scope.Job);
        Assert.Equal(allowEmpty, report.AllowEmpty); Assert.Equal(allowEmpty, report.Succeeded);
        Assert.Equal(allowEmpty ? OfficeBatchRunState.Completed : OfficeBatchRunState.Failed, report.State);
        Assert.Equal(0, report.Summary.Selected); Assert.Equal(0, report.Summary.Failed);
    }

    [Fact]
    public void StrictReaderAndFinishedDeletionPreserveMalformedAndForeignContent() {
        using var scope = new RunDirectory(); string json = scope.Report().ToJson();
        Assert.Throws<JsonException>(() => OfficeBatchReport.Parse(json.Insert(1, "\"extra\":true,")));
        Assert.Throws<JsonException>(() => OfficeBatchReport.Parse(json.Insert(1, "\"state\":\"Running\",")));
        Assert.Throws<JsonException>(() => OfficeBatchReport.Parse(json.Replace("\"Running\"", "0")));
        File.WriteAllText(scope.ReportPath, "Foreign content");
        Assert.ThrowsAny<JsonException>(() => OfficeBatchReport.TryDeleteFinished(scope.ReportPath)); Assert.Equal("Foreign content", File.ReadAllText(scope.ReportPath));
        File.WriteAllText(scope.ReportPath, json); File.WriteAllText(scope.ReportPath + ".lease", new string('a', 32));
        Assert.Throws<InvalidDataException>(() => OfficeBatchReport.TryDeleteFinished(scope.ReportPath)); Assert.Equal(json, File.ReadAllText(scope.ReportPath));
        using (var file = new FileStream(scope.ReportPath, FileMode.Create, FileAccess.Write)) file.SetLength(OfficeBatchReportSerializer.MaximumReportBytes + 1L);
        Assert.Throws<InvalidDataException>(() => OfficeBatchReport.Load(scope.ReportPath));
    }

    private sealed class RunDirectory : IDisposable {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-run-report-" + Guid.NewGuid().ToString("N"));
        internal string Input => Path.Combine(Root, "input");
        internal string Output => Path.Combine(Root, "output");
        internal string JobPath => Path.Combine(Root, "job.json");
        internal string ReportPath => Path.Combine(Root, "run.json");
        internal OfficeBatchOptions Job { get; private set; }
        internal RunDirectory(bool checkpointed = false) {
            Directory.CreateDirectory(Input);
            Job = new() { InputDirectory = Input, OutputDirectory = Output,
                ResumeFrom = checkpointed ? Path.Combine(Root, "state", "plan.json") : null };
            SaveJob();
        }
        internal void SaveJob() { Job.Save(JobPath, overwrite: File.Exists(JobPath)); Job = OfficeBatchOptions.Load(JobPath); Job.ReportPath = ReportPath; }
        internal string Source(string name, string content) { string path = Path.Combine(Input, name); File.WriteAllText(path, content); return path; }
        internal OfficeBatchReport Report() => new() {
            Summary = new(0, 0, 0, 0, false), RunId = Guid.NewGuid().ToString("N"), StartedUtc = DateTimeOffset.UtcNow, JobPath = JobPath,
            JobSha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(JobPath))), OutputDirectory = Output,
            ResumeFrom = Job.ResumeFrom, State = OfficeBatchRunState.Running
        };
        public void Dispose() { if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true); }
    }

    private sealed class BatchProgress(Action<OfficeBatchItemReport> action) : IProgress<OfficeBatchItemReport> {
        public void Report(OfficeBatchItemReport value) => action(value);
    }
    private sealed class CallbackGuard(Func<string, bool, bool> callback) : IOfficePublicationGuard {
        public Task<bool> CanPublishAsync(string path, bool directory, CancellationToken cancellationToken) => Task.FromResult(callback(path, directory));
    }
}
