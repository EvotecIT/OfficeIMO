using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioConversionRunStoreTests {
    [Fact]
    public void HistoryPreferenceStopsNewReportsAndRestorationWhileExplicitClearPreservesDocuments() {
        using var files = new TestFiles();
        var preferences = new StudioPreferencesService(new JsonStudioPreferencesStore(files.Paths.PreferencesPath));
        var history = new StudioDocumentHistory(files.Paths, preferences);
        var store = history.ConversionRuns;
        string reportPath = WriteReport(store, files, DateTimeOffset.UtcNow);
        byte[] reportBytes = File.ReadAllBytes(reportPath);
        string unrelated = Path.Combine(files.Paths.ConversionRunsRoot, "notes.json");
        File.WriteAllText(unrelated, "Unrelated file");

        history.SetRememberHistory(false);

        Assert.False(store.CanRetainReport);
        Assert.Null(store.CreateReportPath());
        Assert.Empty(store.ReadLatest().Entries);
        Assert.Equal(reportBytes, File.ReadAllBytes(reportPath));
        var result = history.Clear();
        Assert.True(result.ConversionRuns); Assert.True(result.Succeeded);
        Assert.False(File.Exists(reportPath));
        Assert.Equal("Unrelated file", File.ReadAllText(unrelated));
        Assert.Equal("Keep the job", File.ReadAllText(files.Job));
        Assert.Equal("Keep the output", File.ReadAllText(files.Output));
        Assert.Equal("Keep the checkpoint", File.ReadAllText(files.Checkpoint));
        Assert.Equal("Keep the recovery", File.ReadAllText(files.Recovery));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnreadableReportIsIsolatedAndReportedWithoutHidingValidAttempts(bool oversized) {
        using var files = new TestFiles();
        var store = new StudioConversionRunStore(files.Paths.ConversionRunsRoot, () => true);
        string valid = WriteReport(store, files, DateTimeOffset.UtcNow.AddMinutes(-1));
        string invalid = store.CreateReportPath()!;
        if (oversized) {
            using var stream = new FileStream(invalid, FileMode.CreateNew, FileAccess.Write);
            stream.SetLength((long)OfficeBatchReportSerializer.MaximumReportBytes + 1);
        } else File.WriteAllText(invalid, "{not a conversion report");

        var restored = store.ReadLatest();

        Assert.Equal(valid, Assert.Single(restored.Entries).ReportPath);
        Assert.NotNull(restored.Error);
        Assert.Contains(Path.GetFileName(invalid), restored.Error);
        Assert.InRange(restored.Error.Length, 1, 1024);
        Assert.False(store.ClearFinished());
        Assert.False(File.Exists(valid));
        Assert.True(File.Exists(invalid));
    }

    [Fact]
    public void RetentionKeepsLatestFiftyAttemptsAndRejectsRemovalOutsideItsOwnStore() {
        using var files = new TestFiles();
        var store = new StudioConversionRunStore(files.Paths.ConversionRunsRoot, () => true);
        DateTimeOffset started = DateTimeOffset.UtcNow.AddHours(-1);
        string[] reports = Enumerable.Range(0, StudioConversionRunStore.MaximumEntries + 2)
            .Select(index => WriteReport(store, files, started.AddSeconds(index))).ToArray();

        Assert.True(store.TrimFinished());

        Assert.Equal(StudioConversionRunStore.MaximumEntries, store.ReadLatest().Entries.Count);
        Assert.False(File.Exists(reports[0])); Assert.False(File.Exists(reports[1]));
        Assert.Equal(reports.Skip(2).Reverse(), store.ReadLatest().Entries.Select(entry => entry.ReportPath));
        Assert.False(store.RemoveFinished(files.Job));
        Assert.Equal("Keep the job", File.ReadAllText(files.Job));
        Assert.True(store.RemoveFinished(reports[0]));
    }

    [Fact]
    public async Task ClearPreservesLiveReportAndMidRunOptOutRemovesOnlyItsHistory() {
        using var files = new TestFiles();
        var preferences = new StudioPreferencesService(new JsonStudioPreferencesStore(files.Paths.PreferencesPath));
        var history = new StudioDocumentHistory(files.Paths, preferences);
        string source = Path.Combine(files.Root, "source");
        Directory.CreateDirectory(source);
        File.WriteAllText(Path.Combine(source, "one.txt"), "A recorded document");
        string jobPath = Path.Combine(files.Root, "run.json"), output = Path.Combine(files.Root, "run-output");
        var definition = new OfficeBatchOptions { InputDirectory = "source", OutputDirectory = "run-output" };
        (definition).Save(jobPath);
        string reportPath = history.ConversionRuns.CreateReportPath()!;
        var guard = new PausedPdfPublicationGuard();
        Task<OfficeBatchReport> execution = OfficeConverter.ConvertAsync(OfficeBatchOptions.Load(jobPath) with {
            ReportPath = reportPath, PublicationGuard = guard, RetainReport = () => history.ConversionRuns.CanRetainReport
        });
        try {
            await guard.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
            Assert.Equal(OfficeBatchRunState.Running, Assert.Single(history.ConversionRuns.ReadLatest().Entries).Report.State);
            Assert.True(history.ConversionRuns.ClearFinished());
            Assert.True(File.Exists(reportPath));
            Assert.False(history.ConversionRuns.RemoveFinished(reportPath));
            history.SetRememberHistory(false);
        } finally {
            guard.Release.TrySetResult();
            await execution.WaitAsync(TimeSpan.FromSeconds(10));
        }

        var report = await execution.WaitAsync(TimeSpan.FromSeconds(10));

        Assert.Equal(OfficeBatchRunState.Completed, report.State);
        Assert.False(File.Exists(reportPath));
        Assert.True(File.Exists(Path.Combine(output, "one.pdf")));
        Assert.True(File.Exists(jobPath));
        Assert.Equal("Keep the recovery", File.ReadAllText(files.Recovery));
    }

    private static string WriteReport(StudioConversionRunStore store, TestFiles files, DateTimeOffset started) {
        string path = store.CreateReportPath()!;
        var report = new OfficeBatchReport {
            RunId = Guid.NewGuid().ToString("N"), StartedUtc = started, FinishedUtc = started.AddSeconds(1),
            JobPath = files.Job, JobSha256 = new string('A', 64), OutputDirectory = Path.GetDirectoryName(files.Output)!,
            ResumeFrom = files.Checkpoint, State = OfficeBatchRunState.Completed,
            Summary = new(0, 0, 0, 0, false), Succeeded = true
        };
        File.WriteAllText(path, OfficeBatchReportSerializer.Serialize(report));
        File.SetLastWriteTimeUtc(path, started.UtcDateTime);
        return path;
    }

    private sealed class TestFiles : IDisposable {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-run-store-" + Guid.NewGuid().ToString("N"));
        internal StudioDataPaths Paths { get; }
        internal string Job { get; }
        internal string Output { get; }
        internal string Checkpoint { get; }
        internal string Recovery { get; }
        internal TestFiles() {
            Directory.CreateDirectory(Root);
            Paths = new StudioDataPaths(Path.Combine(Root, "profile"));
            Job = Path.Combine(Root, "job.json"); File.WriteAllText(Job, "Keep the job");
            Output = Path.Combine(Root, "output", "one.pdf"); Directory.CreateDirectory(Path.GetDirectoryName(Output)!); File.WriteAllText(Output, "Keep the output");
            Checkpoint = Path.Combine(Root, "checkpoint", "receipt.json"); Directory.CreateDirectory(Path.GetDirectoryName(Checkpoint)!); File.WriteAllText(Checkpoint, "Keep the checkpoint");
            Recovery = Path.Combine(Paths.WorkflowRecoveryRoot, "recovery.bin"); Directory.CreateDirectory(Paths.WorkflowRecoveryRoot); File.WriteAllText(Recovery, "Keep the recovery");
        }
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }
}
