using System.Text.Json;
using System.Text.Json.Nodes;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OfficeBatchOptionsTests {
    [Fact]
    public void DefinitionRequiresSchemaAndRejectsOversizedOrEmptyPaths() {
        using var scope = new JobDirectory();
        Assert.Throws<NotSupportedException>(() => OfficeBatchOptions.Parse("{}", scope.Root));
        Assert.Throws<NotSupportedException>(() => OfficeBatchOptions.Parse("{\"schema\":\"officeimo.conversion.job.v99\"}", scope.Root));
        Assert.Throws<InvalidDataException>(() => OfficeBatchOptions.Parse(new string(' ', OfficeBatchOptionsSerializer.MaximumJobBytes + 1), scope.Root));
        Assert.Throws<InvalidDataException>(() => (scope.Job() with { OutputDirectory = string.Empty }).ToJson(scope.Root));
    }

    [Fact]
    public void SavedJobRetainsRelativePathsAndPortableSettingsWithoutProcessDirectory() {
        using var scope = new JobDirectory();
        var request = scope.Job() with { ResumeFrom = Path.Combine(scope.State, "plan.json"),
            InputPaths = [Path.Combine(scope.Input, "notes.txt")], MaxParallelism = 3,
            ConversionOptions = new() { PlainText = new PdfTextImportOptions { EncodingName = "utf-16", TabSize = 4 } } };
        string json = request.ToJson(scope.Root);
        Assert.DoesNotContain(scope.Root.Replace("\\", "\\\\"), json);
        OfficeBatchOptions restored = OfficeBatchOptions.Parse(json, scope.Root);
        Assert.Equal(scope.Input, restored.InputDirectory); Assert.Equal(scope.Output, restored.OutputDirectory);
        Assert.Equal(request.ResumeFrom, restored.ResumeFrom); Assert.Equal(request.InputPaths, restored.InputPaths);
        Assert.Equal(3, restored.MaxParallelism); Assert.Equal("utf-16", restored.ConversionOptions.PlainText!.EncodingName);
        Assert.Equal(4, restored.ConversionOptions.PlainText.TabSize);
    }

    [Fact]
    public void CapturedJobUsesPortableSeparatorsWhenRebased() {
        using var scope = new JobDirectory(); string source = Path.Combine(scope.Input, "nested");
        var request = scope.Job() with { InputDirectory = source, InputPaths = [Path.Combine(source, "notes.txt")],
            OutputDirectory = Path.Combine(scope.Output, "pdf"), ResumeFrom = Path.Combine(scope.State, "checkpoints.json") };
        string jobDirectory = Path.Combine(scope.Root, "jobs", "saved"); string json = request.ToJson(jobDirectory);
        using (var document = JsonDocument.Parse(json)) {
            JsonElement batch = document.RootElement.GetProperty("batch");
            Assert.Equal("../../input/nested", batch.GetProperty("inputDirectory").GetString());
            Assert.Equal("../../input/nested/notes.txt", batch.GetProperty("inputPaths")[0].GetString());
            Assert.Equal("../../output/pdf", batch.GetProperty("outputDirectory").GetString());
            Assert.Equal("../../state/checkpoints.json", batch.GetProperty("resumeFrom").GetString());
        }
        string relocated = Path.Combine(scope.Root, "relocated");
        OfficeBatchOptions restored = OfficeBatchOptions.Parse(json, Path.Combine(relocated, "jobs", "saved"));
        Assert.Equal(Path.Combine(relocated, "input", "nested"), restored.InputDirectory);
        Assert.Equal(Path.Combine(restored.InputDirectory!, "notes.txt"), Assert.Single(restored.InputPaths!));
        Assert.Equal(Path.Combine(relocated, "output", "pdf"), restored.OutputDirectory);
        Assert.Equal(Path.Combine(relocated, "state", "checkpoints.json"), restored.ResumeFrom);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    public void DefinitionRejectsUnknownDuplicateSecretAndNumericFields(int invalidCase) {
        using var scope = new JobDirectory(); string json = scope.Job().ToJson(scope.Root);
        if (invalidCase == 1) json = json.Insert(1, "\"schema\":\"officeimo.conversion.job.v2\",");
        else {
            JsonNode data = JsonNode.Parse(json)!; JsonObject batch = data["batch"]!.AsObject();
            if (invalidCase == 0) batch["unknownSetting"] = true;
            if (invalidCase == 2) batch["pdfPassword"] = "must-not-be-stored";
            if (invalidCase == 3) batch["outputProfile"] = 0;
            if (invalidCase == 4) data["settings"]!["nativeOptions"] = new JsonObject();
            json = data.ToJsonString();
        }
        Assert.Throws<JsonException>(() => OfficeBatchOptions.Parse(json, scope.Root));
    }

    [Fact]
    public async Task RuntimePasswordsRequireReferencesAndAdvancedSettingsAreNeverDropped() {
        using var scope = new JobDirectory(); var request = scope.Job();
        request.ConversionOptions.PdfLoad = new() { Password = "runtime-value" };
        Assert.Throws<NotSupportedException>(() => request.ToJson(scope.Root)); request.ConversionOptions.PdfLoad = null;
        request.ConversionOptions.Word = new() { Title = "A custom title" };
        Assert.Throws<NotSupportedException>(() => request.ToJson(scope.Root)); request.ConversionOptions.Word = null;
        request.ConversionOptions.PlainText = new() { PdfOptions = new PdfCreateOptions { DefaultFontSize = global::OfficeIMO.OfficeLength.Points(16) } };
        Assert.Throws<NotSupportedException>(() => request.ToJson(scope.Root)); request.ConversionOptions.PlainText = null;
        request.PdfPasswordEnvironmentVariable = "OFFICEIMO_TEST_PASSWORD";
        string json = request.ToJson(scope.Root); Assert.Contains("OFFICEIMO_TEST_PASSWORD", json); Assert.DoesNotContain("runtime-value", json);
        OfficeBatchOptions restored = OfficeBatchOptions.Parse(json, scope.Root);
        await Assert.ThrowsAsync<InvalidOperationException>(() => OfficeConverter.ConvertAsync(restored));
        restored.ResolveEnvironmentVariable = name => name == "OFFICEIMO_TEST_PASSWORD" ? "runtime-value" : null;
        OfficeBatchReport result = await OfficeConverter.ConvertAsync(restored);
        Assert.Equal(0, result.Summary.Selected); Assert.Null(restored.ConversionOptions.PdfLoad);
    }

    [Fact]
    public async Task PlanningFindsUnsupportedInputsCollisionsAndLimitsWithoutCreatingState() {
        using var scope = new JobDirectory(); scope.Source("notes.txt", "Real source"); scope.Source("unsupported.bin", "Unselected");
        var job = scope.Job() with { ResumeFrom = Path.Combine(scope.State, "plan.json"), SourcePasswordEnvironmentVariable = "NOT_RESOLVED_WHILE_PLANNING" };
        OfficeBatchPlan plan = await OfficeConverter.PlanAsync(job);
        Assert.True(plan.CanRun); Assert.Equal(1, plan.Selected); Assert.Equal(1, plan.Skipped); Assert.Equal(2, plan.Discovered);
        Assert.False(Directory.Exists(scope.Output)); Assert.False(Directory.Exists(scope.State));
        Assert.Equal("txt-pdf", Assert.Single(plan.Items, item => item.Disposition == "Selected").RouteId);
        job.ResumeFrom = null; Directory.CreateDirectory(scope.Output);
        string existing = Path.Combine(scope.Output, "notes.pdf"); File.WriteAllText(existing, "Keep this destination");
        plan = await OfficeConverter.PlanAsync(job);
        Assert.False(plan.CanRun); Assert.Equal(1, plan.Collisions); Assert.Equal(1, plan.Failed);
        Assert.Equal("Keep this destination", File.ReadAllText(existing));
        job.MaximumInputBytes = 1; plan = await OfficeConverter.PlanAsync(job);
        Assert.False(plan.CanRun); Assert.Contains(plan.Items, item => item.Disposition == "Blocked" && item.Message.Contains("byte limit"));
    }

    [Fact]
    public async Task PlanBoundsObservationsAndReportsEmptyOrMissingSelections() {
        using var scope = new JobDirectory();
        for (int index = 0; index < 260; index++) scope.Source(index + ".bin", "ignored");
        OfficeBatchPlan plan = await OfficeConverter.PlanAsync(scope.Job());
        Assert.False(plan.CanRun); Assert.Equal(0, plan.Selected); Assert.Equal(0, plan.Failed); Assert.Equal(260, plan.Skipped);
        Assert.Equal(250, plan.Items.Count); Assert.True(plan.ItemsTruncated); Assert.Contains(plan.Diagnostics, note => note.Contains("No input files"));
        plan = await OfficeConverter.PlanAsync(scope.Job() with { InputDirectory = Path.Combine(scope.Root, "missing") });
        Assert.False(plan.CanRun); Assert.Equal(1, plan.Failed); Assert.Empty(plan.Items);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfficeConverter.PlanAsync(scope.Job(), cancellationToken: cancelled.Token));
    }

    [Fact]
    public async Task EmptyInputPlanCreatesNoOutputOrCheckpointState() {
        using var scope = new JobDirectory();
        OfficeBatchPlan plan = await OfficeConverter.PlanAsync(scope.Job() with { ResumeFrom = Path.Combine(scope.State, "plan.json") });
        Assert.False(plan.CanRun); Assert.Equal(0, plan.Discovered); Assert.Equal(0, plan.Selected); Assert.Equal(0, plan.Failed); Assert.Empty(plan.Items);
        Assert.Contains(plan.Diagnostics, note => note.Contains("No input files"));
        Assert.False(Directory.Exists(scope.Output)); Assert.False(Directory.Exists(scope.State));
    }

    [Fact]
    public async Task SavedJobExecutesAndReusesOnlyVerifiedCheckpointArtifacts() {
        using var scope = new JobDirectory(); scope.Source("notes.txt", "Saved conversion job proof");
        var job = scope.Job() with { ResumeFrom = Path.Combine(scope.State, "plan.json") };
        string path = Path.Combine(scope.Root, "conversion.officeimo.json"); job.Save(path);
        OfficeBatchReport first = await OfficeConverter.ConvertAsync(OfficeBatchOptions.Load(path));
        Assert.Equal(1, first.Summary.Completed); Assert.Equal(0, first.Summary.Failed);
        string output = Path.Combine(scope.Output, "notes.pdf"); byte[] original = File.ReadAllBytes(output);
        using (var pdf = PdfDocument.Load(output)) Assert.Contains("Saved conversion job proof", pdf.ExtractText());
        OfficeBatchReport resumed = await OfficeConverter.ConvertAsync(OfficeBatchOptions.Load(path));
        Assert.Equal(1, resumed.Summary.Reused); Assert.Equal(original, File.ReadAllBytes(output));
    }

    [Fact]
    public async Task PlanAppliesRouteProfilesAndRejectsDuplicateDestinations() {
        using var scope = new JobDirectory(); string first = scope.Source("notes.txt", "First source");
        OfficeBatchPlan plan = await OfficeConverter.PlanAsync(scope.Job() with { OutputProfile = OfficeConversionProfile.PrintReady });
        Assert.False(plan.CanRun); Assert.Equal(1, plan.Failed); Assert.Contains(plan.Items, item => item.Disposition == "Blocked" && item.Message.Contains("output profile"));
        string other = Path.Combine(scope.Root, "other"); Directory.CreateDirectory(other); string second = Path.Combine(other, "notes.txt"); File.WriteAllText(second, "Different source");
        plan = await OfficeConverter.PlanAsync(scope.Job() with { InputDirectory = null, InputPaths = [first, second], ResumeFrom = Path.Combine(scope.State, "plan.json") });
        Assert.False(plan.CanRun); Assert.Equal(1, plan.Collisions); Assert.Equal(1, plan.Failed); Assert.Equal(2, plan.Selected);
        Assert.False(Directory.Exists(scope.Output)); Assert.False(Directory.Exists(scope.State));
    }

    [Fact]
    public async Task CancellationRetainsCompletedOutputAndResumesRemainingFiles() {
        using var scope = new JobDirectory(); scope.Source("first.txt", "First retained output"); scope.Source("second.txt", "Second retained output");
        var job = scope.Job() with { ResumeFrom = Path.Combine(scope.State, "plan.json"), MaxParallelism = 1 };
        using var cancellation = new CancellationTokenSource(); var progress = new CancelAfterFirstCompletion(cancellation);
        OfficeBatchReport interrupted = await OfficeConverter.ConvertAsync(job, progress, cancellation.Token);
        Assert.True(interrupted.Summary.Cancelled); Assert.Equal(1, interrupted.Summary.Completed);
        string output = Assert.IsType<string>(progress.OutputPath); byte[] firstBytes = File.ReadAllBytes(output); Assert.NotEmpty(firstBytes);
        OfficeBatchReport resumed = await OfficeConverter.ConvertAsync(job);
        Assert.False(resumed.Summary.Cancelled); Assert.Equal(2, resumed.Summary.Completed); Assert.Equal(1, resumed.Summary.Reused); Assert.Equal(0, resumed.Summary.Failed);
        Assert.Equal(firstBytes, File.ReadAllBytes(output)); Assert.Equal(2, Directory.GetFiles(scope.Output, "*.pdf").Length);
    }

    [Fact]
    public async Task DiscoveryLimitKeepsVerifiedOutputsAndTerminalFailureEvidence() {
        using var scope = new JobDirectory(); scope.Source("first.txt", "First source"); scope.Source("second.txt", "Second source");
        var job = scope.Job() with { ResumeFrom = Path.Combine(scope.State, "plan.json"), MaxParallelism = 1, MaximumFiles = 1 };
        OfficeBatchReport aborted = await OfficeConverter.ConvertAsync(job);
        Assert.Contains("file limit", aborted.Summary.FailureSummary); Assert.Equal(1, aborted.Summary.Completed); Assert.Equal(0, aborted.Summary.Failed); Assert.False(aborted.Summary.Cancelled);
        string output = Assert.Single(Directory.GetFiles(scope.Output, "*.pdf")); byte[] original = File.ReadAllBytes(output);
        Assert.Contains("file limit", OfficeBatchReport.Parse(aborted.ToJson()).Summary.FailureSummary);
        job.MaximumFiles = 2; OfficeBatchReport resumed = await OfficeConverter.ConvertAsync(job);
        Assert.Null(resumed.Summary.FailureSummary); Assert.Equal(2, resumed.Summary.Completed); Assert.Equal(1, resumed.Summary.Reused);
        Assert.Equal(original, File.ReadAllBytes(output)); Assert.Null(OfficeBatchReport.Parse(resumed.ToJson()).Summary.FailureSummary);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DiscoveryFailureDoesNotRelabelSiblingCancellationAsCallerCancellation(bool checkpointed) {
        using var scope = new JobDirectory(); scope.Source("first.txt", "First source"); scope.Source("second.txt", "Second source");
        var job = scope.Job() with { ResumeFrom = checkpointed ? Path.Combine(scope.State, "plan.json") : null, MaxParallelism = 2, MaximumFiles = 1 };
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(30)); var observed = new CapturedItems();
        OfficeBatchSummary result = await OfficeBatchExecutor.RunAsync(new WaitForCancellation(), job, observed, timeout.Token);
        Assert.False(timeout.IsCancellationRequested); Assert.Contains("file limit", result.FailureSummary); Assert.False(result.Cancelled);
        Assert.Equal(1, result.Selected); Assert.Equal(0, result.Completed); Assert.Equal(0, result.Failed);
        Assert.Equal(OfficeBatchItemStatus.Cancelled, Assert.Single(observed.Items).Status); Assert.Empty(Directory.GetFiles(scope.Output, "*.pdf"));
    }

    [Fact]
    public void SavingProtectsInputsAndDefinitions() {
        using var scope = new JobDirectory(); string source = scope.Source("notes.txt", "Keep original"); var job = scope.Job();
        Assert.Throws<InvalidDataException>(() => job.Save(source, overwrite: true)); Assert.Equal("Keep original", File.ReadAllText(source));
        string path = Path.Combine(scope.Root, "O'Brien $report.officeimo.json"); job.Save(path); string before = File.ReadAllText(path); job.Name = "Changed";
        Assert.Throws<IOException>(() => job.Save(path)); Assert.Equal(before, File.ReadAllText(path)); job.Save(path, overwrite: true);
        Assert.Equal("Changed", OfficeBatchOptions.Load(path).Name);
    }

    [Theory]
    [InlineData(0x27)]
    [InlineData(0x2018)]
    [InlineData(0x2019)]
    [InlineData(0x201A)]
    [InlineData(0x201B)]
    public void CopiedCommandsKeepPowerShellApostrophesInsideTheLiteral(int apostropheCode) {
        char apostrophe = (char)apostropheCode; string path = "O" + apostrophe + "Brien $report; & value.json";
        string literal = "'O" + apostrophe + apostrophe + "Brien $report; & value.json'";
        Assert.Equal("Invoke-OfficeConversionJob -Path " + literal, OfficeBatchCommands.GetPowerShellCommand(path));
        string cliLiteral = OperatingSystem.IsWindows() ? literal : apostrophe == '\'' ? "'O'\\''Brien $report; & value.json'" : "'" + path + "'";
        Assert.Equal("officeimo workflow job run " + cliLiteral, OfficeBatchCommands.GetCliCommand(path));
    }

    private sealed class JobDirectory : IDisposable {
        internal string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-conversion-job-" + Guid.NewGuid().ToString("N"));
        internal string Input => Path.Combine(Root, "input"); internal string Output => Path.Combine(Root, "output"); internal string State => Path.Combine(Root, "state");
        internal JobDirectory() => Directory.CreateDirectory(Input);
        internal OfficeBatchOptions Job() => new() { InputDirectory = Input, OutputDirectory = Output };
        internal string Source(string name, string content) { string path = Path.Combine(Input, name); File.WriteAllText(path, content); return path; }
        public void Dispose() { if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true); }
    }
    private sealed class CancelAfterFirstCompletion(CancellationTokenSource cancellation) : IProgress<OfficeBatchItemReport> {
        internal string? OutputPath { get; private set; }
        public void Report(OfficeBatchItemReport item) { if (item.Status == OfficeBatchItemStatus.Completed) { OutputPath = item.OutputPath; cancellation.Cancel(); } }
    }
    private sealed class CapturedItems : IProgress<OfficeBatchItemReport> {
        internal System.Collections.Concurrent.ConcurrentQueue<OfficeBatchItemReport> Items { get; } = new();
        public void Report(OfficeBatchItemReport item) => Items.Enqueue(item);
    }
    private sealed class WaitForCancellation : IOfficeConversionExecutor {
        private readonly OfficeConversionExecutor _inner = new();
        public IReadOnlyList<OfficeConversionExecutionRoute> ConversionRoutes => _inner.ConversionRoutes;
        public bool SupportsDirectoryPackageInput(string routeId) => _inner.SupportsDirectoryPackageInput(routeId);
        public async Task<OfficeConversionExecutionResult> RunAsync(OfficeConversionExecutionRequest request, CancellationToken cancellationToken = default) {
            try { await Task.Delay(Timeout.Infinite, cancellationToken); }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) { return await _inner.RunAsync(request, cancellationToken); }
            throw new InvalidOperationException("Discovery failure must cancel the waiting sibling.");
        }
    }
}
