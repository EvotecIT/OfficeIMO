using System.Runtime.InteropServices;
using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Workflows.IWork;
using Xunit;

namespace OfficeIMO.Workflows.IWork.Tests;

public sealed class IWorkWorkflowTests {
    [Theory]
    [InlineData("pages", "pages-docx", "docx")]
    [InlineData("numbers", "numbers-xlsx", "xlsx")]
    [InlineData("key", "keynote-pptx", "pptx")]
    public async Task Registered_routes_publish_reopened_destinations_with_typed_source_evidence(string source, string route, string target) {
        using var files = new Files(source, target);
        OfficeWorkflowRunner runner = IWorkWorkflow.CreateRunner();
        Assert.Contains(runner.ConversionRoutes, item => item.Id == route && item.CanExecute);
        OfficeWorkflowResult result = await runner.RunAsync(files.Request(route));
        Assert.True(result.Succeeded, result.Summary);
        Assert.True(File.Exists(files.Output));
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "OutputReopened");
        var snapshot = Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "SourceSnapshot");
        Assert.Equal(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(files.Input))), snapshot.Details["sha256"], ignoreCase: true);
        OfficeWorkflowConversionEvidence evidence = Assert.IsType<OfficeWorkflowConversionEvidence>(result.ConversionEvidence);
        Assert.Equal("EditableReconstruction", evidence.Facts["projectionKind"]);
        Assert.Contains(evidence.FidelityDiagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Unassessed);
        Assert.Throws<InvalidOperationException>(evidence.RequireNoLoss);
        Assert.Equal(new FileInfo(files.Input).Length, result.InputBytes);
    }

    [Fact]
    public async Task Default_runner_keeps_iWork_opt_in() {
        using var files = new Files("numbers", "xlsx");
        var runner = new OfficeWorkflowRunner();
        Assert.DoesNotContain(runner.ConversionRoutes, route => route.Id == "numbers-xlsx");
        OfficeWorkflowResult result = await runner.RunAsync(files.Request("numbers-xlsx"));
        Assert.Equal(OfficeWorkflowFailureKind.ValidationFailed, result.FailureKind);
        Assert.False(File.Exists(files.Output));
    }

    [Fact]
    public async Task Configured_policy_is_captured_and_unknown_visual_coverage_is_rejected() {
        using var files = new Files("pages", "docx");
        var policy = new IWorkConversionOptions { Mode = IWorkConversionMode.VisualOnly, RequireCompleteVisualCoverage = true };
        var reading = new IWorkReadOptions();
        OfficeWorkflowRunner runner = IWorkWorkflow.CreateRunner(reading, policy);
        policy.RequireCompleteVisualCoverage = false;
        policy.Mode = IWorkConversionMode.EditableOnly;
        reading.MaximumPackageBytes = 1;
        OfficeWorkflowResult result = await runner.RunAsync(files.Request("pages-docx"));
        Assert.False(result.Succeeded);
        Assert.Contains("complete", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(files.Output));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Input_and_output_limits_keep_existing_destination_intact(bool inputLimit) {
        using var files = new Files("numbers", "xlsx");
        byte[] existing = [1, 2, 3];
        File.WriteAllBytes(files.Output, existing);
        OfficeWorkflowRequest request = files.Request("numbers-xlsx");
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        if (inputLimit) request.Limits.MaximumInputBytes = 1;
        else request.Limits.MaximumOutputBytes = 32;
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.Equal(existing, File.ReadAllBytes(files.Output));
        Assert.Equal(2, Directory.GetFiles(files.Root).Length);
    }

    [Fact]
    public async Task Source_mutation_before_publication_is_rejected() {
        using var files = new Files("numbers", "xlsx");
        OfficeWorkflowRequest request = files.Request("numbers-xlsx");
        request.PublicationGuard = new MutatingGuard(files.Input);
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.False(result.Succeeded);
        Assert.False(File.Exists(files.Output));
        Assert.Single(Directory.GetFiles(files.Root));
    }

    [Fact]
    public async Task Provider_zip_stream_uses_the_same_converter_and_publication_contract() {
        using var files = new Files("numbers", "xlsx");
        OfficeWorkflowRequest request = files.Request("numbers-xlsx");
        request.InputPath = "https://provider.invalid/budget";
        request.InputStream = new OfficeWorkflowStreamInput("budget.numbers", token => {
            token.ThrowIfCancellationRequested();
            return Task.FromResult<Stream>(File.OpenRead(files.Input));
        });
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.True(result.Succeeded, result.Summary);
        Assert.NotNull(result.ConversionEvidence);
    }

    [Fact]
    public async Task Cancellation_does_not_publish() {
        using var files = new Files("numbers", "xlsx");
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(files.Request("numbers-xlsx"), cancellationToken: cancelled.Token);
        Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status);
        Assert.False(File.Exists(files.Output));
    }

    [Fact]
    public async Task Registered_converter_cannot_publish_an_invalid_destination() {
        using var files = new Files("pages", "docx");
        var converter = new OfficeWorkflowConversionRegistration("pages-docx", (input, output, limits, token) => {
            output.Write(new byte[] { 1, 2, 3 });
            return new OfficeWorkflowConversionEvidence(new EmptyReport());
        });
        var runner = new OfficeWorkflowRunner(null, null, conversions: [converter]);
        OfficeWorkflowResult result = await runner.RunAsync(files.Request("pages-docx"));
        Assert.False(result.Succeeded);
        Assert.False(File.Exists(files.Output));
        Assert.Single(Directory.GetFiles(files.Root));
    }

    [Fact]
    public void Registration_rejects_duplicate_and_builtin_owners() {
        Assert.Throws<ArgumentException>(() => new OfficeWorkflowConversionRegistration("docx-pdf", (_, _, _, _) => new(new EmptyReport())));
        var registrations = IWorkWorkflow.CreateRegistrations();
        Assert.Throws<ArgumentException>(() => new OfficeWorkflowRunner(null, null, conversions: [registrations[0], registrations[0]]));
    }

    [Theory]
    [InlineData("pages", "pages-docx", "docx")]
    [InlineData("numbers", "numbers-xlsx", "xlsx")]
    [InlineData("key", "keynote-pptx", "pptx")]
    public async Task Unix_special_inputs_are_rejected_without_waiting_for_a_writer(string source, string route, string target) {
        if (OperatingSystem.IsWindows()) return;
        using var files = new Files(source, target);
        File.Delete(files.Input);
        Assert.Equal(0, MkFifo(files.Input, 0x180));
        OfficeWorkflowResult result = await Task.Run(() => IWorkWorkflow.CreateRunner().RunAsync(files.Request(route))).WaitAsync(TimeSpan.FromSeconds(2));
        Assert.False(result.Succeeded);
        Assert.Equal(OfficeWorkflowFailureKind.UnsupportedInput, result.FailureKind);
        Assert.False(File.Exists(files.Output));
    }

    [Theory]
    [InlineData("pdf", "pdf-docx", "docx")]
    [InlineData("docx", "docx-pdf", "pdf")]
    public async Task Builtin_conversion_siblings_reject_Unix_special_inputs(string source, string route, string target) {
        if (OperatingSystem.IsWindows()) return;
        using var files = new Files("numbers", target);
        string special = Path.ChangeExtension(files.Input, "." + source);
        Assert.Equal(0, MkFifo(special, 0x180));
        OfficeWorkflowRequest request = files.Request(route);
        request.InputPath = special;
        OfficeWorkflowResult result = await Task.Run(() => new OfficeWorkflowRunner().RunAsync(request)).WaitAsync(TimeSpan.FromSeconds(2));
        Assert.False(result.Succeeded);
        Assert.False(File.Exists(files.Output));
    }

    [Fact]
    public async Task A_selected_local_symlink_keeps_its_regular_source_identity() {
        if (OperatingSystem.IsWindows()) return;
        using var files = new Files("numbers", "xlsx");
        string selected = Path.Combine(files.Root, "selected.numbers");
        File.CreateSymbolicLink(selected, files.Input);
        OfficeWorkflowRequest request = files.Request("numbers-xlsx");
        request.InputPath = selected;
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.True(result.Succeeded, result.Summary);
        Assert.True(File.Exists(files.Output));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Cancellation_during_capture_or_output_validation_does_not_publish(bool duringValidation) {
        using var files = new Files("numbers", "xlsx");
        using var cancellation = new CancellationTokenSource();
        OfficeWorkflowRequest request = files.Request("numbers-xlsx");
        if (!duringValidation) request.InputStream = new OfficeWorkflowStreamInput("source.numbers", _ =>
            Task.FromResult<Stream>(new CancellingInput(File.ReadAllBytes(files.Input), cancellation)));
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request,
            new InlineProgress(update => { if (duringValidation && update.Stage == "validate-output") cancellation.Cancel(); }), cancellation.Token);
        Assert.Equal(OfficeWorkflowStatus.Cancelled, result.Status);
        Assert.False(File.Exists(files.Output));
        Assert.Single(Directory.GetFiles(files.Root));
    }

    [Fact]
    public async Task Unconfirmed_provider_output_retains_fidelity_and_reopenable_recovery() {
        using var files = new Files("numbers", "xlsx");
        var store = new OfficeWorkflowOutputRecoveryStore(Path.Combine(files.Root, "recovery"));
        OfficeWorkflowRequest request = files.Request("numbers-xlsx");
        request.OutputPath = "content://provider/selected";
        request.ConflictPolicy = OfficeWorkflowConflictPolicy.Replace;
        request.OutputStream = new OfficeWorkflowStreamOutput("selected.xlsx", _ => Task.FromResult<Stream>(new MemoryStream()),
            _ => throw new IOException("Provider write failed."), store);
        OfficeWorkflowResult result = await IWorkWorkflow.CreateRunner().RunAsync(request);
        Assert.Equal(OfficeWorkflowStatus.Unconfirmed, result.Status);
        Assert.True(result.ConversionEvidence!.HasLoss);
        var recovery = Assert.Single(store.GetRecoveries());
        await store.VerifyAsync(recovery);
        using (var workbook = OfficeIMO.Excel.ExcelDocument.Load(recovery.FilePath)) Assert.Single(workbook.Sheets);
        store.Discard(recovery);
    }

    private sealed class CancellingInput(byte[] bytes, CancellationTokenSource cancellation) : MemoryStream(bytes) {
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token) {
            cancellation.Cancel();
            token.ThrowIfCancellationRequested();
            return base.ReadAsync(buffer, offset, count, token);
        }
    }

    private sealed class InlineProgress(Action<OfficeWorkflowProgress> report) : IProgress<OfficeWorkflowProgress> {
        public void Report(OfficeWorkflowProgress value) => report(value);
    }

    [DllImport("libc", EntryPoint = "mkfifo", SetLastError = true)]
    private static extern int MkFifo(string path, uint mode);

    private sealed class EmptyReport : IOfficeConversionReport {
        public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => Array.Empty<OfficeConversionFidelityDiagnostic>();
        public bool HasLoss => false;
        public void RequireNoLoss() { }
    }

    private sealed class MutatingGuard(string path) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string outputPath, bool isDirectory, CancellationToken token) {
            File.AppendAllText(path, "changed");
            return ValueTask.FromResult(true);
        }
    }

    private sealed class Files : IDisposable {
        public Files(string source, string target) {
            Root = Path.Combine(Path.GetTempPath(), "officeimo-iwork-workflow-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Root);
            Input = Path.Combine(Root, "source." + source);
            Output = Path.Combine(Root, "output." + target);
            File.Copy(Path.Combine(AppContext.BaseDirectory, "Corpus", source == "key" ? "tabledeck.key" : "simple." + source), Input);
        }
        public string Root { get; }
        public string Input { get; }
        public string Output { get; }
        public OfficeWorkflowRequest Request(string route) => new() { InputPath = Input, OutputPath = Output,
            Operation = OfficeWorkflowOperation.Convert, ConversionRouteId = route };
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }
}
