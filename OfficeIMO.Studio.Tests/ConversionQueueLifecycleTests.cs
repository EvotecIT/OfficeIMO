using Avalonia.Threading;
using System.Runtime.InteropServices;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class ConversionQueueLifecycleTests {
    [Fact]
    public async Task RetryAndNewPendingJobsNeverRepeatCompletedOutputs() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            using var scope = new Scope();
            string first = scope.Write("first");
            string second = Path.Combine(scope.Root, "second.html");
            IReadOnlyList<string> selection = [first, second];
            using var model = Create(() => selection, scope.Root);
            await model.AddFilesCommand.ExecuteAsync(null);
            await model.RunQueueCommand.ExecuteAsync(null);
            Assert.True(model.Jobs[0].State == ConversionJobState.Completed, model.Jobs[0].Summary + " " + string.Join("; ", model.Jobs[0].Diagnostics.Select(diagnostic => diagnostic.Message)));
            Assert.Equal(ConversionJobState.Failed, model.Jobs[1].State);
            byte[] original = File.ReadAllBytes(model.Jobs[0].OutputPath!);
            var firstDiagnostics = model.Jobs[0].Diagnostics;
            Assert.False(model.CanRun);
            Assert.True(model.CanRetryFailed);

            scope.Write("second");
            await model.RetryFailedCommand.ExecuteAsync(null);
            Assert.All(model.Jobs, job => Assert.Equal(ConversionJobState.Completed, job.State));
            Assert.Equal(original, File.ReadAllBytes(model.Jobs[0].OutputPath!));
            Assert.Same(firstDiagnostics, model.Jobs[0].Diagnostics);
            await model.RunQueueCommand.ExecuteAsync(null);
            await model.RetryFailedCommand.ExecuteAsync(null);
            Assert.Equal(2, Directory.GetFiles(scope.Root, "*.pdf").Length);

            selection = [scope.Write("third")];
            await model.AddFilesCommand.ExecuteAsync(null);
            Assert.True(model.CanRun);
            await model.RunQueueCommand.ExecuteAsync(null);
            Assert.Equal(3, Directory.GetFiles(scope.Root, "*.pdf").Length);
            Assert.Same(firstDiagnostics, model.Jobs[0].Diagnostics);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task CancelledBatchRetainsCompletedOutputAndRetriesOnlyUncommittedJobs() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            using var scope = new Scope();
            string[] sources = [scope.Write("first"), scope.Write("second"), scope.Write("third")];
            ConversionWorkbenchViewModel? model = null;
            bool cancelSecond = true;
            var guard = new PublicationGuard((path, _) => {
                if (cancelSecond && Path.GetFileName(path) == "second.pdf") model!.CancelCommand.Execute(null);
                return ValueTask.FromResult(true);
            });
            model = Create(() => sources, scope.Root, guard: guard);
            using (model) {
                await model.AddFilesCommand.ExecuteAsync(null);
                await model.RunQueueCommand.ExecuteAsync(null);
                Assert.Equal(ConversionJobState.Completed, model.Jobs[0].State);
                Assert.All(model.Jobs.Skip(1), job => Assert.Equal(ConversionJobState.Cancelled, job.State));
                Assert.Single(Directory.GetFiles(scope.Root, "*.pdf"));
                cancelSecond = false;
                await model.RetryFailedCommand.ExecuteAsync(null);
                Assert.All(model.Jobs, job => Assert.Equal(ConversionJobState.Completed, job.State));
                Assert.Equal(3, Directory.GetFiles(scope.Root, "*.pdf").Length);
                Assert.False(File.Exists(Path.Combine(scope.Root, "first (1).pdf")));
            }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task OldProgressCannotReplaceTerminalResults() {
        using var scope = new Scope();
        string source = scope.Write("first");
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var runner = new ObservingRunner();
            using var model = Create(() => [source], scope.Root, runner);
            await model.AddFilesCommand.ExecuteAsync(null);
            await model.RunQueueCommand.ExecuteAsync(null);
            string status = model.Status;
            var job = Assert.Single(model.Jobs);
            runner.Progress!.Report(new OfficeWorkflowProgress(job.Id, "execute", "obsolete progress", 0.1));
            await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
            Assert.Equal(status, model.Status);
            Assert.Equal(ConversionJobState.Completed, job.State);
            Assert.Equal(1D, job.ProgressFraction);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task LostResultRequiresCheckingOutputBeforeAnotherAttempt() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            using var scope = new Scope();
            string source = scope.Write("first");
            using var model = Create(() => [source], scope.Root, new ObservingRunner { LoseResult = true });
            await model.AddFilesCommand.ExecuteAsync(null);
            await model.RunQueueCommand.ExecuteAsync(null);
            Assert.True(File.Exists(Path.Combine(scope.Root, "first.pdf")));
            Assert.Equal(ConversionJobState.Unconfirmed, Assert.Single(model.Jobs).State);
            Assert.False(model.CanRun);
            Assert.False(model.CanRetryFailed);
            await model.RetryFailedCommand.ExecuteAsync(null);
            Assert.Single(Directory.GetFiles(scope.Root, "*.pdf"));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task QueueDeduplicatesAliasesAcrossPickerCallsAndPreservesCaseSensitiveInputs() {
        using var scope = new Scope();
        string source = scope.Write("report");
        string upper = Path.Combine(scope.Root, "REPORT.html");
        bool sameCaseIdentity = File.Exists(upper);
        if (!sameCaseIdentity) File.WriteAllText(upper, "<p>Distinct uppercase input</p>");
        string alias = Path.Combine(scope.Root, "alias.html");
        File.CreateSymbolicLink(alias, source);
        string hardLink = Path.Combine(scope.Root, "hard-link.html");
        Assert.True(OperatingSystem.IsWindows() ? CreateHardLink(hardLink, source, IntPtr.Zero) : Link(source, hardLink) == 0);
        using var model = Create(() => [source, alias, hardLink, upper], scope.Root);
        await model.AddFilesCommand.ExecuteAsync(null);
        await model.AddFilesCommand.ExecuteAsync(null);
        Assert.Equal(sameCaseIdentity ? 1 : 2, model.Jobs.Count);
    }

    [DllImport("kernel32.dll", EntryPoint = "CreateHardLinkW", CharSet = CharSet.Unicode, SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CreateHardLink(string newFile, string existingFile, IntPtr securityAttributes);

    [DllImport("libc", EntryPoint = "link", SetLastError = true)]
    private static extern int Link(string existingFile, string newFile);

    private static ConversionWorkbenchViewModel Create(Func<IReadOnlyList<string>> sources, string root,
        IOfficeWorkflowRunner? runner = null, IOfficeWorkflowPublicationGuard? guard = null) {
        var model = new ConversionWorkbenchViewModel(_ => Task.FromResult(sources()),
            _ => Task.FromResult<string?>(root), runner, localizer: null, publicationGuard: guard);
        model.SelectedRoute = model.Routes.Single(route => route.Route.Id == "html-pdf");
        model.OutputFolder = root;
        return model;
    }

    private sealed class ObservingRunner : IOfficeWorkflowRunner {
        private readonly OfficeWorkflowRunner _owner = new();
        internal IProgress<OfficeWorkflowProgress>? Progress { get; private set; }
        internal bool LoseResult { get; init; }
        public Task<OfficeWorkflowResult> RunAsync(OfficeWorkflowRequest request, IProgress<OfficeWorkflowProgress>? progress = null,
            CancellationToken cancellationToken = default) => _owner.RunAsync(request, progress, cancellationToken);
        public async Task<IReadOnlyList<OfficeWorkflowResult>> RunBatchAsync(IEnumerable<OfficeWorkflowRequest> requests,
            IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) {
            Progress = progress;
            var results = await _owner.RunBatchAsync(requests, progress, cancellationToken);
            if (LoseResult) throw new IOException("Result delivery interrupted");
            return results;
        }
    }

    private sealed class PublicationGuard(Func<string, CancellationToken, ValueTask<bool>> check) : IOfficeWorkflowPublicationGuard {
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) => check(path, cancellationToken);
    }

    private sealed class Scope : IDisposable {
        public string Root { get; } = Path.Combine(Path.GetTempPath(), "officeimo-queue-" + Guid.NewGuid().ToString("N"));
        public Scope() => Directory.CreateDirectory(Root);
        public string Write(string name) {
            string path = Path.Combine(Root, name + ".html");
            File.WriteAllText(path, "<html><body><p>" + name + "</p></body></html>");
            return path;
        }
        public void Dispose() => Directory.Delete(Root, recursive: true);
    }
}
