using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class WorkflowViewModelTests {
    [Fact]
    public void ShellSwitchesBetweenHomeReaderConversionAndHealthWithoutChangingDocumentState() {
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));

        Assert.True(viewModel.IsHomeMode);
        viewModel.ShowPdfWorkspaceCommand.Execute(null);
        Assert.True(viewModel.IsPdfWorkspaceMode);
        Assert.False(viewModel.IsHomeMode);
        viewModel.ShowConversionWorkbenchCommand.Execute(null);
        Assert.True(viewModel.IsConversionMode);
        Assert.False(viewModel.IsPdfWorkspaceMode);
        viewModel.ShowDocumentHealthCommand.Execute(null);
        Assert.True(viewModel.IsDocumentHealthMode);
        viewModel.ShowPdfWorkspaceCommand.Execute(null);
        Assert.True(viewModel.IsPdfWorkspaceMode);
        Assert.False(viewModel.HasDocument);
        viewModel.ShowHomeCommand.Execute(null);
        Assert.True(viewModel.IsHomeMode);
    }

    [Fact]
    public async Task ConversionWorkbenchRunsMatchingFilesAndSurfacesReopenEvidence() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "source.html");
        await File.WriteAllTextAsync(input, "<!doctype html><html><body><h1>Studio conversion</h1></body></html>");
        using var viewModel = new ConversionWorkbenchViewModel(
            _ => Task.FromResult<IReadOnlyList<string>>([input]),
            _ => Task.FromResult<string?>(scope.Path));
        viewModel.SelectedRoute = viewModel.Routes.Single(route => route.Route.Id == "html-pdf");

        await viewModel.AddFilesCommand.ExecuteAsync(null);
        await viewModel.ChooseOutputFolderCommand.ExecuteAsync(null);
        await viewModel.RunQueueCommand.ExecuteAsync(null);

        ConversionJobViewModel job = Assert.Single(viewModel.Jobs);
        Assert.Equal("Completed", job.Status);
        Assert.NotNull(job.OutputPath);
        Assert.True(File.Exists(job.OutputPath));
        Assert.Contains(job.Diagnostics, diagnostic => diagnostic.Code == "OutputReopened");
        Assert.Contains("1 completed", viewModel.Status, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConversionWorkbenchSkipsFilesThatDoNotMatchSelectedRoute() {
        using var scope = new TestDirectory();
        string path = Path.Combine(scope.Path, "notes.txt");
        await File.WriteAllTextAsync(path, "not a Word document");
        using var viewModel = new ConversionWorkbenchViewModel(
            _ => Task.FromResult<IReadOnlyList<string>>([path]),
            _ => Task.FromResult<string?>(scope.Path));
        viewModel.SelectedRoute = viewModel.Routes.Single(route => route.Route.Id == "docx-pdf");

        await viewModel.AddFilesCommand.ExecuteAsync(null);

        Assert.Empty(viewModel.Jobs);
        Assert.Contains("No files matched", viewModel.Status, StringComparison.Ordinal);
    }

    [Fact]
    public async Task DocumentHealthInspectionPopulatesReadableBeforeReport() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "source.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Content(content => content.Item(item =>
            item.Paragraph(paragraph => paragraph.Text("Studio health")))))).Save(input);
        using var viewModel = new DocumentHealthViewModel(
            _ => Task.FromResult<string?>(input),
            _ => Task.FromResult<string?>(scope.Path));

        await viewModel.ChooseInputCommand.ExecuteAsync(null);
        await viewModel.RunCommand.ExecuteAsync(null);

        Assert.Equal("Completed", viewModel.Status);
        Assert.Contains("1 page", viewModel.BeforeSummary, StringComparison.Ordinal);
        Assert.Equal("No artifact was written.", viewModel.AfterSummary);
        Assert.False(viewModel.HasOutput);
    }

    [Fact]
    public async Task RepairPlanRemainsReportOnly() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "source.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(600, 800))).Save(input);
        using var viewModel = new DocumentHealthViewModel(
            _ => Task.FromResult<string?>(input),
            _ => Task.FromResult<string?>(scope.Path));
        viewModel.SelectedOperation = viewModel.Operations.Single(operation => operation.Value == OfficeWorkflowOperation.RepairPlan);
        await viewModel.ChooseInputCommand.ExecuteAsync(null);

        await viewModel.RunCommand.ExecuteAsync(null);

        Assert.Equal("Completed", viewModel.Status);
        Assert.False(viewModel.HasOutput);
        Assert.Contains("not needed", viewModel.Summary, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task ConversionQueueCannotMutateWhileRunningAndDisposeCancelsWithoutRacingFinally() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "source.html");
        await File.WriteAllTextAsync(input, "<!doctype html><html><body><p>Queued</p></body></html>");
        var runner = new BlockingWorkflowRunner();
        var viewModel = new ConversionWorkbenchViewModel(
            _ => Task.FromResult<IReadOnlyList<string>>([input]),
            _ => Task.FromResult<string?>(scope.Path),
            runner);
        viewModel.SelectedRoute = viewModel.Routes.Single(route => route.Route.Id == "html-pdf");
        await viewModel.AddFilesCommand.ExecuteAsync(null);

        Task run = viewModel.RunQueueCommand.ExecuteAsync(null);
        await runner.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.True(viewModel.IsBusy);
        Assert.False(viewModel.AddFilesCommand.CanExecute(null));
        Assert.False(viewModel.RemoveSelectedCommand.CanExecute(null));
        Assert.False(viewModel.ClearQueueCommand.CanExecute(null));
        viewModel.RemoveSelectedCommand.Execute(null);
        Assert.Single(viewModel.Jobs);

        viewModel.Dispose();
        await run.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.False(viewModel.IsBusy);
        Assert.Equal("Cancelled", Assert.Single(viewModel.Jobs).Status);
    }

    [Fact]
    public async Task DocumentHealthDisposeCancelsWithoutRacingFinally() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "source.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(600, 800))).Save(input);
        var runner = new BlockingWorkflowRunner();
        var viewModel = new DocumentHealthViewModel(
            _ => Task.FromResult<string?>(input),
            _ => Task.FromResult<string?>(scope.Path),
            runner);
        await viewModel.ChooseInputCommand.ExecuteAsync(null);

        Task run = viewModel.RunCommand.ExecuteAsync(null);
        await runner.Started.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.True(viewModel.IsBusy);

        viewModel.Dispose();
        await run.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.False(viewModel.IsBusy);
        Assert.Equal("Cancelled", viewModel.Status);
    }

    private sealed class BlockingWorkflowRunner : IOfficeWorkflowRunner {
        private readonly OfficeWorkflowRunner _owner = new();

        public TaskCompletionSource Started { get; } = new(TaskCreationOptions.RunContinuationsAsynchronously);

        public async Task<OfficeWorkflowResult> RunAsync(
            OfficeWorkflowRequest request,
            IProgress<OfficeWorkflowProgress>? progress = null,
            CancellationToken cancellationToken = default) {
            Started.TrySetResult();
            try {
                await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
            } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                // Delegate to the owner so the host observes its normal typed cancelled result.
            }
            return await _owner.RunAsync(request, progress, cancellationToken);
        }

        public async Task<IReadOnlyList<OfficeWorkflowResult>> RunBatchAsync(
            IEnumerable<OfficeWorkflowRequest> requests,
            IProgress<OfficeWorkflowProgress>? progress = null,
            CancellationToken cancellationToken = default) {
            Started.TrySetResult();
            try {
                await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken);
            } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                // Delegate to the owner so the host observes its normal typed cancelled batch.
            }
            return await _owner.RunBatchAsync(requests, progress, cancellationToken);
        }
    }

    private sealed class TestDirectory : IDisposable {
        public TestDirectory() {
            Path = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "officeimo-studio-workflows-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public void Dispose() {
            try {
                Directory.Delete(Path, recursive: true);
            } catch (IOException) {
                // Best effort for transient package handles on Windows.
            } catch (UnauthorizedAccessException) {
                // Best effort for transient package handles on Windows.
            }
        }
    }
}
