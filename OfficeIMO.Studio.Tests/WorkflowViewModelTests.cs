using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class WorkflowViewModelTests {
    [Fact]
    public async Task PrintPreviewBuildsRenderedMultiPageSheets() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "print.pdf");
        PdfDocument.Create(document => {
            document.Page(page => page.Size(400, 600).Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("First")))));
            document.Page(page => page.Size(600, 400).Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Second")))));
            document.Page(page => page.Size(420, 620).Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Third")))));
        }).Save(input);

        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            using var viewModel = new PrintPreviewViewModel(_ => Task.FromResult<string?>(input));
            await viewModel.ChooseInputCommand.ExecuteAsync(null);
            viewModel.SelectedPagesPerSheet = viewModel.PagesPerSheetChoices.Single(choice => choice.Value == 2);
            await viewModel.BuildPreviewCommand.ExecuteAsync(null);

            Assert.True(viewModel.HasPreview);
            Assert.Equal(2, viewModel.Sheets.Count);
            Assert.Equal(2, viewModel.Sheets[0].Placements.Count);
            Assert.Single(viewModel.Sheets[1].Placements);
            Assert.All(viewModel.Sheets.SelectMany(sheet => sheet.Placements), placement => Assert.NotNull(placement.Image));
            Assert.Equal("Print preview ready", viewModel.Status);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task PrintPreviewRejectsAnUnboundedAllPagesRequest() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "large-print.pdf");
        PdfDocument.Create(document => {
            for (int pageNumber = 0; pageNumber <= PrintPreviewViewModel.MaximumPreviewPages; pageNumber++) {
                document.Page(page => page.Size(400, 600));
            }
        }).Save(input);

        using var viewModel = new PrintPreviewViewModel(_ => Task.FromResult<string?>(input)) { InputPath = input };
        await viewModel.BuildPreviewCommand.ExecuteAsync(null);

        Assert.False(viewModel.HasPreview);
        Assert.Empty(viewModel.Sheets);
        Assert.Contains("limited", viewModel.Status, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void ShellSwitchesBetweenPrimaryDestinationsAndDocumentModesWithoutChangingDocumentState() {
        using var viewModel = new MainWindowViewModel(_ => Task.FromResult<string?>(null));

        Assert.True(viewModel.IsHomeMode);
        viewModel.ShowToolsCommand.Execute(null);
        Assert.True(viewModel.IsToolsMode);
        viewModel.ShowOcrCommand.Execute(null);
        Assert.True(viewModel.IsOcrMode);
        viewModel.ShowJobsCommand.Execute(null);
        Assert.True(viewModel.IsJobsMode);
        viewModel.ShowPdfWorkspaceCommand.Execute(null);
        Assert.True(viewModel.IsPdfWorkspaceMode);
        Assert.False(viewModel.IsHomeMode);
        viewModel.ShowAnnotateModeCommand.Execute(null);
        Assert.True(viewModel.IsAnnotateDocumentMode);
        viewModel.ShowPagesModeCommand.Execute(null);
        Assert.True(viewModel.IsPagesDocumentMode);
        viewModel.ShowConversionWorkbenchCommand.Execute(null);
        Assert.True(viewModel.IsConversionMode);
        Assert.False(viewModel.IsPdfWorkspaceMode);
        viewModel.ShowPrintPreviewCommand.Execute(null);
        Assert.True(viewModel.IsOutputMode);
        Assert.True(viewModel.OutputWorkbench.IsPrintPreview);
        viewModel.ShowPageExportCommand.Execute(null);
        Assert.True(viewModel.IsOutputMode);
        Assert.True(viewModel.OutputWorkbench.IsPageExport);
        viewModel.ShowAssemblyCommand.Execute(null);
        Assert.True(viewModel.OutputWorkbench.IsAssembly);
        viewModel.ShowDocumentHealthCommand.Execute(null);
        Assert.True(viewModel.IsDocumentHealthMode);
        Assert.Equal(OfficeWorkflowOperation.Repair, viewModel.DocumentHealth.SelectedOperation.Value);
        viewModel.ShowInspectCommand.Execute(null);
        Assert.Equal(OfficeWorkflowOperation.Inspect, viewModel.DocumentHealth.SelectedOperation.Value);
        viewModel.ShowRepairPlanCommand.Execute(null);
        Assert.Equal(OfficeWorkflowOperation.RepairPlan, viewModel.DocumentHealth.SelectedOperation.Value);
        viewModel.ShowSanitizeCommand.Execute(null);
        Assert.Equal(OfficeWorkflowOperation.Sanitize, viewModel.DocumentHealth.SelectedOperation.Value);
        viewModel.ShowPdfWorkspaceCommand.Execute(null);
        Assert.True(viewModel.IsPdfWorkspaceMode);
        Assert.False(viewModel.HasDocument);
        viewModel.ShowHomeCommand.Execute(null);
        Assert.True(viewModel.IsHomeMode);
    }

    [Fact]
    public async Task OutputWorkbenchExportsPagesAndAssemblesOrderedPdfs() {
        using var scope = new TestDirectory();
        string first = Path.Combine(scope.Path, "first.pdf");
        string second = Path.Combine(scope.Path, "second.pdf");
        string imageFolder = Path.Combine(scope.Path, "images");
        string assembled = Path.Combine(scope.Path, "assembled.pdf");
        PdfDocument.Create(document => {
            document.Page(page => page.Size(400, 600));
            document.Page(page => page.Size(420, 600));
        }).Save(first);
        PdfDocument.Create(document => document.Page(page => page.Size(500, 700))).Save(second);

        using var pageExport = new PageImageExportViewModel(
            _ => Task.FromResult<string?>(first),
            _ => Task.FromResult<string?>(imageFolder));
        await pageExport.ChooseInputCommand.ExecuteAsync(null);
        await pageExport.ChooseOutputDirectoryCommand.ExecuteAsync(null);
        pageExport.Pages = "2";
        await pageExport.ExportCommand.ExecuteAsync(null);

        Assert.True(pageExport.HasOutput);
        Assert.Single(Directory.GetFiles(pageExport.PublishedDirectory!, "*.png"));

        using var assembly = new PdfAssemblyViewModel(
            _ => Task.FromResult<IReadOnlyList<string>>([first, second]),
            _ => Task.FromResult<string?>(null),
            _ => Task.FromResult<string?>(assembled));
        await assembly.AddFilesCommand.ExecuteAsync(null);
        await assembly.ChooseOutputCommand.ExecuteAsync(null);
        await assembly.RunCommand.ExecuteAsync(null);

        Assert.True(assembly.HasOutput);
        Assert.Equal(3, PdfReadDocument.Open(File.ReadAllBytes(assembly.PublishedPath!)).Pages.Count);
    }

    [Fact]
    public async Task GuidedRepairProgressesThroughFilesOptionsAndReviewWithHonestLabels() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "source.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(600, 800))).Save(input);
        using var viewModel = new DocumentHealthViewModel(
            _ => Task.FromResult<string?>(input),
            _ => Task.FromResult<string?>(scope.Path));

        viewModel.PrepareRepairWorkflow();
        Assert.Equal("Repair PDF", viewModel.WorkbenchTitle);
        Assert.Equal("Run repair", viewModel.RunActionLabel);
        Assert.Equal(GuidedWorkflowStep.Files, viewModel.CurrentStep);
        Assert.False(viewModel.ContinueCommand.CanExecute(null));

        await viewModel.ChooseInputCommand.ExecuteAsync(null);
        Assert.True(viewModel.ContinueCommand.CanExecute(null));
        viewModel.ContinueCommand.Execute(null);
        Assert.Equal(GuidedWorkflowStep.Options, viewModel.CurrentStep);
        viewModel.ContinueCommand.Execute(null);
        Assert.Equal(GuidedWorkflowStep.Review, viewModel.CurrentStep);
        Assert.EndsWith("source.repaired.pdf", viewModel.OutputPreviewPath, StringComparison.OrdinalIgnoreCase);

        viewModel.BackCommand.Execute(null);
        Assert.Equal(GuidedWorkflowStep.Options, viewModel.CurrentStep);
    }

    [Fact]
    public async Task CompareRequiresBothFilesAndUsesComparisonOutput() {
        using var scope = new TestDirectory();
        string first = Path.Combine(scope.Path, "first.pdf");
        string second = Path.Combine(scope.Path, "second.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(600, 800))).Save(first);
        PdfDocument.Create(compose => compose.Page(page => page.Size(600, 800))).Save(second);
        var picks = new Queue<string>([first, second]);
        using var viewModel = new DocumentHealthViewModel(
            _ => Task.FromResult<string?>(picks.Dequeue()),
            _ => Task.FromResult<string?>(scope.Path));

        viewModel.PrepareWorkflow(OfficeWorkflowOperation.Compare);
        await viewModel.ChooseInputCommand.ExecuteAsync(null);
        Assert.False(viewModel.ContinueCommand.CanExecute(null));
        await viewModel.ChooseComparisonCommand.ExecuteAsync(null);

        Assert.True(viewModel.ContinueCommand.CanExecute(null));
        Assert.Equal("Compare PDFs", viewModel.WorkbenchTitle);
        Assert.Equal("Run compare", viewModel.RunActionLabel);
        Assert.EndsWith("first.comparison.html", viewModel.OutputPreviewPath, StringComparison.OrdinalIgnoreCase);
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
    public async Task ConversionQueueEnforcesOwnerBatchLimitBeforeRun() {
        string[] paths = Enumerable.Range(1, OfficeWorkflowRunner.MaximumBatchRequestCount + 1)
            .Select(index => Path.Combine(Path.GetTempPath(), $"source-{index}.docx"))
            .ToArray();
        using var viewModel = new ConversionWorkbenchViewModel(
            _ => Task.FromResult<IReadOnlyList<string>>(paths),
            _ => Task.FromResult<string?>(null));
        viewModel.SelectedRoute = viewModel.Routes.Single(route => route.Route.Id == "docx-pdf");

        await viewModel.AddFilesCommand.ExecuteAsync(null);

        Assert.Equal(OfficeWorkflowRunner.MaximumBatchRequestCount, viewModel.Jobs.Count);
        Assert.Contains("limited", viewModel.Status, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(OfficeWorkflowRunner.MaximumBatchRequestCount.ToString(), viewModel.Status, StringComparison.Ordinal);
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
    public async Task FailedRunClearsPreviousHealthEvidenceAndUsesFailureState() {
        using var scope = new TestDirectory();
        string input = Path.Combine(scope.Path, "source.pdf");
        PdfDocument.Create(compose => compose.Page(page => page.Size(600, 800))).Save(input);
        using var viewModel = new DocumentHealthViewModel(
            _ => Task.FromResult<string?>(input),
            _ => Task.FromResult<string?>(scope.Path));
        await viewModel.ChooseInputCommand.ExecuteAsync(null);
        await viewModel.RunCommand.ExecuteAsync(null);
        Assert.True(viewModel.HasHealthReport);
        Assert.True(viewModel.IsResultSuccessful);

        viewModel.InputPath = Path.Combine(scope.Path, "missing.pdf");
        await viewModel.RunCommand.ExecuteAsync(null);

        Assert.True(viewModel.IsResultFailed);
        Assert.False(viewModel.IsResultSuccessful);
        Assert.False(viewModel.HasHealthReport);
        Assert.Equal("—", viewModel.BeforeSummary);
        Assert.Equal("—", viewModel.AfterSummary);
        Assert.Contains("failed", viewModel.Summary, StringComparison.OrdinalIgnoreCase);
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

        viewModel.PrepareWorkflow(OfficeWorkflowOperation.Compare);

        Assert.Equal(OfficeWorkflowOperation.Inspect, viewModel.SelectedOperation.Value);

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
