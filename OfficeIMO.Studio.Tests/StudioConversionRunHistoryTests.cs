using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioConversionRunHistoryTests {
    [Theory]
    [InlineData(900, 650)]
    [InlineData(1400, 850)]
    public async Task SavedPartialFailureSurvivesRestartAndReviewedRetryRetainsBothAttempts(int width, int height) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string root = services.Paths.Root, source = Path.Combine(root, "source"), output = Path.Combine(root, "output");
            Directory.CreateDirectory(source);
            File.WriteAllText(Path.Combine(source, "one.txt"), "Keep the first successful output");
            string broken = Path.Combine(source, "two.docx");
            File.WriteAllText(broken, "A document that needs repair");
            string path = Path.Combine(root, "reviewed-job.json");
            OfficeConversionJobSerializer.SaveFile(new OfficeConversionJob {
                InputDirectory = "source", OutputDirectory = "output", CheckpointDirectory = "checkpoints",
                MaximumConcurrency = 1, RetryFailed = true
            }, path);
            var dialogs = new Dialogs { OpenPath = path };
            using (var initial = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) { FileDialogs = dialogs }) {
                await initial.ConversionWorkbench.BatchExport.LoadJobCommand.ExecuteAsync(null);
                await initial.ConversionWorkbench.BatchExport.RunCommand.ExecuteAsync(null);
                Assert.Equal(OfficeWorkflowStatus.Failed, Assert.Single(services.Jobs.Entries).Outcome);
            }
            string firstPath = Assert.Single(Directory.GetFiles(services.Paths.ConversionRunsRoot, "*.json"));
            byte[] firstReport = File.ReadAllBytes(firstPath), firstPdf = File.ReadAllBytes(Path.Combine(output, "one.txt.pdf"));
            var restarted = StudioApplicationServices.Create(services.Paths);
            var first = Assert.Single(restarted.Jobs.Entries);
            Assert.True(first.HasRunReport); Assert.True(first.CanOpenOutput); Assert.True(first.IsFailed);
            Assert.Equal(2, first.ReportItems.Count);
            Assert.Contains(first.ReportItems, item => item.InputName == "two.docx" && item.IsFailed && !string.IsNullOrWhiteSpace(item.Summary));
            var window = new MainWindow(restarted) { Width = width, Height = height };
            try {
                window.Show();
                window.ViewModel.ShowJobsCommand.Execute(null);
                Layout(window);
                var jobsView = window.GetVisualDescendants().OfType<StudioJobsView>().Single();
                var list = jobsView.GetVisualDescendants().OfType<ListBox>().Single();
                list.SelectedItem = first; list.ScrollIntoView(first); Layout(window);
                var details = jobsView.GetVisualDescendants().OfType<Expander>().Single(item => ReferenceEquals(item.DataContext, first));
                details.IsExpanded = true; Layout(window);
                var review = jobsView.GetVisualDescendants().OfType<Button>().Single(item => ReferenceEquals(item.Command, window.ViewModel.Jobs.ReviewJobCommand) && ReferenceEquals(item.CommandParameter, first));
                AssertVisibleInside(window, review);
                var outcome = jobsView.GetVisualDescendants().OfType<TextBlock>().Single(item => ReferenceEquals(item.DataContext, first) && item.Text == first.Status);
                AssertVisibleInside(window, outcome);
                Assert.True(review.TranslatePoint(default, window)!.Value.Y < details.TranslatePoint(default, window)!.Value.Y);
                Capture(window, $"durable-jobs-{width}-failed.png");
                var failedInput = jobsView.GetVisualDescendants().OfType<TextBlock>().Single(item => item.Text == "two.docx");
                var failedRow = failedInput.GetVisualAncestors().OfType<Border>().First();
                failedRow.BringIntoView(); Layout(window); AssertVisibleInside(window, failedRow);
                Capture(window, $"durable-jobs-{width}-file-details.png");
                using (var repaired = OfficeIMO.Word.WordDocument.Create(broken)) {
                    repaired.AddParagraph("Repaired second source"); repaired.Save();
                }
                await ClickAsync(window, window.ViewModel.Jobs.ReviewJobCommand, first);
                Assert.True(window.ViewModel.IsConversionMode);
                Assert.True(window.ViewModel.ConversionWorkbench.IsFolderJobReviewVisible);
                Assert.Equal(path, window.ViewModel.ConversionWorkbench.BatchExport.SavedJobPath);
                Assert.Single(restarted.Jobs.Entries);
                Assert.False(File.Exists(Path.Combine(output, "two.docx.pdf")));
                await window.ViewModel.ConversionWorkbench.BatchExport.PlanJobCommand.ExecuteAsync(null);
                Assert.Contains("Selected 2", window.ViewModel.ConversionWorkbench.BatchExport.PlanSummary);
                await ClickAsync(window, window.ViewModel.ConversionWorkbench.BatchExport.RunCommand);
                Assert.Equal(2, restarted.Jobs.Entries.Count);
                var retry = restarted.Jobs.Entries[0];
                Assert.True(retry.IsSucceeded, retry.Summary);
                Assert.Equal(1, retry.RunReport!.Batch!.Summary.Reused);
                Assert.Equal(2, retry.RunReport.Batch.Summary.Completed);
                Assert.Equal(0, retry.RunReport.Batch.Summary.Failed);
                var retryReport = OfficeConversionJobRunReportSerializer.ReadFile(retry.RunReportPath!);
                Assert.Single(retryReport.Batch!.Items, item => item.Reused);
                Assert.Single(retryReport.Batch.Items, item => item.Status == OfficeWorkflowStatus.Completed && !item.Reused);
                Assert.Equal(firstPdf, File.ReadAllBytes(Path.Combine(output, "one.txt.pdf")));
                Assert.Equal(firstReport, File.ReadAllBytes(firstPath));
                Assert.Contains("Repaired second source", OfficeIMO.Pdf.PdfDocument.Load(Path.Combine(output, "two.docx.pdf")).Read().Text);
                window.ViewModel.ShowJobsCommand.Execute(null); Layout(window);
                list.SelectedItem = retry; list.ScrollIntoView(retry); Layout(window);
                Capture(window, $"durable-jobs-{width}-retried.png");
                window.ViewModel.Jobs.ClearFinishedCommand.Execute(null);
                Assert.Empty(restarted.Jobs.Entries);
                Assert.Empty(Directory.GetFiles(restarted.Paths.ConversionRunsRoot, "*.json"));
                Assert.Equal(firstPdf, File.ReadAllBytes(Path.Combine(output, "one.txt.pdf")));
                Assert.True(Directory.Exists(Path.Combine(root, "checkpoints")));
                Assert.True(File.Exists(path));
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public void FileResultExcerptKeepsSummaryWarningsAndErrorsWithoutInformationalDetails() {
        var item = new OfficeConversionBatchItemResult("failed.docx", null, OfficeWorkflowStatus.Failed, false,
            "The source document could not be read.", [
                new("capture", "The source was captured internally."),
                new("stage", "The artifact was staged for one filesystem move."),
                new("publish", "The publication boundary was checked."),
                new("font", "A required font is unavailable.", OfficeWorkflowDiagnosticSeverity.Warning),
                new("source", "The document archive is damaged.", OfficeWorkflowDiagnosticSeverity.Error)
            ]);
        var displayed = new StudioConversionRunItem(item, OfficeIMO.Studio.Infrastructure.Localization.StudioLocalization.Current);
        Assert.Equal(string.Join(Environment.NewLine, "The source document could not be read.",
            "A required font is unavailable.", "The document archive is damaged."), displayed.Summary);
        Assert.True(displayed.IsFailed);
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("changed")]
    [InlineData("corrupt")]
    [InlineData("oversized")]
    public async Task HistoricalReviewRejectsUnavailableOrChangedDefinitionAndPreservesTheWorkbench(string condition) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string root = services.Paths.Root, source = Path.Combine(root, "source"), path = Path.Combine(root, "job.json");
            Directory.CreateDirectory(source);
            File.WriteAllText(Path.Combine(source, "one.txt"), "Reviewed source");
            OfficeConversionJobSerializer.SaveFile(new OfficeConversionJob { InputDirectory = "source", OutputDirectory = "output" }, path);
            var dialogs = new Dialogs { OpenPath = path };
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) { FileDialogs = dialogs };
            var batch = model.ConversionWorkbench.BatchExport;
            await batch.LoadJobCommand.ExecuteAsync(null); await batch.RunCommand.ExecuteAsync(null);
            var attempt = Assert.Single(services.Jobs.Entries);
            batch.InputDirectory = "reviewed-input"; batch.OutputDirectory = "reviewed-output"; batch.TabSize = 6;
            if (condition == "missing") File.Delete(path);
            else if (condition == "corrupt") File.WriteAllText(path, "{not a conversion job");
            else if (condition == "oversized") File.WriteAllText(path, new string('x', OfficeConversionJobSerializer.MaximumJobBytes + 1));
            else {
                var changed = OfficeConversionJobSerializer.ReadFile(path); changed.MaximumFiles = 20;
                OfficeConversionJobSerializer.SaveFile(changed, path, overwrite: true);
            }
            model.ShowJobsCommand.Execute(null);
            await model.Jobs.ReviewJobCommand.ExecuteAsync(attempt);
            Assert.Contains(condition == "missing" ? "no longer available" : "changed after this attempt", model.Jobs.ActionError);
            Assert.True(model.IsJobsMode); Assert.Equal("reviewed-input", batch.InputDirectory);
            Assert.Equal("reviewed-output", batch.OutputDirectory); Assert.Equal(6m, batch.TabSize);
            Assert.Single(services.Jobs.Entries);
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task UnsavedAndHistoryOptedOutRunsStaySessionOnly(bool savedButPrivate) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string source = Path.Combine(services.Paths.Root, "source"), output = Path.Combine(services.Paths.Root, "output");
            Directory.CreateDirectory(source);
            File.WriteAllText(Path.Combine(source, "one.txt"), "Session-only source");
            string path = Path.Combine(services.Paths.Root, "job.json");
            OfficeConversionJobSerializer.SaveFile(new OfficeConversionJob { InputDirectory = "source", OutputDirectory = "output" }, path);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) { FileDialogs = new Dialogs { OpenPath = path } };
            var batch = model.ConversionWorkbench.BatchExport;
            if (savedButPrivate) {
                await batch.LoadJobCommand.ExecuteAsync(null);
                services.DocumentHistory.SetRememberHistory(false);
            } else { batch.InputDirectory = source; batch.OutputDirectory = output; }
            await batch.RunCommand.ExecuteAsync(null);
            Assert.True(Assert.Single(services.Jobs.Entries).IsSucceeded, batch.Status);
            Assert.False(Assert.Single(services.Jobs.Entries).HasRunReport);
            Assert.False(Directory.Exists(services.Paths.ConversionRunsRoot));
            Assert.Empty(StudioApplicationServices.Create(services.Paths).Jobs.Entries);
            Assert.True(File.Exists(Path.Combine(output, "one.txt.pdf")));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task MissingRecordedOutputCannotOpenOrRevealAnotherLocation() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string source = Path.Combine(services.Paths.Root, "source"), output = Path.Combine(services.Paths.Root, "output"), path = Path.Combine(services.Paths.Root, "job.json");
            Directory.CreateDirectory(source); File.WriteAllText(Path.Combine(source, "one.txt"), "Source stays available");
            OfficeConversionJobSerializer.SaveFile(new OfficeConversionJob { InputDirectory = "source", OutputDirectory = "output" }, path);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) { FileDialogs = new Dialogs { OpenPath = path } };
            await model.ConversionWorkbench.BatchExport.LoadJobCommand.ExecuteAsync(null);
            await model.ConversionWorkbench.BatchExport.RunCommand.ExecuteAsync(null);
            var attempt = Assert.Single(services.Jobs.Entries);
            var item = Assert.Single(attempt.ReportItems);
            File.Delete(item.OutputPath!); Directory.Delete(output);
            int opened = 0, revealed = 0;
            using var jobs = new StudioJobsViewModel(services.Jobs, (_, _) => { opened++; return Task.CompletedTask; }, revealFolder: _ => { revealed++; return Task.CompletedTask; });
            await jobs.OpenOutputCommand.ExecuteAsync(attempt);
            Assert.Contains("no longer available", jobs.ActionError);
            await jobs.RevealOutputCommand.ExecuteAsync(attempt);
            Assert.Contains("no longer available", jobs.ActionError);
            await jobs.OpenReportOutputCommand.ExecuteAsync(item);
            Assert.Contains("no longer available", jobs.ActionError);
            Assert.Equal(0, opened); Assert.Equal(0, revealed);
            var restored = Assert.Single(StudioApplicationServices.Create(services.Paths).Jobs.Entries);
            Assert.False(restored.CanOpenOutput); Assert.False(restored.HasOutput);
            Assert.Contains("no longer available", restored.OutputAvailabilityNote);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ReportStorageFailureKeepsCompletedOutputAndShowsRecordingError() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string source = Path.Combine(services.Paths.Root, "source"), output = Path.Combine(services.Paths.Root, "output"), path = Path.Combine(services.Paths.Root, "job.json");
            Directory.CreateDirectory(source); File.WriteAllText(Path.Combine(source, "one.txt"), "A committed output survives a history failure");
            OfficeConversionJobSerializer.SaveFile(new OfficeConversionJob { InputDirectory = "source", OutputDirectory = "output" }, path);
            File.WriteAllText(services.Paths.ConversionRunsRoot, "Keep this blocked history location");
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) { FileDialogs = new Dialogs { OpenPath = path } };
            await model.ConversionWorkbench.BatchExport.LoadJobCommand.ExecuteAsync(null);
            await model.ConversionWorkbench.BatchExport.RunCommand.ExecuteAsync(null);
            var attempt = Assert.Single(services.Jobs.Entries);
            Assert.True(attempt.IsSucceeded, attempt.Summary); Assert.True(attempt.CanOpenOutput);
            Assert.Contains("could not be fully saved", attempt.RecordingError);
            Assert.Contains("could not be fully saved", model.ConversionWorkbench.BatchExport.Status);
            Assert.Contains("committed output", OfficeIMO.Pdf.PdfDocument.Load(Path.Combine(output, "one.txt.pdf")).Read().Text);
            Assert.Equal("Keep this blocked history location", File.ReadAllText(services.Paths.ConversionRunsRoot));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task AbandonedRecordedRunRestoresAsInterruptedWithoutExecutingOrClaimingAnOutput() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var services = ((App)Application.Current!).Services;
            string path = Path.Combine(services.Paths.Root, "job.json");
            Directory.CreateDirectory(services.Paths.Root);
            OfficeConversionJobSerializer.SaveFile(new OfficeConversionJob { InputDirectory = "source", OutputDirectory = "output" }, path);
            _ = OfficeConversionJobSerializer.ReadFile(path, out string fingerprint);
            string reportPath = services.DocumentHistory.ConversionRuns.CreateReportPath()!;
            File.WriteAllText(reportPath, OfficeConversionJobRunReportSerializer.Serialize(new OfficeConversionJobRunReport {
                RunId = Guid.NewGuid().ToString("N"), StartedUtc = DateTimeOffset.UtcNow,
                JobPath = path, JobSha256 = fingerprint, OutputDirectory = Path.Combine(services.Paths.Root, "output"),
                State = OfficeConversionJobRunState.Running
            }));
            var restored = Assert.Single(StudioApplicationServices.Create(services.Paths).Jobs.Entries);
            Assert.Equal("Interrupted", restored.Status); Assert.False(restored.IsActive);
            Assert.True(restored.CanReviewJob); Assert.False(restored.CanCancel); Assert.False(restored.HasOutput);
            Assert.Contains("without a final result", restored.Summary);
            Assert.False(Directory.Exists(Path.Combine(services.Paths.Root, "output")));
            return true;
        }, CancellationToken.None);
    }

    private static async Task ClickAsync(Window window, IAsyncRelayCommand command, object? parameter = null) {
        Layout(window);
        var button = window.GetVisualDescendants().OfType<Button>().Single(item => ReferenceEquals(item.Command, command) && ReferenceEquals(item.CommandParameter, parameter));
        button.BringIntoView(); Layout(window);
        Assert.True(button.IsEffectivelyVisible); Assert.True(button.IsEffectivelyEnabled);
        Point point = button.TranslatePoint(new Point(button.Bounds.Width / 2, button.Bounds.Height / 2), window)!.Value;
        Assert.InRange(point.X, 0, window.Bounds.Width); Assert.InRange(point.Y, 0, window.Bounds.Height);
        window.MouseDown(point, MouseButton.Left); window.MouseUp(point, MouseButton.Left);
        if (command.ExecutionTask is { } execution) await execution;
        Layout(window);
    }

    private static void Layout(Window window) {
        window.UpdateLayout(); Dispatcher.UIThread.RunJobs(); AvaloniaHeadlessPlatform.ForceRenderTimerTick(); window.UpdateLayout();
    }

    private static void AssertVisibleInside(Window window, Control control) {
        Assert.True(control.IsEffectivelyVisible);
        Point point = control.TranslatePoint(default, window)!.Value;
        Assert.InRange(point.X, 0, window.Bounds.Width - control.Bounds.Width);
        Assert.InRange(point.Y, 0, window.Bounds.Height - control.Bounds.Height);
    }

    private static void Capture(Window window, string name) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output); Layout(window);
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }

    private sealed class Dialogs : IStudioFileDialogs {
        internal string? OpenPath { get; init; }
        public Task<string?> PickOpenFileAsync(string title, StudioFileType type, CancellationToken cancellationToken) => Task.FromResult(OpenPath);
        public Task<string?> PickSaveFileAsync(string title, string suggestedName, StudioFileType type, CancellationToken cancellationToken) => Task.FromResult<string?>(null);
        public Task<string?> PickFolderAsync(string title, CancellationToken cancellationToken) => Task.FromResult<string?>(null);
    }
}
