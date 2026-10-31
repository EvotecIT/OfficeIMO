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
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioSavedFolderJobTests {
    [Theory]
    [InlineData(900, 650)]
    [InlineData(1400, 850)]
    public async Task SavedFolderJobLoadsPlansRunsAndReusesVerifiedOutputs(int width, int height) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string root = services.Paths.Root;
            string source = Path.Combine(root, "source"), output = Path.Combine(root, "output"), state = Path.Combine(root, "state");
            string jobPath = Path.Combine(root, "monthly.oijob.json");
            Directory.CreateDirectory(source);
            File.WriteAllText(Path.Combine(source, "one.txt"), "Saved Studio job text");
            var dialogs = new Dialogs { SavePath = jobPath, OpenPath = jobPath };
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) { FileDialogs = dialogs };
            var batch = model.ConversionWorkbench.BatchExport;
            string? copied = null; batch.CopyCommandText = text => { copied = text; return Task.CompletedTask; };
            batch.InputDirectory = source; batch.OutputDirectory = output; batch.CheckpointDirectory = state; batch.TabSize = 4;
            Assert.True(model.ConversionWorkbench.AddDroppedPaths([Path.Combine(source, "one.txt")]));
            var retainedQueueJob = Assert.Single(model.ConversionWorkbench.Jobs);
            var view = new ConversionWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show();
                var folderPanel = view.FindControl<Expander>("BatchExportPanel")!;
                folderPanel.IsExpanded = true;
                await ClickAsync(window, view, batch.SaveJobCommand);
                Assert.True(model.ConversionWorkbench.IsFolderJobReviewVisible);
                Assert.True(view.FindControl<Grid>("FolderJobReviewPanel")!.IsEffectivelyVisible);
                Assert.False(view.FindControl<Grid>("QueuePanel")!.IsEffectivelyVisible);
                Assert.False(view.FindControl<ComboBox>("OrdinaryQueueRoute")!.IsEffectivelyVisible);
                Assert.False(model.ConversionWorkbench.HasRunBar);
                Assert.True(File.Exists(jobPath), batch.Status);
                Assert.False(Directory.Exists(output)); Assert.False(Directory.Exists(state));
                await ClickAsync(window, view, batch.CopyCliCommand);
                Assert.Equal(OfficeBatchCommands.GetCliCommand(jobPath), copied);
                batch.TabSize = 3;
                Assert.False(batch.CopyCliCommand.CanExecute(null));
                await ClickAsync(window, view, batch.LoadJobCommand);
                Assert.Equal(4m, batch.TabSize); Assert.Equal(source, batch.InputDirectory);
                Assert.False(Directory.Exists(output)); Assert.False(Directory.Exists(state));
                await ClickAsync(window, view, batch.PlanJobCommand);
                Assert.Contains("Selected 1", batch.PlanSummary);
                Assert.Contains("one.pdf", batch.PlanSummary);
                Assert.False(Directory.Exists(output)); Assert.False(Directory.Exists(state));
                var counts = view.FindControl<TextBlock>("FolderJobPlanCounts")!;
                Assert.Contains("Selected 1", counts.Text);
                AssertVisibleInside(window, counts);
                var planItems = view.FindControl<ItemsControl>("FolderJobPlanItems")!;
                var proposedOutput = planItems.GetVisualDescendants().OfType<TextBlock>().Single(item => item.Text == Path.Combine(output, "one.pdf"));
                AssertVisibleInside(window, proposedOutput);
                Capture(window, $"saved-job-plan-{width}.png");
                await ClickAsync(window, view, batch.RunCommand);
                string pdf = Path.Combine(output, "one.pdf");
                Assert.True(File.Exists(pdf), batch.Status);
                Assert.Contains("Saved Studio job text", OfficeIMO.Pdf.PdfDocument.Load(pdf).Read().Text);
                byte[] original = File.ReadAllBytes(pdf);
                await ClickAsync(window, view, batch.RunCommand);
                Assert.Contains("1 reused", batch.Status);
                Assert.Equal(original, File.ReadAllBytes(pdf));
                Assert.Single(Directory.GetFiles(output));
                await ClickAsync(window, view, batch.CopyPowerShellCommand);
                Assert.Equal(OfficeBatchCommands.GetPowerShellCommand(jobPath), copied);
                folderPanel.IsExpanded = false; window.UpdateLayout();
                Assert.True(view.FindControl<Grid>("QueuePanel")!.IsEffectivelyVisible);
                Assert.True(view.FindControl<ComboBox>("OrdinaryQueueRoute")!.IsEffectivelyVisible);
                Assert.Same(retainedQueueJob, Assert.Single(model.ConversionWorkbench.Jobs));
                Assert.Equal(ConversionJobState.Queued, retainedQueueJob.State);
                folderPanel.IsExpanded = true; window.UpdateLayout();
                Assert.True(view.FindControl<Grid>("FolderJobReviewPanel")!.IsEffectivelyVisible);
                Assert.Contains("Selected 1", counts.Text);
                Capture(window, $"saved-job-finished-{width}.png");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void AssertVisibleInside(Window window, Control control) {
        window.UpdateLayout(); Dispatcher.UIThread.RunJobs(); AvaloniaHeadlessPlatform.ForceRenderTimerTick(); window.UpdateLayout();
        Assert.True(control.IsEffectivelyVisible);
        Point point = control.TranslatePoint(default, window)!.Value;
        Assert.InRange(point.X, 0, window.Bounds.Width - control.Bounds.Width);
        Assert.InRange(point.Y, 0, window.Bounds.Height - control.Bounds.Height);
    }

    [Fact]
    public async Task LoadingAndSavingRetainsPortableSettingsAndRebasesMovedJobPaths() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string root = services.Paths.Root, originalDirectory = Path.Combine(root, "jobs"), copiedDirectory = Path.Combine(root, "other-jobs");
            Directory.CreateDirectory(originalDirectory); Directory.CreateDirectory(copiedDirectory);
            string source = Path.Combine(root, "source"), output = Path.Combine(root, "output");
            Directory.CreateDirectory(source);
            string originalPath = Path.Combine(originalDirectory, "source.json"), copiedPath = Path.Combine(copiedDirectory, "copy.json");
            var original = new OfficeBatchOptions {
                InputDirectory = source, OutputDirectory = output, Recursive = false, Extensions = [".txt"],
                MaxParallelism = 3, MaximumFiles = 50, MaximumInputBytes = 1024 * 1024, MaximumOutputBytes = 2 * 1024 * 1024,
                MaximumXmlCharactersInPart = 4096, OutputProfile = OfficeConversionProfile.Lightweight,
                ConflictPolicy = OfficePublicationPolicy.Rename,
                ConversionOptions = new OfficeConversionOptions {
                CompressPdfOutput = true,                PlainText = new OfficeIMO.Pdf.PdfTextImportOptions { TabSize = 5 }
            }
            };
            original.Name = "Monthly documents";
            (original).Save(originalPath);
            var dialogs = new Dialogs { OpenPath = originalPath, SavePath = copiedPath };
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) { FileDialogs = dialogs };
            var batch = model.ConversionWorkbench.BatchExport;
            await batch.LoadJobCommand.ExecuteAsync(null);
            Assert.Equal(5m, batch.TabSize); Assert.Contains(".txt", batch.JobSettingsSummary);
            Assert.Contains("3 concurrent", batch.JobSettingsSummary);
            batch.TextEncoding = "utf-8";
            await batch.SaveJobCommand.ExecuteAsync(null);
            Assert.True(File.Exists(copiedPath), batch.Status);
            var saved = OfficeBatchOptions.Load(copiedPath);
            var actual = saved;
            Assert.Equal(source, actual.InputDirectory); Assert.Equal(output, actual.OutputDirectory);
            Assert.Equal("Monthly documents", saved.Name);
            Assert.Equal(original.Recursive, actual.Recursive); Assert.Equal(original.Extensions, actual.Extensions);
            Assert.Equal(original.MaxParallelism, actual.MaxParallelism); Assert.Equal(original.MaximumFiles, actual.MaximumFiles);
            Assert.Equal(original.MaximumInputBytes, actual.MaximumInputBytes); Assert.Equal(original.MaximumOutputBytes, actual.MaximumOutputBytes);
            Assert.Equal(original.MaximumXmlCharactersInPart, actual.MaximumXmlCharactersInPart);
            Assert.Equal(original.OutputProfile, actual.OutputProfile); Assert.Equal(original.ConflictPolicy, actual.ConflictPolicy);
            Assert.True(actual.ConversionOptions.CompressPdfOutput); Assert.Equal("utf-8", actual.ConversionOptions.PlainText!.EncodingName);
            Assert.Equal(5, actual.ConversionOptions.PlainText.TabSize);
            Assert.False(Directory.Exists(output));
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RejectedJobLoadPreservesReviewedFieldsAndDoesNotExecute(bool secretReference) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string root = services.Paths.Root, path = Path.Combine(root, "rejected.json"), output = Path.Combine(root, "output");
            Directory.CreateDirectory(root);
            if (secretReference) (new OfficeBatchOptions {
                InputDirectory = "source", OutputDirectory = "output",
                PdfPasswordEnvironmentVariable = "STUDIO_TEST_PASSWORD"
            }).Save(path);
            else File.WriteAllText(path, "{\"schema\":\"unsupported.schema\"}");
            var dialogs = new Dialogs { OpenPath = path };
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) { FileDialogs = dialogs };
            var batch = model.ConversionWorkbench.BatchExport;
            batch.InputDirectory = "reviewed-input"; batch.OutputDirectory = "reviewed-output";
            batch.CheckpointDirectory = "reviewed-state"; batch.TabSize = 6;
            await batch.LoadJobCommand.ExecuteAsync(null);
            Assert.Equal("reviewed-input", batch.InputDirectory); Assert.Equal("reviewed-output", batch.OutputDirectory);
            Assert.Equal("reviewed-state", batch.CheckpointDirectory); Assert.Equal(6m, batch.TabSize);
            Assert.False(batch.HasSavedJob); Assert.False(batch.IsBusy); Assert.False(Directory.Exists(output));
            Assert.Contains(secretReference ? "password references" : "Unsupported", batch.Status);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task InvalidFolderJobSettingsReportTheErrorWithoutStartingExecution() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string root = services.Paths.Root, source = Path.Combine(root, "source"), output = Path.Combine(root, "output"), state = Path.Combine(root, "state");
            Directory.CreateDirectory(source);
            string input = Path.Combine(source, "one.txt");
            File.WriteAllText(input, "Keep the source document");
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var batch = model.ConversionWorkbench.BatchExport;
            batch.InputDirectory = source; batch.OutputDirectory = output; batch.CheckpointDirectory = state;
            // The text box accepts pasted values; the shared saved-job contract bounds encoding names.
            batch.TextEncoding = new string('x', 129);
            await batch.RunCommand.ExecuteAsync(null);
            Assert.Contains("Text encoding", batch.Status);
            Assert.False(batch.IsBusy); Assert.True(batch.RunCommand.CanExecute(null));
            Assert.Empty(model.Jobs.History.Entries);
            Assert.False(Directory.Exists(output)); Assert.False(Directory.Exists(state));
            Assert.Equal("Keep the source document", File.ReadAllText(input));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ExternalSavedJobChangeBlocksCopyAndUnrelatedSaveDestinationIsPreserved() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string root = services.Paths.Root, path = Path.Combine(root, "job.json"), unrelated = Path.Combine(root, "keep.txt");
            Directory.CreateDirectory(Path.Combine(root, "source"));
            (new OfficeBatchOptions { InputDirectory = "source", OutputDirectory = "output" }).Save(path);
            File.WriteAllText(unrelated, "Keep this original document");
            var dialogs = new Dialogs { OpenPath = path, SavePath = unrelated };
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) { FileDialogs = dialogs };
            var batch = model.ConversionWorkbench.BatchExport;
            string? copied = null; batch.CopyCommandText = text => { copied = text; return Task.CompletedTask; };
            await batch.LoadJobCommand.ExecuteAsync(null);
            var changed = OfficeBatchOptions.Load(path);
            changed.MaxParallelism = 4;
            (changed).Save(path, overwrite: true);
            await batch.CopyCliCommand.ExecuteAsync(null);
            Assert.Null(copied); Assert.Contains("changed outside Studio", batch.Status);
            await batch.SaveJobCommand.ExecuteAsync(null);
            Assert.Equal("Keep this original document", File.ReadAllText(unrelated));
            Assert.Equal(path, batch.SavedJobPath);
            Assert.False(Directory.Exists(Path.Combine(root, "output")));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task FolderDiscoveryFailureKeepsPublishedOutputAvailableAndMarksJobFailed() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string root = services.Paths.Root, source = Path.Combine(root, "source"), output = Path.Combine(root, "output");
            Directory.CreateDirectory(source);
            File.WriteAllText(Path.Combine(source, "one.txt"), "Retained document one");
            File.WriteAllText(Path.Combine(source, "two.txt"), "Retained document two");
            string path = Path.Combine(root, "limited-job.json");
            (new OfficeBatchOptions {
                InputDirectory = "source", OutputDirectory = "output", MaximumFiles = 1, MaxParallelism = 1
            }).Save(path);
            var dialogs = new Dialogs { OpenPath = path };
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services) { FileDialogs = dialogs };
            var batch = model.ConversionWorkbench.BatchExport;
            await batch.LoadJobCommand.ExecuteAsync(null);
            await batch.RunCommand.ExecuteAsync(null);
            string saved = Assert.Single(Directory.GetFiles(output, "*.pdf"));
            Assert.Contains("Retained document", OfficeIMO.Pdf.PdfDocument.Load(saved).Read().Text);
            Assert.Contains("1 completed", batch.Status); Assert.Contains("0 failed", batch.Status);
            Assert.Contains("file limit", batch.Status);
            var record = Assert.Single(model.Jobs.History.Entries);
            Assert.Equal(OfficeWorkflowStatus.Failed, record.Outcome);
            Assert.True(record.IsFailed); Assert.True(record.CanOpenOutput);
            Assert.Equal(output, record.OutputPath);
            return true;
        }, CancellationToken.None);
    }

    private static async Task ClickAsync(Window window, ConversionWorkbenchView view, IAsyncRelayCommand command) {
        // Expanding a themed Expander attaches its content presenter during layout. Realize the
        // actual buttons before discovering them, then scroll and click through the rendered tree.
        window.UpdateLayout(); Dispatcher.UIThread.RunJobs(); AvaloniaHeadlessPlatform.ForceRenderTimerTick(); window.UpdateLayout();
        Button button = view.GetVisualDescendants().OfType<Button>().Single(item => ReferenceEquals(item.Command, command));
        button.BringIntoView(); window.UpdateLayout(); Dispatcher.UIThread.RunJobs(); AvaloniaHeadlessPlatform.ForceRenderTimerTick(); window.UpdateLayout();
        Assert.True(button.IsEffectivelyVisible); Assert.True(button.IsEnabled);
        Point location = button.TranslatePoint(new Point(button.Bounds.Width / 2, button.Bounds.Height / 2), window)!.Value;
        Assert.InRange(location.X, 0, window.Width); Assert.InRange(location.Y, 0, window.Height);
        window.MouseDown(location, MouseButton.Left); window.MouseUp(location, MouseButton.Left);
        if (command.ExecutionTask is { } execution) await execution;
        FlushLayout(window);
    }

    private static void Capture(Window window, string name) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        FlushLayout(window);
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }

    private static void FlushLayout(Window window) {
        window.UpdateLayout(); Dispatcher.UIThread.RunJobs(); AvaloniaHeadlessPlatform.ForceRenderTimerTick(); window.UpdateLayout();
    }

    private sealed class Dialogs : IStudioFileDialogs {
        internal string? OpenPath { get; set; }
        internal string? SavePath { get; set; }
        public Task<string?> PickOpenFileAsync(string title, StudioFileType type, CancellationToken cancellationToken) => Task.FromResult(OpenPath);
        public Task<string?> PickSaveFileAsync(string title, string suggestedName, StudioFileType type, CancellationToken cancellationToken) => Task.FromResult(SavePath);
        public Task<string?> PickFolderAsync(string title, CancellationToken cancellationToken) => Task.FromResult<string?>(null);
    }
}
