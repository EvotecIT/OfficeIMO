using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioExtractionPreviewTests {
    [Fact]
    public async Task CancellationAtTheCpuGateLeavesNoOutputOrRecovery() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf"), output = Path.Combine(services.Paths.Root, "output.pdf");
            CreatePdf().Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(output), reviewPageExtraction: _ => Task.FromResult(true));
            await model.OpenDocumentAsync(source); model.SelectAllPagesCommand.Execute(null);
            using var blocker = await Features.Workspace.PdfWorkspace.OpenAsync(source, CancellationToken.None);
            var acquired = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            Task<bool> holding = blocker.RunNonDetachableCpuWorkAsync(() => {
                acquired.SetResult(); release.Task.GetAwaiter().GetResult(); return true;
            }, CancellationToken.None);
            try {
                await acquired.Task;
                Task pending = model.ExtractSelectedCommand.ExecuteAsync(null);
                var job = Assert.Single(services.Jobs.Entries);
                job.CancelCommand.Execute(null); await pending;
                Assert.Equal(services.Localizer.GetOrDefault("Workflow.Status.Cancelled", "Cancelled"), job.Status);
                Assert.False(job.IsActive); Assert.False(job.HasOutput); Assert.False(job.HasRecovery);
                Assert.False(File.Exists(output));
            } finally { release.SetResult(); await holding; }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ReviewUsesEditedWorkspaceAndRejectsStaleApproval(bool stale) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf"), output = Path.Combine(services.Paths.Root, "output.pdf");
            CreatePdf().Save(source);
            byte[] original = File.ReadAllBytes(source);
            MainWindowViewModel? model = null;
            PageExtractionPreviewViewModel? completed = null;
            using (model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => { model!.ClearPageSelectionCommand.Execute(null); return Task.FromResult<string?>(output); },
                reviewPageExtraction: async preview => {
                    Assert.Equal(new[] { 1, 2 }, preview.SelectedPages);
                    preview.PageRange = "4,1,4";
                    Assert.True(preview.CanApply);
                    if (stale) {
                        model!.SetOrganizerSelection([model.OrganizerPages[0]]);
                        await model.RotateRightCommand.ExecuteAsync(null);
                    }
                    return true;
                }, showPageExtractionResult: result => { completed = result; return Task.CompletedTask; })) {
                await model.OpenDocumentAsync(source);
                model.SetOrganizerSelection([model.OrganizerPages[0]]);
                await model.DuplicateSelectedCommand.ExecuteAsync(null);
                model.SetOrganizerSelection([model.OrganizerPages[0], model.OrganizerPages[1]]);
                await model.ExtractSelectedCommand.ExecuteAsync(null);
                Assert.Equal(original, File.ReadAllBytes(source));
                Assert.True(model.IsDirty); Assert.Equal(4, model.Pages.Count); Assert.Equal(source, model.DocumentPath);
                if (stale) {
                    Assert.NotNull(model.ErrorMessage); Assert.Null(completed); Assert.Empty(services.Jobs.Entries);
                    Assert.False(File.Exists(output));
                } else {
                    Assert.Null(model.ErrorMessage); Assert.NotNull(completed); Assert.True(completed.CanOpenOutput);
                    Assert.Equal(new[] { 220D, 200D, 220D }, PdfDocument.Load(File.ReadAllBytes(output)).Inspect().Pages.Select(page => page.Width));
                    var job = Assert.Single(services.Jobs.Entries); Assert.False(job.IsActive); Assert.True(job.HasOutput);
                    await model.UndoCommand.ExecuteAsync(null);
                    Assert.Equal(3, model.Pages.Count); // Extraction creates no edit or undo entry.
                }
            }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ProviderExtractionRequiresConsentAndReportsRecoverableInterruptedWrites(bool consent, bool fail) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf"); CreatePdf().Save(source);
            byte[] original = File.ReadAllBytes(source);
            var output = new TestStorageFile("content://documents/extracted", []) { FailWrite = fail };
            string selected = await services.Storage.RegisterAsync(output.Item, default);
            PageExtractionPreviewViewModel? completed = null;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(selected), reviewPageExtraction: preview => { preview.PageRange = "3,1"; return Task.FromResult(true); },
                confirmProviderWrite: _ => Task.FromResult(consent), showPageExtractionResult: async result => {
                    completed = result;
                    var dialog = new PageExtractionDialog(result) { Width = 380, Height = 460 };
                    try { dialog.Show(); await Layout(dialog); Capture(dialog, $"extract-provider-{fail}"); }
                    finally { dialog.Close(); }
                });
            await model.OpenDocumentAsync(source); model.SelectAllPagesCommand.Execute(null);
            await model.ExtractSelectedCommand.ExecuteAsync(null);
            Assert.Equal(original, File.ReadAllBytes(source)); Assert.False(model.IsDirty);
            if (!consent) {
                Assert.Null(completed); Assert.Equal(0, output.Writes); Assert.Empty(services.Jobs.Entries);
            } else {
                Assert.NotNull(completed); Assert.False(completed.CanRevealOutput);
                var job = Assert.Single(services.Jobs.Entries); Assert.False(job.IsActive);
                Assert.Equal(!fail, completed.CanOpenOutput); Assert.Equal(fail, completed.HasRecovery);
                Assert.Equal(!fail, job.HasOutput); Assert.Equal(fail, job.HasRecovery);
                byte[] bytes;
                if (fail) {
                    Assert.NotNull(model.ErrorMessage);
                    var recovery = Assert.Single(services.WorkflowRecovery.GetRecoveries());
                    await services.WorkflowRecovery.VerifyAsync(recovery); bytes = File.ReadAllBytes(recovery.FilePath);
                } else bytes = output.Bytes;
                Assert.Equal(new[] { 220D, 200D }, PdfDocument.Load(bytes).Inspect().Pages.Select(page => page.Width));
            }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(380, 460, false)]
    [InlineData(580, 560, true)]
    public async Task RenderedPageOrderReviewValidatesInputAndOpensTheSavedResult(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf"); CreatePdf().Save(source);
            string destination = Path.Combine(services.Paths.Root, new string('x', 140) + ".pdf");
            var owner = new Window { Width = 960, Height = 620 };
            string? opened = null;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(destination),
                reviewPageExtraction: preview => new PageExtractionDialog(preview) { Width = width, Height = height }.ShowDialog<bool>(owner),
                showPageExtractionResult: result => new PageExtractionDialog(result) { Width = width, Height = height }.ShowDialog(owner),
                openDocumentInTab: (path, _) => { opened = path; return Task.CompletedTask; });
            try {
                owner.Show(); await model.OpenDocumentAsync(source); model.SelectAllPagesCommand.Execute(null);
                Task pending = model.ExtractSelectedCommand.ExecuteAsync(null);
                var dialog = Assert.Single(owner.OwnedWindows.OfType<PageExtractionDialog>());
                var preview = Assert.IsType<PageExtractionPreviewViewModel>(dialog.DataContext);
                var input = dialog.FindControl<TextBox>("PageRangeInput")!;
                foreach (string invalid in new[] { "", "0", "1-100001", "4", "abc" }) {
                    input.Text = invalid; await Layout(dialog); Assert.False(preview.CanApply); Assert.Empty(preview.Pages);
                }
                Capture(dialog, $"extract-invalid-{width}-{dark}");
                input.Text = "3,1,3"; await Layout(dialog);
                Assert.True(preview.CanApply); Assert.Equal(new[] { 3, 1, 3 }, preview.SelectedPages);
                Assert.True(dialog.FindControl<ListBox>("PageOrder")!.Bounds.Height >= 60);
                Capture(dialog, $"extract-preview-{width}-{dark}");
                Click(dialog, services.Localizer.Get("Organizer.ExtractCreate"));
                DateTime deadline = DateTime.UtcNow.AddSeconds(10);
                PageExtractionDialog? resultDialog = null;
                while (DateTime.UtcNow < deadline) {
                    resultDialog = owner.OwnedWindows.OfType<PageExtractionDialog>().SingleOrDefault();
                    if (resultDialog?.DataContext is PageExtractionPreviewViewModel { HasResult: true }) break;
                    await Task.Delay(10);
                }
                Assert.NotNull(resultDialog); Assert.True(preview.HasResult); await Layout(resultDialog);
                Capture(resultDialog, $"extract-result-{width}-{dark}");
                Click(resultDialog, services.Localizer.Get("Organizer.ExtractOpen"));
                await preview.OpenOutputCommand.ExecutionTask!; Assert.Equal(destination, opened);
                Click(resultDialog, services.Localizer.Get("Common.Close")); await pending;
                Assert.Null(model.ErrorMessage);
                Assert.Equal(new[] { 220D, 200D, 220D }, PdfDocument.Load(File.ReadAllBytes(destination)).Inspect().Pages.Select(page => page.Width));
                string? captureRoot = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_EXTRACTION_CAPTURE_DIR");
                if (!string.IsNullOrWhiteSpace(captureRoot)) File.Copy(destination, Path.Combine(captureRoot, $"extracted-{width}.pdf"), overwrite: true);
            } finally { foreach (var dialog in owner.OwnedWindows.ToArray()) dialog.Close(); owner.Close(); }
            return true;
        }, CancellationToken.None);
    }
    private static PdfDocument CreatePdf() => PdfDocument.Create(document => {
        for (int index = 0; index < 3; index++) { int width = 200 + index * 10; document.Page(page => page.Size(width, 300)); }
    });
    private static async Task Layout(Window window) { window.UpdateLayout(); await Task.Delay(30); window.UpdateLayout(); }
    private static void Click(Window window, string label) {
        var button = window.GetVisualDescendants().OfType<Button>().Single(item => Equals(item.Content, label));
        Assert.True(button.IsVisible); Assert.True(button.IsEnabled); button.Focus();
        button.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });
    }
    private static void Capture(Window window, string name) {
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_EXTRACTION_CAPTURE_DIR");
        if (string.IsNullOrEmpty(root)) return;
        Directory.CreateDirectory(root); using var bitmap = window.CaptureRenderedFrame();
        Assert.NotNull(bitmap); bitmap.Save(Path.Combine(root, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
