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

public sealed class StudioSplitPreviewTests {
    [Fact]
    public async Task CancellingWhileWaitingForTheCpuSlotReportsNoPublication() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf");
            CreatePdf().Save(source);
            string output = Path.Combine(services.Paths.Root, "output");
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickOutputFolder: _ => Task.FromResult<string?>(output), reviewPageSplit: _ => Task.FromResult(true));
            await model.OpenDocumentAsync(source);
            using var blocker = await Features.Workspace.PdfWorkspace.OpenAsync(source, CancellationToken.None);
            var acquired = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            Task<bool> holding = blocker.RunNonDetachableCpuWorkAsync(() => {
                acquired.SetResult(); release.Task.GetAwaiter().GetResult(); return true;
            }, CancellationToken.None);
            try {
                await acquired.Task;
                Task pending = model.SplitCommand.ExecuteAsync(null);
                var job = Assert.Single(services.Jobs.Entries);
                job.CancelCommand.Execute(null);
                await pending;
                Assert.Equal(services.Localizer.GetOrDefault("Workflow.Status.Cancelled", "Cancelled"), job.Status);
                Assert.False(job.HasOutput); Assert.False(job.HasRecovery); Assert.False(job.IsActive);
                Assert.False(Directory.Exists(output));
            } finally { release.SetResult(); await holding; }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ProviderSplitRequiresConsentAndRetainsVerifiedFilesAndRecovery(bool consent, bool failSecond) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf");
            CreatePdf().Save(source);
            var folder = new StudioProviderOutputFolderTests.OutputFolder { FailSecondCreation = failSecond };
            string? selected = await services.Storage.RegisterFolderAsync([folder.Item], default);
            PageSplitPreviewViewModel? completed = null;
            string? opened = null;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickOutputFolder: _ => Task.FromResult(selected), reviewPageSplit: preview => { preview.PagesPerPart = "2"; return Task.FromResult(true); },
                confirmProviderWrite: _ => Task.FromResult(consent),
                showPageSplitResult: result => { completed = result; return Task.CompletedTask; },
                openDocumentInTab: (path, _) => { opened = path; return Task.CompletedTask; });
            await model.OpenDocumentAsync(source);
            await model.SplitCommand.ExecuteAsync(null);
            if (!consent) {
                Assert.Equal(0, folder.Creations); Assert.Empty(services.Jobs.Entries); Assert.Null(completed);
            } else {
                Assert.NotNull(completed);
                Assert.Equal(failSecond ? 1 : 2, completed.Files.Count);
                var job = Assert.Single(services.Jobs.Entries);
                Assert.False(job.IsActive); Assert.True(job.HasOutput); Assert.Equal(failSecond, job.HasRecovery);
                foreach (var file in completed.Files) {
                    var stored = folder.Files.Values.Single(item => item.Location.AbsoluteUri == file.Path);
                    Assert.Equal(file.PageCount, PdfDocument.Load(stored.Bytes).Inspect().PageCount);
                }
                if (failSecond) await services.WorkflowRecovery.VerifyAsync(Assert.Single(job.Recoveries));
                await model.Jobs.OpenOutputCommand.ExecuteAsync(job);
                Assert.Equal(completed.Files[0].Path, opened);
            }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PreviewUsesDirtyPagesAndRejectsChangesWhileReviewing(bool changeDuringReview) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf");
            CreatePdf().Save(source);
            string folder = Path.Combine(services.Paths.Root, "output");
            PageSplitPreviewViewModel? completed = null;
            MainWindowViewModel? model = null;
            using (model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickOutputFolder: _ => Task.FromResult<string?>(folder),
                reviewPageSplit: async preview => {
                    Assert.Equal(5, preview.PageCount);
                    preview.PagesPerPart = "2";
                    Assert.Equal(new[] { 2, 2, 1 }, preview.Parts.Select(part => part.PageCount));
                    if (changeDuringReview) await model!.RotateRightCommand.ExecuteAsync(null);
                    return true;
                }, showPageSplitResult: result => { completed = result; return Task.CompletedTask; })) {
                await model.OpenDocumentAsync(source);
                model.SetOrganizerSelection([model.OrganizerPages[0]]);
                await model.DuplicateSelectedCommand.ExecuteAsync(null);
                await model.SplitCommand.ExecuteAsync(null);
                Assert.True(model.IsDirty);
                Assert.Equal(4, PdfDocument.Load(File.ReadAllBytes(source)).Inspect().PageCount);
                if (changeDuringReview) {
                    Assert.NotNull(model.ErrorMessage); Assert.Null(completed); Assert.Empty(services.Jobs.Entries);
                    Assert.False(Directory.Exists(folder));
                } else {
                    Assert.Null(model.ErrorMessage);
                    Assert.NotNull(completed);
                    Assert.Equal(3, completed.Files.Count);
                    Assert.Equal(new[] { 200D, 200D, 210D, 220D, 230D }, completed.Files
                        .SelectMany(file => PdfDocument.Load(File.ReadAllBytes(file.Path)).Inspect().Pages).Select(page => page.Width));
                    var job = Assert.Single(services.Jobs.Entries);
                    Assert.False(job.IsActive); Assert.True(job.HasOutput); Assert.False(job.HasRecovery);
                }
            }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(380, 400, false)]
    [InlineData(580, 540, true)]
    public async Task RenderedPreviewRejectsInvalidCountsAndOpensVerifiedResults(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf");
            CreatePdf().Save(source);
            var owner = new Window { Width = 960, Height = 620 };
            string? opened = null;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickOutputFolder: _ => Task.FromResult<string?>(Path.Combine(services.Paths.Root, "output")),
                reviewPageSplit: preview => new PageSplitDialog(preview) { Width = width, Height = height }.ShowDialog<bool>(owner),
                showPageSplitResult: result => new PageSplitDialog(result) { Width = width, Height = height }.ShowDialog(owner),
                openDocumentInTab: (path, _) => { opened = path; return Task.CompletedTask; });
            try {
                owner.Show();
                await model.OpenDocumentAsync(source);
                Task pending = model.SplitCommand.ExecuteAsync(null);
                var dialog = Assert.Single(owner.OwnedWindows.OfType<PageSplitDialog>());
                var preview = Assert.IsType<PageSplitPreviewViewModel>(dialog.DataContext);
                var input = dialog.FindControl<TextBox>("PartSizeInput")!;
                foreach (string invalid in new[] { "abc", "0", "1.5", "", "99999999999999999999" }) {
                    input.Text = invalid; await Layout(dialog);
                    Assert.False(preview.CanApply); Assert.Empty(preview.Parts);
                }
                Capture(dialog, $"split-invalid-{width}-{dark}");
                input.Text = "2"; await Layout(dialog);
                Assert.True(preview.CanApply); Assert.Equal(2, preview.Parts.Count);
                Capture(dialog, $"split-preview-{width}-{dark}");
                Click(dialog, services.Localizer.Get("Organizer.SplitCreate"));
                DateTime deadline = DateTime.UtcNow.AddSeconds(10);
                PageSplitDialog? resultDialog = null;
                while (DateTime.UtcNow < deadline) {
                    resultDialog = owner.OwnedWindows.OfType<PageSplitDialog>().SingleOrDefault();
                    if (resultDialog?.DataContext is PageSplitPreviewViewModel { HasResult: true }) break;
                    await Task.Delay(10);
                }
                Assert.NotNull(resultDialog);
                await Layout(resultDialog);
                Assert.True(preview.HasResult);
                Assert.Equal(2, preview.Files.Count);
                Capture(resultDialog, $"split-results-{width}-{dark}");
                Click(resultDialog, services.Localizer.Get("Organizer.SplitOpen"));
                await preview.OpenFileCommand.ExecutionTask!;
                Assert.Equal(preview.Files[0].Path, opened);
                Click(resultDialog, services.Localizer.Get("Common.Close"));
                await pending;
                Assert.Null(model.ErrorMessage);
            } finally {
                foreach (var dialog in owner.OwnedWindows.ToArray()) dialog.Close();
                owner.Close();
            }
            return true;
        }, CancellationToken.None);
    }
    private static PdfDocument CreatePdf() => PdfDocument.Create(document => {
        for (int index = 0; index < 4; index++) { int width = 200 + index * 10; document.Page(page => page.Size(width, 300)); }
    });
    private static async Task Layout(Window window) { window.UpdateLayout(); await Task.Delay(30); window.UpdateLayout(); }
    private static void Click(Window window, string label) {
        var button = window.GetVisualDescendants().OfType<Button>().Single(item => Equals(item.Content, label));
        Assert.True(button.IsVisible); Assert.True(button.IsEnabled);
        button.Focus(); button.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });
    }
    private static void Capture(Window window, string name) {
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_SPLIT_CAPTURE_DIR");
        if (string.IsNullOrEmpty(root)) return;
        Directory.CreateDirectory(root);
        using var bitmap = window.CaptureRenderedFrame();
        Assert.NotNull(bitmap); bitmap.Save(Path.Combine(root, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
