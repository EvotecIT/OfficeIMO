using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioProtectionPreviewTests {
    [Fact]
    public async Task CancellationAtTheCpuGateLeavesNoOutputOrRecovery() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf"), output = Path.Combine(services.Paths.Root, "output.pdf");
            CreatePdf().Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(output), reviewProtection: _ => Task.FromResult(true));
            await model.OpenDocumentAsync(source); model.ProtectUserPassword = "reader"; model.ProtectConfirmPassword = "reader";
            using var blocker = await Features.Workspace.PdfWorkspace.OpenAsync(source, CancellationToken.None);
            var acquired = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            Task<bool> holding = blocker.RunNonDetachableCpuWorkAsync(() => {
                acquired.SetResult(); release.Task.GetAwaiter().GetResult(); return true;
            }, CancellationToken.None);
            try {
                await acquired.Task;
                Task pending = model.SaveProtectedCopyCommand.ExecuteAsync(null);
                var job = Assert.Single(services.Jobs.Entries); job.CancelCommand.Execute(null); await pending;
                Assert.False(job.IsActive); Assert.False(job.HasOutput); Assert.False(job.HasRecovery);
                Assert.Equal(services.Localizer.GetOrDefault("Workflow.Status.Cancelled", "Cancelled"), job.Status);
                Assert.False(File.Exists(output)); Assert.Equal("reader", model.ProtectUserPassword);
            } finally { release.SetResult(); await holding; }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CapturedSettingsExportEditsAndRejectStaleReview(bool stale) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf"), output = Path.Combine(services.Paths.Root, "protected.pdf");
            CreatePdf().Save(source); byte[] original = File.ReadAllBytes(source);
            MainWindowViewModel? model = null; PdfProtectionPreviewViewModel? completed = null;
            using (model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => {
                    model!.ProtectUserPassword = "later-reader"; model.ProtectConfirmPassword = "later-reader";
                    model.ProtectOwnerPassword = "later-owner"; model.ProtectAllowCopy = true;
                    return Task.FromResult<string?>(output);
                }, reviewProtection: async preview => {
                    Assert.Contains(preview.Details, value => value == services.Localizer.Format("Protection.Restricted", services.Localizer.Get("DocumentWorkspace.CopyContent")));
                    Assert.DoesNotContain("captured-reader", string.Join(' ', preview.Details));
                    if (stale) await model!.RotateRightCommand.ExecuteAsync(null);
                    return true;
                }, showProtectionResult: result => { completed = result; return Task.CompletedTask; })) {
                await model.OpenDocumentAsync(source); model.SetOrganizerSelection([model.OrganizerPages[0]]);
                await model.DuplicateSelectedCommand.ExecuteAsync(null);
                model.ProtectUserPassword = "captured-reader"; model.ProtectConfirmPassword = "captured-reader";
                model.ProtectOwnerPassword = "captured-owner"; model.ProtectAllowCopy = false;
                await model.SaveProtectedCopyCommand.ExecuteAsync(null);
                Assert.Equal(original, File.ReadAllBytes(source)); Assert.True(model.IsDirty); Assert.Equal(3, model.Pages.Count);
                Assert.Equal("later-reader", model.ProtectUserPassword); Assert.Equal("later-owner", model.ProtectOwnerPassword);
                if (stale) {
                    Assert.NotNull(model.ErrorMessage); Assert.Null(completed); Assert.Empty(services.Jobs.Entries); Assert.False(File.Exists(output));
                } else {
                    Assert.Null(model.ErrorMessage); Assert.NotNull(completed); Assert.True(completed.CanOpenOutput); Assert.NotNull(completed.Verification);
                    var document = PdfDocument.Load(File.ReadAllBytes(output), new PdfLoadOptions { Password = "captured-owner" });
                    Assert.Equal(3, document.Inspect().Pages.Count); Assert.True(document.Inspect().Security.HasEncryption);
                    Assert.Throws<PdfInvalidPasswordException>(() => PdfDocument.Load(File.ReadAllBytes(output), new PdfLoadOptions { Password = "later-reader" }).Inspect());
                    var job = Assert.Single(services.Jobs.Entries); Assert.False(job.IsActive); Assert.True(job.HasOutput);
                    await model.UndoCommand.ExecuteAsync(null); Assert.Equal(2, model.Pages.Count); Assert.False(model.CanUndo);
                }
            }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ProviderProtectionUsesConsentAndRetainsVerifiedRecovery(bool consent, bool fail) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf"); CreatePdf().Save(source);
            var output = new TestStorageFile("content://documents/protected", []) { FailWrite = fail };
            string destination = await services.Storage.RegisterAsync(output.Item, default);
            PdfProtectionPreviewViewModel? completed = null;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(destination), reviewProtection: _ => Task.FromResult(true),
                confirmProviderWrite: _ => Task.FromResult(consent), showProtectionResult: async result => {
                    completed = result; var dialog = new PdfProtectionDialog(result) { Width = 380, Height = 440 };
                    try { dialog.Show(); await Layout(dialog); Capture(dialog, $"protection-provider-{fail}"); } finally { dialog.Close(); }
                });
            await model.OpenDocumentAsync(source); model.ProtectUserPassword = "reader"; model.ProtectConfirmPassword = "reader";
            await model.SaveProtectedCopyCommand.ExecuteAsync(null);
            Assert.False(model.IsDirty);
            if (!consent) { Assert.Null(completed); Assert.Equal(0, output.Writes); Assert.Empty(services.Jobs.Entries); }
            else {
                Assert.NotNull(completed); Assert.Equal(!fail, completed.CanOpenOutput); Assert.False(completed.CanRevealOutput);
                Assert.Equal(fail, completed.HasRecovery);
                var job = Assert.Single(services.Jobs.Entries); Assert.False(job.IsActive); Assert.Equal(fail, job.HasRecovery);
                byte[] bytes;
                if (fail) { var recovery = Assert.Single(services.WorkflowRecovery.GetRecoveries()); await services.WorkflowRecovery.VerifyAsync(recovery); bytes = File.ReadAllBytes(recovery.FilePath); }
                else bytes = output.Bytes;
                Assert.Equal(2, PdfDocument.Load(bytes, new PdfLoadOptions { Password = "reader" }).Inspect().Pages.Count);
            }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(380, 440, false)]
    [InlineData(580, 600, true)]
    public async Task RenderedReviewCreatesProtectedAndUnencryptedCopies(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf"); CreatePdf().Save(source);
            string destination = Path.Combine(services.Paths.Root, new string('x', 140) + ".pdf");
            var owner = new Window { Width = 960, Height = 620 }; string? opened = null;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickSavePdf: _ => Task.FromResult<string?>(destination),
                reviewProtection: preview => new PdfProtectionDialog(preview) { Width = width, Height = height }.ShowDialog<bool>(owner),
                showProtectionResult: result => new PdfProtectionDialog(result) { Width = width, Height = height }.ShowDialog(owner),
                openDocumentInTab: (path, _) => { opened = path; return Task.CompletedTask; },
                promptPdfPassword: (_, _, _) => Task.FromResult<string?>("owner"));
            try {
                owner.Show(); await model.OpenDocumentAsync(source);
                model.ProtectUserPassword = "reader"; model.ProtectConfirmPassword = "reader"; model.ProtectOwnerPassword = "owner";
                model.ProtectAllowCopy = false;
                for (int operation = 0; operation < 2; operation++) {
                    if (operation == 1) { await model.OpenDocumentAsync(destination); destination = Path.Combine(services.Paths.Root, "unencrypted.pdf"); model.CurrentOwnerPassword = "owner"; }
                    Task pending = operation == 0 ? model.SaveProtectedCopyCommand.ExecuteAsync(null) : model.SaveDecryptedCopyCommand.ExecuteAsync(null);
                    var dialog = Assert.Single(owner.OwnedWindows.OfType<PdfProtectionDialog>());
                    var preview = Assert.IsType<PdfProtectionPreviewViewModel>(dialog.DataContext);
                    await Layout(dialog); Capture(dialog, $"protection-preview-{operation}-{width}-{dark}");
                    var scroll = dialog.FindControl<ScrollViewer>("ReviewContent")!;
                    Assert.True(scroll.Bounds.Height > 100); scroll.ScrollToEnd(); await Layout(dialog);
                    Capture(dialog, $"protection-preview-end-{operation}-{width}-{dark}");
                    Click(dialog, services.Localizer.Get("Protection.Create"));
                    DateTime deadline = DateTime.UtcNow.AddSeconds(15); PdfProtectionDialog? resultDialog = null;
                    while (DateTime.UtcNow < deadline) {
                        resultDialog = owner.OwnedWindows.OfType<PdfProtectionDialog>().SingleOrDefault();
                        if (resultDialog?.DataContext is PdfProtectionPreviewViewModel { HasResult: true }) break;
                        await Task.Delay(10);
                    }
                    Assert.NotNull(resultDialog); Assert.True(preview.HasResult); Assert.True(preview.CanOpenOutput);
                    resultDialog.FindControl<ScrollViewer>("ReviewContent")!.ScrollToEnd(); await Layout(resultDialog);
                    Capture(resultDialog, $"protection-result-{operation}-{width}-{dark}");
                    Click(resultDialog, services.Localizer.Get("Organizer.ExtractOpen")); await preview.OpenOutputCommand.ExecutionTask!;
                    Assert.Equal(destination, opened); Click(resultDialog, services.Localizer.Get("Common.Close")); await pending;
                    Assert.Null(model.ErrorMessage);
                    var info = PdfDocument.Load(File.ReadAllBytes(destination), operation == 0 ? new PdfLoadOptions { Password = "owner" } : null).Inspect();
                    Assert.Equal(2, info.Pages.Count); Assert.Equal(operation == 0, info.Security.HasEncryption);
                }
            } finally { foreach (var dialog in owner.OwnedWindows.ToArray()) dialog.Close(); owner.Close(); }
            return true;
        }, CancellationToken.None);
    }
    private static PdfDocument CreatePdf() => PdfDocument.Create(document => { document.Page(page => page.Size(200, 300)); document.Page(page => page.Size(210, 300)); });
    private static async Task Layout(Window window) { window.UpdateLayout(); await Task.Delay(30); window.UpdateLayout(); }
    private static void Click(Window window, string label) {
        var button = window.GetVisualDescendants().OfType<Button>().Single(item => Equals(item.Content, label));
        Assert.True(button.IsVisible); Assert.True(button.IsEnabled);
        var point = button.TranslatePoint(default, window)!.Value;
        Assert.InRange(point.Y + button.Bounds.Height, 0, window.Bounds.Height);
        button.Focus(); button.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Enter });
    }
    private static void Capture(Window window, string name) {
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_PROTECTION_CAPTURE_DIR"); if (string.IsNullOrEmpty(root)) return;
        Directory.CreateDirectory(root); using var bitmap = window.CaptureRenderedFrame(); Assert.NotNull(bitmap);
        bitmap.Save(Path.Combine(root, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
