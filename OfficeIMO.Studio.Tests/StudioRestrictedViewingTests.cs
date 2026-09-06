using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioRestrictedViewingTests {
    [Fact]
    public async Task AccessibilitySearchRetainsMultipleMatchesPerPageAndHonorsCancellation() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "multipage-restricted.pdf");
            PdfDocument.Create(document => {
                for (int i = 0; i < 60; i++) document.Page(page => page.Content(content => content.Text("Needle and needle")));
            }, new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("reader") {
                OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.Accessibility
            })).Save(source);
            using var workspace = await PdfWorkspace.OpenAsync(source, default, services.Recovery, "reader");
            var reader = PdfDocumentSession.FromWorkspace(workspace);
            var hits = await reader.SearchAsync("needle", default);
            Assert.Equal(Enumerable.Range(1, 60).SelectMany(page => new[] { page, page }), hits.Select(hit => hit.PageNumber));
            Assert.Empty(await reader.SearchAsync("missing", default));
            using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.SearchAsync("needle", cancelled.Token));
            using var cancellation = new CancellationTokenSource();
            var progress = new CancelSearchProgress(cancellation);
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.SearchAsync("needle", cancellation.Token, progress));
            Assert.Equal(1, progress.Reports);
            return true;
        }, CancellationToken.None);
    }

    private sealed class CancelSearchProgress(CancellationTokenSource cancellation) : IProgress<double> {
        public int Reports { get; private set; }
        public void Report(double value) { Reports++; cancellation.Cancel(); }
    }

    [Theory]
    [InlineData(960, 620, false, PdfStandardPermissions.None)]
    [InlineData(1280, 800, true, PdfStandardPermissions.Accessibility)]
    public async Task PasswordDialogOpensRestrictedPagesInTheRealDocumentTab(int width, int height, bool dark, PdfStandardPermissions permissions) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "Restricted document.pdf");
            PdfDocument.Create(document => {
                document.Page(page => page.Size(400, 500).Content(content => content.Text("Visible protected document")));
                document.Page(page => page.Size(400, 500).Content(content => content.Text("Second protected page")));
            }, new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("reader") {
                OwnerPassword = "owner", AllowedPermissions = permissions
            })).Save(source);
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                Task opening = window.TabHost.OpenDocumentAsync(source);
                foreach (string password in new[] { "incorrect", "reader" }) {
                    DateTime deadline = DateTime.UtcNow.AddSeconds(10);
                    while (!window.OwnedWindows.OfType<PdfPasswordDialog>().Any() && !opening.IsCompleted && DateTime.UtcNow < deadline) await Task.Delay(10);
                    var dialog = Assert.Single(window.OwnedWindows.OfType<PdfPasswordDialog>());
                    dialog.UpdateLayout();
                    Assert.Single(dialog.GetVisualDescendants().OfType<TextBox>()).Text = password;
                    var open = dialog.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, services.Localizer.Get("Common.Open")));
                    open.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                }
                await opening;
                var model = window.ViewModel;
                Assert.Null(model.ErrorMessage); Assert.Equal(2, model.Pages.Count);
                Assert.Equal(permissions != PdfStandardPermissions.None, model.SearchCommand.CanExecute(null));
                Assert.False(model.CanExtractPages);
                foreach (var page in model.Pages) {
                    page.AttachToViewport(); await page.EnsureRenderedAsync();
                    Assert.Null(page.RenderError); Assert.NotNull(page.PageImage);
                    Assert.NotNull(page.Scene); Assert.Null(page.Scene.Interactions);
                }
                model.ZoomInCommand.Execute(null);
                model.NextPageCommand.Execute(null);
                model.FitPageCommand.Execute(null);
                foreach (var page in model.Pages) {
                    await page.EnsureRenderedAsync(); Assert.Null(page.RenderError);
                }
                window.UpdateLayout(); await Task.Delay(80); window.UpdateLayout();
                Assert.False(string.IsNullOrWhiteSpace(model.SecurityWarning));
                Assert.Contains(window.GetVisualDescendants().OfType<TextBlock>(), text => text.IsEffectivelyVisible && text.Text == model.SecurityWarning);
                window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!.FocusSearch();
                model.SearchQuery = "Visible";
                if (model.CanSearchDocument) {
                    await model.SearchCommand.ExecuteAsync(null);
                    Assert.Single(model.SearchResults);
                    Assert.Single(model.Pages[0].SearchHighlights);
                    Assert.NotNull(model.Pages[0].ActiveSearchHighlight);
                    model.Pages[0].AttachToViewport(); await model.Pages[0].EnsureRenderedAsync();
                    Assert.NotNull(model.Pages[0].Scene); Assert.Null(model.Pages[0].Scene!.Interactions);
                } else {
                    Assert.Empty(model.SearchResults);
                    Assert.Equal(services.Localizer.Get("Capability.SearchRestricted"), model.SearchPosition);
                }
                window.UpdateLayout(); await Task.Delay(80); window.UpdateLayout();
                string? root = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_RESTRICTED_CAPTURE_DIR");
                if (!string.IsNullOrEmpty(root)) {
                    Directory.CreateDirectory(root); using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
                    frame.Save(Path.Combine(root, $"restricted-{width}-{dark}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
            } finally { foreach (var dialog in window.OwnedWindows.ToArray()) dialog.Close(); window.Close(); window.TabHost.Dispose(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(PdfStandardPermissions.None)]
    [InlineData(PdfStandardPermissions.Accessibility)]
    [InlineData(PdfStandardPermissions.CopyContents)]
    public async Task UserPasswordViewingKeepsLogicalContentAndSearchPermissionBound(PdfStandardPermissions permissions) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "restricted.pdf");
            PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("Visible protected document"))), new PdfOptions().SetEncryption(new PdfStandardEncryptionOptions("reader") {
                OwnerPassword = "owner", AllowedPermissions = permissions
            })).Save(source);
            byte[] original = File.ReadAllBytes(source);
            using var workspace = await PdfWorkspace.OpenAsync(source, default, services.Recovery, "reader");
            bool contentAllowed = permissions.HasFlag(PdfStandardPermissions.CopyContents);
            Assert.Equal(contentAllowed, workspace.ViewInfo.CanExtractContent);
            Assert.Equal(contentAllowed, workspace.DocumentInfo is not null);
            var reader = PdfDocumentSession.FromWorkspace(workspace);
            var scene = await reader.LoadPageSceneAsync(1, default);
            Assert.Equal(contentAllowed, scene.Interactions is not null);
            if (!contentAllowed) {
                Assert.True(scene.RequiresRasterFallback);
                Assert.Empty(scene.Drawing.Elements);
            }
            var rendered = await reader.RenderPageAsync(1, 1, default);
            Assert.NotNull(rendered);
            if (permissions == PdfStandardPermissions.None)
                await Assert.ThrowsAsync<InvalidOperationException>(() => reader.SearchAsync("Visible", default));
            else Assert.Single(await reader.SearchAsync("Visible", default));
            Assert.False(workspace.CanExtractPages); Assert.False(workspace.CanChangeEncryption("reader"));
            Assert.True(workspace.CanChangeEncryption("owner"));
            Assert.Equal(original, File.ReadAllBytes(source)); Assert.False(workspace.IsDirty);
            return true;
        }, CancellationToken.None);
    }
}
