using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class ComparisonNavigationTests {
    [Fact]
    public async Task ComparesOneHundredA4PagesWithoutExhaustingPixelBudget() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string primary = Path.Combine(services.Paths.Root, "hundred-pages.pdf");
            string other = Path.Combine(services.Paths.Root, "hundred-pages-comparison.pdf");
            PdfDocument.Create(document => {
                for (int page = 0; page < 100; page++) document.Page(p => p.Size(595, 842));
            }).Save(primary);
            File.Copy(primary, other);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await model.OpenDocumentAsync(primary);
            await model.OpenComparisonDocumentAsync(other);
            await model.ComparePagesCommand.ExecuteAsync(null);
            Assert.Equal("No rendered differences found.", model.ComparisonSummary);
            Assert.Empty(model.ComparisonDifferences);
            return true;
        }, default);
    }

    [Theory]
    [InlineData(960, 640, false)]
    [InlineData(1280, 800, true)]
    public async Task NavigateChangedResizedAndUnpairedPagesAndInvalidateOnMutation(int width, int height, bool extraComparisonPage) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = extraComparisonPage ? StudioThemePreference.Dark : StudioThemePreference.Light });
            string primary = Path.Combine(services.Paths.Root, "current.pdf");
            string other = Path.Combine(services.Paths.Root, "comparison.pdf");
            Create(primary, "Original content", false, !extraComparisonPage);
            Create(other, "Revised content", true, extraComparisonPage);
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(primary);
                var model = window.ViewModel;
                await model.OpenComparisonDocumentAsync(other);
                await model.ComparePagesCommand.ExecuteAsync(null);
                Assert.Equal(3, model.ComparisonDifferences.Count);
                Assert.Equal(new[] { 2, 3, 4 }, model.ComparisonDifferences.Select(item => item.PageNumber));
                Assert.Equal(2, model.SelectedPage!.PageNumber);
                Assert.Equal(2, model.ComparisonSelectedPage!.PageNumber);
                Assert.NotNull(model.ComparisonDifferenceImage);
                Assert.Equal(model.ComparisonDifferenceImage.PixelSize.Width * model.Zoom, model.ComparisonDifferenceWidth);
                Assert.NotNull(model.SelectedComparisonDifference!.Comparison!.ChangedBounds);
                window.UpdateLayout();
                Capture(window, "comparison-changed-" + width);
                model.NextPageCommand.Execute(null);
                Assert.Equal(3, model.SelectedComparisonDifference!.PageNumber);
                Assert.True(model.SelectedComparisonDifference.Comparison!.HasSizeDifference);
                window.UpdateLayout(); Capture(window, "comparison-ordinary-next-" + width);
                model.PreviousPageCommand.Execute(null);
                Assert.Equal(2, model.SelectedComparisonDifference!.PageNumber);
                await model.ActivateComparisonPageLinkAsync("FirstPage");
                Assert.Equal(1, model.SelectedPage!.PageNumber);
                Assert.Null(model.SelectedComparisonDifference);
                Assert.Null(model.ComparisonDifferenceImage);
                Assert.False(model.ShowComparisonDifferenceImage);
                for (int attempt = 0; attempt < 100 && (model.SelectedPage!.IsRendering || model.ComparisonSelectedPage!.IsRendering); attempt++)
                    await Task.Delay(20);
                window.UpdateLayout(); Capture(window, "comparison-unchanged-first-" + width);
                model.NextComparisonDifferenceCommand.Execute(null);
                Assert.Equal(2, model.SelectedComparisonDifference!.PageNumber);
                model.NextComparisonDifferenceCommand.Execute(null);
                Assert.True(model.SelectedComparisonDifference!.Comparison!.HasSizeDifference);
                model.NextComparisonDifferenceCommand.Execute(null);
                Assert.Equal(4, model.SelectedComparisonDifference!.PageNumber);
                Assert.Null(model.ComparisonDifferenceImage);
                if (extraComparisonPage) { Assert.Null(model.SelectedPage); Assert.Equal(4, model.ComparisonSelectedPage!.PageNumber); }
                else { Assert.Equal(4, model.SelectedPage!.PageNumber); Assert.Null(model.ComparisonSelectedPage); }
                window.UpdateLayout(); Capture(window, "comparison-unpaired-" + width);
                model.PreviousComparisonDifferenceCommand.Execute(null);
                Assert.Equal(3, model.SelectedPage!.PageNumber);
                model.SetOrganizerSelection([model.OrganizerPages[0]]);
                await model.DuplicateSelectedCommand.ExecuteAsync(null);
                Assert.True(model.IsDirty, model.ErrorMessage);
                Assert.Empty(model.ComparisonDifferences);
                Assert.Null(model.ComparisonDifferenceImage);
                Assert.True(model.IsComparisonOpen);
                model.CloseComparisonCommand.Execute(null);
                Assert.False(model.IsComparisonOpen);
            } finally { window.Close(); }
            return true;
        }, default);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CancelOrCloseDoesNotWaitForAnOccupiedCpuGate(bool close) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string source = Path.Combine(services.Paths.Root, "current.pdf");
            string other = Path.Combine(services.Paths.Root, "other.pdf");
            Create(source, "Original", false, false);
            Create(other, "Changed", false, false);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await model.OpenDocumentAsync(source);
            await model.OpenComparisonDocumentAsync(other);
            using var blocker = await PdfWorkspace.OpenAsync(source, default);
            var acquired = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            var release = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
            Task<bool> holding = blocker.RunNonDetachableCpuWorkAsync(() => {
                acquired.SetResult(); release.Task.GetAwaiter().GetResult(); return true;
            }, default);
            try {
                await acquired.Task;
                Task pending = model.ComparePagesCommand.ExecuteAsync(null);
                Assert.True(model.IsComparingPages);
                if (close) model.CloseComparisonCommand.Execute(null);
                else model.CancelPageComparisonCommand.Execute(null);
                await pending.WaitAsync(TimeSpan.FromSeconds(3));
                Assert.False(model.IsComparingPages);
                Assert.False(holding.IsCompleted);
                Assert.Empty(model.ComparisonDifferences);
                Assert.Null(model.ComparisonDifferenceImage);
                Assert.Equal(!close, model.IsComparisonOpen);
            } finally { release.TrySetResult(); await holding; }
            return true;
        }, default);
    }

    private static void Create(string path, string text, bool resize, bool extra) => PdfDocument.Create(c => {
        c.Page(p => p.Size(300, 400).Content(content => content.Item(i => i.Paragraph(t => t.Text("Unchanged first page")))));
        c.Page(p => p.Size(300, 400).Content(content => content.Item(i => i.Paragraph(t => t.Text(text)))));
        c.Page(p => p.Size(resize ? 320 : 300, 400));
        if (extra) c.Page(p => p.Size(300, 400).Content(content => content.Item(i => i.Paragraph(t => t.Text("Unpaired last page")))));
    }).Save(path);

    private static void Capture(Window window, string name) {
        using var image = window.CaptureRenderedFrame();
        Assert.NotNull(image);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        image.Save(Path.Combine(output, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
