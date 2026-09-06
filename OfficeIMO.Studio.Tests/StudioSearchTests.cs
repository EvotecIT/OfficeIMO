using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Drawing;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioSearchTests {
    [Theory]
    [InlineData(960, 620, false, false)]
    [InlineData(1280, 800, true, true)]
    public async Task SearchNavigatesIndividualOccurrencesAndClearsStaleHighlights(int width, int height, bool dark, bool rotate) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "Find occurrences.pdf");
            var document = PdfDocument.Create(compose => {
                compose.Page(page => page.Size(400, 500).Content(content => content.Text("Needle and needle")));
                compose.Page(page => page.Size(400, 500).Content(content => content.Text("Another needle")));
            });
            if (rotate) document = document.Pages.Rotate(90);
            document.Save(source);
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show(); await window.TabHost.OpenDocumentAsync(source);
                var model = window.ViewModel;
                var view = window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!;
                window.KeyPress(Key.F, OperatingSystem.IsMacOS() ? RawInputModifiers.Meta : RawInputModifiers.Control, PhysicalKey.None, null);
                window.UpdateLayout(); await Task.Delay(50); window.UpdateLayout();
                var box = view.FindControl<TextBox>("SearchBox")!;
                Assert.True(box.IsFocused);
                model.SearchQuery = "needle";
                await model.SearchCommand.ExecuteAsync(null);
                Assert.Equal(3, model.SearchResults.Count);
                Assert.Equal(new[] { 1, 1, 2 }, model.SearchResults.Select(hit => hit.PageNumber));
                Assert.Same(model.SearchResults[0], model.SelectedSearchResult);
                Assert.Equal(2, model.Pages[0].SearchHighlights.Count);
                Assert.NotEqual(model.SearchResults[0].Bounds, model.SearchResults[1].Bounds);
                window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Same(model.SearchResults[1], model.SelectedSearchResult);
                window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Equal(2, model.SelectedPage!.PageNumber);
                window.KeyPress(Key.Enter, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Same(model.SearchResults[0], model.SelectedSearchResult);
                window.KeyPress(Key.Enter, RawInputModifiers.Shift, PhysicalKey.None, null);
                Assert.Same(model.SearchResults[2], model.SelectedSearchResult);
                Assert.True(box.IsFocused);
                model.SelectedSearchResult = model.SearchResults[1];
                foreach (var page in model.Pages) { page.AttachToViewport(); await page.EnsureRenderedAsync(); Assert.Null(page.RenderError); }
                window.UpdateLayout(); await Task.Delay(100); window.UpdateLayout();
                Assert.Contains(window.GetVisualDescendants().OfType<OfficeIMO.Studio.Features.Reader.PdfPageCanvas>(), canvas => canvas.IsEffectivelyVisible && canvas.ActiveSearchHighlight.HasValue);
                Capture(window, $"search-{width}-{dark}-{rotate}");
                if (rotate) AssertRenderedInk(model.Pages[0], model.SearchResults[1].Bounds);
                model.SearchQuery = "missing";
                Assert.Empty(model.SearchResults); Assert.Null(model.SelectedSearchResult);
                Assert.Null(model.OperationStatus);
                Assert.All(model.Pages, page => { Assert.Empty(page.SearchHighlights); Assert.Null(page.ActiveSearchHighlight); });
                await model.SearchCommand.ExecuteAsync(null); Assert.Empty(model.SearchResults);
                window.UpdateLayout(); Capture(window, $"search-empty-{width}-{dark}");
                model.SearchQuery = "needle"; await model.SearchCommand.ExecuteAsync(null);
                model.SetOrganizerSelection(model.OrganizerPages);
                await model.RotateRightCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage); Assert.Empty(model.SearchResults);
                Assert.All(model.Pages, page => Assert.Empty(page.SearchHighlights));
                await model.SearchCommand.ExecuteAsync(null); Assert.Equal(3, model.SearchResults.Count);
                box.Focus(); window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Empty(model.SearchResults); Assert.Empty(model.SearchQuery); Assert.False(box.IsFocused);
            } finally { window.Close(); window.TabHost.Dispose(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task QueryChangesDiscardAnInFlightSearch(bool clear) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "Search snapshot.pdf");
            PdfDocument.Create(compose => compose.Page(page => page.Content(content => content.Text("Needle needle")))).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await model.OpenDocumentAsync(source);
            model.SearchQuery = "needle";
            Task searching = model.SearchCommand.ExecuteAsync(null);
            Assert.True(model.IsWorkspaceBusy);
            if (clear) model.ClearSearchCommand.Execute(null);
            else model.SearchQuery = "changed";
            await searching;
            Assert.Null(model.ErrorMessage); Assert.Empty(model.SearchResults); Assert.Null(model.SelectedSearchResult);
            Assert.All(model.Pages, page => Assert.Empty(page.SearchHighlights));
            model.SearchQuery = "needle";
            await model.SearchCommand.ExecuteAsync(null);
            Assert.Equal(2, model.SearchResults.Count);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ImageOnlyPageExplainsTheOcrRequirementWithoutInventingMatches() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "Scanned page.pdf");
            byte[] image = PdfDocument.Create(compose => compose.Page(page => page.Content(text => text.Text("Scanned needle"))))
                .Render.Pages(PdfPageSelection.From(1), new PdfPageRenderOptions { Format = PdfPageRenderFormat.Png, Dpi = 72 }).Single().Bytes!;
            PdfDocument.Create(compose => compose.Page(page => page.Content(content => content.Image(image, 400, 500)))).Save(source);
            var window = new MainWindow(services) { Width = 960, Height = 620 };
            try {
                window.Show(); await window.TabHost.OpenDocumentAsync(source);
                var model = window.ViewModel;
                window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!.FocusSearch();
                model.SearchQuery = "needle"; await model.SearchCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage); Assert.Empty(model.SearchResults);
                Assert.Contains("OCR", model.SearchPosition);
                model.Pages[0].AttachToViewport(); await model.Pages[0].EnsureRenderedAsync();
                window.UpdateLayout(); await Task.Delay(80); window.UpdateLayout();
                Capture(window, "search-scanned");
            } finally { window.Close(); window.TabHost.Dispose(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(ReaderLayoutMode.SinglePage)]
    [InlineData(ReaderLayoutMode.TwoPage)]
    [InlineData(ReaderLayoutMode.Grid)]
    public async Task SearchRevealsTheSelectedOccurrenceInEachReaderLayout(ReaderLayoutMode layout) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "Search layouts.pdf");
            PdfDocument.Create(compose => {
                for (int page = 0; page < 12; page++) compose.Page(item => item.Size(400, 500).Content(content => content.Text("Needle on this page")));
            }).Save(source);
            var window = new MainWindow(services) { Width = 1280, Height = 800 };
            try {
                window.Show(); await window.TabHost.OpenDocumentAsync(source);
                var model = window.ViewModel;
                model.SelectedReaderLayoutChoice = model.ReaderLayoutChoices.Single(choice => choice.Mode == layout);
                window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!.FocusSearch();
                model.SearchQuery = "needle"; await model.SearchCommand.ExecuteAsync(null);
                model.SelectedSearchResult = model.SearchResults[10];
                window.UpdateLayout(); await Task.Delay(100); window.UpdateLayout();
                var page = model.Pages[10]; page.AttachToViewport(); await page.EnsureRenderedAsync();
                window.UpdateLayout(); await Task.Delay(80); window.UpdateLayout();
                Assert.Equal(11, model.SelectedPage!.PageNumber);
                var canvas = Assert.Single(window.GetVisualDescendants().OfType<OfficeIMO.Studio.Features.Reader.PdfPageCanvas>()
                    .Where(canvas => canvas.IsEffectivelyVisible && canvas.DataContext == page));
                Assert.Equal(model.SelectedSearchResult.Bounds, canvas.ActiveSearchHighlight);
                var location = canvas.TranslatePoint(new Point(0, 0), window)!.Value;
                double highlightedY = location.Y + model.SelectedSearchResult.Bounds.Center.Y * canvas.Bounds.Height / page.Scene!.Drawing.Height;
                Assert.InRange(highlightedY, 200, 750);
                Capture(window, $"search-layout-{layout}");
                var box = window.FindControl<DocumentWorkspaceView>("DocumentWorkspace")!.FindControl<TextBox>("SearchBox")!;
                box.Focus(); window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.False(box.IsFocused);
            } finally { window.Close(); window.TabHost.Dispose(); }
            return true;
        }, CancellationToken.None);
    }

    private static void AssertRenderedInk(PdfPageViewModel page, Rect region) {
        Assert.NotNull(page.PageImage);
        using var output = new MemoryStream(); page.PageImage.Save(output, Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
        Assert.True(OfficeRasterImageDecoder.TryDecode(output.ToArray(), out var image));
        double xScale = image!.Width / page.Scene!.Drawing.Width, yScale = image.Height / page.Scene.Drawing.Height;
        int ink = 0;
        for (int y = Math.Max(0, (int)(region.Top * yScale)); y < Math.Min(image.Height, (int)Math.Ceiling(region.Bottom * yScale)); y++)
            for (int x = Math.Max(0, (int)(region.Left * xScale)); x < Math.Min(image.Width, (int)Math.Ceiling(region.Right * xScale)); x++) {
                var pixel = image.GetPixel(x, y);
                if (pixel.A > 0 && pixel.R < 100 && pixel.G < 100 && pixel.B < 100) ink++;
            }
        Assert.True(ink > 10, "The highlighted occurrence must contain rendered glyphs.");
    }

    private static void Capture(Window window, string name) {
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_SEARCH_CAPTURE_DIR");
        if (string.IsNullOrEmpty(root)) return;
        Directory.CreateDirectory(root);
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        frame.Save(Path.Combine(root, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
