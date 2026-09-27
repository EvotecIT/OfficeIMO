using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using PdfDocument = OfficeIMO.Pdf.PdfDocument;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfPageViewTests {
    private static readonly byte[] TinyPng = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    [Fact]
    public async Task PendingFormFocusIsAppliedWhenVirtualizedPageViewAttaches() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            string root = Path.Combine(Path.GetTempPath(), "officeimo-form-late-page-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            try {
                string source = Path.Combine(root, "form.pdf");
                PdfDocument.Create(compose => compose.Page(page => page.Content(content =>
                    content.Item(item => item.TextField("Name", value: "Initial"))))).Save(source);
                using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
                await model.OpenDocumentAsync(source);
                model.ShowFormsModeCommand.Execute(null);
                PdfPageViewModel page = Assert.Single(model.Pages);
                page.AttachToViewport();
                await WaitUntilAsync(() => page.Scene is not null);
                page.ShowInlineFormField(Assert.Single(model.FormFields), focus: true);
                Assert.True(page.FocusInlineFormEditorRequested);

                var view = new PdfPageView { DataContext = page };
                var window = new Window { Width = 800, Height = 900, Content = view };
                try {
                    window.Show();
                    window.UpdateLayout();
                    TextBox editor = view.FindControl<TextBox>("InlineFormText")!;
                    await WaitUntilAsync(() => editor.IsFocused);
                    Assert.False(page.FocusInlineFormEditorRequested);
                    Assert.Equal("Initial", editor.Text);
                    string? folder = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                    if (!string.IsNullOrWhiteSpace(folder)) {
                        Directory.CreateDirectory(folder);
                        using var frame = window.CaptureRenderedFrame();
                        Assert.NotNull(frame);
                        frame.Save(Path.Combine(folder, "late-form-page-focus.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                    }
                } finally { window.Close(); }
            } finally { Directory.Delete(root, recursive: true); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task AttachingAfterDataContextStartsPageRendering() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            int renderCalls = 0;
            using var coordinator = new PageRenderCoordinator((page, scale, _) => {
                Interlocked.Increment(ref renderCalls);
                return Task.FromResult(
                    new PdfRenderedPage(page, scale, TinyPng, 1, 1, TimeSpan.Zero, Array.Empty<string>()));
            });
            using var sceneCoordinator = new PageSceneCoordinator((page, _) =>
                Task.FromResult(TestPdfPageScenes.Create(page, requiresRasterFallback: true)));
            using var viewModel = new PdfPageViewModel(1, 612, 792, 0, 1D, sceneCoordinator, coordinator);
            var view = new PdfPageView { DataContext = viewModel };
            var window = new Window { Content = view };

            try {
                window.Show();
                window.Measure(new Size(800, 600));
                window.Arrange(new Rect(0, 0, 800, 600));
                await WaitUntilAsync(() => viewModel.PageImage is not null && !viewModel.IsRendering);

                Assert.Equal(1, renderCalls);
                Assert.NotNull(viewModel.PageImage);
            } finally {
                window.Close();
            }

            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task PendingRedactionAreasBindsToThePageCanvas() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            using var renderCoordinator = new PageRenderCoordinator((page, scale, _) =>
                Task.FromResult(new PdfRenderedPage(page, scale, TinyPng, 1, 1, TimeSpan.Zero, Array.Empty<string>())));
            using var sceneCoordinator = new PageSceneCoordinator((page, _) =>
                Task.FromResult(TestPdfPageScenes.Create(page)));
            using var viewModel = new PdfPageViewModel(1, 612, 792, 0, 1D, sceneCoordinator, renderCoordinator) {
                PendingRedactionAreas = new[] { new Rect(42D, 64D, 180D, 36D) }
            };
            var view = new PdfPageView { DataContext = viewModel };
            var window = new Window { Content = view };

            try {
                window.Show();
                window.Measure(new Size(800, 900));
                window.Arrange(new Rect(0, 0, 800, 900));

                PdfPageCanvas canvas = Assert.IsType<PdfPageCanvas>(view.FindControl<PdfPageCanvas>("PageCanvas"));
                Assert.Equal(viewModel.PendingRedactionAreas, canvas.PendingRedactionAreas);
            } finally {
                window.Close();
            }

            return Task.CompletedTask;
        }, CancellationToken.None);
    }

    private static async Task WaitUntilAsync(Func<bool> condition) {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        while (!condition()) await Task.Delay(10, timeout.Token);
    }
}
