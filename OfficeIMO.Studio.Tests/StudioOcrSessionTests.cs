using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Ocr;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioOcrSessionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task RetryAndNewWorkPreserveCompletedItemsOriginalSources(bool newWork) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string first = Path.Combine(services.Paths.Root, "scan.pdf");
            string second = Path.Combine(services.Paths.Root, "scan-searchable.pdf");
            byte[] original = PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).ToBytes();
            File.WriteAllBytes(first, original); File.WriteAllBytes(second, original);
            IReadOnlyList<string> selection = newWork ? [second] : [first, second];
            var engine = new DelegateOcrEngine("source-fixture", (_, _) => Task.FromResult(new OcrResult {
                Text = "Recognized", Spans = [new OcrTextSpan { Text = "Recognized", Confidence = 1,
                    Level = OcrTextSpanLevel.Word, CoordinateUnit = OcrCoordinateUnit.Points,
                    Region = new OcrRegion { X = 20, Y = 60, Width = 90, Height = 12 } }]
            }));
            using var model = new OcrSessionViewModel(_ => Task.FromResult(selection),
                _ => Task.FromResult<string?>(services.Paths.Root), services.Localizer,
                createEngine: (_, _, _) => Task.FromResult<IOcrEngine>(engine));
            await model.AddFilesCommand.ExecuteAsync(null);
            model.OutputFolder = services.Paths.Root; model.ReplaceExisting = true;
            async Task Approve(Task running) {
                OcrReviewViewModel? previous = null;
                var deadline = DateTime.UtcNow.AddSeconds(30);
                while (!running.IsCompleted && DateTime.UtcNow < deadline) {
                    var review = model.PdfReview;
                    if (review is not null && review != previous) {
                        previous = review;
                        await review.PreviewTask;
                        review.CommitCommand.Execute(null);
                    }
                    await Task.WhenAny(running, Task.Delay(10));
                }
                Assert.True(running.IsCompleted, model.Status);
                await running;
            }
            await Approve(model.RunCommand.ExecuteAsync(null));
            Assert.Equal(OfficeWorkflowStatus.Completed, model.Items.Single(item => item.InputPath == second).Status);
            if (newWork) { selection = [first]; await model.AddFilesCommand.ExecuteAsync(null); }
            else Assert.Equal(OfficeWorkflowStatus.Failed, model.Items[0].Status);
            await Approve(newWork ? model.RunCommand.ExecuteAsync(null) : model.RetryCommand.ExecuteAsync(null));
            Assert.Equal(OfficeWorkflowStatus.Failed, model.Items.Single(item => item.InputPath == first).Status);
            Assert.Equal(original, File.ReadAllBytes(first));
            Assert.Equal(original, File.ReadAllBytes(second));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task SameNamedSourcesGetDistinctStableOutputs() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string[] sources = [Path.Combine(services.Paths.Root, "first", "scan.png"),
                Path.Combine(services.Paths.Root, "second", "scan.png")];
            foreach (string source in sources) {
                Directory.CreateDirectory(Path.GetDirectoryName(source)!);
                File.WriteAllBytes(source, OfficeRasterImageEncoder.Encode(new OfficeRasterImage(20, 30, OfficeColor.White), OfficeImageExportFormat.Png));
            }
            var engine = new DelegateOcrEngine("naming-fixture", (_, _) => Task.FromResult(new OcrResult { Text = "Recognized" }));
            using var model = new OcrSessionViewModel(_ => Task.FromResult<IReadOnlyList<string>>(sources),
                _ => Task.FromResult<string?>(services.Paths.Root), services.Localizer,
                createEngine: (_, _, _) => Task.FromResult<IOcrEngine>(engine));
            await model.AddFilesCommand.ExecuteAsync(null);
            Assert.Equal(new[] { "scan.png.txt", "scan.png (2).txt" }, model.Items.Select(item => item.OutputName));
            model.OutputFolder = services.Paths.Root;
            var running = model.RunCommand.ExecuteAsync(null);
            ImageOcrReviewViewModel? previous = null;
            for (int index = 0; index < 2; index++) {
                await WaitFor(() => model.ImageReview is not null && model.ImageReview != previous, running, model);
                previous = model.ImageReview!;
                await previous.PreviewTask;
                previous.Text = "Reviewed " + index;
                previous.CommitCommand.Execute(null);
            }
            await running;
            Assert.All(model.Items, item => Assert.Equal(OfficeWorkflowStatus.Completed, item.Status));
            Assert.Equal("Reviewed 0", File.ReadAllText(model.Items[0].OutputPath!));
            Assert.Equal("Reviewed 1", File.ReadAllText(model.Items[1].OutputPath!));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task ImageReviewCanChangeTiffPagesWithoutLosingTextCorrections() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "pages.tiff");
            File.WriteAllBytes(source, OfficeTiffCodec.EncodePages([
                new OfficeRasterImage(20, 30, OfficeColor.White), new OfficeRasterImage(30, 20, OfficeColor.White)
            ]));
            var engine = new DelegateOcrEngine("tiff-fixture", (_, _) => Task.FromResult(new OcrResult { Text = "Two pages" }));
            using var model = new OcrSessionViewModel(_ => Task.FromResult<IReadOnlyList<string>>([source]),
                _ => Task.FromResult<string?>(services.Paths.Root), services.Localizer,
                createEngine: (_, _, _) => Task.FromResult<IOcrEngine>(engine));
            await model.AddFilesCommand.ExecuteAsync(null);
            model.OutputFolder = services.Paths.Root;
            var running = model.RunCommand.ExecuteAsync(null);
            await WaitFor(() => model.ImageReview is not null, running, model);
            var review = model.ImageReview!;
            await review.PreviewTask;
            Assert.Equal(2, review.Frames.Count);
            Assert.Equal(new PixelSize(20, 30), review.Preview!.PixelSize);
            review.Text = "Corrected text from both pages";
            review.SelectedFrame = review.Frames[1];
            await review.PreviewTask;
            Assert.Equal(new PixelSize(30, 20), review.Preview!.PixelSize);
            Assert.Equal("Corrected text from both pages", review.Text);
            review.CommitCommand.Execute(null);
            await running;
            Assert.Equal(OfficeWorkflowStatus.Completed, model.Items[0].Status);
            Assert.Equal(review.Text, File.ReadAllText(model.Items[0].OutputPath!));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task SessionIsReachableInTheCompactDesktopShell() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var services = ((App)Application.Current!).Services;
            var window = new MainWindow(services) { Width = 960, Height = 620 };
            try {
                window.Show();
                window.ViewModel.ShowOcrCommand.Execute(null);
                window.UpdateLayout();
                var selector = window.GetVisualDescendants().OfType<TabControl>().Single();
                selector.SelectedIndex = 1;
                window.UpdateLayout();
                var view = window.GetVisualDescendants().OfType<OcrSessionView>().Single();
                Assert.Same(window.ViewModel.OcrSession, view.DataContext);
                Assert.True(view.IsEffectivelyVisible);
                var languages = view.GetVisualDescendants().OfType<Expander>().Single();
                languages.IsExpanded = true;
                Capture(window, "ocr-session-shell-960-expanded");
                var run = view.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, window.ViewModel.OcrSession.RunCommand));
                Assert.True(run.TranslatePoint(new Point(0, run.Bounds.Height), window)?.Y <= window.ClientSize.Height);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 800, true)]
    public async Task MixedSessionReviewsTextRetainsCompletedOutputAndRetriesOnlyCancelledFile(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string image = Path.Combine(services.Paths.Root, "Invoice.png");
            string pdf = Path.Combine(services.Paths.Root, "Scanned page.pdf");
            byte[] visualSource = PdfDocument.Create(document => document.Page(page => page.Size(360, 480).Canvas(canvas =>
                canvas.Text([new PdfTextRun("Invoice 4827")], 30, 60, 280, 30, fontSize: 24)
                    .Text([new PdfTextRun("Reviewed searchable statement")], 30, 120, 300, 25, fontSize: 16)))).ToBytes();
            byte[] png = PdfDocument.Load(visualSource).Render.Pages(PdfPageSelection.From(1),
                new PdfPageRenderOptions { Format = PdfPageRenderFormat.Png, Dpi = 72 }).Single().Bytes!;
            File.WriteAllBytes(image, png);
            byte[] originalPdf = PdfDocument.Create(document => document.Page(page => page.Size(360, 480))).ToBytes();
            File.WriteAllBytes(pdf, originalPdf);
            int calls = 0;
            var engine = new DelegateOcrEngine("session-fixture", (_, _) => {
                calls++;
                return Task.FromResult(new OcrResult { Text = "Invoice 4827\nReviewed searchable statement", Provider = "Fixture", Language = "eng", Confidence = 0.96,
                    Spans = [new OcrTextSpan { Text = "Reviewed text", Level = OcrTextSpanLevel.Word, Confidence = 0.96,
                        CoordinateUnit = OcrCoordinateUnit.Points, Region = new OcrRegion { X = 30, Y = 60, Width = 100, Height = 20 } }] });
            });
            using var model = new OcrSessionViewModel(_ => Task.FromResult<IReadOnlyList<string>>([image, pdf]),
                _ => Task.FromResult<string?>(services.Paths.Root), services.Localizer, jobs: services.Jobs,
                createEngine: (_, _, _) => Task.FromResult<IOcrEngine>(engine));
            var window = new Window { Width = width, Height = height, Content = new OcrSessionView { DataContext = model } };
            try {
                window.Show();
                Assert.False(model.HasItems);
                Assert.False(model.CanRun);
                Assert.False(model.RemoveSelectedCommand.CanExecute(null));
                Capture(window, $"ocr-session-empty-{width}-{dark}");
                var add = Assert.Single(window.GetVisualDescendants().OfType<Button>(), button =>
                    ReferenceEquals(button.Command, model.AddFilesCommand) && button.Classes.Contains("primary"));
                Point addPoint = add.TranslatePoint(new Point(add.Bounds.Width / 2, add.Bounds.Height / 2), window)!.Value;
                window.MouseDown(addPoint, Avalonia.Input.MouseButton.Left);
                window.MouseUp(addPoint, Avalonia.Input.MouseButton.Left);
                if (model.AddFilesCommand.ExecutionTask is { } intake) await intake;
                Assert.True(model.HasItems);
                Assert.True(model.RemoveSelectedCommand.CanExecute(null));
                Assert.False(model.CanRun);
                Assert.False(string.IsNullOrWhiteSpace(model.SetupHint));
                await model.ChooseOutputFolderCommand.ExecuteAsync(null);
                Assert.True(model.CanRun);
                Assert.Empty(model.SetupHint);
                Capture(window, $"ocr-session-setup-{width}-{dark}");
                var running = model.RunCommand.ExecuteAsync(null);
                await WaitFor(() => model.ImageReview is not null, running, model);
                var review = model.ImageReview!;
                await review.PreviewTask;
                Assert.True(review.CommitCommand.CanExecute(null), review.PreviewError);
                Assert.False(model.CanRun);
                review.Text = "Invoice 4827\nCorrected zażółć 🚀";
                Capture(window, $"ocr-session-image-review-{width}-{dark}");
                review.IsZoomed = true;
                Capture(window, $"ocr-session-image-review-{width}-{dark}-pixels");
                review.IsZoomed = false;
                var save = window.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, review.CommitCommand));
                Assert.True(save.IsEffectivelyVisible);
                Assert.True(save.TranslatePoint(new Point(0, save.Bounds.Height), window)?.Y <= window.ClientSize.Height);
                review.CommitCommand.Execute(null);
                await WaitFor(() => model.PdfReview is not null, running, model);
                await model.PdfReview!.PreviewTask;
                string textOutput = Path.Combine(services.Paths.Root, "Invoice.png.txt");
                Assert.Equal("Invoice 4827\nCorrected zażółć 🚀", File.ReadAllText(textOutput));
                await WaitFor(() => model.Items[0].Status == OfficeWorkflowStatus.Completed, running, model);
                Assert.Contains(services.Jobs.Entries, entry => entry.HasOutput && entry.OutputPath == textOutput);
                model.PdfReview.CancelCommand.Execute(null);
                await running;
                Assert.Equal(OfficeWorkflowStatus.Cancelled, model.Items[1].Status);
                Assert.True(model.CanRetry);
                byte[] firstOutput = File.ReadAllBytes(textOutput);
                var retry = model.RetryCommand.ExecuteAsync(null);
                await WaitFor(() => model.PdfReview is not null, retry, model);
                await model.PdfReview!.PreviewTask;
                model.PdfReview.CommitCommand.Execute(null);
                await retry;
                Assert.Equal(3, calls);
                Assert.All(model.Items, item => Assert.Equal(OfficeWorkflowStatus.Completed, item.Status));
                Assert.Equal(firstOutput, File.ReadAllBytes(textOutput));
                Assert.Equal(png, File.ReadAllBytes(image));
                Assert.Equal(originalPdf, File.ReadAllBytes(pdf));
                Assert.Contains("Reviewed text", PdfReadDocument.Open(File.ReadAllBytes(model.Items[1].OutputPath!)).ExtractText());
                Assert.Equal(0, services.Jobs.ActiveCount);
                Capture(window, $"ocr-session-completed-{width}-{dark}");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
    private static async Task WaitFor(Func<bool> condition, Task running, OcrSessionViewModel model) {
        var deadline = DateTime.UtcNow.AddSeconds(30);
        while (!condition() && !running.IsCompleted && DateTime.UtcNow < deadline) await Task.Delay(10);
        Assert.True(condition(), model.Status);
    }
    private static void Capture(Window window, string name) {
        window.UpdateLayout();
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? folder = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(folder)) return;
        Directory.CreateDirectory(folder);
        frame.Save(Path.Combine(folder, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
