using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioOcrReviewTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 800, true)]
    public async Task ReviewKeepsPageChoicesUntilExplicitCommitAndCancellationPreservesOutput(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "OCR review source.pdf");
            string output = Path.Combine(services.Paths.Root, "OCR reviewed output.pdf");
            PdfDocument.Create(compose => {
                for (int page = 0; page < 2; page++)
                    compose.Page(item => item.Size(600, 800).Canvas(canvas =>
                        canvas.Text([new PdfTextRun("Native text")], 72, 66, 400, 25, fontSize: 12)));
            }).Save(source);
            byte[] original = File.ReadAllBytes(source);
            var native = PdfPageInteractionMap.Create(original, 1).TextRegions;
            double left = native.Min(region => region.Quad.Left), top = native.Min(region => region.Quad.Top);
            double right = native.Max(region => region.Quad.Right), bottom = native.Max(region => region.Quad.Bottom);
            var engine = new DelegateOcrEngine("fixture", (_, _) => Task.FromResult(new OcrResult {
                Provider = "Review fixture", Language = "eng", Spans = [
                    Word("Native", left, top, right - left, bottom - top, 0.99),
                    Word("Selected searchable text", 72, 180, 140, 14, 0.98),
                    Word("Excluded after review", 72, 220, 140, 14, 0.91),
                    Word("Uncertain word", 72, 260, 90, 14, 0.15)
                ]
            }));
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(source), services: services,
                ocrService: new StudioOcrPublicationTests.EngineService(engine));
            var model = shell.OcrWorkbench;
            var window = new Window { Width = width, Height = height, Content = new SearchablePdfOcrView { DataContext = shell } };
            try {
                window.Show(); model.InputPath = source; model.OutputPath = output;
                window.UpdateLayout();
                using (var setup = window.CaptureRenderedFrame()) {
                    string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                    if (!string.IsNullOrWhiteSpace(evidence)) {
                        Directory.CreateDirectory(evidence);
                        setup!.Save(Path.Combine(evidence, $"ocr-setup-{width}-{(dark ? "dark" : "light")}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                    }
                }
                var running = model.RunCommand.ExecuteAsync(null);
                var review = await WaitForReview(model, running);
                Assert.False(File.Exists(output));
                Assert.False(model.RunCommand.CanExecute(null));
                Assert.Equal(2, review.Pages.Count);
                Assert.Equal(2, review.Words.Count(word => !word.IsEligible));
                review.Words.Single(word => word.Text == "Excluded after review").IsIncluded = false;
                review.SelectedPage = review.Pages[1];
                await review.PreviewTask;
                review.ExcludePageCommand.Execute(null);
                review.SelectedPage = review.Pages[0];
                await review.PreviewTask;
                Assert.False(review.Words.Single(word => word.Text == "Excluded after review").IsIncluded);
                review.SelectedWord = review.Words.Single(word => word.Text == "Selected searchable text");
                window.UpdateLayout();
                using (var frame = window.CaptureRenderedFrame()) {
                    Assert.NotNull(frame);
                    string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                    if (!string.IsNullOrWhiteSpace(evidence)) {
                        Directory.CreateDirectory(evidence);
                        frame.Save(Path.Combine(evidence, $"ocr-review-{width}-{(dark ? "dark" : "light")}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                    }
                }
                review.IsZoomed = true;
                review.SelectedWord = review.Words.Single(word => word.Text == "Uncertain word");
                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), Avalonia.Threading.DispatcherPriority.Background);
                var zoom = window.GetVisualDescendants().OfType<Avalonia.Controls.Shapes.Rectangle>().Single(item => item.Name == "ZoomSelection");
                Assert.Equal(90, zoom.Bounds.Width);
                using (var frame = window.CaptureRenderedFrame()) {
                    string? evidence = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                    if (!string.IsNullOrWhiteSpace(evidence)) frame!.Save(Path.Combine(evidence, $"ocr-review-{width}-{(dark ? "dark" : "light")}-zoom.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
                review.CommitCommand.Execute(null);
                await running;
                Assert.False(model.HasReview);
                Assert.True(model.HasOutput, model.ErrorMessage);
                var saved = PdfReadDocument.Open(File.ReadAllBytes(output));
                Assert.Contains("Selected searchable text", saved.Pages[0].ExtractText());
                Assert.DoesNotContain("Excluded after review", saved.ExtractText());
                Assert.DoesNotContain("Uncertain word", saved.ExtractText());
                Assert.DoesNotContain("Selected searchable text", saved.Pages[1].ExtractText());
                byte[] published = File.ReadAllBytes(output);
                string? artifactFolder = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrWhiteSpace(artifactFolder)) File.WriteAllBytes(Path.Combine(artifactFolder, $"ocr-reviewed-{width}.pdf"), published);
                model.ReplaceExistingOutput = true;
                running = model.RunCommand.ExecuteAsync(null);
                review = await WaitForReview(model, running);
                review.CancelCommand.Execute(null);
                await running;
                Assert.False(model.HasOutput);
                Assert.False(model.HasReview);
                Assert.Equal("OCR cancelled", model.Status);
                Assert.Equal(published, File.ReadAllBytes(output));
                Assert.Equal(original, File.ReadAllBytes(source));
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static async Task<OcrReviewViewModel> WaitForReview(SearchablePdfOcrViewModel model, Task running) {
        var deadline = DateTime.UtcNow.AddSeconds(30);
        while (model.Review is null && !running.IsCompleted && DateTime.UtcNow < deadline) await Task.Delay(10);
        Assert.NotNull(model.Review);
        await model.Review.PreviewTask;
        Assert.True(model.Review.CommitCommand.CanExecute(null), model.Review.PreviewError);
        return model.Review;
    }

    private static OcrTextSpan Word(string text, double x, double y, double width, double height, double confidence) => new() {
        Text = text, Level = OcrTextSpanLevel.Word, Confidence = confidence,
        CoordinateUnit = OcrCoordinateUnit.Points, Region = new OcrRegion { X = x, Y = y, Width = width, Height = height }
    };
}
