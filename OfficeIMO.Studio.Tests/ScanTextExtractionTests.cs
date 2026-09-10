using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class ScanTextExtractionTests {
    [Theory]
    [InlineData(960, 640)]
    [InlineData(1440, 900)]
    public async Task ReviewedTextCopiesWithoutOutputAndCancellationLeavesSourceUnchanged(int width, int height) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "quick-text.pdf");
            PdfDocument.Create(doc => {
                doc.Page(page => page.Size(300, 300));
                doc.Page(page => page.Size(300, 300));
            }).Save(source);
            byte[] original = File.ReadAllBytes(source);
            var recognition = new RecognitionService();
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                scanTextRecognition: recognition, canPublishPath: _ => throw new InvalidOperationException("Text extraction must not publish."));
            var model = shell.OcrWorkbench;
            var view = new SearchablePdfOcrView { DataContext = shell };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show(); model.InputPath = source; model.OutputPath = string.Empty;
                model.Scan.PageNumber = 2;
                Assert.True(model.ExtractTextCommand.CanExecute(null));
                Assert.False(model.RunCommand.CanExecute(null));
                var running = model.ExtractTextCommand.ExecuteAsync(null);
                var review = await WaitForReview(model, running);
                Assert.Equal(2, Assert.Single(review.Pages).Number);
                Assert.Equal("Use selected text", review.CommitLabel);
                Assert.False(model.ExtractTextCommand.CanExecute(null));
                review.Words.Single(word => word.Text == "Discard").IsIncluded = false;
                review.CommitCommand.Execute(null); await running;
                Assert.Equal("Copy this", model.ExtractedText);
                Assert.False(model.HasOutput);
                Assert.Null(model.ErrorMessage);
                Assert.Equal(original, File.ReadAllBytes(source));
                window.UpdateLayout();
                var textBox = view.FindControl<TextBox>("ExtractedTextBox")!;
                view.FindControl<Button>("CopyExtractedTextButton")!.BringIntoView();
                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), Avalonia.Threading.DispatcherPriority.Background);
                Assert.True(textBox.IsReadOnly);
                Assert.Equal("Copy this", textBox.Text);
                view.FindControl<Button>("CopyExtractedTextButton")!.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
                for (int retry = 0; retry < 100 && model.Status != "Reviewed text copied."; retry++) await Task.Delay(10);
                Assert.Equal("Reviewed text copied.", model.Status);
                Assert.Equal("Copy this", await window.Clipboard!.TryGetTextAsync());
                using (var frame = window.CaptureRenderedFrame()) {
                    Assert.NotNull(frame);
                    string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                    if (!string.IsNullOrEmpty(output)) {
                        Directory.CreateDirectory(output);
                        frame.Save(Path.Combine(output, $"quick-text-{width}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                    }
                }
                running = model.ExtractTextCommand.ExecuteAsync(null);
                review = await WaitForReview(model, running);
                review.CancelCommand.Execute(null); await running;
                Assert.False(model.HasExtractedText);
                Assert.False(model.HasReview);
                Assert.False(model.IsBusy);
                Assert.Equal(original, File.ReadAllBytes(source));
                Assert.Equal(2, recognition.Calls);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static async Task<OcrReviewViewModel> WaitForReview(SearchablePdfOcrViewModel model, Task running) {
        var deadline = DateTime.UtcNow.AddSeconds(30);
        while (model.Review == null && !running.IsCompleted && DateTime.UtcNow < deadline) await Task.Delay(10);
        Assert.True(model.Review != null, model.ErrorMessage);
        await model.Review!.PreviewTask;
        Assert.True(model.Review.CommitCommand.CanExecute(null), model.Review.PreviewError);
        return model.Review;
    }

    private sealed class RecognitionService : IScanTextRecognitionService {
        public int Calls { get; private set; }
        public Task<PdfSearchableOcrReview> PrepareAsync(byte[] source, SearchablePdfOcrOptions options, CancellationToken token) {
            Calls++;
            var engine = new DelegateOcrEngine("quick-text", (_, _) => Task.FromResult(new OcrResult {
                Provider = "fixture", Language = "eng", Spans = [Word("Copy this", 20), Word("Discard", 60)]
            }));
            return PdfDocument.Load(source).PrepareSearchableOcrAsync(engine, options.Pdf, token);
        }
        private static OcrTextSpan Word(string text, double y) => new() {
            Text = text, Confidence = .95, Level = OcrTextSpanLevel.Word, CoordinateUnit = OcrCoordinateUnit.Points,
            Region = new OcrRegion { X = 20, Y = y, Width = 80, Height = 12 }
        };
    }
}
