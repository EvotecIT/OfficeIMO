using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Ocr;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioOcrPublicationTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 820, true)]
    public async Task OcrShowsVerifiedOutputAndLatePublicationFailure(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            string source = Path.Combine(services.Paths.Root, "Scanned document.pdf");
            string output = Path.Combine(services.Paths.Root, "Searchable document.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).Save(source);
            bool allowed = true;
            bool denyAfterRecognition = false;
            var engine = new DelegateOcrEngine("fixture", (_, _) => {
                if (denyAfterRecognition) allowed = false;
                return Task.FromResult(new OcrResult { Provider = "fixture", Spans = [new OcrTextSpan {
                    Text = "Recognized", Level = OcrTextSpanLevel.Word, Confidence = 1,
                    CoordinateUnit = OcrCoordinateUnit.Points,
                    Region = new OcrRegion { X = 20, Y = 30, Width = 80, Height = 12 }
                }] });
            });
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(source),
                ocrService: new EngineService(engine), services: services,
                canPublishPath: _ => allowed,
                publicationGuard: new StudioWorkflowPublicationGuard((_, _) => allowed));
            var model = shell.OcrWorkbench;
            var window = new Window { Width = width, Height = height, Content = new SearchablePdfOcrView { DataContext = shell } };
            try {
                window.Show();
                model.InputPath = source;
                model.OutputPath = output;
                await RunThroughReviewAsync(model);
                Assert.True(model.HasOutput, model.ErrorMessage);
                Assert.Equal("Recognized", PdfReadDocument.Open(File.ReadAllBytes(output)).Pages.Single().ExtractText().Trim());
                Assert.True(services.Jobs.Entries[0].HasOutput);
                Capture(window, $"ocr-publication-{width}-{(dark ? "dark" : "light")}-completed.png");
                byte[] saved = File.ReadAllBytes(output);
                model.ReplaceExistingOutput = true;
                denyAfterRecognition = true;
                await RunThroughReviewAsync(model);
                Assert.False(model.HasOutput);
                Assert.True(model.HasError);
                Assert.Equal("Failed", services.Jobs.Entries[0].Status);
                Assert.False(services.Jobs.Entries[0].HasOutput);
                Assert.Equal(saved, File.ReadAllBytes(output));
                Capture(window, $"ocr-publication-{width}-{(dark ? "dark" : "light")}-denied.png");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 820, true)]
    public async Task ProviderOcrRequiresConsentAndRetainsAnInterruptedOutput(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            byte[] sourceBytes = PdfDocument.Create(document => document.Page(page => page.Size(300, 300))).ToBytes();
            var input = new TestStorageFile("content://documents/ocr-source", sourceBytes, "Selected scanned document.pdf");
            var output = new TestStorageFile("content://documents/ocr-output", sourceBytes, "Selected searchable document.pdf");
            string source = await services.Storage.RegisterAsync(input.Item, default);
            string destination = await services.Storage.RegisterAsync(output.Item, default);
            bool consent = false;
            bool revoke = false;
            int recognized = 0;
            var engine = new DelegateOcrEngine("fixture", (_, _) => {
                recognized++;
                if (revoke) input.DenyRead = true;
                return Task.FromResult(new OcrResult { Provider = "fixture", Spans = [new OcrTextSpan {
                    Text = "Recognized", Level = OcrTextSpanLevel.Word, Confidence = 1,
                    CoordinateUnit = OcrCoordinateUnit.Points,
                    Region = new OcrRegion { X = 20, Y = 30, Width = 80, Height = 12 }
                }] });
            });
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(source),
                pickSavePdf: _ => Task.FromResult<string?>(destination), services: services,
                ocrService: new EngineService(engine), confirmWorkflowProviderWrite: _ => Task.FromResult(consent));
            var model = shell.OcrWorkbench;
            var window = new Window { Width = width, Height = height, Content = new SearchablePdfOcrView { DataContext = shell } };
            try {
                window.Show();
                model.UseDocument(Path.Combine(services.Paths.Root, "Previous local input.pdf"));
                Assert.NotEmpty(model.OutputPath);
                await model.ChooseInputCommand.ExecuteAsync(null);
                Assert.Equal(input.Name, model.InputName);
                Assert.Empty(model.OutputPath);
                Assert.False(model.RunCommand.CanExecute(null));
                await model.ChooseOutputCommand.ExecuteAsync(null);
                Assert.Equal(output.Name, model.OutputName);
                await RunThroughReviewAsync(model);
                Assert.Equal(0, recognized);
                Assert.Equal(0, output.Writes);
                Assert.False(model.HasOutput);
                consent = true;
                await RunThroughReviewAsync(model);
                Assert.True(model.HasOutput, model.ErrorMessage);
                Assert.Equal(destination, model.PublishedPath);
                Assert.Equal("Recognized", PdfReadDocument.Open(output.Bytes).Pages.Single().ExtractText().Trim());
                Assert.Equal(0, input.Writes);
                Assert.Equal(input.Reads, input.ClosedReads);
                Capture(window, $"ocr-provider-{width}-{(dark ? "dark" : "light")}-completed.png");
                byte[] saved = output.Bytes.ToArray();
                revoke = true;
                await RunThroughReviewAsync(model);
                Assert.False(model.HasOutput);
                Assert.Equal(saved, output.Bytes);
                Assert.Equal(1, output.Writes);
                Assert.Equal("Failed", services.Jobs.Entries[0].Status);
                input.DenyRead = false;
                revoke = false;
                output.FailWrite = true;
                await RunThroughReviewAsync(model);
                Assert.False(model.HasOutput);
                Assert.True(model.HasRecovery);
                Assert.Equal("Check output", services.Jobs.Entries[0].Status);
                Assert.NotNull(services.Jobs.Entries[0].Recovery);
                var retained = Assert.Single(services.WorkflowRecovery.GetRecoveries());
                await services.WorkflowRecovery.VerifyAsync(retained);
                Assert.Equal("Recognized", PdfReadDocument.Open(File.ReadAllBytes(retained.FilePath)).Pages.Single().ExtractText().Trim());
                Assert.Equal(2, output.Writes);
                Assert.Equal(sourceBytes, input.Bytes);
                Capture(window, $"ocr-provider-{width}-{(dark ? "dark" : "light")}-recovery.png");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    internal sealed class EngineService(IOcrEngine engine) : ISearchablePdfOcrService {
        public async Task<SearchablePdfOcrOutcome> MakeSearchableAsync(string inputPath, string outputPath,
            SearchablePdfOcrOptions options, CancellationToken cancellationToken) {
            var result = await new OfficeWorkflowRunner().MakePdfSearchableAsync(new() {
                InputPath = inputPath, OutputPath = outputPath, Ocr = options.Pdf,
                InputStream = options.InputStream, OutputStream = options.OutputStream,
                PublicationGuard = options.PublicationGuard, ReviewAsync = options.ReviewAsync,
                ConflictPolicy = options.OutputConflictPolicy == OfficeConversionFileConflictPolicy.Replace
                    ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail
            }, engine, cancellationToken);
            return new(result.AddedWordCount, result.ModifiedPages, result.Provider, result);
        }
    }

    private static async Task RunThroughReviewAsync(SearchablePdfOcrViewModel model) {
        var running = model.RunCommand.ExecuteAsync(null);
        var deadline = DateTime.UtcNow.AddSeconds(30);
        while (!running.IsCompleted && DateTime.UtcNow < deadline) {
            if (model.Review is { } review) {
                await review.PreviewTask;
                Assert.True(review.CommitCommand.CanExecute(null), review.PreviewError);
                review.CommitCommand.Execute(null);
            }
            await Task.WhenAny(running, Task.Delay(10));
        }
        await running.WaitAsync(TimeSpan.FromSeconds(5));
    }

    private static void Capture(Window window, string name) {
        window.UpdateLayout();
        window.GetVisualDescendants().OfType<ScrollViewer>().First().ScrollToEnd();
        window.UpdateLayout();
        Assert.Contains(window.GetVisualDescendants().OfType<TextBlock>(), text =>
            text.Text is "Searchable PDF created" or "OCR could not finish" or "Check output");
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
