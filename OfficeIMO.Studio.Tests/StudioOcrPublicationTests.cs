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
                await model.RunCommand.ExecuteAsync(null);
                Assert.True(model.HasOutput, model.ErrorMessage);
                Assert.Equal("Recognized", PdfReadDocument.Open(File.ReadAllBytes(output)).Pages.Single().ExtractText().Trim());
                Assert.True(services.Jobs.Entries[0].HasOutput);
                Capture(window, $"ocr-publication-{width}-{(dark ? "dark" : "light")}-completed.png");
                byte[] saved = File.ReadAllBytes(output);
                model.ReplaceExistingOutput = true;
                denyAfterRecognition = true;
                await model.RunCommand.ExecuteAsync(null);
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

    private sealed class EngineService(IOcrEngine engine) : ISearchablePdfOcrService {
        public async Task<SearchablePdfOcrOutcome> MakeSearchableAsync(string inputPath, string outputPath,
            SearchablePdfOcrOptions options, CancellationToken cancellationToken) {
            var result = await new OfficeWorkflowRunner().MakePdfSearchableAsync(new() {
                InputPath = inputPath, OutputPath = outputPath, Ocr = options.Pdf,
                PublicationGuard = options.PublicationGuard,
                ConflictPolicy = options.OutputConflictPolicy == OfficeConversionFileConflictPolicy.Replace
                    ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail
            }, engine, cancellationToken);
            return new(result.AddedWordCount, result.ModifiedPages, result.Provider, result);
        }
    }

    private static void Capture(Window window, string name) {
        window.UpdateLayout();
        window.GetVisualDescendants().OfType<ScrollViewer>().First().ScrollToEnd();
        window.UpdateLayout();
        Assert.Contains(window.GetVisualDescendants().OfType<TextBlock>(), text =>
            text.Text is "Searchable PDF created" or "OCR could not finish");
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
