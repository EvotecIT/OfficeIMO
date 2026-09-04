using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Studio.Features.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class SearchablePdfOcrViewModelTests {
    [Fact]
    public async Task RunProjectsSelectedLanguagesPagesAndSafetyOptionsToTheOcrFacade() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-studio-ocr-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(directory, "scan.pdf");
        string output = Path.Combine(directory, "scan-searchable.pdf");
        var service = new RecordingOcrService(new SearchablePdfOcrOutcome(42, [1, 3], "fixture-ocr"));
        string? opened = null;
        using var viewModel = new SearchablePdfOcrViewModel(
            _ => Task.FromResult<string?>(input),
            _ => Task.FromResult<string?>(directory),
            (path, _) => {
                opened = path;
                return Task.CompletedTask;
            },
            service);

        await viewModel.ChooseInputCommand.ExecuteAsync(null);
        viewModel.Languages.Single(choice => choice.Label == "Polish").IsSelected = true;
        viewModel.Pages = "1,3";
        viewModel.RenderDpi = 200D;
        viewModel.MinimumConfidencePercent = 65D;
        viewModel.ReplaceExistingOutput = true;
        await viewModel.RunCommand.ExecuteAsync(null);

        Assert.Equal(input, service.InputPath);
        Assert.Equal(output, service.OutputPath);
        Assert.NotNull(service.Options);
        Assert.Equal(TesseractOcrLanguage.English | TesseractOcrLanguage.Polish, service.Options!.Languages);
        Assert.Equal("1,3", service.Options.Pdf.ReadOptions.PageSelection!.ToString());
        Assert.Equal(200D, service.Options.Pdf.Dpi);
        Assert.Equal(0.65D, service.Options.Pdf.MinimumConfidence);
        Assert.Equal(OfficeConversionFileConflictPolicy.Replace, service.Options.OutputConflictPolicy);
        Assert.True(viewModel.HasOutput);
        Assert.Contains("42", viewModel.Summary, StringComparison.Ordinal);

        await viewModel.OpenOutputCommand.ExecuteAsync(null);
        Assert.Equal(output, opened);
    }

    [Fact]
    public async Task FailureRemainsActionableAndDoesNotPublishAnOutput() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-studio-ocr-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(directory, "scan.pdf");
        var service = new RecordingOcrService(new InvalidOperationException("Install or configure Tesseract."));
        using var viewModel = new SearchablePdfOcrViewModel(
            _ => Task.FromResult<string?>(input),
            _ => Task.FromResult<string?>(directory),
            service: service);
        viewModel.UseDocument(input);

        await viewModel.RunCommand.ExecuteAsync(null);

        Assert.False(viewModel.HasOutput);
        Assert.True(viewModel.HasError);
        Assert.Equal("Install or configure Tesseract.", viewModel.ErrorMessage);
        Assert.Equal("OCR could not finish", viewModel.Status);
    }

    [Fact]
    public async Task RunRejectsReplacingTheSourcePdf() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-studio-ocr-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(directory, "scan.pdf");
        var service = new RecordingOcrService(new SearchablePdfOcrOutcome(1, [1], "fixture-ocr"));
        using var viewModel = new SearchablePdfOcrViewModel(
            _ => Task.FromResult<string?>(input),
            _ => Task.FromResult<string?>(directory),
            service: service);
        viewModel.UseDocument(input);
        viewModel.OutputPath = input;
        viewModel.ReplaceExistingOutput = true;

        await viewModel.RunCommand.ExecuteAsync(null);

        Assert.Equal(0, service.CallCount);
        Assert.False(viewModel.HasOutput);
        Assert.Equal("Choose an OCR output PDF that is different from the source PDF.", viewModel.ErrorMessage);
    }

    [Fact]
    public async Task RunRejectsAnOutputOwnedByAnotherOpenTab() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-studio-ocr-" + Guid.NewGuid().ToString("N"));
        string input = Path.Combine(directory, "scan.pdf");
        string output = Path.Combine(directory, "already-open.pdf");
        var service = new RecordingOcrService(new SearchablePdfOcrOutcome(1, [1], "fixture-ocr"));
        using var viewModel = new SearchablePdfOcrViewModel(
            _ => Task.FromResult<string?>(input),
            _ => Task.FromResult<string?>(directory),
            service: service,
            canPublishPath: path => !string.Equals(path, output, StringComparison.OrdinalIgnoreCase));
        viewModel.UseDocument(input);
        viewModel.OutputPath = output;
        viewModel.ReplaceExistingOutput = true;

        await viewModel.RunCommand.ExecuteAsync(null);

        Assert.Equal(0, service.CallCount);
        Assert.False(viewModel.HasOutput);
        Assert.Equal(
            "That PDF is already open in another tab. Close it or choose a different output file name.",
            viewModel.ErrorMessage);
    }

    private sealed class RecordingOcrService : ISearchablePdfOcrService {
        private readonly SearchablePdfOcrOutcome? _result;
        private readonly Exception? _exception;

        internal RecordingOcrService(SearchablePdfOcrOutcome result) => _result = result;

        internal RecordingOcrService(Exception exception) => _exception = exception;

        internal string? InputPath { get; private set; }
        internal string? OutputPath { get; private set; }
        internal SearchablePdfOcrOptions? Options { get; private set; }
        internal int CallCount { get; private set; }

        public Task<SearchablePdfOcrOutcome> MakeSearchableAsync(
            string inputPath,
            string outputPath,
            SearchablePdfOcrOptions options,
            CancellationToken cancellationToken) {
            CallCount++;
            InputPath = inputPath;
            OutputPath = outputPath;
            Options = options;
            if (_exception is not null) return Task.FromException<SearchablePdfOcrOutcome>(_exception);
            return Task.FromResult(_result!);
        }
    }
}
