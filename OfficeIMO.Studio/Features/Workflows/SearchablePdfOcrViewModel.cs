using System.Collections.ObjectModel;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Studio.Features.Workflows;

internal sealed record SearchablePdfOcrOutcome(
    int AddedWordCount,
    IReadOnlyList<int> ModifiedPages,
    string? Provider);

internal sealed record SearchablePdfOcrOptions(
    TesseractOcrLanguage Languages,
    bool ProvisionMissingLanguageData,
    OfficeConversionFileConflictPolicy OutputConflictPolicy,
    PdfOcrMergeOptions Pdf);

internal interface ISearchablePdfOcrService {
    Task<SearchablePdfOcrOutcome> MakeSearchableAsync(
        string inputPath,
        string outputPath,
        SearchablePdfOcrOptions options,
        CancellationToken cancellationToken);
}

internal sealed class SearchablePdfOcrService : ISearchablePdfOcrService {
    public async Task<SearchablePdfOcrOutcome> MakeSearchableAsync(
        string inputPath,
        string outputPath,
        SearchablePdfOcrOptions options,
        CancellationToken cancellationToken) {
        TesseractOcrSession session = await TesseractOcr
            .CreateSessionAsync(new TesseractOcrSessionOptions {
                Languages = options.Languages,
                ProvisionMissingLanguageData = options.ProvisionMissingLanguageData
            }, cancellationToken)
            .ConfigureAwait(false);
        options.Pdf.Language = options.Languages.ToTesseractExpression();
        options.Pdf.SourceName = inputPath;
        PdfDocument source = PdfDocument.Load(inputPath);
        PdfSearchableOcrResult result = await source
            .MakeSearchableAsync(session.Engine, options.Pdf, cancellationToken)
            .ConfigureAwait(false);
        await result.Document
            .SaveAsync(outputPath, options.OutputConflictPolicy, cancellationToken)
            .ConfigureAwait(false);
        string? provider = result.Ocr.Pages
            .Select(static page => page.Provider)
            .FirstOrDefault(static value => !string.IsNullOrWhiteSpace(value));
        return new SearchablePdfOcrOutcome(result.AddedWordCount, result.ModifiedPages, provider);
    }
}

public sealed partial class OcrLanguageChoice : ObservableObject {
    internal OcrLanguageChoice(TesseractOcrLanguage value, string label, bool isSelected = false) {
        Value = value;
        Label = label;
        _isSelected = isSelected;
    }

    internal TesseractOcrLanguage Value { get; }

    public string Label { get; }

    [ObservableProperty]
    private bool _isSelected;
}

public sealed partial class SearchablePdfOcrViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<string?>> _pickPdf;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputFolder;
    private readonly Func<string, CancellationToken, Task>? _openDocument;
    private readonly ISearchablePdfOcrService _service;
    private readonly Func<string, bool> _canPublishPath;
    private CancellationTokenSource? _cancellation;
    private string? _automaticOutputPath;

    internal SearchablePdfOcrViewModel(
        Func<CancellationToken, Task<string?>> pickPdf,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        Func<string, CancellationToken, Task>? openDocument = null,
        ISearchablePdfOcrService? service = null,
        Func<string, bool>? canPublishPath = null) {
        _pickPdf = pickPdf ?? throw new ArgumentNullException(nameof(pickPdf));
        _pickOutputFolder = pickOutputFolder ?? throw new ArgumentNullException(nameof(pickOutputFolder));
        _openDocument = openDocument;
        _service = service ?? new SearchablePdfOcrService();
        _canPublishPath = canPublishPath ?? (_ => true);
        Languages = new ObservableCollection<OcrLanguageChoice>(TesseractOcrLanguages.Supported.Select(language => {
            var choice = new OcrLanguageChoice(language, FormatLanguage(language), language == TesseractOcrLanguage.English);
            choice.PropertyChanged += OnLanguagePropertyChanged;
            return choice;
        }));
    }

    public ObservableCollection<OcrLanguageChoice> Languages { get; }

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RunCommand))]
    private string _inputPath = string.Empty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RunCommand))]
    private string _outputPath = string.Empty;

    [ObservableProperty]
    private string _pages = string.Empty;

    [ObservableProperty]
    private double _renderDpi = 150D;

    [ObservableProperty]
    private double _minimumConfidencePercent = 50D;

    [ObservableProperty]
    private bool _provisionMissingLanguageData = true;

    [ObservableProperty]
    private bool _replaceExistingOutput;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanCancel))]
    [NotifyCanExecuteChangedFor(nameof(RunCommand))]
    private bool _isBusy;

    [ObservableProperty]
    private string _status = "Choose a scanned PDF to make its text searchable.";

    [ObservableProperty]
    private string _summary = "No OCR output yet";

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasError))]
    private string? _errorMessage;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasOutput))]
    [NotifyCanExecuteChangedFor(nameof(OpenOutputCommand))]
    private string? _publishedPath;

    public bool CanCancel => IsBusy;

    public bool HasOutput => !string.IsNullOrWhiteSpace(PublishedPath);

    public bool HasError => !string.IsNullOrWhiteSpace(ErrorMessage);

    public string LanguageSummary {
        get {
            string[] selected = Languages.Where(static choice => choice.IsSelected).Select(static choice => choice.Label).ToArray();
            return selected.Length switch {
                0 => "Select at least one language",
                1 => selected[0],
                _ => $"{selected.Length} languages selected"
            };
        }
    }

    private bool CanRun =>
        !IsBusy &&
        !string.IsNullOrWhiteSpace(InputPath) &&
        !string.IsNullOrWhiteSpace(OutputPath) &&
        Languages.Any(static choice => choice.IsSelected);

    internal void UseDocument(string? path) {
        if (!string.IsNullOrWhiteSpace(path)) InputPath = path;
    }

    partial void OnInputPathChanged(string value) {
        string? suggestion = TryCreateOutputPath(value);
        if (suggestion is null) return;
        if (string.IsNullOrWhiteSpace(OutputPath) || PathsEqual(OutputPath, _automaticOutputPath)) {
            _automaticOutputPath = suggestion;
            OutputPath = suggestion;
        }
    }

    [RelayCommand]
    private async Task ChooseInputAsync(CancellationToken cancellationToken) {
        string? path = await _pickPdf(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) UseDocument(path);
    }

    [RelayCommand]
    private async Task ChooseOutputFolderAsync(CancellationToken cancellationToken) {
        string? folder = await _pickOutputFolder(cancellationToken).ConfigureAwait(true);
        if (string.IsNullOrWhiteSpace(folder)) return;
        string inputName = string.IsNullOrWhiteSpace(InputPath)
            ? "searchable"
            : Path.GetFileNameWithoutExtension(InputPath) + "-searchable";
        _automaticOutputPath = Path.Combine(Path.GetFullPath(folder), inputName + ".pdf");
        OutputPath = _automaticOutputPath;
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync() {
        _cancellation?.Dispose();
        using var operation = new CancellationTokenSource();
        _cancellation = operation;
        IsBusy = true;
        PublishedPath = null;
        ErrorMessage = null;
        Status = "Preparing the OCR engine and page renderings…";
        Summary = "OCR is running";

        try {
            string input = Path.GetFullPath(InputPath.Trim());
            string output = Path.GetFullPath(OutputPath.Trim());
            if (PathsEqual(input, output)) {
                throw new InvalidOperationException("Choose an OCR output PDF that is different from the source PDF.");
            }
            if (!_canPublishPath(output)) {
                throw new InvalidOperationException(
                    "That PDF is already open in another tab. Close it or choose a different output file name.");
            }
            TesseractOcrLanguage selectedLanguages = Languages
                .Where(static choice => choice.IsSelected)
                .Aggregate((TesseractOcrLanguage)0, static (current, choice) => current | choice.Value);
            var options = new SearchablePdfOcrOptions(
                selectedLanguages,
                ProvisionMissingLanguageData,
                ReplaceExistingOutput
                    ? OfficeConversionFileConflictPolicy.Replace
                    : OfficeConversionFileConflictPolicy.FailIfExists,
                new PdfOcrMergeOptions {
                    ReadOptions = new PdfReadOptions {
                        PageSelection = string.IsNullOrWhiteSpace(Pages) ? null : PdfPageSelection.Parse(Pages)
                    },
                    Dpi = RenderDpi,
                    MinimumConfidence = MinimumConfidencePercent / 100D
                });
            SearchablePdfOcrOutcome result = await _service
                .MakeSearchableAsync(input, output, options, operation.Token)
                .ConfigureAwait(true);
            PublishedPath = output;
            string pageLabel = result.ModifiedPages.Count == 1 ? "1 page" : $"{result.ModifiedPages.Count:N0} pages";
            Status = result.AddedWordCount > 0
                ? "Searchable PDF created"
                : "PDF created; no new OCR words passed the selected confidence threshold";
            Summary = $"Added {result.AddedWordCount:N0} searchable words across {pageLabel}" +
                      (string.IsNullOrWhiteSpace(result.Provider) ? "." : $" with {result.Provider}.");
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
            Status = "OCR cancelled";
            Summary = "The source PDF was not changed.";
        } catch (Exception ex) {
            Status = "OCR could not finish";
            Summary = "The source PDF was not changed.";
            ErrorMessage = ex.Message;
        } finally {
            IsBusy = false;
            if (ReferenceEquals(_cancellation, operation)) _cancellation = null;
        }
    }

    [RelayCommand]
    private void Cancel() => _cancellation?.Cancel();

    [RelayCommand(CanExecute = nameof(HasOutput))]
    private Task OpenOutputAsync(CancellationToken cancellationToken) =>
        _openDocument is not null && PublishedPath is not null
            ? _openDocument(PublishedPath, cancellationToken)
            : Task.CompletedTask;

    private void OnLanguagePropertyChanged(object? sender, PropertyChangedEventArgs e) {
        if (e.PropertyName != nameof(OcrLanguageChoice.IsSelected)) return;
        OnPropertyChanged(nameof(LanguageSummary));
        RunCommand.NotifyCanExecuteChanged();
    }

    private static string? TryCreateOutputPath(string value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        try {
            string input = Path.GetFullPath(value);
            return Path.Combine(
                Path.GetDirectoryName(input)!,
                Path.GetFileNameWithoutExtension(input) + "-searchable.pdf");
        } catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException) {
            return null;
        }
    }

    private static bool PathsEqual(string? left, string? right) {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right)) return false;
        StringComparison comparison = OperatingSystem.IsWindows()
            ? StringComparison.OrdinalIgnoreCase
            : StringComparison.Ordinal;
        return string.Equals(left, right, comparison);
    }

    private static string FormatLanguage(TesseractOcrLanguage language) {
        string name = language.ToString();
        var label = new System.Text.StringBuilder(name.Length + 4);
        for (int index = 0; index < name.Length; index++) {
            char current = name[index];
            if (index > 0 && char.IsUpper(current) && char.IsLower(name[index - 1])) label.Append(' ');
            label.Append(current);
        }
        return label.ToString();
    }

    public void Dispose() {
        _cancellation?.Cancel();
        foreach (OcrLanguageChoice language in Languages) language.PropertyChanged -= OnLanguagePropertyChanged;
    }
}
