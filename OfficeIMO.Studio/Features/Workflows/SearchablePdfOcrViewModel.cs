using System.Collections.ObjectModel;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Workflows;
using OfficeIMO.Internal;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Workflows;

internal sealed record SearchablePdfOcrOutcome(
    int AddedWordCount,
    IReadOnlyList<int> ModifiedPages,
    string? Provider,
    PdfSearchableWorkflowResult? Workflow = null);

internal sealed record SearchablePdfOcrOptions(
    TesseractOcrLanguage Languages,
    bool ProvisionMissingLanguageData,
    OfficeConversionFileConflictPolicy OutputConflictPolicy,
    PdfOcrMergeOptions Pdf) {
    internal IOfficeWorkflowPublicationGuard? PublicationGuard { get; init; }
    internal OfficeWorkflowStreamInput? InputStream { get; init; }
    internal OfficeWorkflowStreamOutput? OutputStream { get; init; }
    internal Func<PdfSearchableOcrReview, CancellationToken, Task<IReadOnlyList<PdfRecognizedWord>>>? ReviewAsync { get; init; }
}

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
        var pdfOptions = options.Pdf.Clone();
        pdfOptions.Language = options.Languages.ToTesseractExpression();
        pdfOptions.SourceName = options.InputStream?.Name ?? OfficeStorageIdentity.GetFileName(inputPath);
        var request = new PdfSearchableWorkflowRequest {
            InputPath = inputPath,
            OutputPath = outputPath,
            Ocr = pdfOptions,
            InputStream = options.InputStream,
            OutputStream = options.OutputStream,
            ConflictPolicy = options.OutputConflictPolicy == OfficeConversionFileConflictPolicy.Replace
                ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail,
            PublicationGuard = options.PublicationGuard,
            ReviewAsync = options.ReviewAsync
        };
        TesseractOcrSession session = await TesseractOcr
            .CreateSessionAsync(new TesseractOcrSessionOptions {
                Languages = options.Languages,
                ProvisionMissingLanguageData = options.ProvisionMissingLanguageData
            }, cancellationToken)
            .ConfigureAwait(false);
        PdfSearchableWorkflowResult result = await new OfficeWorkflowRunner()
            .MakePdfSearchableAsync(request, session.Engine, cancellationToken).ConfigureAwait(false);
        return new SearchablePdfOcrOutcome(result.AddedWordCount, result.ModifiedPages, result.Provider, result);
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

    internal static IReadOnlyList<OcrLanguageChoice> CreateChoices(IStudioLocalizer localizer) =>
        TesseractOcrLanguages.Supported.Select(language => new OcrLanguageChoice(language,
            localizer.GetOrDefault($"Ocr.Language.{language}", FormatLanguage(language)), language == TesseractOcrLanguage.English)).ToArray();

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
}

public sealed partial class SearchablePdfOcrViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<string?>> _pickPdf;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputFolder;
    private readonly Func<string, CancellationToken, Task>? _openDocument;
    private readonly ISearchablePdfOcrService _service;
    private readonly Func<string, bool> _canPublishPath;
    private readonly IOfficeWorkflowPublicationGuard? _publicationGuard;
    private readonly IStudioLocalizer _localizer;
    private readonly StudioJobHistory? _jobHistory;
    private readonly StudioStorageAccess? _storage;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputPdf;
    private readonly OfficeWorkflowOutputRecoveryStore? _recoveryStore;
    private readonly Func<string, Task<bool>> _confirmProviderWrite;
    private CancellationTokenSource? _cancellation;
    private string? _automaticOutputPath;
    internal Func<string, CancellationToken, Task>? OpenAssistantOutput { get; set; }
    internal bool ReturnToAssistant { get; private set; }
    private bool _disposed;

    internal SearchablePdfOcrViewModel(
        Func<CancellationToken, Task<string?>> pickPdf,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        Func<string, CancellationToken, Task>? openDocument = null,
        ISearchablePdfOcrService? service = null,
        Func<string, bool>? canPublishPath = null,
        IStudioLocalizer? localizer = null,
        StudioJobHistory? jobHistory = null,
        IOfficeWorkflowPublicationGuard? publicationGuard = null,
        StudioStorageAccess? storage = null,
        Func<CancellationToken, Task<string?>>? pickOutputPdf = null,
        OfficeWorkflowOutputRecoveryStore? recoveryStore = null,
        Func<string, Task<bool>>? confirmProviderWrite = null,
        IScanTextRecognitionService? textRecognition = null) {
        _pickPdf = pickPdf ?? throw new ArgumentNullException(nameof(pickPdf));
        _pickOutputFolder = pickOutputFolder ?? throw new ArgumentNullException(nameof(pickOutputFolder));
        _openDocument = openDocument;
        _service = service ?? new SearchablePdfOcrService();
        _textRecognition = textRecognition ?? new ScanTextRecognitionService();
        _canPublishPath = canPublishPath ?? (_ => true);
        _publicationGuard = publicationGuard ?? new OfficeIMO.Studio.Features.Shell.StudioWorkflowPublicationGuard((path, _) => _canPublishPath(path));
        _localizer = localizer ?? StudioLocalization.Current;
        _jobHistory = jobHistory;
        _storage = storage;
        _pickOutputPdf = pickOutputPdf ?? (_ => Task.FromResult<string?>(null));
        _recoveryStore = recoveryStore;
        _confirmProviderWrite = confirmProviderWrite ?? (_ => Task.FromResult(false));
        Scan = new ScanPreparationViewModel(ReadScanSourceAsync, SaveScanCopyAsync, _localizer);
        Scan.PropertyChanged += ScanChanged;
        Languages = new ObservableCollection<OcrLanguageChoice>(OcrLanguageChoice.CreateChoices(_localizer));
        foreach (var choice in Languages) choice.PropertyChanged += OnLanguagePropertyChanged;
        Status = T("Status.Ready", "Choose a scanned PDF to make its text searchable.");
        Summary = T("Summary.Empty", "No OCR output yet");
    }

    public ObservableCollection<OcrLanguageChoice> Languages { get; }
    public ScanPreparationViewModel Scan { get; }

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RunCommand))]
    [NotifyPropertyChangedFor(nameof(InputName))]
    [NotifyCanExecuteChangedFor(nameof(ExtractTextCommand))]
    private string _inputPath = string.Empty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(RunCommand))]
    [NotifyPropertyChangedFor(nameof(OutputName))]
    [NotifyPropertyChangedFor(nameof(IsProviderOutput))]
    private string _outputPath = string.Empty;

    public string InputName => Describe(InputPath);
    public string OutputName => Describe(OutputPath);
    public bool IsProviderOutput => !string.IsNullOrWhiteSpace(OutputPath) && _storage?.UsesProviderPublication(OutputPath) == true;

    [ObservableProperty]
    private bool _hasRecovery;

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
    [NotifyCanExecuteChangedFor(nameof(ExtractTextCommand))]
    [NotifyCanExecuteChangedFor(nameof(RunCommand))]
    private bool _isBusy;

    [ObservableProperty]
    private string _status = string.Empty;

    [ObservableProperty]
    private string _summary = string.Empty;

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
                0 => T("Language.None", "Select at least one language"),
                1 => selected[0],
                _ => _localizer.FormatOrDefault("Ocr.Language.Count", "{0} languages selected", selected.Length)
            };
        }
    }

    private bool CanRun =>
        !IsBusy && !Scan.IsBusy &&
        !string.IsNullOrWhiteSpace(InputPath) &&
        !string.IsNullOrWhiteSpace(OutputPath) &&
        Languages.Any(static choice => choice.IsSelected);

    internal void UseDocument(string? path, bool returnToAssistant = false) {
        if (IsBusy) return;
        if (!string.IsNullOrWhiteSpace(path)) InputPath = path;
        ReturnToAssistant = returnToAssistant;
    }

    partial void OnInputPathChanged(string value) {
        ReturnToAssistant = false;
        if (_isExtractingText) _cancellation?.Cancel();
        ExtractedText = string.Empty;
        Scan.Invalidate(clearSource: true);
        string? suggestion = TryCreateOutputPath(value);
        if (suggestion is null) {
            if (PathsEqual(OutputPath, _automaticOutputPath)) OutputPath = string.Empty;
            _automaticOutputPath = null;
            return;
        }
        if (string.IsNullOrWhiteSpace(OutputPath) || PathsEqual(OutputPath, _automaticOutputPath)) {
            _automaticOutputPath = suggestion;
            OutputPath = suggestion;
        }
    }

    [RelayCommand]
    private async Task ChooseInputAsync(CancellationToken cancellationToken) {
        if (IsBusy) return;
        string? path = await _pickPdf(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) UseDocument(path);
    }

    [RelayCommand]
    private async Task ChooseOutputFolderAsync(CancellationToken cancellationToken) {
        if (IsBusy) return;
        string? folder = await _pickOutputFolder(cancellationToken).ConfigureAwait(true);
        if (string.IsNullOrWhiteSpace(folder)) return;
        string inputName = string.IsNullOrWhiteSpace(InputPath)
            ? "searchable"
            : Path.GetFileNameWithoutExtension(InputName) + "-searchable";
        _automaticOutputPath = Path.Combine(OfficeStorageIdentity.GetLocalPath(folder)
            ?? throw new IOException("Choose a filesystem output folder."), inputName + ".pdf");
        OutputPath = _automaticOutputPath;
    }

    [RelayCommand]
    private async Task ChooseOutputAsync(CancellationToken cancellationToken) {
        if (IsBusy) return;
        string? path = await _pickOutputPdf(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path) && !IsBusy) {
            _automaticOutputPath = null;
            OutputPath = path;
        }
    }

    [RelayCommand(CanExecute = nameof(CanRun))]
    private async Task RunAsync() {
        if (!CanRun) return;
        _cancellation?.Dispose();
        using var operation = new CancellationTokenSource();
        _cancellation = operation;
        IsBusy = true;
        PublishedPath = null;
        HasRecovery = false;
        ErrorMessage = null;
        Status = T("Status.Preparing", "Preparing the OCR engine and page renderings…");
        Summary = T("Summary.Running", "OCR is running");
        StudioJobRecord? job = null;

        try {
            string input = OfficeStorageIdentity.Normalize(InputPath.Trim());
            string output = OfficeStorageIdentity.Normalize(OutputPath.Trim());
            bool providerOutput = _storage?.UsesProviderPublication(output) == true;
            bool providerInput = _storage?.UsesProviderPublication(input) == true;
            if (providerInput || providerOutput ? string.Equals(input, output, StringComparison.Ordinal)
                    : OfficeStorageIdentity.AreEquivalent(input, output)) {
                throw new InvalidOperationException(T("Error.SamePath", "Choose an OCR output PDF that is different from the source PDF."));
            }
            if (!providerInput && !providerOutput && !_canPublishPath(output)) {
                throw new InvalidOperationException(
                    T("Error.OutputOpen", "That PDF is already open in another tab. Close it or choose a different output file name."));
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
                Scan.ApplyTo(new PdfOcrMergeOptions {
                    ReadOptions = new PdfReadOptions {
                        PageSelection = string.IsNullOrWhiteSpace(Pages) ? null : PdfPageSelection.Parse(Pages)
                    },
                    Dpi = RenderDpi,
                    MinimumConfidence = MinimumConfidencePercent / 100D
                })) { PublicationGuard = _publicationGuard, InputStream = _storage?.CreateWorkflowInput(input), ReviewAsync = ReviewWordsAsync };
            if (providerOutput) {
                if (!await _confirmProviderWrite(output).ConfigureAwait(true)) {
                    Status = T("Status.Cancelled", "OCR cancelled");
                    Summary = T("Summary.Empty", "No OCR output yet");
                    return;
                }
                operation.Token.ThrowIfCancellationRequested();
                options = options with {
                    OutputStream = _storage!.CreateWorkflowOutput(output, _recoveryStore
                        ?? throw new IOException("Workflow recovery storage is unavailable.")),
                    OutputConflictPolicy = OfficeConversionFileConflictPolicy.Replace
                };
            }
            job = _jobHistory?.Start(T("Job.Title", "Searchable PDF OCR"), input, output, operation.Cancel);
            using IDisposable? execution = _jobHistory is null ? null : await _jobHistory.EnterAsync(operation.Token).ConfigureAwait(true);
            if (!providerInput && !providerOutput && !_canPublishPath(output)) throw new InvalidOperationException(T("Error.OutputOpen", "That PDF is already open in another tab. Close it or choose a different output file name."));
            job?.Report(new OfficeIMO.Workflows.OfficeWorkflowProgress("ocr", "execute", Status, 0D));
            SearchablePdfOcrOutcome result = await _service
                .MakeSearchableAsync(input, output, options, operation.Token)
                .ConfigureAwait(true);
            HasRecovery = result.Workflow?.Recovery is not null;
            if (result.Workflow is { Status: not OfficeWorkflowStatus.Completed } workflow) {
                Status = workflow.Status == OfficeWorkflowStatus.Cancelled
                    ? T("Status.Cancelled", "OCR cancelled") : workflow.Status == OfficeWorkflowStatus.Unconfirmed
                        ? _localizer.GetOrDefault("Jobs.Unconfirmed", "Check output") : T("Status.Failed", "OCR could not finish");
                Summary = workflow.Summary;
                if (workflow.Status != OfficeWorkflowStatus.Cancelled) ErrorMessage = workflow.Summary;
                job?.Complete(workflow.Status, null, workflow.Summary, workflow.Recovery);
                return;
            }
            PublishedPath = result.Workflow?.OutputPath ?? output;
            string pageLabel = result.ModifiedPages.Count == 1
                ? T("Result.OnePage", "1 page")
                : _localizer.FormatOrDefault("Ocr.Result.Pages", "{0:N0} pages", result.ModifiedPages.Count);
            Status = result.AddedWordCount > 0
                ? T("Status.Completed", "Searchable PDF created")
                : T("Status.NoWords", "PDF created; no searchable words were added");
            Summary = string.IsNullOrWhiteSpace(result.Provider)
                ? _localizer.FormatOrDefault("Ocr.Result.Summary", "Added {0:N0} searchable words across {1}.", result.AddedWordCount, pageLabel)
                : _localizer.FormatOrDefault("Ocr.Result.SummaryWithProvider", "Added {0:N0} searchable words across {1} with {2}.", result.AddedWordCount, pageLabel, result.Provider);
            job?.Complete(OfficeIMO.Workflows.OfficeWorkflowStatus.Completed, PublishedPath, Summary, result.Workflow?.Recovery);
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
            Status = T("Status.Cancelled", "OCR cancelled");
            Summary = T("Summary.SourceUnchanged", "The source PDF was not changed.");
            job?.Complete(OfficeIMO.Workflows.OfficeWorkflowStatus.Cancelled, null, Summary);
        } catch (Exception ex) {
            Status = T("Status.Failed", "OCR could not finish");
            Summary = T("Summary.SourceUnchanged", "The source PDF was not changed.");
            ErrorMessage = ex.Message;
            job?.Unconfirmed(ex.Message);
        } finally {
            Review?.Dispose();
            Review = null;
            IsBusy = false;
            if (ReferenceEquals(_cancellation, operation)) _cancellation = null;
        }
    }

    [RelayCommand]
    private void Cancel() => _cancellation?.Cancel();

    [RelayCommand(CanExecute = nameof(HasOutput))]
    private Task OpenOutputAsync(CancellationToken cancellationToken) =>
        ReturnToAssistant && OpenAssistantOutput is not null && PublishedPath is not null
            ? OpenAssistantOutput(PublishedPath, cancellationToken)
            : _openDocument is not null && PublishedPath is not null
            ? _openDocument(PublishedPath, cancellationToken)
            : Task.CompletedTask;

    private void OnLanguagePropertyChanged(object? sender, PropertyChangedEventArgs e) {
        if (e.PropertyName != nameof(OcrLanguageChoice.IsSelected)) return;
        OnPropertyChanged(nameof(LanguageSummary));
        ExtractTextCommand.NotifyCanExecuteChanged();
        RunCommand.NotifyCanExecuteChanged();
    }

    private string? TryCreateOutputPath(string value) {
        if (string.IsNullOrWhiteSpace(value)) return null;
        try {
            if (_storage?.UsesProviderPublication(value) == true) return null;
            string? input = OfficeStorageIdentity.GetLocalPath(value);
            if (input is null) return null;
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

    private string Describe(string location) {
        if (string.IsNullOrWhiteSpace(location)) return string.Empty;
        try { return _storage?.Describe(location).Name ?? OfficeStorageIdentity.GetFileName(location); } catch (Exception error) when (error is ArgumentException or NotSupportedException or IOException) { return location; }
    }

    public void Dispose() {
        _disposed = true;
        ExtractTextCommand.NotifyCanExecuteChanged();
        _cancellation?.Cancel();
        Review?.Dispose();
        Scan.PropertyChanged -= ScanChanged; Scan.Dispose();
        foreach (OcrLanguageChoice language in Languages) language.PropertyChanged -= OnLanguagePropertyChanged;
    }

    private string T(string suffix, string fallback) =>
        _localizer.GetOrDefault("Ocr." + suffix, fallback);
}
