using System.Collections.ObjectModel;
using System.ComponentModel;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Internal;
using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Ordered PDF/image OCR intake and review over the shared workflow session owner.</summary>
public sealed partial class OcrSessionViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<IReadOnlyList<string>>> _pickFiles;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputFolder;
    private readonly Func<TesseractOcrLanguage, bool, CancellationToken, Task<IOcrEngine>> _createEngine;
    private readonly IStudioLocalizer _localizer;
    private readonly StudioStorageAccess? _storage;
    private readonly StudioJobHistory? _jobs;
    private readonly OfficeWorkflowOutputRecoveryStore? _recovery;
    private readonly IOfficeWorkflowPublicationGuard? _guard;
    private readonly Func<string, Task<bool>> _confirmProvider;
    private readonly Func<string, CancellationToken, Task>? _openOutput;
    private CancellationTokenSource? _cancellation;
    private bool _disposed;

    internal OcrSessionViewModel(Func<CancellationToken, Task<IReadOnlyList<string>>> pickFiles,
        Func<CancellationToken, Task<string?>> pickOutputFolder, IStudioLocalizer localizer,
        StudioStorageAccess? storage = null, StudioJobHistory? jobs = null,
        OfficeWorkflowOutputRecoveryStore? recovery = null, IOfficeWorkflowPublicationGuard? guard = null,
        Func<string, Task<bool>>? confirmProvider = null,
        Func<TesseractOcrLanguage, bool, CancellationToken, Task<IOcrEngine>>? createEngine = null,
        Func<string, CancellationToken, Task>? openOutput = null) {
        _pickFiles = pickFiles; _pickOutputFolder = pickOutputFolder; _localizer = localizer;
        _storage = storage; _jobs = jobs; _recovery = recovery; _guard = guard;
        _confirmProvider = confirmProvider ?? (_ => Task.FromResult(false));
        _createEngine = createEngine ?? CreateEngineAsync;
        _openOutput = openOutput;
        Languages = OcrLanguageChoice.CreateChoices(localizer);
        foreach (var language in Languages) language.PropertyChanged += LanguageChanged;
        Status = T("Ready", "Add PDFs or images. Each file is reviewed before its output is saved.");
    }
    public ObservableCollection<OcrSessionItem> Items { get; } = [];
    public bool HasItems => Items.Count > 0;
    public string SetupHint => !HasItems ? string.Empty
        : string.IsNullOrWhiteSpace(OutputFolder) ? T("Setup.Folder", "Choose an output folder to enable recognition.")
        : !Languages.Any(item => item.IsSelected) ? T("Setup.Language", "Select at least one recognition language.")
        : string.Empty;
    public IReadOnlyList<OcrLanguageChoice> Languages { get; }
    [ObservableProperty] private OcrSessionItem? _selectedItem;
    [ObservableProperty] private string _outputFolder = string.Empty;
    [ObservableProperty] private string _status;
    [ObservableProperty] private bool _provisionMissingLanguageData = true;
    [ObservableProperty] private bool _replaceExisting;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanRun))]
    [NotifyPropertyChangedFor(nameof(CanRetry))]
    private bool _isBusy;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasReview))]
    [NotifyPropertyChangedFor(nameof(HasPdfReview))]
    private OcrReviewViewModel? _pdfReview;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasReview))]
    [NotifyPropertyChangedFor(nameof(HasImageReview))]
    private ImageOcrReviewViewModel? _imageReview;
    public bool HasReview => PdfReview is not null || ImageReview is not null;
    public bool HasPdfReview => PdfReview is not null;
    public bool HasImageReview => ImageReview is not null;
    public bool HasSelectedOutput => SelectedItem?.HasOutput == true;
    public bool CanRemoveSelected => !_disposed && !IsBusy && SelectedItem is not null;
    public bool CanChooseConflictPolicy => string.IsNullOrWhiteSpace(OutputFolder) || _storage?.UsesProviderPublication(OutputFolder) != true;
    partial void OnSelectedItemChanged(OcrSessionItem? oldValue, OcrSessionItem? newValue) {
        if (oldValue is not null) oldValue.PropertyChanged -= SelectedItemChanged;
        if (newValue is not null) newValue.PropertyChanged += SelectedItemChanged;
        OnPropertyChanged(nameof(HasSelectedOutput));
        RemoveSelectedCommand.NotifyCanExecuteChanged();
    }
    private void SelectedItemChanged(object? sender, PropertyChangedEventArgs args) => OnPropertyChanged(nameof(HasSelectedOutput));
    public bool CanRun => !_disposed && !IsBusy && !string.IsNullOrWhiteSpace(OutputFolder) &&
        Languages.Any(item => item.IsSelected) && Items.Any(item => item.Status is null);
    public bool CanRetry => !_disposed && !IsBusy && !string.IsNullOrWhiteSpace(OutputFolder) &&
        Languages.Any(item => item.IsSelected) && Items.Any(item => item.CanRetry);
    partial void OnIsBusyChanged(bool value) => NotifyCommands();
    partial void OnOutputFolderChanged(string value) {
        OnPropertyChanged(nameof(CanChooseConflictPolicy)); NotifyCommands();
    }
    private void LanguageChanged(object? sender, PropertyChangedEventArgs args) => NotifyCommands();

    [RelayCommand]
    private async Task AddFilesAsync(CancellationToken token) {
        if (IsBusy || _disposed) return;
        try {
            var files = await _pickFiles(token).ConfigureAwait(true);
            if (IsBusy || _disposed) return;
            int unsupported = 0;
            foreach (string file in files) {
                if (Items.Count >= OfficeWorkflowRunner.MaximumBatchRequestCount) throw new InvalidOperationException(T("Limit", "The OCR session is full. Remove finished items before adding more files."));
                string source = OfficeStorageIdentity.Normalize(file);
                string name = _storage?.Describe(source).Name ?? OfficeStorageIdentity.GetFileName(source);
                string extension = Path.GetExtension(name).ToLowerInvariant();
                if (extension is not (".pdf" or ".png" or ".jpg" or ".jpeg" or ".bmp" or ".tif" or ".tiff" or ".gif" or ".webp")) { unsupported++; continue; }
                if (Items.Any(item => OfficeStorageIdentity.AreEquivalent(item.InputPath, source))) continue;
                Items.Add(new(source, name, _localizer, Items.Select(item => item.OutputName)));
            }
            SelectedItem ??= Items.FirstOrDefault();
            Status = unsupported == 0 ? T("Ready", "Add PDFs or images. Each file is reviewed before its output is saved.")
                : T("Unsupported", "Some selected files were not PDFs or supported raster images and were not added.");
        } catch (Exception error) { Status = error.Message; }
        NotifyCommands();
    }
    [RelayCommand(CanExecute = nameof(CanRemoveSelected))]
    private void RemoveSelected() {
        if (IsBusy || SelectedItem is null) return;
        Items.Remove(SelectedItem); SelectedItem = Items.FirstOrDefault(); NotifyCommands();
    }
    [RelayCommand]
    private async Task ChooseOutputFolderAsync(CancellationToken token) {
        if (IsBusy || _disposed) return;
        try {
            string? folder = await _pickOutputFolder(token).ConfigureAwait(true);
            if (!IsBusy && !_disposed && !string.IsNullOrWhiteSpace(folder)) OutputFolder = folder;
        } catch (Exception error) { Status = error.Message; }
    }
    [RelayCommand(CanExecute = nameof(CanRun))]
    private Task RunAsync() => RunItemsAsync(Items.Where(item => item.Status is null).ToArray());
    [RelayCommand(CanExecute = nameof(CanRetry))]
    private Task RetryAsync() => RunItemsAsync(Items.Where(item => item.CanRetry).ToArray());
    [RelayCommand] private void Cancel() => _cancellation?.Cancel();
    [RelayCommand]
    private async Task OpenOutputAsync(CancellationToken token) {
        if (SelectedItem is not { HasOutput: true, OutputPath: { } path } || _openOutput is null) return;
        try { await _openOutput(path, token).ConfigureAwait(true); }
        catch (Exception error) { Status = error.Message; }
    }

    private async Task RunItemsAsync(OcrSessionItem[] candidates) {
        if (IsBusy || _disposed || candidates.Length == 0) return;
        using var operation = new CancellationTokenSource();
        _cancellation = operation; IsBusy = true;
        StudioStorageAccess.DirectoryOutputSession? directory = null;
        var history = new Dictionary<string, StudioJobRecord>();
        bool dispatched = false;
        foreach (var item in candidates) item.ResetAttempt();
        try {
            var languages = Languages.Where(item => item.IsSelected).Aggregate((TesseractOcrLanguage)0, (value, item) => value | item.Value);
            bool provision = ProvisionMissingLanguageData;
            bool replace = ReplaceExisting;
            string folder = OutputFolder;
            if (_storage?.UsesProviderPublication(folder) == true) {
                if (!await _confirmProvider(folder).ConfigureAwait(true)) operation.Cancel();
                operation.Token.ThrowIfCancellationRequested();
                directory = _storage.CreateDirectoryOutput(folder, _recovery ?? throw new IOException("Workflow recovery storage is unavailable."));
            }
            var requests = new List<OfficeOcrSessionRequest>();
            foreach (var item in candidates) {
                var destination = directory is null ? null : await directory.ResolveAsync(item.OutputName, operation.Token).ConfigureAwait(true);
                string output = destination?.Location ?? Path.Combine(Path.GetFullPath(folder), item.OutputName);
                var policy = destination is not null || replace ? OfficeWorkflowConflictPolicy.Replace : OfficeWorkflowConflictPolicy.Fail;
                if (item.IsPdf) requests.Add(new(item.Id, new PdfSearchableWorkflowRequest {
                    InputPath = item.InputPath, InputStream = _storage?.CreateWorkflowInput(item.InputPath),
                    OutputPath = output, OutputStream = destination?.Output, ConflictPolicy = policy,
                    PublicationGuard = _guard, Ocr = new() { Language = languages.ToTesseractExpression() },
                    ReviewAsync = (evidence, token) => ReviewPdfAsync(item, evidence, operation, token)
                }));
                else requests.Add(new(item.Id, new ImageOcrWorkflowRequest {
                    InputPath = item.InputPath, InputStream = _storage?.CreateWorkflowInput(item.InputPath),
                    OutputPath = output, OutputStream = destination?.Output, ConflictPolicy = policy,
                    PublicationGuard = _guard, Ocr = new() { Language = languages.ToTesseractExpression() },
                    ReviewAsync = (evidence, token) => ReviewImageAsync(item, evidence, operation, token)
                }));
                if (_jobs is not null) history.Add(item.Id, _jobs.Start(T("Title", "OCR session"), item.InputPath, output, operation.Cancel, candidates.Length > 1));
            }
            using var permit = _jobs is null ? null : await _jobs.EnterAsync(operation.Token).ConfigureAwait(true);
            Status = T("Preparing", "Preparing the OCR runtime…");
            IOcrEngine engine = await _createEngine(languages, provision, operation.Token).ConfigureAwait(true);
            var progress = new Progress<OfficeWorkflowProgress>(update => {
                if (!ReferenceEquals(_cancellation, operation) || _disposed) return;
                var item = candidates.First(entry => entry.Id == update.RequestId);
                if (update.Stage == "execute" && item.Status is null) item.IsRunning = true;
                if (history.TryGetValue(item.Id, out var record)) record.Report(update);
            });
            void Apply(OfficeOcrSessionResult result) {
                var item = candidates.First(entry => entry.Id == result.Id);
                item.Apply(result);
                if (history.TryGetValue(item.Id, out var record)) record.Complete(result.Status, result.OutputPath, result.Summary, result.Recovery);
            }
            var completed = new Progress<OfficeOcrSessionResult>(result => {
                if (ReferenceEquals(_cancellation, operation) && !_disposed) Apply(result);
            });
            dispatched = true;
            var results = await new OfficeWorkflowRunner().RunOcrSessionAsync(requests, engine, progress, completed, operation.Token,
                Items.Where(item => item.HasOutput).Select(item => item.OutputPath!).ToArray(),
                Items.Select(item => new OfficeWorkflowProtectedSource(item.InputPath, _storage?.CreateWorkflowInput(item.InputPath))).ToArray()).ConfigureAwait(true);
            foreach (var result in results) Apply(result);
            Status = _localizer.FormatOrDefault("OcrSession.Outcomes", "Last run: {0} saved · {1} failed · {2} cancelled · {3} need checking. Saved outputs were retained.",
                candidates.Count(item => item.Status == OfficeWorkflowStatus.Completed), candidates.Count(item => item.Status == OfficeWorkflowStatus.Failed),
                candidates.Count(item => item.Status == OfficeWorkflowStatus.Cancelled), candidates.Count(item => item.Status == OfficeWorkflowStatus.Unconfirmed));
        } catch (Exception error) {
            var state = dispatched ? OfficeWorkflowStatus.Unconfirmed : error is OperationCanceledException ? OfficeWorkflowStatus.Cancelled : OfficeWorkflowStatus.Failed;
            foreach (var item in candidates.Where(item => item.Status is null)) {
                item.IsRunning = false; item.Status = state; item.Summary = error.Message;
                if (history.TryGetValue(item.Id, out var record)) record.Complete(state, null, error.Message);
            }
            Status = error is OperationCanceledException ? T("Cancelled", "OCR session cancelled. Completed outputs were retained.") : error.Message;
        } finally {
            PdfReview?.Dispose(); PdfReview = null; ImageReview?.Dispose(); ImageReview = null;
            try { directory?.Dispose(); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) { Status += " " + error.Message; }
            if (ReferenceEquals(_cancellation, operation)) _cancellation = null;
            IsBusy = false; NotifyCommands();
        }
    }
    private static async Task<IOcrEngine> CreateEngineAsync(TesseractOcrLanguage languages, bool provision, CancellationToken token) =>
        (await TesseractOcr.CreateSessionAsync(new() { Languages = languages, ProvisionMissingLanguageData = provision }, token).ConfigureAwait(false)).Engine;
    private void NotifyCommands() {
        OnPropertyChanged(nameof(HasItems)); OnPropertyChanged(nameof(SetupHint));
        OnPropertyChanged(nameof(CanRun)); OnPropertyChanged(nameof(CanRetry));
        RunCommand.NotifyCanExecuteChanged(); RetryCommand.NotifyCanExecuteChanged();
        RemoveSelectedCommand.NotifyCanExecuteChanged();
    }
    private string T(string suffix, string fallback) => _localizer.GetOrDefault("OcrSession." + suffix, fallback);
    public void Dispose() {
        if (_disposed) return;
        _disposed = true; _cancellation?.Cancel(); PdfReview?.Dispose(); ImageReview?.Dispose();
        if (SelectedItem is not null) SelectedItem.PropertyChanged -= SelectedItemChanged;
        foreach (var language in Languages) language.PropertyChanged -= LanguageChanged;
    }
}
