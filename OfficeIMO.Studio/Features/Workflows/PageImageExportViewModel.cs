using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Drawing;
using OfficeIMO.Internal;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed record ImageExportFormatChoice(OfficeImageExportFormat Value, string Label, string Description);

public sealed partial class PageImageExportViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<string?>> _pickPdf;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputFolder;
    private readonly IOfficeOutputWorkflowRunner _runner;
    private readonly IStudioLocalizer _localizer;
    private readonly IOfficeWorkflowPublicationGuard? _publicationGuard;
    private readonly StudioJobHistory? _jobHistory;
    private readonly StudioStorageAccess? _storage;
    private readonly OfficeWorkflowOutputRecoveryStore? _recoveryStore;
    private readonly Func<string, Task<bool>> _confirmProviderWrite;
    private CancellationTokenSource? _cancellation;

    public PageImageExportViewModel(
        Func<CancellationToken, Task<string?>> pickPdf,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeOutputWorkflowRunner? runner = null) : this(pickPdf, pickOutputFolder, runner, null) { }

    internal PageImageExportViewModel(
        Func<CancellationToken, Task<string?>> pickPdf,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeOutputWorkflowRunner? runner,
        IStudioLocalizer? localizer = null,
        IOfficeWorkflowPublicationGuard? publicationGuard = null,
        StudioJobHistory? jobHistory = null,
        StudioStorageAccess? storage = null,
        OfficeWorkflowOutputRecoveryStore? recoveryStore = null, Func<string, Task<bool>>? confirmProviderWrite = null) {
        _pickPdf = pickPdf;
        _pickOutputFolder = pickOutputFolder;
        _runner = runner ?? new OfficeWorkflowRunner();
        _publicationGuard = publicationGuard;
        _jobHistory = jobHistory;
        _storage = storage;
        _recoveryStore = recoveryStore;
        _confirmProviderWrite = confirmProviderWrite ?? (_ => Task.FromResult(false));
        _localizer = localizer ?? StudioLocalization.Current;
        Formats = [
            Format(OfficeImageExportFormat.Png, "PNG", "Lossless raster pages with transparency support."),
            Format(OfficeImageExportFormat.Jpeg, "JPEG", "Compact photographic raster pages."),
            Format(OfficeImageExportFormat.Webp, "WebP", "Compact lossless raster pages."),
            Format(OfficeImageExportFormat.Tiff, "TIFF", "Lossless archival raster pages."),
            Format(OfficeImageExportFormat.Svg, "SVG", "Managed vector page scenes where supported.")
        ];
        SelectedFormat = Formats[0];
        Status = T("Status.Ready", "Choose a PDF and an output folder.");
        Summary = T("Summary.Empty", "No export yet");
    }

    public IReadOnlyList<ImageExportFormatChoice> Formats { get; }

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ExportCommand))]
    [NotifyPropertyChangedFor(nameof(InputName))]
    private string _inputPath = string.Empty;

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(ExportCommand))]
    private string _outputDirectory = string.Empty;

    [ObservableProperty]
    private string _pages = string.Empty;

    [ObservableProperty]
    private ImageExportFormatChoice _selectedFormat = null!;

    [ObservableProperty]
    private double _targetDpi = 144D;

    [ObservableProperty]
    private int _maximumDimension;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanCancel))]
    [NotifyCanExecuteChangedFor(nameof(ExportCommand))]
    private bool _isBusy;

    [ObservableProperty]
    private double _progressFraction;

    [ObservableProperty]
    private string _status = string.Empty;

    [ObservableProperty]
    private string _summary = string.Empty;

    [ObservableProperty]
    private string? _publishedDirectory;

    [ObservableProperty]
    private bool _hasRecovery;

    public bool CanCancel => IsBusy;
    public string InputName => string.IsNullOrWhiteSpace(InputPath) ? string.Empty
        : _storage?.Describe(InputPath).Name ?? OfficeStorageIdentity.GetFileName(InputPath);
    public bool HasOutput => !string.IsNullOrWhiteSpace(PublishedDirectory);
    private bool CanExport => !IsBusy && !string.IsNullOrWhiteSpace(InputPath) && !string.IsNullOrWhiteSpace(OutputDirectory);

    internal void UseDocument(string? path) {
        if (IsBusy || string.IsNullOrWhiteSpace(path)) return;
        InputPath = path;
        if (string.IsNullOrWhiteSpace(OutputDirectory) && _storage?.UsesProviderPublication(path) != true &&
            OfficeStorageIdentity.GetLocalPath(path) is { } localPath) {
            OutputDirectory = Path.Combine(
                Path.GetDirectoryName(localPath)!,
                Path.GetFileNameWithoutExtension(localPath) + " pages");
        }
    }

    [RelayCommand]
    private async Task ChooseInputAsync(CancellationToken cancellationToken) {
        string? path = await _pickPdf(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) UseDocument(path);
    }

    [RelayCommand]
    private async Task ChooseOutputDirectoryAsync(CancellationToken cancellationToken) {
        string? path = await _pickOutputFolder(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) OutputDirectory = path;
    }

    [RelayCommand(CanExecute = nameof(CanExport))]
    private async Task ExportAsync() {
        _cancellation?.Dispose();
        using var operation = new CancellationTokenSource();
        _cancellation = operation;
        IsBusy = true;
        ProgressFraction = 0D;
        PublishedDirectory = null;
        HasRecovery = false;
        OnPropertyChanged(nameof(HasOutput));
        StudioJobRecord? job = null;
        bool ownerStarted = false;
        StudioStorageAccess.DirectoryOutputSession? directoryOutput = null;

        try {
            string destination = OutputDirectory;
            if (_storage?.UsesProviderPublication(destination) == true) {
                if (!await _confirmProviderWrite(destination).ConfigureAwait(true)) {
                    Status = T("Status.Cancelled", "Page export cancelled");
                    return;
                }
                operation.Token.ThrowIfCancellationRequested();
                directoryOutput = _storage.CreateDirectoryOutput(destination, _recoveryStore
                    ?? throw new IOException("Workflow recovery storage is unavailable."));
            }
            var request = new PdfPageImageExportRequest {
                InputPath = InputPath,
                InputStream = _storage?.CreateWorkflowInput(InputPath),
                OutputDirectory = destination,
                DirectoryOutput = directoryOutput?.Output,
                Pages = string.IsNullOrWhiteSpace(Pages) ? null : Pages,
                Format = SelectedFormat.Value,
                TargetDpi = TargetDpi,
                MaximumDimension = MaximumDimension > 0 ? MaximumDimension : null,
                PublicationGuard = _publicationGuard,
                ConflictPolicy = directoryOutput is null ? OfficeWorkflowConflictPolicy.Rename : OfficeWorkflowConflictPolicy.Replace
            };
            job = _jobHistory?.Start(T("Job.Title", "Page image export"), request.InputPath, request.OutputDirectory, operation.Cancel);
            using IDisposable? execution = _jobHistory is null ? null : await _jobHistory.EnterAsync(operation.Token).ConfigureAwait(true);
            var progress = new Progress<OfficeWorkflowProgress>(update => {
                if (!IsBusy || !ReferenceEquals(_cancellation, operation)) return;
                ProgressFraction = update.Fraction;
                Status = _localizer.GetOrDefault($"Workflow.Progress.{update.Stage}", update.Message);
                job?.Report(update);
            });
            ownerStarted = true;
            PdfPageImageExportResult result = await _runner.ExportPdfPagesAsync(request, progress, operation.Token).ConfigureAwait(true);
            job?.CompleteBatch(result.Status, result.OutputDirectory, result.Summary, result.OutputRecoveries, result.Files.Count > 0);
            HasRecovery = result.OutputRecoveries.Count > 0;
            Summary = result.Summary;
            Status = result.Status switch {
                OfficeWorkflowStatus.Completed => T("Status.Completed", "Page images ready"),
                OfficeWorkflowStatus.Cancelled => T("Status.Cancelled", "Page export cancelled"),
                OfficeWorkflowStatus.Unconfirmed => _localizer.GetOrDefault("Jobs.Unconfirmed", "Check output"),
                _ => _localizer.GetOrDefault("PageExport.Status.Failed", result.Summary)
            };
            PublishedDirectory = result.OutputDirectory;
            ProgressFraction = result.Status == OfficeWorkflowStatus.Completed ? 1D : ProgressFraction;
            OnPropertyChanged(nameof(HasOutput));
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
            Status = ownerStarted ? _localizer.GetOrDefault("Jobs.Unconfirmed", "Check output")
                : _localizer.GetOrDefault("Workflow.Status.Cancelled", "Cancelled");
            if (ownerStarted) job?.Unconfirmed(Status);
            else job?.Complete(OfficeWorkflowStatus.Cancelled, null, Status);
        } catch (Exception exception) {
            Status = exception.Message;
            job?.Unconfirmed(exception.Message);
        } finally {
            try { directoryOutput?.Dispose(); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) { Status += " " + error.Message; }
            IsBusy = false;
            if (ReferenceEquals(_cancellation, operation)) _cancellation = null;
        }
    }

    [RelayCommand]
    private void Cancel() => _cancellation?.Cancel();

    public void Dispose() => _cancellation?.Cancel();

    private ImageExportFormatChoice Format(OfficeImageExportFormat value, string label, string description) =>
        new(value, T($"Format.{value}.Label", label), T($"Format.{value}.Description", description));

    private string T(string suffix, string fallback) =>
        _localizer.GetOrDefault("PageExport." + suffix, fallback);
}
