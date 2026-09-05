using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Drawing;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed record ImageExportFormatChoice(OfficeImageExportFormat Value, string Label, string Description);

public sealed partial class PageImageExportViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<string?>> _pickPdf;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputFolder;
    private readonly IOfficeOutputWorkflowRunner _runner;
    private readonly IStudioLocalizer _localizer;
    private CancellationTokenSource? _cancellation;

    public PageImageExportViewModel(
        Func<CancellationToken, Task<string?>> pickPdf,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeOutputWorkflowRunner? runner = null) : this(pickPdf, pickOutputFolder, runner, null) { }

    internal PageImageExportViewModel(
        Func<CancellationToken, Task<string?>> pickPdf,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeOutputWorkflowRunner? runner,
        IStudioLocalizer? localizer = null) {
        _pickPdf = pickPdf;
        _pickOutputFolder = pickOutputFolder;
        _runner = runner ?? new OfficeWorkflowRunner();
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

    public bool CanCancel => IsBusy;
    public bool HasOutput => !string.IsNullOrWhiteSpace(PublishedDirectory);
    private bool CanExport => !IsBusy && !string.IsNullOrWhiteSpace(InputPath) && !string.IsNullOrWhiteSpace(OutputDirectory);

    internal void UseDocument(string? path) {
        if (string.IsNullOrWhiteSpace(path)) return;
        InputPath = path;
        if (string.IsNullOrWhiteSpace(OutputDirectory)) {
            OutputDirectory = Path.Combine(
                Path.GetDirectoryName(path)!,
                Path.GetFileNameWithoutExtension(path) + " pages");
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
        OnPropertyChanged(nameof(HasOutput));

        try {
            var progress = new Progress<OfficeWorkflowProgress>(update => {
                ProgressFraction = update.Fraction;
                Status = _localizer.GetOrDefault($"Workflow.Progress.{update.Stage}", update.Message);
            });
            PdfPageImageExportResult result = await _runner.ExportPdfPagesAsync(new PdfPageImageExportRequest {
                InputPath = InputPath,
                OutputDirectory = OutputDirectory,
                Pages = string.IsNullOrWhiteSpace(Pages) ? null : Pages,
                Format = SelectedFormat.Value,
                TargetDpi = TargetDpi,
                MaximumDimension = MaximumDimension > 0 ? MaximumDimension : null,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Rename
            }, progress, operation.Token).ConfigureAwait(true);
            Summary = result.Summary;
            Status = result.Status switch {
                OfficeWorkflowStatus.Completed => T("Status.Completed", "Page images ready"),
                OfficeWorkflowStatus.Cancelled => T("Status.Cancelled", "Page export cancelled"),
                _ => _localizer.GetOrDefault("PageExport.Status.Failed", result.Summary)
            };
            PublishedDirectory = result.OutputDirectory;
            ProgressFraction = result.Status == OfficeWorkflowStatus.Completed ? 1D : ProgressFraction;
            OnPropertyChanged(nameof(HasOutput));
        } finally {
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
