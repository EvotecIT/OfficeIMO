using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Drawing;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed record ImageExportFormatChoice(OfficeImageExportFormat Value, string Label, string Description);

public sealed partial class PageImageExportViewModel : ObservableObject, IDisposable {
    private readonly Func<CancellationToken, Task<string?>> _pickPdf;
    private readonly Func<CancellationToken, Task<string?>> _pickOutputFolder;
    private readonly IOfficeOutputWorkflowRunner _runner;
    private CancellationTokenSource? _cancellation;

    public PageImageExportViewModel(
        Func<CancellationToken, Task<string?>> pickPdf,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        IOfficeOutputWorkflowRunner? runner = null) {
        _pickPdf = pickPdf;
        _pickOutputFolder = pickOutputFolder;
        _runner = runner ?? new OfficeWorkflowRunner();
        SelectedFormat = Formats[0];
    }

    public IReadOnlyList<ImageExportFormatChoice> Formats { get; } = [
        new(OfficeImageExportFormat.Png, "PNG", "Lossless raster pages with transparency support."),
        new(OfficeImageExportFormat.Jpeg, "JPEG", "Compact photographic raster pages."),
        new(OfficeImageExportFormat.Webp, "WebP", "Compact lossless raster pages."),
        new(OfficeImageExportFormat.Tiff, "TIFF", "Lossless archival raster pages."),
        new(OfficeImageExportFormat.Svg, "SVG", "Managed vector page scenes where supported.")
    ];

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
    private string _status = "Choose a PDF and an output folder.";

    [ObservableProperty]
    private string _summary = "No export yet";

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
                Status = update.Message;
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
                OfficeWorkflowStatus.Completed => "Page images ready",
                OfficeWorkflowStatus.Cancelled => "Page export cancelled",
                _ => result.Summary
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
}
