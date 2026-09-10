using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed record PrintPaperChoice(string Name, PageSize Size);
public sealed record PrintOrientationChoice(PdfPrintOrientation Value, string Label);
public sealed record PrintScaleChoice(PdfPrintScaleMode Value, string Label, string Description);
public sealed record PrintPagesPerSheetChoice(int Value, string Label);

public sealed class PrintPreviewSheetViewModel : IDisposable {
    public PrintPreviewSheetViewModel(PdfRenderedPrintSheet sheet, string label) {
        const double maximumPreviewWidth = 350D;
        double scale = maximumPreviewWidth / sheet.Plan.PaperSize.Width;
        SheetNumber = sheet.Plan.SheetNumber;
        Label = label;
        Width = maximumPreviewWidth;
        Height = sheet.Plan.PaperSize.Height * scale;
        Placements = sheet.Plan.Placements;
        using var stream = new MemoryStream(sheet.GetPng(), writable: false);
        Image = Bitmap.DecodeToWidth(stream, (int)maximumPreviewWidth);
    }

    public int SheetNumber { get; }
    public double Width { get; }
    public double Height { get; }
    public IReadOnlyList<PdfPrintPlacement> Placements { get; }
    public Bitmap Image { get; }
    public string Label { get; }

    public void Dispose() {
        Image.Dispose();
    }
}

public sealed partial class PrintPreviewViewModel : ObservableObject, IDisposable {
    internal const int MaximumPreviewPages = 100;
    private readonly Func<CancellationToken, Task<string?>> _pickPdf;
    private readonly IStudioLocalizer _localizer;
    private readonly StudioStorageAccess _storage;
    private readonly Func<string, CancellationToken, Task<PdfDocument>> _readSnapshot;
    private readonly Func<CancellationToken, Task<string?>> _pickPrintFile;
    private readonly IPdfPrinterService _printers;
    private readonly StudioJobHistory? _jobHistory;
    private PdfPreparedPrintDocument? _preparedPrint;
    private bool _disposed;
    private bool _delivering;
    private CancellationTokenSource? _cancellation;

    public PrintPreviewViewModel(Func<CancellationToken, Task<string?>> pickPdf) : this(pickPdf, null) { }

    internal PrintPreviewViewModel(Func<CancellationToken, Task<string?>> pickPdf, IStudioLocalizer? localizer, StudioStorageAccess? storage = null,
        Func<string, CancellationToken, Task<PdfDocument>>? readSnapshot = null, IPdfPrinterService? printers = null,
        Func<CancellationToken, Task<string?>>? pickPrintFile = null, StudioJobHistory? jobHistory = null) {
        _pickPdf = pickPdf;
        _localizer = localizer ?? StudioLocalization.Current;
        _storage = storage ?? new StudioStorageAccess();
        _readSnapshot = readSnapshot ?? ReadStorageSnapshotAsync;
        _printers = printers ?? new PdfPrinterService();
        _pickPrintFile = pickPrintFile ?? (_ => Task.FromResult<string?>(null));
        _jobHistory = jobHistory;
        PaperChoices = [
            new("A4", PageSizes.A4),
            new(T("Paper.Letter", "Letter"), PageSizes.Letter),
            new(T("Paper.Legal", "Legal"), PageSizes.Legal),
            new("A3", PageSizes.A3)
        ];
        OrientationChoices = [
            Orientation(PdfPrintOrientation.Automatic, "Automatic"),
            Orientation(PdfPrintOrientation.Portrait, "Portrait"),
            Orientation(PdfPrintOrientation.Landscape, "Landscape")
        ];
        ScaleChoices = [
            Scale(PdfPrintScaleMode.Fit, "Fit", "Show the whole page."),
            Scale(PdfPrintScaleMode.ActualSize, "Actual size", "Keep physical page size where it fits."),
            Scale(PdfPrintScaleMode.Fill, "Fill", "Fill each slot and crop overflow.")
        ];
        PagesPerSheetChoices = [
            new(1, T("PagesPerSheet.One", "1 page")),
            new(2, T("PagesPerSheet.Two", "2 pages")),
            new(4, T("PagesPerSheet.Four", "4 pages"))
        ];
        SelectedPaper = PaperChoices[0];
        SelectedOrientation = OrientationChoices[0];
        SelectedScale = ScaleChoices[0];
        SelectedPagesPerSheet = PagesPerSheetChoices[0];
        SelectedDuplex = DuplexChoices[0];
        Status = T("Status.Ready", "Choose a PDF and preview its print sheets.");
        Summary = T("Summary.Empty", "No preview yet");
    }

    public IReadOnlyList<PrintPaperChoice> PaperChoices { get; }

    public IReadOnlyList<PrintOrientationChoice> OrientationChoices { get; }

    public IReadOnlyList<PrintScaleChoice> ScaleChoices { get; }

    public IReadOnlyList<PrintPagesPerSheetChoice> PagesPerSheetChoices { get; }

    public ObservableCollection<PrintPreviewSheetViewModel> Sheets { get; } = new();

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(BuildPreviewCommand))]
    [NotifyPropertyChangedFor(nameof(InputName))]
    private string _inputPath = string.Empty;

    [ObservableProperty]
    private string _pages = string.Empty;

    [ObservableProperty]
    private PrintPaperChoice _selectedPaper = null!;

    [ObservableProperty]
    private PrintOrientationChoice _selectedOrientation = null!;

    [ObservableProperty]
    private PrintScaleChoice _selectedScale = null!;

    [ObservableProperty]
    private PrintPagesPerSheetChoice _selectedPagesPerSheet = null!;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(CanCancel))]
    [NotifyCanExecuteChangedFor(nameof(BuildPreviewCommand))]
    private bool _isBusy;

    [ObservableProperty]
    private double _progressFraction;

    [ObservableProperty]
    private string _status = string.Empty;

    [ObservableProperty]
    private string _summary = string.Empty;

    public bool HasPreview => Sheets.Count > 0;
    public string InputName => string.IsNullOrWhiteSpace(InputPath) ? string.Empty : _storage.Describe(InputPath).Name;
    public bool CanCancel => IsBusy;
    private bool CanBuildPreview => !_disposed && !IsBusy && !string.IsNullOrWhiteSpace(InputPath);

    private async Task<PdfDocument> ReadStorageSnapshotAsync(string path, CancellationToken token) {
        StudioStorageSnapshot snapshot = await _storage.ReadSnapshotAsync(path, token).ConfigureAwait(true);
        return await Task.Run(() => PdfDocument.Load(snapshot.Bytes), token).ConfigureAwait(true);
    }

    internal void UseDocument(string? path) {
        if (IsBusy) return;
        if (!string.IsNullOrWhiteSpace(path)) InputPath = path;
    }

    [RelayCommand]
    private async Task ChooseInputAsync(CancellationToken cancellationToken) {
        string? path = await _pickPdf(cancellationToken).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) InputPath = path;
    }

    [RelayCommand(CanExecute = nameof(CanBuildPreview))]
    private async Task BuildPreviewAsync() {
        _cancellation?.Dispose();
        using var operation = new CancellationTokenSource();
        _cancellation = operation;
        IsBusy = true;
        ProgressFraction = 0D;
        Status = T("Status.Planning", "Planning print sheets");
        ClearPreview();

        try {
            PdfPrintPlanRequest request = new() {
                InputPath = InputPath,
                Pages = string.IsNullOrWhiteSpace(Pages) ? null : Pages,
                PaperSize = SelectedPaper.Size,
                Orientation = SelectedOrientation.Value,
                PagesPerSheet = SelectedPagesPerSheet.Value,
                ScaleMode = SelectedScale.Value
            };
            PdfDocument document = await _readSnapshot(request.InputPath, operation.Token).ConfigureAwait(true);
            ProgressFraction = 0.2D;
            Status = T("Status.Rendering", "Rendering page previews");
            // A3 at 300 DPI needs about 17.4 million pixels in either orientation.
            var options = new PdfPrintRenderOptions { MaximumPages = MaximumPreviewPages, Dpi = PrintDpi, MaximumPixelsPerImage = 20_000_000 };
            using IDisposable? permit = _jobHistory is null ? null : await _jobHistory.EnterAsync(operation.Token).ConfigureAwait(true);
            PdfPreparedPrintDocument prepared = await Task.Run(() => PdfPrintRenderer.Prepare(document, request, options, operation.Token), operation.Token).ConfigureAwait(true);
            operation.Token.ThrowIfCancellationRequested();
            if (_disposed) return;
            PdfPrintPlan plan = prepared.Plan;
            foreach (PdfRenderedPrintSheet sheet in prepared.Sheets) {
                Sheets.Add(new PrintPreviewSheetViewModel(sheet,
                    _localizer.FormatOrDefault("PrintPreview.Sheet.Label", "Sheet {0}", sheet.Plan.SheetNumber)));
            }
            _preparedPrint = prepared;
            OnPropertyChanged(nameof(HasPreview));
            Summary = _localizer.FormatOrDefault(
                "PrintPreview.Summary",
                "{0:N0} {1} · {2:N0} {3}",
                plan.SelectedPages.Count,
                plan.SelectedPages.Count == 1 ? T("Count.Page", "page") : T("Count.Pages", "pages"),
                plan.Sheets.Count,
                plan.Sheets.Count == 1 ? T("Count.Sheet", "sheet") : T("Count.Sheets", "sheets"));
            Status = T("Status.Completed", "Print preview ready") + (prepared.Diagnostics.Count == 0 ? string.Empty : " " + string.Join(" ", prepared.Diagnostics));
            ProgressFraction = 1D;
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
            ClearPreview();
            Status = T("Status.Cancelled", "Print preview cancelled");
            Summary = T("Summary.None", "No preview");
        } catch (Exception ex) {
            ClearPreview();
            Status = _localizer.FormatOrDefault("PrintPreview.Status.Failed", "Print preview failed: {0}", ex.Message);
            Summary = T("Summary.Unavailable", "Preview unavailable");
        } finally {
            IsBusy = false;
            if (ReferenceEquals(_cancellation, operation)) _cancellation = null;
        }
    }

    [RelayCommand]
    private void Cancel() => _cancellation?.Cancel();

    private void ClearPreview() {
        _preparedPrint = null;
        foreach (PrintPreviewSheetViewModel sheet in Sheets) sheet.Dispose();
        Sheets.Clear();
        OnPropertyChanged(nameof(HasPreview));
        PrintCommand.NotifyCanExecuteChanged();
    }

    public void Dispose() {
        _disposed = true;
        _discoveryCancellation?.Cancel();
        _paperSourceCancellation?.Cancel();
        _cancellation?.Cancel();
        ClearPreview();
    }

    private PrintOrientationChoice Orientation(PdfPrintOrientation value, string fallback) =>
        new(value, T($"Orientation.{value}", fallback));

    private PrintScaleChoice Scale(PdfPrintScaleMode value, string label, string description) =>
        new(value, T($"Scale.{value}.Label", label), T($"Scale.{value}.Description", description));

    private string T(string suffix, string fallback) =>
        _localizer.GetOrDefault("PrintPreview." + suffix, fallback);
}
