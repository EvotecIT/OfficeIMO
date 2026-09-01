using System.Collections.ObjectModel;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed record PrintPaperChoice(string Name, PageSize Size);
public sealed record PrintOrientationChoice(PdfPrintOrientation Value, string Label);
public sealed record PrintScaleChoice(PdfPrintScaleMode Value, string Label, string Description);
public sealed record PrintPagesPerSheetChoice(int Value, string Label);

public sealed class PrintPreviewPlacementViewModel : IDisposable {
    public PrintPreviewPlacementViewModel(PdfPrintPlacement placement, Bitmap image, double previewScale) {
        PageNumber = placement.PageNumber;
        Image = image;
        Left = (placement.IsClipped ? placement.SlotX : placement.X) * previewScale;
        Top = (placement.IsClipped ? placement.SlotY : placement.Y) * previewScale;
        Width = (placement.IsClipped ? placement.SlotWidth : placement.Width) * previewScale;
        Height = (placement.IsClipped ? placement.SlotHeight : placement.Height) * previewScale;
        ImageLeft = placement.IsClipped ? (placement.X - placement.SlotX) * previewScale : 0D;
        ImageTop = placement.IsClipped ? (placement.Y - placement.SlotY) * previewScale : 0D;
        ImageWidth = placement.Width * previewScale;
        ImageHeight = placement.Height * previewScale;
        IsClipped = placement.IsClipped;
    }

    public int PageNumber { get; }
    public Bitmap Image { get; }
    public double Left { get; }
    public double Top { get; }
    public double Width { get; }
    public double Height { get; }
    public double ImageLeft { get; }
    public double ImageTop { get; }
    public double ImageWidth { get; }
    public double ImageHeight { get; }
    public bool IsClipped { get; }
    public void Dispose() => Image.Dispose();
}

public sealed class PrintPreviewSheetViewModel : IDisposable {
    public PrintPreviewSheetViewModel(PdfPrintSheet sheet, IReadOnlyList<Bitmap> images) {
        const double maximumPreviewWidth = 350D;
        double scale = maximumPreviewWidth / sheet.PaperSize.Width;
        SheetNumber = sheet.SheetNumber;
        Width = maximumPreviewWidth;
        Height = sheet.PaperSize.Height * scale;
        Placements = sheet.Placements
            .Select((placement, index) => new PrintPreviewPlacementViewModel(placement, images[index], scale))
            .ToArray();
    }

    public int SheetNumber { get; }
    public double Width { get; }
    public double Height { get; }
    public IReadOnlyList<PrintPreviewPlacementViewModel> Placements { get; }
    public string Label => "Sheet " + SheetNumber;

    public void Dispose() {
        foreach (PrintPreviewPlacementViewModel placement in Placements) placement.Dispose();
    }
}

public sealed partial class PrintPreviewViewModel : ObservableObject, IDisposable {
    internal const int MaximumPreviewPages = 100;
    private readonly Func<CancellationToken, Task<string?>> _pickPdf;
    private CancellationTokenSource? _cancellation;

    public PrintPreviewViewModel(Func<CancellationToken, Task<string?>> pickPdf) {
        _pickPdf = pickPdf;
        SelectedPaper = PaperChoices[0];
        SelectedOrientation = OrientationChoices[0];
        SelectedScale = ScaleChoices[0];
        SelectedPagesPerSheet = PagesPerSheetChoices[0];
    }

    public IReadOnlyList<PrintPaperChoice> PaperChoices { get; } = [
        new("A4", PageSizes.A4),
        new("Letter", PageSizes.Letter),
        new("Legal", PageSizes.Legal),
        new("A3", PageSizes.A3)
    ];

    public IReadOnlyList<PrintOrientationChoice> OrientationChoices { get; } = [
        new(PdfPrintOrientation.Automatic, "Automatic"),
        new(PdfPrintOrientation.Portrait, "Portrait"),
        new(PdfPrintOrientation.Landscape, "Landscape")
    ];

    public IReadOnlyList<PrintScaleChoice> ScaleChoices { get; } = [
        new(PdfPrintScaleMode.Fit, "Fit", "Show the whole page."),
        new(PdfPrintScaleMode.ActualSize, "Actual size", "Keep physical page size where it fits."),
        new(PdfPrintScaleMode.Fill, "Fill", "Fill each slot and crop overflow.")
    ];

    public IReadOnlyList<PrintPagesPerSheetChoice> PagesPerSheetChoices { get; } = [
        new(1, "1 page"),
        new(2, "2 pages"),
        new(4, "4 pages")
    ];

    public ObservableCollection<PrintPreviewSheetViewModel> Sheets { get; } = new();

    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(BuildPreviewCommand))]
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
    private string _status = "Choose a PDF and preview its print sheets.";

    [ObservableProperty]
    private string _summary = "No preview yet";

    public bool HasPreview => Sheets.Count > 0;
    public bool CanCancel => IsBusy;
    private bool CanBuildPreview => !IsBusy && !string.IsNullOrWhiteSpace(InputPath);

    internal void UseDocument(string? path) {
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
        Status = "Planning print sheets";
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
            PdfPrintPlan plan = await Task.Run(() => PdfPrintPlanner.Create(request), operation.Token).ConfigureAwait(true);
            if (plan.SelectedPages.Count > MaximumPreviewPages) {
                throw new InvalidOperationException(
                    $"Print preview is limited to {MaximumPreviewPages:N0} pages. Enter a smaller page selection.");
            }
            ProgressFraction = 0.2D;
            Status = "Rendering page previews";
            var options = new PdfImageExportOptions {
                ThumbnailMaxDimension = 350,
                MaximumOutputCount = MaximumPreviewPages
            };
            PdfDocument document = PdfDocument.Load(InputPath);
            IReadOnlyList<OfficeImageExportResult> rendered = await document
                .ToImages(options)
                .Pages(PdfPageSelection.From(plan.SelectedPages.ToArray()))
                .AsPng()
                .ExportAsync(operation.Token)
                .ConfigureAwait(true);
            operation.Token.ThrowIfCancellationRequested();

            int imageIndex = 0;
            foreach (PdfPrintSheet sheet in plan.Sheets) {
                var bitmaps = new List<Bitmap>(sheet.Placements.Count);
                try {
                    for (int index = 0; index < sheet.Placements.Count; index++) {
                        using var stream = new MemoryStream(rendered[imageIndex++].Bytes, writable: false);
                        bitmaps.Add(new Bitmap(stream));
                    }
                    Sheets.Add(new PrintPreviewSheetViewModel(sheet, bitmaps));
                } catch {
                    foreach (Bitmap bitmap in bitmaps) bitmap.Dispose();
                    throw;
                }
                ProgressFraction = 0.2D + (double)imageIndex / rendered.Count * 0.8D;
            }
            OnPropertyChanged(nameof(HasPreview));
            Summary = $"{plan.SelectedPages.Count:N0} {(plan.SelectedPages.Count == 1 ? "page" : "pages")} · {plan.Sheets.Count:N0} {(plan.Sheets.Count == 1 ? "sheet" : "sheets")}";
            Status = "Print preview ready";
            ProgressFraction = 1D;
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
            ClearPreview();
            Status = "Print preview cancelled";
            Summary = "No preview";
        } catch (Exception ex) {
            ClearPreview();
            Status = "Print preview failed: " + ex.Message;
            Summary = "Preview unavailable";
        } finally {
            IsBusy = false;
            if (ReferenceEquals(_cancellation, operation)) _cancellation = null;
        }
    }

    [RelayCommand]
    private void Cancel() => _cancellation?.Cancel();

    private void ClearPreview() {
        foreach (PrintPreviewSheetViewModel sheet in Sheets) sheet.Dispose();
        Sheets.Clear();
        OnPropertyChanged(nameof(HasPreview));
    }

    public void Dispose() {
        _cancellation?.Cancel();
        ClearPreview();
    }
}
