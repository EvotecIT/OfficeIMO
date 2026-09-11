using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed record PrintDuplexChoice(PdfPrintDuplex Value, string Label) {
    public override string ToString() => Label;
}

public sealed partial class PrintPreviewViewModel {
    private CancellationTokenSource? _discoveryCancellation;
    [ObservableProperty] private IReadOnlyList<PdfPrinterInfo> _printerChoices = [];
    [ObservableProperty] private PdfPrinterInfo? _selectedPrinter;
    [ObservableProperty] private bool _isDiscoveringPrinters;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasPrinterDiscoveryError))]
    private string _printerDiscoveryError = string.Empty;
    [ObservableProperty] private int _copies = 1;
    [ObservableProperty] private int _printDpi = 150;
    [ObservableProperty] private PrintDuplexChoice? _selectedDuplex;
    [ObservableProperty] private string _printOutputPath = string.Empty;
    private IReadOnlyList<PrintDuplexChoice>? _duplexChoices;
    public IReadOnlyList<int> PrintDpiChoices { get; } = [150, 300];
    public IReadOnlyList<PrintDuplexChoice> DuplexChoices => _duplexChoices ??= [
        new(PdfPrintDuplex.PrinterDefault, T("Duplex.Default", "Printer default")),
        new(PdfPrintDuplex.SingleSided, T("Duplex.Single", "Single-sided")),
        new(PdfPrintDuplex.LongEdge, T("Duplex.Long", "Double-sided, long edge")),
        new(PdfPrintDuplex.ShortEdge, T("Duplex.Short", "Double-sided, short edge"))
    ];
    public bool RequiresPrintOutput => SelectedPrinter?.RequiresOutputFile == true;
    public bool HasPrinterDiscoveryError => !string.IsNullOrWhiteSpace(PrinterDiscoveryError);
    public bool CanChangePrintSettings => !_disposed && !IsBusy;
    public bool CanPrint => CanChangePrintSettings && !IsDiscoveringPrinters && !IsDiscoveringPaperSources && _preparedPrint is not null && SelectedPrinter is not null &&
        (!RequiresPrintOutput || !string.IsNullOrWhiteSpace(PrintOutputPath));

    protected override void OnPropertyChanged(PropertyChangedEventArgs e) {
        base.OnPropertyChanged(e);
        if (e.PropertyName is nameof(InputPath) or nameof(Pages) or nameof(SelectedPaper) or nameof(SelectedOrientation)
            or nameof(SelectedScale) or nameof(SelectedPagesPerSheet) or nameof(PrintDpi)) InvalidatePreparedSheets();
        if (e.PropertyName is nameof(IsBusy) or nameof(SelectedPrinter) or nameof(PrintOutputPath) or nameof(HasPreview) or nameof(IsDiscoveringPaperSources) or nameof(IsDiscoveringPrinters)) {
            OnPropertyChanged(nameof(CanChangePrintSettings));
            OnPropertyChanged(nameof(CanPrint));
            PrintCommand.NotifyCanExecuteChanged();
        }
        if (e.PropertyName == nameof(SelectedPrinter)) OnPropertyChanged(nameof(RequiresPrintOutput));
    }

    internal void InvalidateDocument(string? path) {
        if (string.IsNullOrWhiteSpace(path) || string.IsNullOrWhiteSpace(InputPath)) return;
        try {
            if (OfficeIMO.Internal.OfficeStorageIdentity.AreEquivalent(path, InputPath)) InvalidatePreparedSheets();
        } catch (Exception error) when (error is ArgumentException or IOException or NotSupportedException) { }
    }

    private void InvalidatePreparedSheets() {
        if (!_delivering) _cancellation?.Cancel();
        ClearPreview();
    }

    [RelayCommand]
    private async Task RefreshPrintersAsync(CancellationToken token) {
        if (_disposed || IsBusy || IsDiscoveringPrinters) return;
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(token);
        _discoveryCancellation = operation;
        IsDiscoveringPrinters = true;
        PrinterDiscoveryError = string.Empty;
        try {
            var printers = await _printers.GetPrintersAsync(operation.Token).ConfigureAwait(true);
            operation.Token.ThrowIfCancellationRequested();
            if (_disposed) return;
            string? previous = SelectedPrinter?.Name;
            PrinterChoices = printers;
            var selected = printers.FirstOrDefault(printer => printer.Name == previous) ?? printers.FirstOrDefault();
            if (Equals(selected, SelectedPrinter)) PaperSourceDiscovery = RefreshPaperSourcesAsync(selected);
            else SelectedPrinter = selected;
            await PaperSourceDiscovery.ConfigureAwait(true);
            if (printers.Count == 0) PrinterDiscoveryError = T("NoPrinters", "No printer queues are installed.");
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) { }
        catch (Exception error) { if (!_disposed) PrinterDiscoveryError = error.Message; }
        finally {
            if (ReferenceEquals(_discoveryCancellation, operation)) _discoveryCancellation = null;
            IsDiscoveringPrinters = false;
        }
    }

    [RelayCommand]
    private async Task ChoosePrintOutputAsync(CancellationToken token) {
        if (!CanChangePrintSettings) return;
        string? path = await _pickPrintFile(token).ConfigureAwait(true);
        if (!string.IsNullOrWhiteSpace(path)) PrintOutputPath = path;
    }

    [RelayCommand(CanExecute = nameof(CanPrint))]
    private async Task PrintAsync(CancellationToken token) {
        if (!CanPrint || _preparedPrint is not { } prepared || SelectedPrinter is not { } printer) return;
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(token);
        _cancellation = operation;
        _delivering = true;
        IsBusy = true;
        StudioJobRecord? job = null;
        try {
            var options = new PdfPrintDeliveryOptions {
                PrinterName = printer.Name, DocumentName = InputName, Copies = Copies,
                Duplex = SelectedDuplex?.Value ?? PdfPrintDuplex.PrinterDefault,
                PaperSourceId = SelectedPaperSource?.Id,
                OutputFilePath = printer.RequiresOutputFile ? PrintOutputPath : null
            };
            job = _jobHistory?.Start(T("JobTitle", "Print reviewed sheets"), InputName, printer.Name, operation.Cancel);
            using IDisposable? permit = _jobHistory is null ? null : await _jobHistory.EnterAsync(operation.Token).ConfigureAwait(true);
            Status = T("Sending", "Sending the reviewed sheets to the printer…");
            PdfPrintSubmission receipt = await _printers.SubmitAsync(prepared, options, operation.Token).ConfigureAwait(true);
            Status = _localizer.FormatOrDefault("PrintPreview.Accepted", "Printer accepted job {0}: {1} sheet(s), {2} copy/copies. Check the printer for completion.",
                receipt.JobId, receipt.SheetCount, receipt.Copies);
            if (!string.IsNullOrWhiteSpace(receipt.CleanupWarning)) Status += " " + receipt.CleanupWarning;
            job?.Complete(OfficeWorkflowStatus.Completed, null, Status);
        } catch (PdfPrintDeliveryException error) {
            Status = error.Message;
            job?.Complete(OfficeWorkflowStatus.Unconfirmed, null, Status);
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
            Status = T("DeliveryCancelled", "Printing cancelled before submission.");
            job?.Complete(OfficeWorkflowStatus.Cancelled, null, Status);
        } catch (Exception error) {
            Status = error.Message;
            job?.Complete(OfficeWorkflowStatus.Failed, null, Status);
        } finally {
            _delivering = false;
            IsBusy = false;
            if (ReferenceEquals(_cancellation, operation)) _cancellation = null;
        }
    }
}
